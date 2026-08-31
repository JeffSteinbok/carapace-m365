/**
 * MailStore — optional SQLite cache for incoming mail.
 *
 * Enabled by setting M365_MAIL_STORE_PATH to a file path.
 * When disabled (path is empty/unset), all methods are no-ops.
 *
 * Schema is append-only; new columns are added with ALTER TABLE so
 * existing databases are forward-compatible without migration scripts.
 */

import { createRequire } from "node:module";
import type { DatabaseSync, StatementSync, SQLInputValue } from "node:sqlite";

const requireSqlite = createRequire(import.meta.url);

/**
 * Loads `node:sqlite` on demand.
 *
 * The module is only available on Node 22.5+, and the store is optional, so
 * this is resolved lazily rather than imported at module load. That keeps the
 * disabled path a true no-op on runtimes without it — importing statically
 * would make every consumer require Node 22.5+ even with the store switched
 * off.
 */
function loadSqlite(): { DatabaseSync: new (path: string) => DatabaseSync } {
  try {
    return requireSqlite("node:sqlite") as { DatabaseSync: new (path: string) => DatabaseSync };
  } catch (cause) {
    throw new Error(
      "M365_MAIL_STORE_PATH is set but `node:sqlite` is unavailable. The mail store requires Node 22.5 or newer; unset the variable to run without it.",
      { cause },
    );
  }
}

export interface StoredMessage {
  /** Microsoft Graph message ID — use for reply/forward API calls */
  graph_id: string;
  /** RFC 2822 Message-ID header — use for In-Reply-To threading */
  internet_message_id: string | null;
  subject: string;
  sender_email: string;
  sender_name: string;
  received_at: string | null;
  has_attachments: number; // 0 | 1
  body_preview: string | null;
  /** Raw JSON of the full Graph message object */
  raw_json: string | null;
  stored_at: string;
}

export interface MessageRow {
  graph_id: string;
  internet_message_id: string | null;
  subject: string;
  sender_email: string;
  sender_name: string;
  received_at: string | null;
  has_attachments: number;
  body_preview: string | null;
  stored_at: string;
}

const CREATE_TABLE = `
CREATE TABLE IF NOT EXISTS messages (
  graph_id             TEXT PRIMARY KEY,
  internet_message_id  TEXT,
  subject              TEXT NOT NULL DEFAULT '',
  sender_email         TEXT NOT NULL DEFAULT '',
  sender_name          TEXT NOT NULL DEFAULT '',
  received_at          TEXT,
  has_attachments      INTEGER NOT NULL DEFAULT 0,
  body_preview         TEXT,
  raw_json             TEXT,
  stored_at            TEXT NOT NULL
) STRICT;
`;

const CREATE_INDEXES = [
  `CREATE INDEX IF NOT EXISTS idx_messages_received_at   ON messages (received_at DESC);`,
  `CREATE INDEX IF NOT EXISTS idx_messages_sender_email  ON messages (sender_email);`,
  `CREATE INDEX IF NOT EXISTS idx_messages_internet_id   ON messages (internet_message_id);`,
];

const UPSERT = `
INSERT INTO messages
  (graph_id, internet_message_id, subject, sender_email, sender_name,
   received_at, has_attachments, body_preview, raw_json, stored_at)
VALUES
  (:graph_id, :internet_message_id, :subject, :sender_email, :sender_name,
   :received_at, :has_attachments, :body_preview, :raw_json, :stored_at)
ON CONFLICT (graph_id) DO UPDATE SET
  internet_message_id = excluded.internet_message_id,
  subject             = excluded.subject,
  sender_email        = excluded.sender_email,
  sender_name         = excluded.sender_name,
  received_at         = excluded.received_at,
  has_attachments     = excluded.has_attachments,
  body_preview        = excluded.body_preview,
  raw_json            = excluded.raw_json,
  stored_at           = excluded.stored_at;
`;

export class MailStore {
  private readonly db: DatabaseSync | null = null;
  private readonly upsertStmt: StatementSync | null = null;

  constructor(path: string | undefined) {
    if (!path) return; // disabled

    const { DatabaseSync } = loadSqlite();
    this.db = new DatabaseSync(path);
    this.db.exec("PRAGMA journal_mode=WAL;");
    this.db.exec("PRAGMA synchronous=NORMAL;");
    this.db.exec(CREATE_TABLE);
    for (const idx of CREATE_INDEXES) this.db.exec(idx);
    this.upsertStmt = this.db.prepare(UPSERT);
  }

  get enabled(): boolean {
    return this.db !== null;
  }

  /**
   * Store (or update) a message. Safe to call multiple times for the same graph_id.
   */
  store(msg: StoredMessage): void {
    if (!this.upsertStmt) return;
    this.upsertStmt.run({
      graph_id: msg.graph_id,
      internet_message_id: msg.internet_message_id ?? null,
      subject: msg.subject,
      sender_email: msg.sender_email,
      sender_name: msg.sender_name,
      received_at: msg.received_at ?? null,
      has_attachments: msg.has_attachments,
      body_preview: msg.body_preview ?? null,
      raw_json: msg.raw_json ?? null,
      stored_at: msg.stored_at,
    });
  }

  /**
   * Look up a message by its Microsoft Graph ID.
   * Fast PK lookup — use this before making a Graph API call for reply/fwd.
   */
  getByGraphId(graphId: string): MessageRow | null {
    if (!this.db) return null;
    const stmt = this.db.prepare(
      `SELECT graph_id, internet_message_id, subject, sender_email, sender_name,
              received_at, has_attachments, body_preview, stored_at
       FROM messages WHERE graph_id = ?`,
    );
    return (stmt.get(graphId) as MessageRow | undefined) ?? null;
  }

  /**
   * Look up a message by its RFC 2822 Message-ID header value.
   * Used for In-Reply-To threading.
   */
  getByInternetMessageId(internetMessageId: string): MessageRow | null {
    if (!this.db) return null;
    const stmt = this.db.prepare(
      `SELECT graph_id, internet_message_id, subject, sender_email, sender_name,
              received_at, has_attachments, body_preview, stored_at
       FROM messages WHERE internet_message_id = ? LIMIT 1`,
    );
    return (stmt.get(internetMessageId) as MessageRow | undefined) ?? null;
  }

  /**
   * Search messages by sender email, subject substring, and/or date range.
   * Returns up to `limit` rows ordered by received_at DESC.
   */
  search(options: {
    sender_email?: string;
    subject_contains?: string;
    since?: string;  // ISO datetime
    before?: string; // ISO datetime
    limit?: number;
  }): MessageRow[] {
    if (!this.db) return [];

    const conditions: string[] = [];
    const bindings: SQLInputValue[] = [];

    if (options.sender_email) {
      conditions.push("sender_email = ?");
      bindings.push(options.sender_email.toLowerCase());
    }
    if (options.subject_contains) {
      conditions.push("subject LIKE ? ESCAPE '\\'");
      bindings.push(`%${options.subject_contains.replace(/[%_\\]/g, "\\$&")}%`);
    }
    if (options.since) {
      conditions.push("received_at >= ?");
      bindings.push(options.since);
    }
    if (options.before) {
      conditions.push("received_at < ?");
      bindings.push(options.before);
    }

    const where = conditions.length ? `WHERE ${conditions.join(" AND ")}` : "";
    const limit = Math.min(options.limit ?? 50, 200);

    const stmt = this.db.prepare(
      `SELECT graph_id, internet_message_id, subject, sender_email, sender_name,
              received_at, has_attachments, body_preview, stored_at
       FROM messages ${where}
       ORDER BY received_at DESC
       LIMIT ${limit}`,
    );
    return stmt.all(...bindings) as unknown as MessageRow[];
  }

  /**
   * List recent messages, ordered by received_at DESC.
   */
  recent(limit = 20): MessageRow[] {
    if (!this.db) return [];
    const stmt = this.db.prepare(
      `SELECT graph_id, internet_message_id, subject, sender_email, sender_name,
              received_at, has_attachments, body_preview, stored_at
       FROM messages ORDER BY received_at DESC LIMIT ?`,
    );
    return stmt.all(Math.min(limit, 200)) as unknown as MessageRow[];
  }

  /**
   * Total number of stored messages.
   */
  count(): number {
    if (!this.db) return 0;
    const stmt = this.db.prepare("SELECT COUNT(*) AS n FROM messages");
    const row = stmt.get() as { n: number };
    return row.n;
  }

  close(): void {
    this.db?.close();
  }
}
