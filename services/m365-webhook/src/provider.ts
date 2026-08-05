import { mkdirSync, writeFileSync } from "node:fs";
import { basename, join } from "node:path";
import type {
  AttachmentMeta,
  MailEnvelope,
  MailProviderClient,
} from "carapace-mail-runtime";
import { GraphClient } from "./graph.js";

function contentTypeAllowed(contentType: string, allowedTypes: string[]): boolean {
  return allowedTypes.some((allowed) =>
    allowed.endsWith("/*")
      ? contentType.startsWith(allowed.slice(0, -1))
      : contentType === allowed,
  );
}

export class OutlookProviderClient implements MailProviderClient {
  constructor(
    private readonly graph: GraphClient,
    private readonly logger: (message: string) => void,
  ) {}

  async fetchBody(envelope: MailEnvelope): Promise<MailEnvelope> {
    if (envelope.body_text != null || envelope.body_html != null) return envelope;
    this.logger(`fetchBody: re-fetching message ${envelope.message_id}`);
    const message = await this.graph.fetchMessage(envelope.message_id);
    const bodyHtml =
      message.body?.contentType?.toLowerCase() === "html"
        ? (message.body.content ?? null)
        : null;
    const bodyText =
      message.body?.contentType?.toLowerCase() === "text"
        ? (message.body.content ?? null)
        : (message.bodyPreview ?? null);
    return {
      ...envelope,
      body_text: bodyText,
      body_html: bodyHtml,
      raw: {
        ...(envelope.raw ?? {}),
        ...(message as unknown as Record<string, unknown>),
      },
    };
  }

  async listAttachments(envelope: MailEnvelope): Promise<AttachmentMeta[]> {
    if (!envelope.has_attachments) return [];
    const attachments = await this.graph.fetchAttachments(envelope.message_id);
    return attachments.map((attachment) => ({
      name: attachment.name ?? "attachment",
      content_type: attachment.contentType ?? "application/octet-stream",
      is_inline: attachment.isInline ?? false,
      content_id: attachment.contentId ?? null,
    }));
  }

  async downloadAttachments(
    envelope: MailEnvelope,
    outputDir: string,
    options?: {
      content_types?: string[] | null;
      inline_only?: boolean | null;
      include_body_html?: boolean;
    },
  ): Promise<string[]> {
    mkdirSync(outputDir, { recursive: true });
    const saved: string[] = [];
    if (options?.include_body_html) {
      const html = envelope.body_html
        ?? ((envelope.raw as Record<string, unknown> | undefined)?.body as
          | { content?: string }
          | undefined)?.content;
      if (html) {
        writeFileSync(join(outputDir, "body.html"), html, "utf8");
        saved.push("body.html");
      }
    }
    if (!envelope.has_attachments) return saved;

    const attachments = await this.graph.fetchAttachments(envelope.message_id);
    for (const attachment of attachments) {
      if (!attachment.contentBytes) continue;
      const inline = attachment.isInline ?? false;
      if (options?.inline_only === true && !inline) continue;
      if (options?.inline_only === false && inline) continue;
      const contentType = attachment.contentType ?? "application/octet-stream";
      if (options?.content_types
        && !contentTypeAllowed(contentType, options.content_types)) continue;
      const filename = basename(attachment.name ?? `attachment-${attachment.id}.bin`);
      writeFileSync(
        join(outputDir, filename),
        Buffer.from(attachment.contentBytes, "base64"),
      );
      saved.push(filename);
    }
    return saved;
  }
}
