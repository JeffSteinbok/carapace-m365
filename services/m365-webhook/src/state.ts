import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  writeFileSync,
} from "node:fs";
import { dirname } from "node:path";

export interface M365WebhookState {
  refreshToken?: string;
  subscriptionId?: string;
  expirationDateTime?: string;
  notificationUrl?: string;
  clientState?: string;
}

export class StateStore {
  constructor(private readonly path: string) {}

  load(): M365WebhookState {
    if (!existsSync(this.path)) return {};
    try {
      return JSON.parse(readFileSync(this.path, "utf8")) as M365WebhookState;
    } catch (error) {
      throw new Error(`Could not read Microsoft 365 webhook state at ${this.path}: ${error}`);
    }
  }

  update(patch: Partial<M365WebhookState>): M365WebhookState {
    const state = { ...this.load(), ...patch };
    this.save(state);
    return state;
  }

  save(state: M365WebhookState): void {
    const directory = dirname(this.path);
    mkdirSync(directory, { recursive: true, mode: 0o700 });
    const temporary = `${this.path}.${process.pid}.tmp`;
    writeFileSync(temporary, `${JSON.stringify(state, null, 2)}\n`, {
      encoding: "utf8",
      mode: 0o600,
    });
    try {
      chmodSync(temporary, 0o600);
    } catch {
      // Windows and some filesystems do not implement POSIX permissions.
    }
    renameSync(temporary, this.path);
    try {
      chmodSync(this.path, 0o600);
    } catch {
      // Best effort only where restrictive permissions are supported.
    }
  }
}
