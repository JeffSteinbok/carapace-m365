import { randomUUID } from "node:crypto";
import {
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

export const DEFAULT_DIRECT_TOKEN_STATE_PATH = join(
  homedir(),
  ".openclaw",
  "state",
  "m365-direct-token.json",
);

export interface DirectTokenState {
  refreshToken?: string;
}

export class DirectTokenStateStore {
  constructor(private readonly path: string) {}

  load(): DirectTokenState {
    if (!existsSync(this.path)) return {};
    try {
      const value = JSON.parse(readFileSync(this.path, "utf8")) as unknown;
      if (!value || typeof value !== "object" || Array.isArray(value)) {
        throw new Error("state must be a JSON object");
      }
      const refreshToken = (value as Record<string, unknown>).refreshToken;
      if (refreshToken !== undefined && typeof refreshToken !== "string") {
        throw new Error("refreshToken must be a string");
      }
      return refreshToken ? { refreshToken } : {};
    } catch (error) {
      throw new Error(`Could not read Microsoft 365 direct-token state at ${this.path}: ${error}`);
    }
  }

  save(state: DirectTokenState): void {
    const directory = dirname(this.path);
    mkdirSync(directory, { recursive: true, mode: 0o700 });
    try {
      chmodSync(directory, 0o700);
    } catch {
      // Windows and some filesystems do not implement POSIX permissions.
    }

    const temporary = `${this.path}.${process.pid}.${randomUUID()}.tmp`;
    try {
      writeFileSync(temporary, `${JSON.stringify(state, null, 2)}\n`, {
        encoding: "utf8",
        mode: 0o600,
      });
      try {
        chmodSync(temporary, 0o600);
      } catch {
        // Best effort only where restrictive permissions are supported.
      }
      renameSync(temporary, this.path);
      try {
        chmodSync(this.path, 0o600);
      } catch {
        // Best effort only where restrictive permissions are supported.
      }
    } catch (error) {
      try {
        unlinkSync(temporary);
      } catch {
        // Ignore cleanup failures and report the original write error.
      }
      throw new Error(`Could not write Microsoft 365 direct-token state at ${this.path}: ${error}`);
    }
  }
}
