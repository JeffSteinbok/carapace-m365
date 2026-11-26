import { linkSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { loadServiceConfig } from "../src/config.js";

let scratch: string;

function fixtureAccount(id: string) {
  return {
    statePath: join(scratch, id + "-state.json"),
    rulesPath: join(scratch, id + "-rules.json"),
    pipelineWorkspace: join(scratch, id + "-workspace"),
    logPath: join(scratch, id + ".log"),
    brokerSecretEnv: `${id}_SECRET`,
    clientStateEnv: `${id}_STATE`,
    features: ["mail-write"],
    allowedAgents: [id] as string[] | undefined,
    notify: { channel: "test", target: id },
    legacyTokenPath: id === "first",
  };
}

function load(first = fixtureAccount("first"), second = fixtureAccount("second")) {
  writeFileSync(join(scratch, "accounts.json"), JSON.stringify({ accounts: { first, second } }));
  return loadServiceConfig().accounts;
}

describe("multi-account service config", () => {
  beforeEach(() => {
    scratch = mkdtempSync(join(tmpdir(), "m365-account-fixture-"));
    vi.stubEnv("M365_ACCOUNTS_PATH", join(scratch, "accounts.json"));
    vi.stubEnv("M365_WEBHOOK_URL", "https://example.test/webhook");
    vi.stubEnv("M365_CLIENT_ID", "test-client");
    vi.stubEnv("M365_MAIL_STORE_PATH", join(scratch, "legacy.sqlite"));
    vi.stubEnv("first_SECRET", "a-secret");
    vi.stubEnv("first_STATE", "a-state");
    vi.stubEnv("second_SECRET", "b-secret");
    vi.stubEnv("second_STATE", "b-state");
  });

  afterEach(() => {
    vi.unstubAllEnvs();
    rmSync(scratch, { recursive: true, force: true });
  });

  it("inherits the existing cache only for the legacy token account", () => {
    const [first, second] = load();
    expect(first?.mailStorePath).toBe(join(scratch, "legacy.sqlite"));
    expect(second?.mailStorePath).toBeUndefined();
  });

  it("requires an explicit agent allowlist", () => {
    expect(() => load(undefined, { ...fixtureAccount("second"), allowedAgents: undefined })).toThrow(/agent allowlist/);
  });

  it("rejects duplicate credential material across accounts", () => {
    vi.stubEnv("second_SECRET", "a-secret");
    expect(() => load()).toThrow(/Shared Microsoft 365 account credential/);
  });

  it("rejects a shared path even across settings of different types", () => {
    expect(() => load(undefined, { ...fixtureAccount("second"), logPath: fixtureAccount("first").statePath }))
      .toThrow(/Shared Microsoft 365 account path/);
  });

  it("rejects two hard links to the same state file", () => {
    const first = fixtureAccount("first");
    const second = fixtureAccount("second");
    writeFileSync(first.statePath, "{}");
    linkSync(first.statePath, second.statePath);
    expect(() => load(first, second)).toThrow(/Shared Microsoft 365 account path/);
  });

  it("rejects more than one legacy token route", () => {
    expect(() => load(undefined, { ...fixtureAccount("second"), legacyTokenPath: true })).toThrow(/Multiple legacy token routes/);
  });
});
