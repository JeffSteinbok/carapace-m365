import { afterEach, describe, expect, it } from "vitest";
import { existsSync, rmSync } from "node:fs";
import { join, resolve } from "node:path";
import type { MailEnvelope } from "carapace-mail-runtime";
import type { GraphClient } from "../src/graph.js";
import { OutlookProviderClient } from "../src/provider.js";

const scratch = resolve("services", "m365-webhook", ".test-provider");

afterEach(() => {
  rmSync(scratch, { recursive: true, force: true });
});

const envelope: MailEnvelope = {
  message_id: "msg-1",
  provider: "outlook",
  account_id: "m365",
  mailbox_id: "inbox",
  sender_name: "USPS",
  sender_email: "USPSInformeddelivery@email.informeddelivery.usps.com",
  subject: "Your Daily Digest",
  body_html: '<img src="cid:scan-1">',
  // Graph reports false when the only attachments are inline.
  has_attachments: false,
};

const graph = {
  fetchAttachments: async () => [
    {
      id: "a1",
      name: "scan-1.jpg",
      contentType: "image/jpeg",
      isInline: true,
      contentId: "scan-1",
      contentBytes: Buffer.from("jpeg").toString("base64"),
    },
  ],
} as unknown as GraphClient;

describe("OutlookProviderClient", () => {
  it("downloads inline attachments when Graph reports hasAttachments=false", async () => {
    const provider = new OutlookProviderClient(graph, () => {});
    const saved = await provider.downloadAttachments(envelope, scratch, {
      content_types: ["image/*"],
      include_body_html: true,
    });
    expect(saved).toEqual(["body.html", "scan-1.jpg"]);
    expect(existsSync(join(scratch, "scan-1.jpg"))).toBe(true);
  });

  it("lists inline attachments when Graph reports hasAttachments=false", async () => {
    const provider = new OutlookProviderClient(graph, () => {});
    const listed = await provider.listAttachments(envelope);
    expect(listed).toEqual([
      { name: "scan-1.jpg", content_type: "image/jpeg", is_inline: true, content_id: "scan-1" },
    ]);
  });
});

