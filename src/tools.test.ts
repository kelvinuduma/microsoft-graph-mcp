import { test } from "node:test";
import assert from "node:assert/strict";
import { Config } from "./config.js";
import { scopesFor } from "./scopes.js";
import { MS_GRAPH_TOOLS, getTools } from "./tools.js";

function cfg(overrides: Partial<Config> = {}): Config {
  return {
    clientId: "x",
    tenantId: "organizations",
    profile: "corporate",
    isCorporateProfile: true,
    isConsumer: false,
    toolPrefix: "",
    enableOutbound: false,
    writeDirs: [],
    readDirs: [],
    denyPaths: [],
    allowExecutableWrites: false,
    maxTransferBytes: 1,
    timeZone: "UTC",
    ...overrides,
  };
}

const names = (c: Config) => getTools(c).map((t) => t.name);

test("tool names are unique", () => {
  const all = MS_GRAPH_TOOLS.map((t) => t.name);
  assert.equal(new Set(all).size, all.length);
});

test("outbound tools are hidden by default and shown when enabled", () => {
  assert.ok(!names(cfg()).includes("outlook_send_mail"));
  assert.ok(!names(cfg()).includes("teams_send_chat_message"));
  assert.ok(names(cfg()).includes("outlook_create_reply_draft"));
  assert.ok(names(cfg({ enableOutbound: true })).includes("outlook_send_mail"));
});

test("consumer accounts do not see SharePoint or Teams tools", () => {
  const consumer = names(cfg({ isConsumer: true, enableOutbound: true }));
  assert.ok(
    !consumer.some(
      (n) => n.startsWith("sharepoint_") || n.startsWith("teams_"),
    ),
  );
  assert.ok(consumer.includes("outlook_send_mail"));
});

test("prefix is applied and internal metadata is not exposed", () => {
  const tools = getTools(cfg({ toolPrefix: "personal_" }));
  assert.ok(tools.every((t) => t.name.startsWith("personal_")));
  assert.ok(tools.every((t) => !("meta" in t)));
  assert.ok(
    tools.every(
      (t) => t.annotations && typeof t.annotations.readOnlyHint === "boolean",
    ),
  );
});

test("send scopes are requested only when outbound is enabled", () => {
  assert.ok(!scopesFor(cfg()).includes("Mail.Send"));
  assert.ok(scopesFor(cfg()).includes("Chat.Read"));
  assert.ok(scopesFor(cfg({ enableOutbound: true })).includes("Mail.Send"));
  assert.ok(
    scopesFor(cfg({ enableOutbound: true })).includes("Chat.ReadWrite"),
  );
  assert.ok(
    !scopesFor(cfg({ isConsumer: true })).some(
      (s) => s.startsWith("Chat") || s.startsWith("Sites"),
    ),
  );
});
