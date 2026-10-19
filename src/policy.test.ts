import { test } from "node:test";
import assert from "node:assert/strict";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { Config } from "./config.js";
import {
  PolicyError,
  assertGraphUrl,
  clampTop,
  isSharedWithOthers,
  odataString,
  resolveReadPath,
  resolveWritePath,
  seg,
  siteSeg,
} from "./policy.js";

const sandbox = fs.mkdtempSync(path.join(os.tmpdir(), "graph-mcp-test-"));
const allowed = path.join(sandbox, "allowed");
const outside = path.join(sandbox, "outside");
fs.mkdirSync(allowed, { recursive: true });
fs.mkdirSync(outside, { recursive: true });

function cfg(overrides: Partial<Config> = {}): Config {
  return {
    clientId: "x",
    tenantId: "common",
    profile: "corporate",
    isCorporateProfile: true,
    isConsumer: false,
    toolPrefix: "",
    enableOutbound: false,
    writeDirs: [allowed],
    readDirs: [allowed],
    denyPaths: [],
    allowExecutableWrites: false,
    maxTransferBytes: 1024,
    timeZone: "UTC",
    ...overrides,
  };
}

test.after(() => fs.rmSync(sandbox, { recursive: true, force: true }));

test("write: accepts a new file inside an allowed root", () => {
  const p = path.join(allowed, "sub", "report.pdf");
  assert.equal(resolveWritePath(p, cfg()), path.resolve(p));
});

test("write: rejects relative paths and traversal out of the root", () => {
  assert.throws(() => resolveWritePath("report.pdf", cfg()), PolicyError);
  assert.throws(
    () => resolveWritePath(path.join(allowed, "..", "outside", "x.pdf"), cfg()),
    /outside the allowed folders/,
  );
});

test("write: rejects protected folders and credential-like names", () => {
  assert.throws(
    () =>
      resolveWritePath(path.join(allowed, ".ssh", "authorized_keys"), cfg()),
    /protected location/,
  );
  assert.throws(
    () =>
      resolveWritePath(
        path.join(allowed, "AppData", "Roaming", "x.txt"),
        cfg(),
      ),
    /protected location/,
  );
  assert.throws(
    () => resolveWritePath(path.join(allowed, ".env"), cfg()),
    /credential/,
  );
  assert.throws(
    () =>
      resolveWritePath(
        path.join(allowed, ".microsoft-graph-mcp", "x.txt"),
        cfg(),
      ),
    /token cache/,
  );
});

test("write: protected names are judged relative to the allowed root", () => {
  // e.g. the Windows temp folder lives under AppData but is explicitly allowed
  const root = path.join(sandbox, "AppData", "Local", "Temp");
  fs.mkdirSync(root, { recursive: true });
  const p = path.join(root, "ok.pdf");
  assert.equal(
    resolveWritePath(p, cfg({ writeDirs: [allowed, root] })),
    path.resolve(p),
  );
});

test("write: blocks executables unless allowed", () => {
  const p = path.join(allowed, "setup.ps1");
  assert.throws(() => resolveWritePath(p, cfg()), /executable/);
  assert.equal(
    resolveWritePath(p, cfg({ allowExecutableWrites: true })),
    path.resolve(p),
  );
});

test("write: refuses to overwrite unless asked", () => {
  const p = path.join(allowed, "exists.txt");
  fs.writeFileSync(p, "x");
  assert.throws(() => resolveWritePath(p, cfg()), /already exists/);
  assert.equal(resolveWritePath(p, cfg(), true), path.resolve(p));
});

test("write: honours configured deny fragments, case-insensitively on Windows/macOS", () => {
  const p = path.join(allowed, "Strategy", "plan.docx");
  const deny = process.platform === "linux" ? "Strategy" : "strategy";
  assert.throws(
    () => resolveWritePath(p, cfg({ denyPaths: [deny] })),
    /deny rule/,
  );
});

test("read: enforces roots, deny rules and size limit", () => {
  const ok = path.join(allowed, "small.txt");
  fs.writeFileSync(ok, "hello");
  assert.equal(resolveReadPath(ok, cfg()), fs.realpathSync.native(ok));

  const big = path.join(allowed, "big.bin");
  fs.writeFileSync(big, Buffer.alloc(2048));
  assert.throws(() => resolveReadPath(big, cfg()), /limit/);

  const out = path.join(outside, "secret.txt");
  fs.writeFileSync(out, "x");
  assert.throws(
    () => resolveReadPath(out, cfg()),
    /outside the allowed folders/,
  );

  const key = path.join(allowed, "id_ed25519");
  fs.writeFileSync(key, "x");
  assert.throws(() => resolveReadPath(key, cfg()), /credential/);
});

test("seg: encodes path segments so IDs cannot change the request path", () => {
  assert.equal(seg("AAMk/../users?x=1", "id"), "AAMk%2F..%2Fusers%3Fx%3D1");
  assert.throws(() => seg("", "id"), PolicyError);
  assert.throws(() => seg(42, "id"), PolicyError);
  assert.equal(
    siteSeg("contoso.sharepoint.com,abc,def"),
    "contoso.sharepoint.com,abc,def",
  );
  assert.equal(siteSeg("a/b"), "a%2Fb");
});

test("assertGraphUrl: only Microsoft Graph links are followed", () => {
  const good = "https://graph.microsoft.com/v1.0/me/messages?$skip=10";
  assert.equal(assertGraphUrl(good, "messages"), new URL(good).toString());
  for (const bad of [
    "https://evil.example/v1.0/me/messages",
    "https://graph.microsoft.com.evil.example/v1.0/me/messages",
    "https://user:pw@graph.microsoft.com/v1.0/me/messages",
    "http://graph.microsoft.com/v1.0/me/messages",
  ]) {
    assert.throws(() => assertGraphUrl(bad, "messages"), PolicyError, bad);
  }
});

test("assertGraphUrl: a next page cannot become a different query", () => {
  const drive =
    "https://graph.microsoft.com/v1.0/drives/b!abc/items/root/children?$skiptoken=X&$top=50";
  assert.equal(assertGraphUrl(drive, "children"), new URL(drive).toString());
  // resource switch: mail search smuggled into a OneDrive tool
  assert.throws(
    () =>
      assertGraphUrl(
        'https://graph.microsoft.com/v1.0/me/messages?$search="password"',
        "children",
      ),
    /next page of 'children'/,
  );
  // same resource, but with expansion of attachment bytes
  assert.throws(
    () =>
      assertGraphUrl(
        "https://graph.microsoft.com/v1.0/me/messages?$skip=10&$expand=attachments",
        "messages",
      ),
    /unexpected option/,
  );
  assert.throws(
    () =>
      assertGraphUrl(
        "https://graph.microsoft.com/beta/me/messages?$skip=10",
        "messages",
      ),
    PolicyError,
  );
});

test("seg: rejects dot segments", () => {
  assert.throws(() => seg(".", "id"), PolicyError);
  assert.throws(() => seg("..", "id"), PolicyError);
});

test("write: rejects alternate data streams, device names and trailing dots", () => {
  const c = cfg({ allowExecutableWrites: false });
  for (const name of [
    "evil.bat::$DATA",
    "x.ps1::$DATA",
    "evil.hta:x",
    "note.txt:hidden",
    "CON",
    "nul.txt",
    "com1.log",
    "evil.bat.",
    "evil.bat ",
  ]) {
    assert.throws(
      () => resolveWritePath(path.join(allowed, name), c),
      PolicyError,
      name,
    );
  }
});

test("write: refuses agent and tool config files that would run code", () => {
  for (const name of [
    ".claude.json",
    ".gitconfig",
    ".wslconfig",
    ".pypirc",
    "AGENTS.md",
    "CLAUDE.md",
    "SOUL.md",
    ".cursorrules",
  ]) {
    assert.throws(
      () => resolveWritePath(path.join(allowed, name), cfg(), true),
      /credential|shell profile/,
      name,
    );
  }
});

test("isSharedWithOthers: shapes observed on a real personal OneDrive", () => {
  const owner = {
    roles: ["owner"],
    link: { webUrl: "https://x" },
    grantedToV2: { user: { id: "a" } },
  };
  const anonEdit = {
    roles: ["write"],
    link: { type: "edit", scope: "anonymous" },
    inheritedFrom: { id: "p" },
  };
  const orgView = {
    roles: ["read"],
    link: { type: "view", scope: "organization" },
  };
  const direct = { roles: ["read"], grantedToV2: { user: { id: "b" } } };
  assert.equal(isSharedWithOthers([owner]), false);
  assert.equal(isSharedWithOthers([{ roles: ["owner"], link: {} }]), false);
  assert.equal(isSharedWithOthers([anonEdit, owner]), true);
  assert.equal(isSharedWithOthers([owner, orgView]), true);
  assert.equal(isSharedWithOthers([owner, direct]), true);
  assert.equal(isSharedWithOthers([]), false);
  assert.equal(isSharedWithOthers(undefined), false);
});

test("clampTop and odataString", () => {
  assert.equal(clampTop(undefined, 10, 50), 10);
  assert.equal(clampTop(500, 10, 50), 50);
  assert.equal(clampTop(-3, 10, 50), 1);
  assert.equal(odataString("O'Brien"), "O''Brien");
});
