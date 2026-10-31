import { Client } from "@microsoft/microsoft-graph-client";
import { loadConfig } from "./config.js";
import { MSGraphAuth, getMsalCachePath } from "./auth.js";
import { GraphHandlers } from "./handlers.js";

/**
 * Interactive sign-in and health check, run in a terminal:
 *   npm run auth                 sign in (device code) and verify
 *   npm run auth -- --sign-out   forget the cached sign-in for this profile
 *
 * Choose the profile with MICROSOFT_PROFILE (default: corporate).
 */
async function run(): Promise<void> {
  const cfg = loadConfig();
  const auth = new MSGraphAuth(cfg);

  if (process.argv.includes("--sign-out")) {
    await auth.clearCache();
    console.log(
      `Signed out of profile '${cfg.profile}' (${getMsalCachePath(cfg)}).`,
    );
    return;
  }

  console.log(`=== Microsoft Graph MCP sign-in (profile: ${cfg.profile}) ===`);
  await auth.initialize();
  const token = await auth.signInInteractive();

  const client = Client.initWithMiddleware({
    authProvider: { getAccessToken: async () => token },
  });
  const report = await new GraphHandlers(client, auth, cfg).execute(
    "graph_health_check",
    {},
  );

  console.log("");
  console.log(
    `Account      : ${report.account?.displayName} <${report.account?.mail}>`,
  );
  console.log(`Account type : ${report.accountType}`);
  console.log(
    `Outbound     : ${report.outboundToolsEnabled ? "enabled" : "disabled (drafts only)"}`,
  );
  console.log(
    `Mailbox      : ${report.mailbox?.status}${report.mailbox?.unread !== undefined ? ` (${report.mailbox.unread} unread)` : ""}`,
  );
  if (report.licences?.skus)
    console.log(`Licences     : ${report.licences.skus.join(", ") || "none"}`);
  console.log(`Token cache  : ${getMsalCachePath(cfg)} (encrypted)`);
  console.log(`Verdict      : ${report.verdict}`);
  process.exitCode = report.mailbox?.status === "ok" ? 0 : 2;
}

run().catch((error: any) => {
  console.error(`\n[Sign-in failed] ${error?.message || error}`);
  process.exit(1);
});
