import { Config } from "./config.js";

/**
 * Delegated scopes, aligned to the implemented tools (least privilege).
 * Send-capable scopes are requested only when outbound tools are enabled.
 * Note: Entra access tokens carry every scope the user has ever consented to
 * for the app, so this is not a hard boundary if Mail.Send was consented
 * earlier. The tool gating is the control; for a hard boundary use an app
 * registration that has never been granted Mail.Send.
 */
export function scopesFor(cfg: Config): string[] {
  const scopes = [
    "User.Read",
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Files.ReadWrite",
  ];
  if (!cfg.isConsumer) {
    scopes.push(
      "Sites.Read.All",
      cfg.enableOutbound ? "Chat.ReadWrite" : "Chat.Read",
    );
  }
  if (cfg.enableOutbound) scopes.push("Mail.Send");
  return scopes;
}
