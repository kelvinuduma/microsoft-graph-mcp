/**
 * Turns Graph SDK errors into short, actionable messages for the model and
 * the user. Raw error objects are logged to stderr only.
 */

const HINTS: Record<string, string> = {
  MailboxNotEnabledForRESTAPI:
    "The mailbox is inactive or soft-deleted. This usually means the user's Exchange Online licence was removed or has expired. An administrator must reassign the licence (a soft-deleted mailbox is recoverable for about 30 days).",
  MailboxNotSupportedForRESTAPI:
    "This mailbox type (for example on-premises or a legacy shared mailbox) is not available through Microsoft Graph.",
  ErrorAccessDenied:
    "Access denied. The signed-in account lacks permission for this item, or the scope was not consented.",
  Authorization_RequestDenied:
    "Insufficient privileges. This needs a scope or directory role the signed-in account does not have.",
  InvalidAuthenticationToken:
    "The access token was rejected. Sign in again (`npm run auth`).",
  ErrorItemNotFound:
    "The item was not found. Check the ID; it may have been moved or deleted.",
  itemNotFound:
    "The item was not found. Check the ID; it may have been moved or deleted.",
  ErrorInvalidIdMalformed:
    "The ID is malformed. Use an ID returned by a list or search tool.",
  activityLimitReached:
    "Microsoft Graph is throttling requests. Wait a minute and retry.",
  TooManyRequests:
    "Microsoft Graph is throttling requests. Wait a minute and retry.",
  ErrorQuotaExceeded: "The mailbox is over its storage quota.",
};

export function describeError(error: any): string {
  const code: string | undefined = error?.code || error?.body?.code;
  const status: number | undefined = error?.statusCode;
  const requestId: string | undefined =
    error?.requestId || error?.headers?.["request-id"];
  const base = error?.message ? String(error.message) : String(error);

  const parts = [base];
  if (code && HINTS[code]) parts.push(`Hint: ${HINTS[code]}`);
  const meta = [
    code && `code=${code}`,
    status && `status=${status}`,
    requestId && `requestId=${requestId}`,
  ].filter(Boolean);
  if (meta.length) parts.push(`(${meta.join(", ")})`);
  return parts.join(" ");
}

export function errorCode(error: any): string | undefined {
  return error?.code || error?.body?.code;
}
