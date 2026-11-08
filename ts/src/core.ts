// Public communication core. No MCP dependency: reusable from CLI, service tests or another TypeScript application.
export {
  Outlook,
  ApiError,
  listFolders,
  resolveFolderId,
  type Obj,
  type Fetch,
} from "./api.ts";
export { Tokens, TokenError, loaderFromEnv, type Account } from "./tokens.ts";
export {
  ServiceOutlook,
  httpGateway,
  sshGateway,
  type GatewayCall,
} from "./service.ts";
export {
  listEmails,
  listEmailsPage,
  searchEmails,
  searchEmailsPage,
  getMessage,
  readThread,
  readThreadPage,
  listAttachments,
  getAttachment,
  type EmailList,
  type ListOptions,
} from "./mail.ts";
export { listCalendars, listCalendarEvents } from "./calendar.ts";
export {
  sendEmail,
  manageEmail,
  parseRecipients,
  type SendOptions,
} from "./send.ts";
export { listUnread, type UnreadMailbox } from "./unread.ts";
