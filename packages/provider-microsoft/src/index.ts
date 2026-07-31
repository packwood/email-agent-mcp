// @usejunior/provider-microsoft — Microsoft Graph API email provider
export { GraphEmailProvider, RealGraphApiClient, GraphApiError, type GraphApiClient, type DeltaResult } from './email-graph-provider.js';
export {
  MatonGraphApiClient,
  loadMatonOutlookConnections,
  resolveMatonOutlookConnection,
  matonGraphUrl,
  type MatonOutlookConnection,
} from './maton.js';
export { DelegatedAuthManager, ClientCredentialsAuthManager, listConfiguredMailboxes, listConfiguredMailboxesWithMetadata, loadMailboxMetadata, toFilesystemSafeKey, getConfigDir, GRAPH_SCOPES, isAuthError } from './auth.js';
export type { MailboxMetadata } from './auth.js';
export {
  handleValidationToken,
  isDuplicateNotification,
  checkSubscriptionExists,
  healthCheckEndpoint,
  createInboxSubscription,
} from './subscriptions.js';
