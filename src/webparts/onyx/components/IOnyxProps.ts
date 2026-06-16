import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IOnyxProps {
  adminFlowUrl: string;
  adminEmails: string;
  context: WebPartContext;
  currentUserEmail: string;
  currentUserId?: number;
  dataSiteUrl: string;
  staffDirectorySiteUrl: string;
  staffDirectoryListName: string;
  userDisplayName: string;
}
