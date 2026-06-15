import { MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * NotificationService handles sending Outlook emails and Teams chat messages
 * using Microsoft Graph API via MSGraphClientV3.
 */
export class NotificationService {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Send approval notification via Outlook email and Teams chat
   */
  public async sendMaterialRequestApprovalNotification(params: {
    requestTitle: string;
    project: string;
    amount?: number;
    approverEmail: string;
    approverId?: string;
    deepLink: string;
  }): Promise<void> {
    const { requestTitle, project, amount, approverEmail, approverId, deepLink } = params;

    const subject = `Approval Required: ${requestTitle}`;
    const bodyContent = this.buildNotificationBody(requestTitle, project, amount, deepLink);

    console.log('[NotificationService] Sending approval notification to:', approverEmail);

    // 1. Send Outlook Email
    await this.sendOutlookEmail(approverEmail, subject, bodyContent);

    // 2. Send Teams Chat Message (resolve user ID if not provided)
    let resolvedUserId = approverId;
    if (!resolvedUserId && approverEmail) {
      resolvedUserId = await this.resolveUserIdByEmail(approverEmail);
    }

    if (resolvedUserId) {
      await this.sendTeamsChatMessage(resolvedUserId, bodyContent);
    } else {
      console.warn('[NotificationService] Could not resolve approver for Teams chat. Email was sent.');
    }
  }

  private buildNotificationBody(
    requestTitle: string,
    project: string,
    amount: number | undefined,
    deepLink: string
  ): string {
    return `
      <p>A new Material Request requires your approval.</p>
      <ul>
        <li><strong>Request:</strong> ${requestTitle}</li>
        <li><strong>Project:</strong> ${project}</li>
        ${amount ? `<li><strong>Estimated Cost:</strong> $${amount}</li>` : ''}
      </ul>
      <p>
        <a href="${deepLink}">Open in Bonnedo Enterprise Platform</a>
      </p>
    `;
  }

  private async sendOutlookEmail(to: string, subject: string, htmlBody: string): Promise<void> {
    try {
      await this.graphClient
        .api('/me/sendMail')
        .post({
          message: {
            subject,
            body: {
              contentType: 'HTML',
              content: htmlBody,
            },
            toRecipients: [
              {
                emailAddress: {
                  address: to,
                },
              },
            ],
          },
        });
    } catch (error) {
      console.error('Failed to send Outlook email:', error);
    }
  }

  private async sendTeamsChatMessage(userId: string, message: string): Promise<void> {
    try {
      // Create a new chat with the user (1:1)
      const chat = await this.graphClient.api('/chats').post({
        chatType: 'oneOnOne',
        members: [
          {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            roles: ['owner'],
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`,
          },
        ],
      });

      // Send message to the newly created chat
      await this.graphClient.api(`/chats/${chat.id}/messages`).post({
        body: {
          contentType: 'html',
          content: message,
        },
      });

      console.log('[NotificationService] Teams chat message sent successfully');
    } catch (error) {
      console.error('[NotificationService] Failed to send Teams chat message:', error);
    }
  }

  /**
   * Resolve Azure AD user ID from email address
   */
  private async resolveUserIdByEmail(email: string): Promise<string | undefined> {
    try {
      const result = await this.graphClient
        .api('/users')
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select('id,displayName,mail,userPrincipalName')
        .top(1)
        .get();

      if (result.value && result.value.length > 0) {
        console.log('[NotificationService] Resolved approver:', result.value[0]);
        return result.value[0].id;
      }
    } catch (error) {
      console.warn('[NotificationService] Could not resolve user by email:', error);
    }
    return undefined;
  }
}
