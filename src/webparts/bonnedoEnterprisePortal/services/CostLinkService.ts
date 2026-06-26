import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export interface ICostTransaction {
  projectCode: string;
  phase?: string;
  transactionType: 'PO Commitment' | 'Material Issue' | 'Vendor Invoice' | 'Budget Transfer' | 'Other';
  amount: number;
  vendorCode?: string;
  vendorName?: string;
  referenceId: string;
  referenceType: string;
  description: string;
}

/**
 * Logs a cost transaction to the Cost_Transactions register for integrated cost tracking.
 * Links procurement/material events to the project budget.
 */
export async function logCostTransaction(
  spHttpClient: SPHttpClient,
  pageContext: PageContext,
  transaction: ICostTransaction
): Promise<boolean> {
  try {
    const listUrl = `${pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Cost_Transactions')/items`;

    const body = {
      Title: `${transaction.transactionType} - ${transaction.referenceId}`,
      Project_Code: transaction.projectCode,
      Phase: transaction.phase || '',
      Transaction_Type: transaction.transactionType,
      Amount: transaction.amount,
      Vendor_Code: transaction.vendorCode || '',
      Vendor_Name: transaction.vendorName || '',
      Reference_ID: transaction.referenceId,
      Reference_Type: transaction.referenceType,
      Description: transaction.description,
      Transaction_Date: new Date().toISOString(),
    };

    const response: SPHttpClientResponse = await spHttpClient.post(
      listUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
        },
        body: JSON.stringify(body),
      }
    );

    if (!response.ok) {
      if (response.status === 404) {
        console.warn(`[CostLinkService] Cost_Transactions list not found. Cost transaction not logged. Provision the list to enable cost tracking.`);
      } else {
        console.warn(`[CostLinkService] Failed to log cost transaction: ${response.status}`);
      }
      return false;
    }

    console.log(`[CostLinkService] Cost transaction logged: ${transaction.transactionType} - ${transaction.amount}`);
    return true;
  } catch (error) {
    console.error('[CostLinkService] Error logging cost transaction:', error);
    return false;
  }
}
