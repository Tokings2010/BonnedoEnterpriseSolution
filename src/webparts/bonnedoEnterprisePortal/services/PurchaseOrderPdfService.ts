import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { NotificationService } from './NotificationService';
import { ReportTemplateService } from './ReportTemplateService';

export interface IPOData {
  PO_Number: string;
  Vendor?: string;
  VendorEmail?: string;
  Project_Code?: string;
  TotalAmount: number;
  Items?: Array<{
    Description: string;
    Quantity: number;
    UnitPrice: number;
    Total: number;
  }>;
  Requested_By?: string;
  Delivery_Date?: string;
  Description?: string;
}

export class PurchaseOrderPdfService {
  private spHttpClient: SPHttpClient;
  private pageContext: PageContext;

  constructor(spHttpClient: SPHttpClient, pageContext: PageContext) {
    this.spHttpClient = spHttpClient;
    this.pageContext = pageContext;
  }

  public async generateAndDistributePO(
    poData: IPOData,
    graphClient?: MSGraphClientV3
  ): Promise<void> {
    try {
      const pdfBlob = await this.generateProfessionalPdf(poData);
      const fileServerRelativeUrl = await this.uploadPdfToLibrary(poData.PO_Number, pdfBlob);
      await this.sendPoNotifications(poData, fileServerRelativeUrl, graphClient);

      console.log(`[PO PDF Service] Successfully generated: ${poData.PO_Number}`);
    } catch (error) {
      console.error('[PO PDF Service] Error:', error);
      throw error;
    }
  }

  private async generateProfessionalPdf(poData: IPOData): Promise<Blob> {
    const html = this.buildPoHtmlTemplate(poData);

    const container = document.createElement('div');
    container.innerHTML = html;
    container.style.position = 'absolute';
    container.style.left = '-9999px';
    container.style.width = '210mm';
    document.body.appendChild(container);

    try {
      const canvas = await html2canvas(container, {
        scale: 2,
        useCORS: true,
        logging: false,
      });

      const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4',
      });

      const imgData = canvas.toDataURL('image/png');
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
      return pdf.output('blob');
    } finally {
      document.body.removeChild(container);
    }
  }

  private buildPoHtmlTemplate(poData: IPOData): string {
    const templateService = new ReportTemplateService();
    let template = templateService.getTemplate('PurchaseOrder') || '';

    const itemsHtml = poData.Items?.map(item => `
      <tr>
        <td style="padding:6px; border:1px solid #ccc;">${item.Description}</td>
        <td style="padding:6px; border:1px solid #ccc; text-align:center;">${item.Quantity}</td>
        <td style="padding:6px; border:1px solid #ccc; text-align:right;">$${item.UnitPrice.toFixed(2)}</td>
        <td style="padding:6px; border:1px solid #ccc; text-align:right;">$${item.Total.toFixed(2)}</td>
      </tr>
    `).join('') || '';

    const subtotal = poData.TotalAmount;
    const vat = subtotal * 0.05;
    const wht = subtotal * 0.02;
    const grandTotal = subtotal + vat - wht;

    // Replace placeholders
    template = template
      .replace(/{{PO_Number}}/g, poData.PO_Number)
      .replace(/{{Date}}/g, new Date().toLocaleDateString())
      .replace(/{{Project_Code}}/g, poData.Project_Code || 'N/A')
      .replace(/{{Vendor_Name}}/g, poData.Vendor || 'N/A')
      .replace(/{{Delivery_Date}}/g, poData.Delivery_Date ? new Date(poData.Delivery_Date).toLocaleDateString() : 'TBD')
      .replace(/{{Line_Items}}/g, itemsHtml)
      .replace(/{{Subtotal}}/g, `$${subtotal.toFixed(2)}`)
      .replace(/{{VAT}}/g, `$${vat.toFixed(2)}`)
      .replace(/{{WHT}}/g, `$${wht.toFixed(2)}`)
      .replace(/{{Total_Amount}}/g, `$${grandTotal.toFixed(2)}`)
      .replace(/{{Amount_In_Words}}/g, this.numberToWords(grandTotal));

    return template;
  }

  private numberToWords(num: number): string {
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' })
      .format(num)
      .replace('$', '');
  }

  private async uploadPdfToLibrary(poNumber: string, blob: Blob): Promise<string> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const libraryName = 'Shared Documents';
    const folderPath = 'Procurement/PO';
    const fullFolderPath = `${libraryName}/${folderPath}`;
    const fileName = `${poNumber}.pdf`;

    try {
      // Ensure folder path exists
      await this.ensureFolderExists(libraryName, folderPath);

      // Upload PDF
      const uploadUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(fullFolderPath)}')/Files/add(overwrite=true,url='${fileName}')`;

      const response = await this.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/pdf',
        },
        body: blob,
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`PDF upload failed: ${response.statusText} - ${errorText}`);
      }

      const result = await response.json();
      return result.d.ServerRelativeUrl;
    } catch (error) {
      console.error('[PO PDF Service] Upload error:', error);
      throw error;
    }
  }

  /**
   * Recursively ensure folder path exists in the document library
   */
  private async ensureFolderExists(libraryName: string, folderPath: string): Promise<void> {
    const webUrl = this.pageContext.web.absoluteUrl;
    const segments = folderPath.split('/').filter(s => s.length > 0);
    let currentPath = libraryName;

    for (const segment of segments) {
      currentPath += `/${segment}`;
      const folderUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(currentPath)}')`;

      try {
        const checkResponse = await this.spHttpClient.get(folderUrl, SPHttpClient.configurations.v1);

        if (!checkResponse.ok) {
          // Create folder
          const parentFolder = currentPath.substring(0, currentPath.lastIndexOf('/'));
          const folderName = segment;

          const createUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(parentFolder)}')/Folders`;

          const createResponse = await this.spHttpClient.post(createUrl, SPHttpClient.configurations.v1, {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              '__metadata': { type: 'SP.Folder' },
              'Name': folderName,
            }),
          });

          if (!createResponse.ok) {
            const errorText = await createResponse.text();
            throw new Error(`Failed to create folder '${segment}': ${errorText}`);
          }
        }
      } catch (error) {
        console.error(`[PO PDF Service] Error ensuring folder ${currentPath}:`, error);
        throw error;
      }
    }
  }

  private async sendPoNotifications(
    poData: IPOData,
    pdfUrl: string,
    graphClient?: MSGraphClientV3
  ): Promise<void> {
    if (!graphClient) return;

    const notificationService = new NotificationService(graphClient);
    const deepLink = `${window.location.origin}${window.location.pathname}?po=${poData.PO_Number}`;

    if (poData.VendorEmail) {
      await notificationService.sendMaterialRequestApprovalNotification({
        requestTitle: `Purchase Order ${poData.PO_Number}`,
        project: poData.Project_Code || 'N/A',
        amount: poData.TotalAmount,
        approverEmail: poData.VendorEmail,
        deepLink,
      });
    }
  }
}
