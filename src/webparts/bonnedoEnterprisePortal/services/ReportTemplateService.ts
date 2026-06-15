export interface IReportTemplate {
  name: string;
  html: string;
}

export class ReportTemplateService {
  private templates: Map<string, string> = new Map();

  constructor() {
    this.registerDefaultTemplates();
  }

  private registerDefaultTemplates() {
    // Purchase Order Template
    this.templates.set('PurchaseOrder', `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; padding: 30px; width: 210mm; color: #222;">
        <div style="text-align:center; border-bottom: 3px solid #003366; padding-bottom:15px;">
          <h1 style="margin:0; color:#003366; font-size:24px;">BONNEDO ENTERPRISE</h1>
          <p style="margin:4px 0; font-size:11px;">ISO 9001:2015 Certified Company</p>
        </div>

        <h2 style="text-align:center; margin:20px 0 10px; color:#003366;">PURCHASE ORDER</h2>

        <div style="display:flex; justify-content:space-between; margin-bottom:20px; font-size:12px;">
          <div>
            <strong>PO Number:</strong> {{PO_Number}}<br/>
            <strong>Date:</strong> {{Date}}<br/>
            <strong>Project Code:</strong> {{Project_Code}}
          </div>
          <div style="text-align:right;">
            <strong>Vendor:</strong> {{Vendor_Name}}<br/>
            <strong>Delivery Date:</strong> {{Delivery_Date}}
          </div>
        </div>

        <table style="width:100%; border-collapse:collapse; margin-bottom:20px; font-size:11px;">
          <thead>
            <tr style="background-color:#003366; color:white;">
              <th style="padding:8px; border:1px solid #003366; text-align:left;">Description</th>
              <th style="padding:8px; border:1px solid #003366; text-align:center;">Qty</th>
              <th style="padding:8px; border:1px solid #003366; text-align:right;">Unit Price</th>
              <th style="padding:8px; border:1px solid #003366; text-align:right;">Total</th>
            </tr>
          </thead>
          <tbody>{{Line_Items}}</tbody>
        </table>

        <div style="text-align:right; font-size:12px;">
          <div>Subtotal: <strong>{{Subtotal}}</strong></div>
          <div>VAT (5%): <strong>{{VAT}}</strong></div>
          <div>WHT (2%): <strong>-{{WHT}}</strong></div>
          <div style="margin-top:8px; font-size:14px; border-top:2px solid #003366; padding-top:6px;">
            <strong>Grand Total: {{Total_Amount}}</strong>
          </div>
        </div>

        <div style="margin-top:25px; font-size:11px;">
          <strong>Amount in Words:</strong> {{Amount_In_Words}} only.
        </div>

        <div style="margin-top:35px; font-size:10px; border-top:1px solid #ccc; padding-top:10px;">
          <strong>Terms &amp; Conditions:</strong><br/>
          1. Payment within 30 days of delivery.<br/>
          2. Goods subject to inspection and approval.<br/>
          3. This PO is valid for 90 days from date of issue.
        </div>
      </div>
    `);

    // Payment Request Template
    this.templates.set('PaymentRequest', `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; padding: 30px; width: 210mm;">
        <div style="text-align:center; border-bottom: 3px solid #003366; padding-bottom:15px;">
          <h1 style="margin:0; color:#003366;">BONNEDO ENTERPRISE</h1>
          <p style="margin:4px 0; font-size:11px;">ISO 9001:2015 Certified | Finance Division</p>
        </div>

        <h2 style="text-align:center; margin:20px 0;">PAYMENT REQUEST</h2>

        <div style="margin-bottom:20px; font-size:12px;">
          <strong>PR Number:</strong> {{PR_Number}}<br/>
          <strong>Vendor:</strong> {{Vendor_Name}}<br/>
          <strong>Amount:</strong> {{Total_Amount}}<br/>
          <strong>Requested By:</strong> {{Requested_By}}
        </div>

        <div style="margin-top:30px; font-size:11px;">
          <strong>Notes:</strong> {{Notes}}
        </div>

        <div style="margin-top:40px; font-size:10px; border-top:1px solid #ccc; padding-top:10px;">
          <strong>Authorization:</strong><br/>
          Manager: ________________ &nbsp;&nbsp; Finance Lead: ________________ &nbsp;&nbsp; Director: ________________
        </div>
      </div>
    `);
  }

  public getTemplate(name: string): string | undefined {
    return this.templates.get(name);
  }

  public registerTemplate(name: string, html: string) {
    this.templates.set(name, html);
  }
}
