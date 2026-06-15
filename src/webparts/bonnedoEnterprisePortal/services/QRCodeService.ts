/**
 * QR Code Service for generating QR codes
 * Uses QR Server API for generating QR codes
 */

export class QRCodeService {
  /**
   * Generate QR code URL for given data
   * @param data - Data to encode in QR code
   * @param size - Size of QR code (default: 200)
   * @returns URL to QR code image
   */
  public static generateQRCodeUrl(data: string, size: number = 200): string {
    if (!data) {
      return '';
    }

    // Using QR Server API (free, no authentication required)
    const encodedData = encodeURIComponent(data);
    return `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodedData}`;
  }

  /**
   * Generate QR code data for a procurement record
   * @param recordType - Type of record (MR, PR, PO, GRN)
   * @param recordId - ID of the record
   * @param recordNumber - Record number/title
   * @returns QR code data string
   */
  public static generateProcurementQRData(
    recordType: 'MR' | 'PR' | 'PO' | 'GRN',
    recordId: number,
    recordNumber: string
  ): string {
    return `${recordType}-${recordId}-${recordNumber}`;
  }

  /**
   * Generate full QR code URL for procurement record
   * @param recordType - Type of record
   * @param recordId - ID of the record
   * @param recordNumber - Record number/title
   * @param size - Size of QR code
   * @returns Full QR code URL
   */
  public static generateProcurementQRCodeUrl(
    recordType: 'MR' | 'PR' | 'PO' | 'GRN',
    recordId: number,
    recordNumber: string,
    size: number = 200
  ): string {
    const qrData = this.generateProcurementQRData(recordType, recordId, recordNumber);
    return this.generateQRCodeUrl(qrData, size);
  }
}
