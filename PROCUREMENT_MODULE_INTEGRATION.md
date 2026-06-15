# Procurement Module - Integration Guide

## Quick Start

### 1. Import the Component

```typescript
import ProcurementModule from './ProcurementModule';
```

### 2. Add to EnterpriseLayout

Update your EnterpriseLayout or module selector to include ProcurementModule:

```typescript
import ProcurementModule from './ProcurementModule';

// In your module rendering logic
if (selectedTopMenu === 'procurement') {
  return (
    <ProcurementModule
      spHttpClient={spHttpClient}
      pageContext={pageContext}
      userDisplayName={userDisplayName}
    />
  );
}
```

### 3. Create SharePoint Lists

Create four SharePoint lists with the following fields:

#### PRC_Material_Request_Register
```
Title (Text) - Request number
Project_Code (Text) - Project identifier
Material (Text) - Material name
Quantity (Number) - Quantity requested
UOM (Text) - Unit of measure
Request_Date (DateTime) - Request date
Status (Choice) - Draft, Submitted, Approved, Rejected, Completed
Approval_Status (Choice) - Pending, Approved, Rejected
QR_Code (Text) - QR code data
Notes (Note) - Additional notes
```

#### PRC_Purchase_Requisition_Register
```
Title (Text) - Requisition number
Project_Code (Text) - Project identifier
Description (Text) - Item description
Quantity (Number) - Quantity required
UOM (Text) - Unit of measure
EstimatedCost (Number) - Estimated cost
Request_Date (DateTime) - Request date
Status (Choice) - Draft, Submitted, Approved, Rejected, Completed
Approval_Status (Choice) - Pending, Approved, Rejected
QR_Code (Text) - QR code data
```

#### PRC_Purchase_Order_Register
```
Title (Text) - PO number
Vendor (Text) - Vendor name
Description (Text) - Item description
Quantity (Number) - Quantity ordered
UOM (Text) - Unit of measure
UnitPrice (Number) - Unit price
TotalAmount (Number) - Total amount
DeliveryDate (DateTime) - Expected delivery date
Status (Choice) - Draft, Submitted, Approved, Rejected, Completed
Approval_Status (Choice) - Pending, Approved, Rejected
QR_Code (Text) - QR code data
```

#### PRC_Goods_Received_Note_Register
```
Title (Text) - GRN number
PO_Number (Text) - Related PO number
Vendor (Text) - Vendor name
Quantity_Received (Number) - Quantity received
UOM (Text) - Unit of measure
Received_Date (DateTime) - Receipt date
Status (Choice) - Draft, Received, Verified, Completed
QR_Code (Text) - QR code data
Notes (Note) - Additional notes
```

---

## Component Usage

### ProcurementModule

Main component that manages all four submodules.

```typescript
<ProcurementModule
  spHttpClient={spHttpClient}
  pageContext={pageContext}
  userDisplayName={userDisplayName}
/>
```

**Props:**
- `spHttpClient` (SPHttpClient) - SharePoint HTTP client
- `pageContext` (PageContext) - SharePoint page context
- `userDisplayName` (string) - Current user's display name

### ProcurementForm

Reusable form component for creating records.

```typescript
<ProcurementForm
  formType="MR"
  listName="PRC_Material_Request_Register"
  spHttpClient={spHttpClient}
  pageContext={pageContext}
  onSubmitSuccess={(itemId) => console.log('Created:', itemId)}
  onCancel={() => console.log('Cancelled')}
/>
```

**Props:**
- `formType` ('MR' | 'PR' | 'PO' | 'GRN') - Type of record
- `listName` (string) - SharePoint list name
- `spHttpClient` (SPHttpClient) - HTTP client
- `pageContext` (PageContext) - Page context
- `onSubmitSuccess` (callback) - Called after successful creation
- `onCancel` (callback) - Called when cancelled

### ProcurementTable

Reusable table component with search, filter, and sort.

```typescript
<ProcurementTable
  listName="PRC_Material_Request_Register"
  recordType="MR"
  columns={columns}
  spHttpClient={spHttpClient}
  pageContext={pageContext}
  onRowSelected={(record) => console.log('Selected:', record)}
/>
```

**Props:**
- `listName` (string) - SharePoint list name
- `recordType` ('MR' | 'PR' | 'PO' | 'GRN') - Record type
- `columns` (IColumn[]) - Column definitions
- `spHttpClient` (SPHttpClient) - HTTP client
- `pageContext` (PageContext) - Page context
- `onRowSelected` (callback) - Called when row selected
- `filterQuery` (string, optional) - OData filter
- `pageSize` (number, optional) - Items per page

### ProcurementDetailsPanel

Right-side panel for viewing record details.

```typescript
<ProcurementDetailsPanel
  isOpen={true}
  record={selectedRecord}
  recordType="MR"
  onDismiss={() => console.log('Closed')}
  onApprove={() => console.log('Approved')}
  onReject={() => console.log('Rejected')}
  showApprovalButtons={true}
/>
```

**Props:**
- `isOpen` (boolean) - Panel visibility
- `record` (any) - Record to display
- `recordType` ('MR' | 'PR' | 'PO' | 'GRN') - Record type
- `onDismiss` (callback) - Called when closed
- `onApprove` (callback, optional) - Approve action
- `onReject` (callback, optional) - Reject action
- `showApprovalButtons` (boolean, optional) - Show buttons

---

## Features

### Search
- Search by record number
- Search by project code
- Search by material/vendor/description
- Real-time filtering

### Filtering
- Filter by status
- Dropdown selector
- Multiple status options

### Sorting
- Click column headers to sort
- Ascending/descending toggle
- Visual sort indicators

### Mobile Responsive
- Desktop: DetailsList view
- Mobile (< 768px): Card layout
- Touch-friendly interface
- Readable text sizes

### QR Code
- Generated using QR Server API
- Encodes: RecordType-RecordID-RecordNumber
- 200x200px display
- No external dependencies

### Approval Buttons
- Approve button (green)
- Reject button (gray)
- Only shown for pending records
- Callback support

---

## Form Fields

### Material Request Form
- Project Code (required)
- Material (required)
- Quantity (required)
- Unit of Measure (required)
- Notes (optional)

### Purchase Requisition Form
- Project Code (required)
- Description (required)
- Quantity (required)
- Unit of Measure (required)
- Estimated Cost (optional)
- Notes (optional)

### Purchase Order Form
- Vendor (required)
- Description (optional)
- Quantity (required)
- Unit of Measure (required)
- Unit Price (optional)
- Notes (optional)

### Goods Received Note Form
- PO Number (required)
- Vendor (required)
- Quantity Received (required)
- Unit of Measure (required)
- Notes (optional)

---

## Validation

### Form Validation
- Project Code: Required
- Quantity: Must be > 0
- Material/Description/Vendor: Required based on form type
- UOM: Required

### Error Messages
- User-friendly error messages
- Displayed in MessageBar
- Clear guidance for fixing

---

## SharePoint Integration

### Create Record
```typescript
POST /_api/web/lists/getByTitle('ListName')/items
Body: {
  Title: 'MR-1234567890',
  Project_Code: 'PROJ001',
  Material: 'Steel Pipe',
  Quantity: 100,
  UOM: 'PCS',
  Status: 'Draft',
  Approval_Status: 'Pending'
}
```

### Read Records
```typescript
GET /_api/web/lists/getByTitle('ListName')/items
?$filter=Status eq 'Draft'
&$select=*
```

### Update Record
```typescript
MERGE /_api/web/lists/getByTitle('ListName')/items(1)
Body: {
  Status: 'Approved',
  Approval_Status: 'Approved'
}
```

---

## Mobile Responsive Behavior

### Desktop (> 768px)
- DetailsList with all columns
- Full column headers
- Hover effects
- Sorting indicators

### Mobile (< 768px)
- Card layout
- Key information displayed
- Full-width cards
- Touch-friendly spacing
- Stacked buttons

### Breakpoint
- 768px (iPad portrait width)
- Automatic detection
- Responsive to window resize

---

## Error Handling

### Common Errors

**List not found:**
```
Error: Failed to fetch data: Not Found
```
Solution: Verify list name is correct

**Permission denied:**
```
Error: Failed to create item: Forbidden
```
Solution: Check user has edit permissions

**Invalid data:**
```
Error: Validation failed
```
Solution: Check form field values

---

## Performance Tips

1. **Use pagination** - Set appropriate pageSize
2. **Filter data** - Use filterQuery to reduce dataset
3. **Limit columns** - Only display necessary columns
4. **Memoization** - Services are memoized
5. **Lazy loading** - Data loaded on demand

---

## Accessibility

- Keyboard navigation supported
- ARIA labels for screen readers
- Semantic HTML structure
- High contrast support
- Focus management

---

## Browser Support

- Chrome (latest)
- Edge (latest)
- Firefox (latest)
- Safari (latest)

---

## Dependencies

- @fluentui/react: ^8.125.5
- @microsoft/sp-http: 1.19.0
- @microsoft/sp-page-context: 1.19.0
- react: 17.0.1

---

## Troubleshooting

### Records not loading
- Check list name is correct
- Verify user has read permissions
- Check network connectivity

### Form not submitting
- Verify all required fields are filled
- Check user has edit permissions
- Review browser console for errors

### QR code not displaying
- Check internet connectivity (QR Server API)
- Verify record has ID and Title
- Check Image component is rendering

### Mobile layout not switching
- Check window width detection
- Verify CSS media queries
- Test with browser DevTools

---

## Next Steps

1. Create SharePoint lists
2. Add ProcurementModule to EnterpriseLayout
3. Test create workflow
4. Test read/search/filter workflow
5. Test mobile responsive layout
6. Deploy to production

---

## Support

For issues or questions:
1. Check this documentation
2. Review code comments
3. Check SharePoint list configuration
4. Review browser console for errors
5. Check network requests in DevTools
