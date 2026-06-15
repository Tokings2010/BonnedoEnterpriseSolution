# ApprovalPanel Integration Guide

## Overview

The ApprovalPanel component provides a complete invoice approval workflow integrated with SharePoint. This guide explains how to integrate it into your application.

## Quick Start

### 1. Import the Component

```typescript
import ApprovalPanel from './ApprovalPanel';
import { IApprovalRecord } from '../models/DataModels';
```

### 2. Add State Management

```typescript
const [selectedRecord, setSelectedRecord] = React.useState<IApprovalRecord | null>(null);
const [isPanelOpen, setIsPanelOpen] = React.useState(false);
```

### 3. Render the Component

```typescript
<ApprovalPanel
  isOpen={isPanelOpen}
  record={selectedRecord}
  listName="APInvoices"
  spHttpClient={spHttpClient}
  pageContext={pageContext}
  userDisplayName={userDisplayName}
  onDismiss={() => setIsPanelOpen(false)}
  onApprovalComplete={(updatedRecord) => {
    // Handle completion
  }}
/>
```

## Integration with DataGrid

### Complete Example

```typescript
import DataGrid from './DataGrid';
import ApprovalPanel from './ApprovalPanel';
import { IApprovalRecord } from '../models/DataModels';

const FinanceModule: React.FC<IFinanceModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
}) => {
  const [selectedRecord, setSelectedRecord] = React.useState<IApprovalRecord | null>(null);
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [refreshKey, setRefreshKey] = React.useState(0);

  const columns = [
    { key: 'Title', name: 'Invoice', fieldName: 'Title', minWidth: 150 },
    { key: 'Vendor', name: 'Vendor', fieldName: 'Vendor', minWidth: 150 },
    { key: 'Amount', name: 'Amount', fieldName: 'Amount', minWidth: 120 },
    { key: 'Approval_Status', name: 'Status', fieldName: 'Approval_Status', minWidth: 100 },
  ];

  const handleRowSelected = (record: IApprovalRecord) => {
    setSelectedRecord(record);
    setIsPanelOpen(true);
  };

  const handleApprovalComplete = (updatedRecord: IApprovalRecord) => {
    // Refresh data grid
    setRefreshKey((prev) => prev + 1);
    setSelectedRecord(null);
  };

  return (
    <>
      <DataGrid
        key={refreshKey}
        listName="APInvoices"
        columns={columns}
        filterQuery="Approval_Status eq 'Pending'"
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        onRowSelected={handleRowSelected}
      />

      <ApprovalPanel
        isOpen={isPanelOpen}
        record={selectedRecord}
        listName="APInvoices"
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        userDisplayName={userDisplayName}
        onDismiss={() => setIsPanelOpen(false)}
        onApprovalComplete={handleApprovalComplete}
      />
    </>
  );
};
```

## SharePoint List Configuration

### Create the APInvoices List

1. **Go to SharePoint Site**
2. **Create New List** with these columns:

| Column Name | Type | Required | Notes |
|------------|------|----------|-------|
| Title | Text | Yes | Invoice number |
| InvoiceNumber | Text | No | Specific invoice ID |
| Vendor | Text | No | Vendor name |
| Amount | Number | No | Invoice amount |
| FundType | Text | No | Cost center/fund |
| Approval_Status | Choice | Yes | Pending, Approved, Rejected |
| Current_Approver | Text | No | Current approver name |
| Approval_History | Note | No | JSON approval history |
| Comments | Note | No | Additional comments |

### Create Choice Field

**Approval_Status** choices:
- Pending
- Approved
- Rejected

## Component Features

### Display Fields

The panel displays these read-only fields:
- Invoice Number
- Vendor
- Amount (formatted with currency)
- Fund Type
- Current Status (color-coded)

### Approval History

Shows all previous approvals/rejections with:
- Action (Approve/Reject badge)
- Approver name
- Timestamp
- Comment

### User Actions

1. **Add Comment** (optional for approve, required for reject)
2. **Approve Button** - Approves invoice
3. **Reject Button** - Rejects invoice (requires comment)

## SharePoint Updates

### On Approval

```typescript
{
  Approval_Status: 'Approved',
  Current_Approver: userDisplayName,
  Approval_History: JSON.stringify([...history, newAction])
}
```

### On Rejection

```typescript
{
  Approval_Status: 'Rejected',
  Current_Approver: userDisplayName,
  Approval_History: JSON.stringify([...history, newAction])
}
```

## Approval History Format

The component stores approval history as JSON:

```json
[
  {
    "action": "approve",
    "comment": "Looks good",
    "timestamp": "2024-01-15T10:30:00.000Z",
    "approver": "John Doe"
  },
  {
    "action": "reject",
    "comment": "Missing documentation",
    "timestamp": "2024-01-15T11:00:00.000Z",
    "approver": "Jane Smith"
  }
]
```

## Error Handling

### Validation Errors

**Rejection without comment:**
```
Error: "Comment is required when rejecting an invoice."
```

### SharePoint Errors

**Permission denied:**
```
Error: "Failed to update list item: Forbidden"
```

**List not found:**
```
Error: "Failed to update list item: Not Found"
```

## Styling Customization

### Override Styles

Edit `ApprovalPanel.module.scss`:

```scss
.panelContent {
  padding: 20px;
  // Custom styling
}

.historyContainer {
  background-color: custom-color;
  // Custom styling
}

.buttonGroup {
  // Custom button styling
}
```

## State Management

### Component State

```typescript
interface IApprovalPanelState {
  comment: string;              // User's comment
  isLoading: boolean;           // Loading state
  isSubmitting: boolean;        // Submission state
  error: string | null;         // Error message
  successMessage: string | null; // Success message
  approvalHistory: IApprovalAction[]; // Parsed history
}
```

## Callbacks

### onDismiss

Called when panel closes:
```typescript
onDismiss={() => {
  setIsPanelOpen(false);
  setSelectedRecord(null);
}}
```

### onApprovalComplete

Called after successful approval/rejection:
```typescript
onApprovalComplete={(updatedRecord) => {
  console.log('Updated record:', updatedRecord);
  // Refresh data
  setRefreshKey((prev) => prev + 1);
}}
```

## Performance Tips

1. **Memoize Service**: SharePointService is memoized
2. **Lazy Parse History**: History is parsed only when record changes
3. **Efficient Updates**: Only necessary fields are updated
4. **Scrollable Content**: Long histories don't impact performance

## Accessibility

- Keyboard navigation supported
- ARIA labels for screen readers
- Semantic HTML structure
- High contrast support
- Focus management

## Troubleshooting

### Panel doesn't open
- Check `isOpen` prop is true
- Verify `record` is not null
- Ensure component is rendered

### Updates not saving
- Verify user has edit permissions
- Check list name is correct
- Review network requests in DevTools

### History not displaying
- Verify Approval_History field is Note type
- Check JSON format is valid
- Ensure history is being saved

### Comment validation not working
- Check comment field binding
- Verify validation logic
- Check error message display

## Advanced Usage

### Custom Approval Logic

Extend the component to add custom logic:

```typescript
const handleApprove = async () => {
  // Custom pre-approval logic
  if (record.Amount > 10000) {
    // Require additional approval
  }
  
  // Call original approval
  await originalApprove();
};
```

### Multiple Approvers

Implement multi-level approval:

```typescript
const approvalLevels = [
  { level: 1, approver: 'Manager' },
  { level: 2, approver: 'Director' },
  { level: 3, approver: 'CFO' }
];
```

### Approval Routing

Route approvals based on amount:

```typescript
const getApprover = (amount: number) => {
  if (amount > 50000) return 'CFO';
  if (amount > 10000) return 'Director';
  return 'Manager';
};
```

## Integration Checklist

- [ ] Import ApprovalPanel component
- [ ] Create APInvoices SharePoint list
- [ ] Add required fields to list
- [ ] Update FinanceModule to use ApprovalPanel
- [ ] Test approval workflow
- [ ] Test rejection workflow
- [ ] Verify SharePoint updates
- [ ] Test error scenarios
- [ ] Verify approval history
- [ ] Test with different user roles

## Next Steps

1. Create SharePoint list with required fields
2. Update FinanceModule to integrate ApprovalPanel
3. Test approval and rejection workflows
4. Verify SharePoint updates
5. Deploy to production

## Support

For issues:
1. Check APPROVALPANEL_README.md
2. Review code comments
3. Check SharePoint list configuration
4. Review browser console for errors
5. Check network requests in DevTools
