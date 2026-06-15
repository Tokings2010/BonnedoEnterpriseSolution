# ApprovalPanel Component Documentation

## Overview

The `ApprovalPanel` component is a Fluent UI Panel-based component designed for invoice approval workflows. It displays invoice details, approval history, and provides a user-friendly interface for approving or rejecting invoices with comments.

## Features

- **Right-side Panel**: Opens from the right side of the screen
- **Invoice Details Display**: Shows invoice number, vendor, amount, fund type, and status
- **Approval History**: Displays all previous approval actions with timestamps and comments
- **Comment Input**: Text area for adding comments (required for rejection)
- **Approve/Reject Buttons**: Action buttons to approve or reject invoices
- **Validation**: Requires comment when rejecting
- **SharePoint Integration**: Updates list items with approval status and history
- **Loading States**: Shows spinner during submission
- **Error Handling**: Displays user-friendly error messages
- **Success Feedback**: Confirms successful approval/rejection

## Props

```typescript
interface IApprovalPanelProps {
  // Required
  isOpen: boolean;                    // Panel visibility state
  record: IApprovalRecord | null;     // Invoice record to approve
  listName: string;                   // SharePoint list name
  spHttpClient: SPHttpClient;         // SPFx HTTP client
  pageContext: PageContext;           // SPFx page context
  userDisplayName: string;            // Current user's display name

  // Callbacks
  onDismiss: () => void;              // Called when panel closes
  onApprovalComplete?: (record: IApprovalRecord) => void;  // Called after approval/rejection
}
```

## Data Models

### IApprovalRecord

```typescript
interface IApprovalRecord extends IListItem {
  Title: string;
  InvoiceNumber?: string;
  Vendor?: string;
  Amount?: number;
  FundType?: string;
  Approval_Status?: string;           // 'Pending', 'Approved', 'Rejected'
  Current_Approver?: string;
  Approval_History?: string;          // JSON string of approval actions
  Comments?: string;
}
```

### IApprovalAction

```typescript
interface IApprovalAction {
  action: 'approve' | 'reject';
  comment: string;
  timestamp: string;                  // ISO format
  approver: string;
}
```

## Usage Example

### Basic Usage

```typescript
import ApprovalPanel from './ApprovalPanel';
import { IApprovalRecord } from '../models/DataModels';

const MyComponent: React.FC<IMyComponentProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
}) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [selectedRecord, setSelectedRecord] = React.useState<IApprovalRecord | null>(null);

  const handleRowSelected = (record: IApprovalRecord) => {
    setSelectedRecord(record);
    setIsOpen(true);
  };

  const handleApprovalComplete = (updatedRecord: IApprovalRecord) => {
    console.log('Approval completed:', updatedRecord);
    // Refresh data grid or update state
  };

  return (
    <>
      <DataGrid
        listName="APInvoices"
        columns={columns}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        onRowSelected={handleRowSelected}
      />

      <ApprovalPanel
        isOpen={isOpen}
        record={selectedRecord}
        listName="APInvoices"
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        userDisplayName={userDisplayName}
        onDismiss={() => setIsOpen(false)}
        onApprovalComplete={handleApprovalComplete}
      />
    </>
  );
};
```

## SharePoint List Setup

### Required Fields

Create a SharePoint list with these fields:

| Field Name | Type | Required | Description |
|-----------|------|----------|-------------|
| Title | Text | Yes | Invoice number or identifier |
| InvoiceNumber | Text | No | Specific invoice number |
| Vendor | Text | No | Vendor/supplier name |
| Amount | Number | No | Invoice amount |
| FundType | Text | No | Fund type or cost center |
| Approval_Status | Choice | Yes | Pending, Approved, Rejected |
| Current_Approver | Text | No | Name of current approver |
| Approval_History | Note | No | JSON history of approvals |
| Comments | Note | No | Additional comments |

### Choice Field Values

**Approval_Status:**
- Pending
- Approved
- Rejected

## Component Behavior

### Approval Flow

1. **User Opens Panel**
   - Panel displays invoice details
   - Shows approval history if available
   - Comment field is empty

2. **User Reviews Invoice**
   - Reads invoice details
   - Reviews approval history
   - Decides to approve or reject

3. **User Approves**
   - Optionally adds comment
   - Clicks "Approve" button
   - Component updates SharePoint list
   - Shows success message
   - Panel closes automatically

4. **User Rejects**
   - Must add comment (validation)
   - Clicks "Reject" button
   - Component updates SharePoint list
   - Shows success message
   - Panel closes automatically

### Validation Rules

- **Rejection requires comment**: If user tries to reject without a comment, an error message is displayed
- **Comment is optional for approval**: User can approve without adding a comment
- **All fields are read-only**: Invoice details cannot be edited in the panel

## SharePoint Updates

### Fields Updated on Approval

```typescript
{
  Approval_Status: 'Approved',
  Current_Approver: userDisplayName,
  Approval_History: JSON.stringify(updatedHistory)
}
```

### Fields Updated on Rejection

```typescript
{
  Approval_Status: 'Rejected',
  Current_Approver: userDisplayName,
  Approval_History: JSON.stringify(updatedHistory)
}
```

### Approval History Format

```typescript
[
  {
    action: 'approve',
    comment: 'Looks good',
    timestamp: '2024-01-15T10:30:00.000Z',
    approver: 'John Doe'
  },
  {
    action: 'reject',
    comment: 'Missing documentation',
    timestamp: '2024-01-15T11:00:00.000Z',
    approver: 'Jane Smith'
  }
]
```

## Styling

The component uses Fluent UI theming and includes custom SCSS for:
- Panel content layout
- Field display styling
- Approval history styling
- Action badges (approve/reject)
- Button group layout
- Comment box styling

### Custom Styling

Override styles in `ApprovalPanel.module.scss`:

```scss
.panelContent {
  // Custom panel content styling
}

.historyContainer {
  // Custom history display styling
}

.buttonGroup {
  // Custom button group styling
}
```

## Error Handling

### Common Errors

1. **Failed to update list item**
   - Cause: Permission issues or network error
   - Display: Error message in MessageBar
   - Action: User can retry

2. **Comment required for rejection**
   - Cause: User tried to reject without comment
   - Display: Error message
   - Action: User must add comment and retry

3. **Invalid approval history**
   - Cause: Corrupted JSON in Approval_History field
   - Display: Component handles gracefully
   - Action: Starts with empty history

## Performance Considerations

- **Memoization**: SharePointService is memoized to prevent unnecessary re-renders
- **Lazy Loading**: Approval history is parsed only when record changes
- **Efficient Updates**: Only necessary fields are updated in SharePoint
- **Scrollable Content**: Long approval histories are scrollable

## Accessibility

- Keyboard navigation support
- ARIA labels for screen readers
- Semantic HTML structure
- High contrast support
- Focus management

## Browser Support

- Chrome (latest)
- Edge (latest)
- Firefox (latest)
- Safari (latest)

## Dependencies

- @fluentui/react: ^8.125.5
- @microsoft/sp-http: 1.19.0
- @microsoft/sp-page-context: 1.19.0
- react: 17.0.1

## Integration with EnterpriseLayout

The ApprovalPanel is designed to work with the EnterpriseLayout and DataGrid:

```typescript
// In FinanceModule or similar
<>
  <DataGrid
    listName="APInvoices"
    columns={columns}
    spHttpClient={spHttpClient}
    pageContext={pageContext}
    onRowSelected={(record) => {
      setSelectedRecord(record as IApprovalRecord);
      setIsPanelOpen(true);
    }}
  />

  <ApprovalPanel
    isOpen={isPanelOpen}
    record={selectedRecord}
    listName="APInvoices"
    spHttpClient={spHttpClient}
    pageContext={pageContext}
    userDisplayName={userDisplayName}
    onDismiss={() => setIsPanelOpen(false)}
    onApprovalComplete={(updatedRecord) => {
      // Refresh data grid
      setSelectedRecord(null);
      // Trigger data refresh
    }}
  />
</>
```

## Troubleshooting

### Panel doesn't open
- Verify `isOpen` prop is true
- Check `record` prop is not null
- Ensure component is rendered

### Updates not saving
- Check user has edit permissions on list
- Verify list name is correct
- Check network connectivity
- Review browser console for errors

### Approval history not displaying
- Verify Approval_History field contains valid JSON
- Check field is of type "Note" or "Text"
- Ensure history is being saved correctly

### Comment validation not working
- Verify comment field is properly bound
- Check validation logic in handleReject
- Ensure error message is displayed

## Future Enhancements

- [ ] Multiple approvers workflow
- [ ] Approval delegation
- [ ] Email notifications
- [ ] Approval templates
- [ ] Bulk approval actions
- [ ] Approval analytics
- [ ] Audit trail
- [ ] Approval routing rules

## Support

For issues or questions:
1. Check this documentation
2. Review code comments in ApprovalPanel.tsx
3. Check SharePoint list configuration
4. Review browser console for errors
5. Check network requests in DevTools
