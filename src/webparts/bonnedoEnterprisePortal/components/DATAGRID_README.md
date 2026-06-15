# DataGrid Component Documentation

## Overview

The `DataGrid` component is a reusable React component built with Fluent UI's `DetailsList` that provides a powerful data display solution for SharePoint lists. It includes pagination, column sorting, row selection, and hover highlighting.

## Features

- **Pagination**: Navigate through large datasets with configurable page sizes
- **Column Sorting**: Click column headers to sort data ascending/descending
- **Row Selection**: Single row selection with callback support
- **Row Hover Highlight**: Visual feedback when hovering over rows
- **Row Double-Click**: Trigger actions on row double-click
- **SharePoint Integration**: Automatic data fetching from SharePoint lists using SPHttpClient
- **Error Handling**: Graceful error display with user-friendly messages
- **Loading States**: Spinner display during data loading
- **Empty States**: Message when no data is available

## Props

### IDataGridProps

```typescript
interface IDataGridProps {
  // Required
  listName: string;                    // Name of the SharePoint list
  columns: IDataGridColumn[];          // Column definitions
  spHttpClient: SPHttpClient;          // SPFx HTTP client
  pageContext: PageContext;            // SPFx page context

  // Optional
  filterQuery?: string;                // OData filter query (e.g., "Status eq 'Active'")
  pageSize?: number;                   // Items per page (default: 10)
  onRowSelected?: (record: IListItem) => void;      // Callback when row is selected
  onRowDoubleClick?: (record: IListItem) => void;   // Callback when row is double-clicked
}
```

### IDataGridColumn

```typescript
interface IDataGridColumn extends IColumn {
  fieldName: string;                   // Field name from SharePoint list
  name: string;                        // Display name
  minWidth?: number;                   // Minimum column width
  maxWidth?: number;                   // Maximum column width
  isResizable?: boolean;               // Allow column resizing
  isSorted?: boolean;                  // Current sort state
  isSortedDescending?: boolean;        // Sort direction
}
```

## Usage Example

### Basic Usage

```typescript
import DataGrid, { IDataGridColumn } from './DataGrid';
import { IListItem } from '../models/DataModels';

const MyComponent: React.FC<IMyComponentProps> = ({
  spHttpClient,
  pageContext,
}) => {
  const columns: IDataGridColumn[] = [
    {
      key: 'Title',
      name: 'Title',
      fieldName: 'Title',
      minWidth: 200,
      isResizable: true,
    },
    {
      key: 'Status',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 100,
      isResizable: true,
    },
    {
      key: 'Amount',
      name: 'Amount',
      fieldName: 'Amount',
      minWidth: 120,
      isResizable: true,
    },
  ];

  const handleRowSelected = (record: IListItem) => {
    console.log('Selected:', record);
  };

  return (
    <DataGrid
      listName="MyList"
      columns={columns}
      spHttpClient={spHttpClient}
      pageContext={pageContext}
      onRowSelected={handleRowSelected}
      pageSize={15}
    />
  );
};
```

### With Filtering

```typescript
<DataGrid
  listName="Invoices"
  columns={columns}
  filterQuery="Status eq 'Pending' and Amount gt 1000"
  spHttpClient={spHttpClient}
  pageContext={pageContext}
  onRowSelected={handleRowSelected}
/>
```

### With Double-Click Handler

```typescript
const handleDoubleClick = (record: IListItem) => {
  // Open detail panel or navigate to record
  console.log('Opening record:', record.ID);
};

<DataGrid
  listName="MyList"
  columns={columns}
  spHttpClient={spHttpClient}
  pageContext={pageContext}
  onRowSelected={handleRowSelected}
  onRowDoubleClick={handleDoubleClick}
/>
```

## SharePoint Service

The `SharePointService` class handles all SharePoint API interactions:

### Methods

#### getListData()
Fetches data from a SharePoint list with pagination and filtering.

```typescript
const items = await sharePointService.getListData(
  'MyList',
  "Status eq 'Active'",
  10,  // top
  0    // skip
);
```

#### getListItemCount()
Gets the total count of items in a list.

```typescript
const count = await sharePointService.getListItemCount(
  'MyList',
  "Status eq 'Active'"
);
```

#### getListFields()
Retrieves all fields from a SharePoint list.

```typescript
const fields = await sharePointService.getListFields('MyList');
```

## OData Filter Examples

### Common Filter Queries

```typescript
// Equals
"Status eq 'Active'"

// Not equals
"Status ne 'Inactive'"

// Greater than
"Amount gt 1000"

// Less than
"Amount lt 5000"

// Greater than or equal
"Amount ge 1000"

// Less than or equal
"Amount le 5000"

// Contains (for text fields)
"substringof('test', Title)"

// Multiple conditions (AND)
"Status eq 'Active' and Amount gt 1000"

// Multiple conditions (OR)
"Status eq 'Active' or Status eq 'Pending'"

// Date comparison
"Created ge datetime'2024-01-01T00:00:00'"

// Starts with
"startswith(Title, 'Invoice')"
```

## Styling

The component uses Fluent UI theming and includes custom SCSS for:
- Row hover effects
- Selected row highlighting
- Column header styling
- Pagination container styling
- Empty and loading states

### Custom Styling

You can override styles by modifying `DataGrid.module.scss`:

```scss
.rowContainer {
  // Custom row styling
  &:hover {
    background-color: custom-color;
  }
}
```

## Integration with EnterpriseLayout

The DataGrid component is designed to work seamlessly with the EnterpriseLayout:

```typescript
// In EnterpriseLayout or a module component
<DataGrid
  listName="FinanceApprovals"
  columns={financeColumns}
  filterQuery="Status eq 'Pending'"
  spHttpClient={this.props.spHttpClient}
  pageContext={this.props.pageContext}
  onRowSelected={(record) => {
    // Open right panel with record details
    setSelectedRecord(record);
    setIsPanelOpen(true);
  }}
/>
```

## Error Handling

The component displays error messages using Fluent UI's `MessageBar`:

```typescript
// Errors are automatically caught and displayed
// Common errors:
// - List not found
// - Invalid filter query
// - Network errors
// - Permission issues
```

## Performance Considerations

1. **Pagination**: Large datasets are automatically paginated to improve performance
2. **Memoization**: SharePointService is memoized to prevent unnecessary re-renders
3. **Lazy Loading**: Data is loaded on-demand based on current page
4. **Column Sorting**: Client-side sorting for better UX

## Accessibility

- Keyboard navigation support via Fluent UI DetailsList
- ARIA labels for screen readers
- Semantic HTML structure
- High contrast support

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

## Troubleshooting

### No data displayed
- Verify list name is correct
- Check filter query syntax
- Ensure user has permissions to access the list

### Sorting not working
- Ensure column fieldName matches SharePoint field name
- Check data types are compatible

### Pagination not showing
- Verify total items exceed page size
- Check pageSize prop is set correctly

### Performance issues
- Reduce page size
- Add more specific filter queries
- Check network performance

## Future Enhancements

- Multi-select rows
- Export to Excel
- Advanced filtering UI
- Column customization
- Inline editing
- Grouping and aggregation
