# Enterprise Portal Integration Guide

## Project Structure

```
src/webparts/bonnedoEnterprisePortal/
├── components/
│   ├── BonnedoEnterprisePortal.tsx          # Main component wrapper
│   ├── BonnedoEnterprisePortal.module.scss  # Main styles
│   ├── EnterpriseLayout.tsx                 # Layout component
│   ├── EnterpriseLayout.module.scss         # Layout styles
│   ├── DataGrid.tsx                         # Data grid component
│   ├── DataGrid.module.scss                 # Data grid styles
│   ├── FinanceModule.tsx                    # Example module
│   ├── IBonnedoEnterprisePortalProps.ts     # Props interface
│   ├── index.ts                             # Component exports
│   └── DATAGRID_README.md                   # DataGrid documentation
├── services/
│   └── SharePointService.ts                 # SharePoint API service
├── models/
│   └── DataModels.ts                        # Type definitions
├── layouts/
├── loc/
├── assets/
└── BonnedoEnterprisePortalWebPart.ts        # Web part class
```

## Component Architecture

### 1. BonnedoEnterprisePortal (Main Component)
- Entry point for the web part
- Passes props to EnterpriseLayout
- Handles SPFx context

### 2. EnterpriseLayout (Layout Component)
- Top navigation with CommandBar
- Side navigation with Nav component
- Main content area
- Right panel for details
- State management for menu selection and record selection

### 3. DataGrid (Reusable Component)
- Displays SharePoint list data
- Handles pagination, sorting, filtering
- Emits row selection events
- Integrates with SharePointService

### 4. SharePointService (Service Layer)
- Encapsulates SharePoint API calls
- Handles data fetching with OData
- Manages list metadata

## Integration Steps

### Step 1: Update BonnedoEnterprisePortalWebPart.ts

Add SPHttpClient to the component props:

```typescript
import { SPHttpClient } from '@microsoft/sp-http';

const element: React.ReactElement<IBonnedoEnterprisePortalProps> = React.createElement(
  BonnedoEnterprisePortal,
  {
    description: this.properties.description,
    isDarkTheme: this._isDarkTheme,
    environmentMessage: this._environmentMessage,
    hasTeamsContext: !!this.context.sdks.microsoftTeams,
    userDisplayName: this.context.pageContext.user.displayName,
    spHttpClient: this.context.spHttpClient,
    pageContext: this.context.pageContext,
  }
);
```

### Step 2: Update IBonnedoEnterprisePortalProps.ts

Add new props:

```typescript
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export interface IBonnedoEnterprisePortalProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
}
```

### Step 3: Update EnterpriseLayout.tsx

Integrate DataGrid into modules:

```typescript
import DataGrid from './DataGrid';
import { IDataGridColumn } from './DataGrid';

// In your module rendering
<DataGrid
  listName="MyList"
  columns={columns}
  spHttpClient={spHttpClient}
  pageContext={pageContext}
  onRowSelected={handleRecordSelect}
/>
```

### Step 4: Create Module Components

Create specific module components (Finance, Procurement, etc.) that use DataGrid:

```typescript
const FinanceModule: React.FC<IModuleProps> = ({
  spHttpClient,
  pageContext,
  onRecordSelected,
}) => {
  const columns: IDataGridColumn[] = [
    // Define columns
  ];

  return (
    <DataGrid
      listName="Invoices"
      columns={columns}
      spHttpClient={spHttpClient}
      pageContext={pageContext}
      onRowSelected={onRecordSelected}
    />
  );
};
```

## Usage Examples

### Example 1: Finance Module with DataGrid

```typescript
import DataGrid, { IDataGridColumn } from './DataGrid';
import { IListItem } from '../models/DataModels';

const FinanceModule: React.FC<IModuleProps> = ({
  spHttpClient,
  pageContext,
}) => {
  const [selectedRecord, setSelectedRecord] = React.useState<IListItem | null>(null);

  const columns: IDataGridColumn[] = [
    {
      key: 'Title',
      name: 'Invoice Number',
      fieldName: 'Title',
      minWidth: 150,
      isResizable: true,
    },
    {
      key: 'Amount',
      name: 'Amount',
      fieldName: 'Amount',
      minWidth: 120,
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
      key: 'DueDate',
      name: 'Due Date',
      fieldName: 'DueDate',
      minWidth: 120,
      isResizable: true,
    },
  ];

  return (
    <DataGrid
      listName="APInvoices"
      columns={columns}
      filterQuery="Status eq 'Pending'"
      pageSize={15}
      spHttpClient={spHttpClient}
      pageContext={pageContext}
      onRowSelected={(record) => {
        setSelectedRecord(record);
        // Trigger panel open in parent
      }}
    />
  );
};
```

### Example 2: Procurement Module with Filtering

```typescript
const ProcurementModule: React.FC<IModuleProps> = ({
  spHttpClient,
  pageContext,
}) => {
  const [filterStatus, setFilterStatus] = React.useState('Active');

  const columns: IDataGridColumn[] = [
    {
      key: 'Title',
      name: 'PO Number',
      fieldName: 'Title',
      minWidth: 150,
      isResizable: true,
    },
    {
      key: 'Vendor',
      name: 'Vendor',
      fieldName: 'Vendor',
      minWidth: 150,
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

  return (
    <DataGrid
      listName="PurchaseOrders"
      columns={columns}
      filterQuery={`Status eq '${filterStatus}'`}
      spHttpClient={spHttpClient}
      pageContext={pageContext}
      onRowSelected={(record) => console.log('Selected PO:', record)}
    />
  );
};
```

## SharePoint List Setup

### Required List Fields

For the Finance module, create a SharePoint list with these fields:

```
Title (Text) - Invoice Number
Amount (Number) - Invoice Amount
Status (Choice) - Pending, Approved, Rejected, Paid
DueDate (Date) - Due Date
Approver (Person) - Approver Name
Description (Text) - Invoice Description
```

### OData Filter Examples

```typescript
// Pending invoices
"Status eq 'Pending'"

// High value invoices
"Amount gt 5000"

// Recent invoices
"Created ge datetime'2024-01-01T00:00:00'"

// Specific approver
"Approver/Title eq 'John Doe'"

// Multiple conditions
"Status eq 'Pending' and Amount gt 1000 and Created ge datetime'2024-01-01T00:00:00'"
```

## Styling and Theming

### Fluent UI Theme Integration

The components automatically use the Fluent UI theme:

```typescript
import { getTheme } from '@fluentui/react';

const theme = getTheme();
// Access theme colors
theme.palette.themePrimary
theme.palette.neutralLight
theme.palette.white
```

### Custom Styling

Override styles in module SCSS files:

```scss
@import '~@fluentui/react/dist/sass/References.scss';

.customRow {
  &:hover {
    background-color: $ms-color-themeLighter;
  }
}
```

## Error Handling

### Common Errors and Solutions

1. **List Not Found**
   - Verify list name matches exactly (case-sensitive)
   - Check user has access to the list

2. **Invalid Filter Query**
   - Validate OData syntax
   - Check field names match SharePoint fields

3. **Permission Denied**
   - Ensure user has read permissions
   - Check list visibility settings

4. **Network Errors**
   - Check SharePoint site availability
   - Verify SPHttpClient configuration

## Performance Optimization

### Best Practices

1. **Use Pagination**
   - Set appropriate pageSize (10-25 items)
   - Reduces initial load time

2. **Filter Data**
   - Use filterQuery to reduce data set
   - Improves sorting and pagination performance

3. **Limit Columns**
   - Only display necessary columns
   - Reduces rendering overhead

4. **Memoization**
   - SharePointService is memoized
   - Prevents unnecessary API calls

## Testing

### Unit Testing Example

```typescript
import { render, screen } from '@testing-library/react';
import DataGrid from './DataGrid';

describe('DataGrid', () => {
  it('should render with columns', () => {
    const columns = [
      { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 100 },
    ];

    render(
      <DataGrid
        listName="TestList"
        columns={columns}
        spHttpClient={mockSpHttpClient}
        pageContext={mockPageContext}
      />
    );

    expect(screen.getByText('Title')).toBeInTheDocument();
  });
});
```

## Deployment

### Build and Package

```bash
# Build the solution
npm run build

# Package for deployment
gulp bundle --ship
gulp package-solution --ship
```

### SharePoint Deployment

1. Upload .sppkg file to App Catalog
2. Approve the app
3. Add web part to SharePoint page
4. Configure list names in property pane

## Support and Troubleshooting

### Debug Mode

Enable debug logging:

```typescript
// In SharePointService.ts
console.log('Fetching data from:', url);
console.log('Response:', data);
```

### Browser DevTools

- Check Network tab for API calls
- Inspect React component state
- Review console for errors

## Future Enhancements

- [ ] Multi-select rows
- [ ] Export to Excel
- [ ] Advanced filtering UI
- [ ] Column customization
- [ ] Inline editing
- [ ] Grouping and aggregation
- [ ] Real-time updates
- [ ] Caching layer
