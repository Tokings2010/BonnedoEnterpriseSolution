import * as React from 'react';
import { getTheme, mergeStyleSets, Pivot, PivotItem } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import FinancePaymentRequests from './FinancePaymentRequests';
import FinanceApprovedPayments from './FinanceApprovedPayments';
import FinanceBudgetTracking from './FinanceBudgetTracking';
import FinanceExpenseRegister from './FinanceExpenseRegister';
import FinancePaymentForm from './FinancePaymentForm';
import ExpenseForm from './ExpenseForm';

export interface IFinanceModuleProps {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
  userDisplayName: string;
  webPartContext?: WebPartContext;
  onBack?: () => void;
}

const FinanceModule: React.FC<IFinanceModuleProps> = ({
  spHttpClient,
  pageContext,
  userDisplayName,
  webPartContext,
  onBack,
}) => {
  const theme = getTheme();
  const [selectedTab, setSelectedTab] = React.useState<string>('payment-requests');
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [isMobileView, setIsMobileView] = React.useState(window.innerWidth < 768);

  // Form panel states
  const [isPaymentFormOpen, setIsPaymentFormOpen] = React.useState(false);
  const [isExpenseFormOpen, setIsExpenseFormOpen] = React.useState(false);

  // Handle window resize for responsive behavior
  React.useEffect(() => {
    const handleResize = (): void => {
      setIsMobileView(window.innerWidth < 768);
    };

    window.addEventListener('resize', handleResize);
    return (): void => {
      window.removeEventListener('resize', handleResize);
    };
  }, []);

  const classNames = mergeStyleSets({
    root: {
      display: 'flex',
      flexDirection: 'column',
      height: '100%',
      backgroundColor: theme.palette.white,
    },
    pivotContainer: {
      flex: 1,
      overflow: 'auto',
    },
  });

  const handleRefresh = (): void => {
    setRefreshKey((prev) => prev + 1);
  };

  const handlePaymentFormSubmit = (): void => {
    setIsPaymentFormOpen(false);
    handleRefresh();
  };

  const handleExpenseFormSubmit = (): void => {
    setIsExpenseFormOpen(false);
    handleRefresh();
  };

  return (
    <div className={classNames.root}>
      <Pivot
        selectedKey={selectedTab}
        onLinkClick={(item?: PivotItem) => {
          if (item) {
            setSelectedTab(item.props.itemKey || 'payment-requests');
          }
        }}
        className={classNames.pivotContainer}
      >
        <PivotItem
          itemKey="payment-requests"
          headerText="Payment Requests"
          itemIcon="Money"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <FinancePaymentRequests
              key={`payment-requests-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onNewRequest={() => setIsPaymentFormOpen(true)}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>

        <PivotItem
          itemKey="approved-payments"
          headerText="Approved Payments"
          itemIcon="CheckMark"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <FinanceApprovedPayments
              key={`approved-payments-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>

        <PivotItem
          itemKey="budget-tracking"
          headerText="Budget Tracking"
          itemIcon="Chart"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <FinanceBudgetTracking
              key={`budget-tracking-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>

        <PivotItem
          itemKey="expense-register"
          headerText="Expense Register"
          itemIcon="Receipt"
        >
          <div style={{ padding: '16px', height: '100%' }}>
            <FinanceExpenseRegister
              key={`expense-register-${refreshKey}`}
              spHttpClient={spHttpClient}
              pageContext={pageContext}
              isMobileView={isMobileView}
              onNewExpense={() => setIsExpenseFormOpen(true)}
              onRefresh={handleRefresh}
            />
          </div>
        </PivotItem>
      </Pivot>

      {/* Payment Request Form Panel */}
      <FinancePaymentForm
        isOpen={isPaymentFormOpen}
        onDismiss={() => setIsPaymentFormOpen(false)}
        onSubmitSuccess={handlePaymentFormSubmit}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        webPartContext={webPartContext}
      />

      {/* Expense Form Panel */}
      <ExpenseForm
        isOpen={isExpenseFormOpen}
        onDismiss={() => setIsExpenseFormOpen(false)}
        onSubmitSuccess={handleExpenseFormSubmit}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        webPartContext={webPartContext}
      />
    </div>
  );
};

export default FinanceModule;
