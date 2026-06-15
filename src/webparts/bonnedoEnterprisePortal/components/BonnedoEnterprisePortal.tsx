import * as React from 'react';
import type { IBonnedoEnterprisePortalProps } from './IBonnedoEnterprisePortalProps';
import EnterpriseLayout from './EnterpriseLayout';

export default class BonnedoEnterprisePortal extends React.Component<IBonnedoEnterprisePortalProps, {}> {
  public render(): React.ReactElement<IBonnedoEnterprisePortalProps> {
    const { isDarkTheme, userDisplayName, spHttpClient, pageContext, webPartContext, hasTeamsContext, userPermissions, portalMode, fullWidth } = this.props;

    return (
      <EnterpriseLayout
        userDisplayName={userDisplayName}
        isDarkTheme={isDarkTheme}
        spHttpClient={spHttpClient}
        pageContext={pageContext}
        webPartContext={webPartContext}
        hasTeamsContext={hasTeamsContext}
        userPermissions={userPermissions}
        portalMode={portalMode}
        fullWidth={fullWidth}
      />
    );
  }
}
