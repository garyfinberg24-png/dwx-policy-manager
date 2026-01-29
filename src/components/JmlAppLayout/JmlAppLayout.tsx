// @ts-nocheck
import * as React from 'react';
import styles from './JmlAppLayout.module.scss';
import { IJmlAppLayoutProps } from './IJmlAppLayoutProps';
import JmlAppHeader from '../JmlAppHeader/JmlAppHeader';
import JmlAppFooter from '../JmlAppFooter/JmlAppFooter';
import { PolicyManagerRole } from '../../services/PolicyRoleService';
import { RoleDetectionService } from '../../services/RoleDetectionService';
import { getHighestPolicyRole } from '../../services/PolicyRoleService';
import { signalAppReady } from '../../utils/SharePointOverrides';
/**
 * DWx App Layout Component
 * Provides a complete page layout with header, content area, and footer
 * Use this to wrap your webpart content for consistent full-page styling
 *
 * Features:
 * - Consistent header with navigation
 * - Branded footer
 *
 * Note: Component files retain Jml naming for import compatibility,
 * but the exported component can be referenced as DwxAppLayout
 */
const DwxAppLayout: React.FC<IJmlAppLayoutProps> = (props) => {
  const {
    // Context
    context,

    // Page Header
    pageTitle,
    pageDescription,
    pageIcon,
    stats,
    breadcrumbs,

    // Navigation
    activeNavKey,
    navItems,
    quickLinks,
    showQuickLinks = true,

    // User/Search/Notifications
    showSearch = true,
    showNotifications = true,
    notificationCount = 0,
    onSearchClick,
    onNotificationsClick,
    showSettings = false,
    onSettingsClick,
    userRole,
    availableRoles,
    onRoleChange,
    policyManagerRole,

    // Footer
    version,
    supportUrl,
    supportText,
    footerLinkGroups,
    compactFooter = true,
    organizationName,

    // Layout
    showHeader = true,
    showFooter = true, // Default true - show DWx branded footer
    maxContentWidth = '1400px',
    contentPadding = '24px',
    children,

  } = props;

  // Auto-detect role if not explicitly provided
  const [detectedRole, setDetectedRole] = React.useState<PolicyManagerRole>(PolicyManagerRole.Admin);

  React.useEffect(() => {
    if (!policyManagerRole && context) {
      // Attempt auto-detection from SharePoint groups
      try {
        const sp = (context as any)._sp || null;
        if (sp) {
          const roleService = new RoleDetectionService(sp);
          roleService.getCurrentUserRoles().then(userRoles => {
            if (userRoles && userRoles.length > 0) {
              setDetectedRole(getHighestPolicyRole(userRoles));
            }
          }).catch(() => {
            // Default to Admin if detection fails (shows all nav items)
            setDetectedRole(PolicyManagerRole.Admin);
          });
        }
      } catch {
        // Default to Admin
        setDetectedRole(PolicyManagerRole.Admin);
      }
    }
  }, [policyManagerRole, context]);

  const effectiveRole = policyManagerRole || detectedRole;

  // Signal app readiness to dismiss loading skeleton and reveal content
  React.useEffect(() => {
    signalAppReady();
  }, []);

  // Custom content wrapper style
  const contentWrapperStyle: React.CSSProperties = {
    maxWidth: maxContentWidth,
    padding: contentPadding
  };

  return (
    <div className={styles.jmlAppLayout}>
      {/* Header */}
      {showHeader && (
        <JmlAppHeader
          context={context}
          pageTitle={pageTitle}
          pageDescription={pageDescription}
          pageIcon={pageIcon}
          stats={stats}
          breadcrumbs={breadcrumbs}
          activeNavKey={activeNavKey}
          navItems={navItems}
          quickLinks={quickLinks}
          showQuickLinks={showQuickLinks}
          showSearch={showSearch}
          showNotifications={showNotifications}
          notificationCount={notificationCount}
          onSearchClick={onSearchClick}
          onNotificationsClick={onNotificationsClick}
          showSettings={showSettings}
          onSettingsClick={onSettingsClick}
          userRole={userRole}
          availableRoles={availableRoles}
          onRoleChange={onRoleChange}
          policyRole={effectiveRole}
        />
      )}

      {/* Main Content Area */}
      <main className={styles.contentArea}>
        <div className={styles.contentWrapper} style={contentWrapperStyle}>
          {children}
        </div>
      </main>

      {/* Footer */}
      {showFooter && (
        <JmlAppFooter
          version={version}
          supportUrl={supportUrl}
          supportText={supportText}
          linkGroups={footerLinkGroups}
          compact={compactFooter}
          organizationName={organizationName}
        />
      )}
    </div>
  );
};

// Export with both names for compatibility
export default DwxAppLayout;
export { DwxAppLayout };
// Legacy alias for backward compatibility
export const JmlAppLayout = DwxAppLayout;
