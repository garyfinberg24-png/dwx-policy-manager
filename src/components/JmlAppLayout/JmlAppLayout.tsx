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
    sp,

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
    dwxHub,

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

  // Auto-detect role if not explicitly provided — default to least privilege (User)
  // Initialize from localStorage cache for instant nav rendering (avoids flash of User-only nav)
  const cachedRole = React.useMemo(() => {
    try {
      const cached = localStorage.getItem('pm_detected_role');
      if (cached && ['User', 'Author', 'Manager', 'Admin'].includes(cached)) {
        return cached as PolicyManagerRole;
      }
    } catch { /* ignore */ }
    return PolicyManagerRole.User;
  }, []);

  const [detectedRole, setDetectedRoleRaw] = React.useState<PolicyManagerRole>(cachedRole);
  const setDetectedRole = React.useCallback((role: PolicyManagerRole) => {
    setDetectedRoleRaw(role);
    try { localStorage.setItem('pm_detected_role', role); } catch { /* ignore */ }
  }, []);

  React.useEffect(() => {
    if (!policyManagerRole && context) {
      // Attempt auto-detection: first from PM_UserProfiles (admin-assigned roles), then SP groups
      try {
        const spInstance = sp || (context as any)._sp || null;
        if (spInstance) {
          // Try admin-assigned roles from PM_UserProfiles first
          const userEmail = context.pageContext?.user?.email || '';
          if (userEmail) {
            spInstance.web.lists.getByTitle('PM_UserProfiles')
              .items.filter("Email eq '" + userEmail.replace(/'/g, "''") + "'")
              .select('PMRole', 'PMRoles')
              .top(1)()
              .then((profiles: any[]) => {
                if (profiles.length > 0) {
                  const profile = profiles[0];
                  // Multi-role: check PMRoles field (semicolon-delimited)
                  const rolesStr = profile.PMRoles || profile.PMRole || 'User';
                  const roles = rolesStr.split(';').map((r: string) => r.trim()).filter(Boolean);
                  // Use highest role for nav filtering
                  const LEVEL: Record<string, number> = { User: 0, Author: 1, Manager: 2, Admin: 3 };
                  const highest = roles.reduce((a: string, b: string) => (LEVEL[b] || 0) > (LEVEL[a] || 0) ? b : a, 'User');
                  setDetectedRole(highest as PolicyManagerRole);
                  return;
                }
                // Fallback to SP group detection
                const roleService = new RoleDetectionService(spInstance);
                roleService.getCurrentUserRoles().then(userRoles => {
                  if (userRoles && userRoles.length > 0) {
                    setDetectedRole(getHighestPolicyRole(userRoles));
                  }
                }).catch(() => setDetectedRole(PolicyManagerRole.User));
              })
              .catch(() => {
                // PM_UserProfiles may not exist — fall back to SP groups
                const roleService = new RoleDetectionService(spInstance);
                roleService.getCurrentUserRoles().then(userRoles => {
                  if (userRoles && userRoles.length > 0) {
                    setDetectedRole(getHighestPolicyRole(userRoles));
                  }
                }).catch(() => setDetectedRole(PolicyManagerRole.User));
              });
          } else {
            const roleService = new RoleDetectionService(spInstance);
            roleService.getCurrentUserRoles().then(userRoles => {
              if (userRoles && userRoles.length > 0) {
                setDetectedRole(getHighestPolicyRole(userRoles));
              }
            }).catch(() => setDetectedRole(PolicyManagerRole.User));
          }
        }
      } catch {
        setDetectedRole(PolicyManagerRole.User);
      }
    }
  }, [policyManagerRole, context]);

  const effectiveRole = policyManagerRole || detectedRole;

  // Signal app readiness to dismiss loading skeleton and reveal content
  React.useEffect(() => {
    signalAppReady();
  }, []);

  // Content wrapper style — only apply overrides if non-default values are passed
  const contentWrapperStyle: React.CSSProperties | undefined =
    (maxContentWidth !== '1400px' || contentPadding !== '24px')
      ? { maxWidth: maxContentWidth, padding: contentPadding }
      : undefined;

  return (
    <div className={styles.jmlAppLayout}>
      {/* Header */}
      {showHeader && (
        <JmlAppHeader
          context={context}
          sp={sp}
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
          dwxHub={dwxHub}
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
