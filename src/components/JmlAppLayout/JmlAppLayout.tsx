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
      // Role detection chain: IsSiteAdmin → PM_UserProfiles → SP group membership → User fallback
      try {
        const spInstance = sp || (context as any)._sp || null;
        if (spInstance) {
          const userEmail = context.pageContext?.user?.email || '';

          // Helper: fall back to SP group-based detection
          const detectViaGroups = (): void => {
            const roleService = new RoleDetectionService(spInstance);
            roleService.getCurrentUserRoles().then(userRoles => {
              if (userRoles && userRoles.length > 0) {
                const role = getHighestPolicyRole(userRoles);
                console.log('[PolicyManager] Role detected via SP groups:', role);
                setDetectedRole(role);
              } else {
                console.log('[PolicyManager] No SP group roles found, defaulting to User');
                setDetectedRole(PolicyManagerRole.User);
              }
            }).catch(() => {
              console.warn('[PolicyManager] SP group detection failed, defaulting to User');
              setDetectedRole(PolicyManagerRole.User);
            });
          };

          // Helper: check PM_UserProfiles then fall back to groups
          const detectViaProfiles = (): void => {
            if (!userEmail) { detectViaGroups(); return; }
            spInstance.web.lists.getByTitle('PM_UserProfiles')
              .items.filter("Email eq '" + userEmail.replace(/'/g, "''") + "'")
              .select('PMRole', 'PMRoles')
              .top(1)()
              .then((profiles: any[]) => {
                if (profiles.length > 0) {
                  const profile = profiles[0];
                  const rolesStr = profile.PMRoles || profile.PMRole || 'User';
                  const roles = rolesStr.split(';').map((r: string) => r.trim()).filter(Boolean);
                  const LEVEL: Record<string, number> = { User: 0, Author: 1, Manager: 2, Admin: 3 };
                  const highest = roles.reduce((a: string, b: string) => (LEVEL[b] || 0) > (LEVEL[a] || 0) ? b : a, 'User');
                  console.log('[PolicyManager] Role detected via PM_UserProfiles:', highest);
                  setDetectedRole(highest as PolicyManagerRole);
                  return;
                }
                console.log('[PolicyManager] No PM_UserProfiles record, falling back to SP groups');
                detectViaGroups();
              })
              .catch(() => {
                console.log('[PolicyManager] PM_UserProfiles unavailable, falling back to SP groups');
                detectViaGroups();
              });
          };

          // Step 1: Check IsSiteAdmin — Site Collection Admins are always Admin
          spInstance.web.currentUser.select('IsSiteAdmin')()
            .then((user: any) => {
              if (user.IsSiteAdmin) {
                console.log('[PolicyManager] User is Site Collection Admin → Admin role');
                setDetectedRole(PolicyManagerRole.Admin);
                return;
              }
              // Step 2: Check PM_UserProfiles → Step 3: SP groups
              detectViaProfiles();
            })
            .catch(() => {
              // IsSiteAdmin check failed — continue with profile/group detection
              console.warn('[PolicyManager] IsSiteAdmin check failed, continuing with profile detection');
              detectViaProfiles();
            });
        }
      } catch {
        console.warn('[PolicyManager] Role detection failed entirely, defaulting to User');
        setDetectedRole(PolicyManagerRole.User);
      }
    }
  }, [policyManagerRole, context]);

  const effectiveRole = policyManagerRole || detectedRole;

  // Signal app readiness to dismiss loading skeleton and reveal content
  React.useEffect(() => {
    signalAppReady();

    // Load and apply custom theme — SP is authoritative, localStorage is fast fallback
    try {
      const { ThemeManager } = require('../../utils/themeManager');
      // Step 1: Apply from localStorage immediately (fast, prevents flash of default theme)
      const stored = ThemeManager.getTheme();
      if (stored && stored.primaryColor && stored.primaryColor !== '#0d9488') {
        ThemeManager.apply(stored);
      }
      // Step 2: Load from SP (authoritative — overrides localStorage for all users)
      if (sp) {
        ThemeManager.loadFromSP(sp).then((spTheme: any) => {
          if (spTheme && spTheme.primaryColor) {
            // Always apply SP theme — it's the admin-configured source of truth
            ThemeManager.apply(spTheme);
            // Sync to localStorage so next page load is instant
            try { localStorage.setItem('pm_custom_theme', JSON.stringify(spTheme)); } catch { /* */ }
          }
        }).catch(() => { /* SP unavailable — localStorage fallback is already applied */ });
      }
    } catch { /* ThemeManager not available — use defaults */ }
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
