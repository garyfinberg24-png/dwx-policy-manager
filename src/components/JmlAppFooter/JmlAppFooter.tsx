// @ts-nocheck
/**
 * JmlAppFooter - Global footer for the Policy Manager application
 * Forest Teal themed footer with branding, links, and version info
 */
import * as React from 'react';
import { IJmlAppFooterProps, IFooterLinkGroup } from './IJmlAppFooterProps';

const footerStyles: Record<string, React.CSSProperties> = {
  footer: {
    background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
    color: '#94a3b8',
    padding: 0,
    marginTop: 'auto',
    width: '100%',
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
  },
  footerCompact: {
    background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
    color: '#94a3b8',
    padding: 0,
    marginTop: 'auto',
    width: '100%',
    fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif"
  },
  topBar: {
    height: 3,
    background: 'linear-gradient(90deg, #0d9488, #14b8a6, #0d9488)',
    width: '100%'
  },
  mainContent: {
    maxWidth: 1400,
    margin: '0 auto',
    padding: '32px 40px 24px'
  },
  mainContentCompact: {
    maxWidth: 1400,
    margin: '0 auto',
    padding: '16px 40px'
  },
  topSection: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    gap: 40,
    flexWrap: 'wrap' as const
  },
  brandSection: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: 8,
    minWidth: 200
  },
  brandName: {
    fontSize: 16,
    fontWeight: 600,
    color: '#f1f5f9',
    display: 'flex',
    alignItems: 'center',
    gap: 8
  },
  brandIcon: {
    width: 24,
    height: 24,
    borderRadius: 4,
    background: 'linear-gradient(135deg, #0d9488, #14b8a6)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: 12,
    color: '#ffffff',
    fontWeight: 700
  },
  brandDescription: {
    fontSize: 12,
    lineHeight: 1.5,
    color: '#64748b',
    maxWidth: 280
  },
  linkGroupsContainer: {
    display: 'flex',
    gap: 40,
    flexWrap: 'wrap' as const
  },
  linkGroup: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: 8,
    minWidth: 140
  },
  linkGroupTitle: {
    fontSize: 12,
    fontWeight: 600,
    color: '#cbd5e1',
    textTransform: 'uppercase' as const,
    letterSpacing: 0.5,
    marginBottom: 4
  },
  link: {
    fontSize: 13,
    color: '#94a3b8',
    textDecoration: 'none',
    transition: 'color 0.15s',
    cursor: 'pointer'
  },
  divider: {
    height: 1,
    background: '#1e293b',
    margin: '20px 0 16px',
    border: 'none'
  },
  dividerCompact: {
    height: 1,
    background: '#1e293b',
    margin: '12px 0',
    border: 'none'
  },
  bottomRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    flexWrap: 'wrap' as const,
    gap: 12
  },
  copyright: {
    fontSize: 11,
    color: '#475569'
  },
  bottomLinks: {
    display: 'flex',
    gap: 16,
    alignItems: 'center'
  },
  bottomLink: {
    fontSize: 11,
    color: '#475569',
    textDecoration: 'none',
    transition: 'color 0.15s',
    cursor: 'pointer'
  },
  versionBadge: {
    fontSize: 10,
    color: '#64748b',
    padding: '2px 8px',
    borderRadius: 10,
    background: 'rgba(255,255,255,0.05)',
    border: '1px solid rgba(255,255,255,0.08)'
  },
  supportLink: {
    fontSize: 12,
    color: '#14b8a6',
    textDecoration: 'none',
    display: 'flex',
    alignItems: 'center',
    gap: 4,
    transition: 'color 0.15s'
  },
  compactRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    flexWrap: 'wrap' as const,
    gap: 12
  },
  compactLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: 12
  }
};

const defaultLinkGroups: IFooterLinkGroup[] = [
  {
    title: 'Policy Manager',
    links: [
      { text: 'Browse Policies', url: '/sites/PolicyManager/SitePages/PolicyHub.aspx' },
      { text: 'My Policies', url: '/sites/PolicyManager/SitePages/MyPolicies.aspx' },
      { text: 'Help Centre', url: '/sites/PolicyManager/SitePages/PolicyHelp.aspx' }
    ]
  },
  {
    title: 'Resources',
    links: [
      { text: 'User Guide', url: '#' },
      { text: 'FAQs', url: '#' },
      { text: 'Contact Support', url: '#' }
    ]
  }
];

export const DwxAppFooter: React.FC<IJmlAppFooterProps> = ({
  version = '1.0.0',
  supportUrl,
  supportText = 'Support',
  linkGroups,
  compact = true,
  organizationName = 'DWx Digital Workplace'
}) => {
  const currentYear = new Date().getFullYear();
  const groups = linkGroups && linkGroups.length > 0 ? linkGroups : defaultLinkGroups;

  if (compact) {
    return (
      <footer style={footerStyles.footerCompact} data-jml-footer="true">
        <div style={footerStyles.topBar} />
        <div style={footerStyles.mainContentCompact}>
          <div style={footerStyles.compactRow}>
            <div style={footerStyles.compactLeft}>
              <div style={footerStyles.brandIcon}>PM</div>
              <span style={{ fontSize: 13, color: '#cbd5e1', fontWeight: 500 }}>Policy Manager</span>
              <span style={footerStyles.versionBadge}>v{version}</span>
              <span style={footerStyles.copyright}>
                © {currentYear} {organizationName}. All rights reserved.
              </span>
            </div>
            <div style={footerStyles.bottomLinks}>
              {supportUrl && (
                <a href={supportUrl} style={footerStyles.supportLink} target="_blank" rel="noopener noreferrer">
                  {supportText}
                </a>
              )}
              <a href="#" style={footerStyles.bottomLink}>Privacy</a>
              <a href="#" style={footerStyles.bottomLink}>Terms</a>
              <span style={{ fontSize: 11, color: '#334155' }}>
                Powered by DWx
              </span>
            </div>
          </div>
        </div>
      </footer>
    );
  }

  return (
    <footer style={footerStyles.footer} data-jml-footer="true">
      <div style={footerStyles.topBar} />
      <div style={footerStyles.mainContent}>
        {/* Top Section: Brand + Link Groups */}
        <div style={footerStyles.topSection}>
          <div style={footerStyles.brandSection}>
            <div style={footerStyles.brandName}>
              <div style={footerStyles.brandIcon}>PM</div>
              Policy Manager
            </div>
            <div style={footerStyles.brandDescription}>
              Enterprise policy management, distribution, and compliance tracking for your organisation.
            </div>
            {supportUrl && (
              <a href={supportUrl} style={footerStyles.supportLink} target="_blank" rel="noopener noreferrer">
                ⓘ {supportText}
              </a>
            )}
          </div>

          <div style={footerStyles.linkGroupsContainer}>
            {groups.map((group, gi) => (
              <div key={gi} style={footerStyles.linkGroup}>
                <div style={footerStyles.linkGroupTitle}>{group.title}</div>
                {group.links.map((link, li) => (
                  <a
                    key={li}
                    href={link.url}
                    style={footerStyles.link}
                    onMouseEnter={e => (e.currentTarget.style.color = '#14b8a6')}
                    onMouseLeave={e => (e.currentTarget.style.color = '#94a3b8')}
                  >
                    {link.text}
                  </a>
                ))}
              </div>
            ))}
          </div>
        </div>

        {/* Divider */}
        <hr style={footerStyles.divider} />

        {/* Bottom Row */}
        <div style={footerStyles.bottomRow}>
          <span style={footerStyles.copyright}>
            © {currentYear} {organizationName}. All rights reserved.
          </span>
          <div style={footerStyles.bottomLinks}>
            <a href="#" style={footerStyles.bottomLink}
              onMouseEnter={e => (e.currentTarget.style.color = '#94a3b8')}
              onMouseLeave={e => (e.currentTarget.style.color = '#475569')}
            >Privacy Policy</a>
            <a href="#" style={footerStyles.bottomLink}
              onMouseEnter={e => (e.currentTarget.style.color = '#94a3b8')}
              onMouseLeave={e => (e.currentTarget.style.color = '#475569')}
            >Terms of Use</a>
            <a href="#" style={footerStyles.bottomLink}
              onMouseEnter={e => (e.currentTarget.style.color = '#94a3b8')}
              onMouseLeave={e => (e.currentTarget.style.color = '#475569')}
            >Accessibility</a>
            <span style={footerStyles.versionBadge}>v{version}</span>
            <span style={{ fontSize: 11, color: '#334155' }}>
              Powered by DWx
            </span>
          </div>
        </div>
      </div>
    </footer>
  );
};

// Export with both names for compatibility
export default DwxAppFooter;
// Legacy alias for backward compatibility
export const JmlAppFooter = DwxAppFooter;
export type { IJmlAppFooterProps };
