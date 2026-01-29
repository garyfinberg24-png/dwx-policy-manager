// @ts-nocheck
import * as React from 'react';
import styles from './PageSubheader.module.scss';

export interface IPageSubheaderProps {
  /** Page title */
  title: string;
  /** Page description/subtitle */
  description?: string;
  /** Fluent UI icon name (e.g., 'Library', 'Settings') */
  iconName?: string;
  /** Custom icon element (overrides iconName) */
  icon?: React.ReactNode;
  /** Optional badge count to display */
  badgeCount?: number;
  /** Optional action buttons on the right side */
  actions?: React.ReactNode;
  /** Optional children rendered below title */
  children?: React.ReactNode;
}

/**
 * DWx Standard Page Subheader Panel
 * Consistent page title area used across all Policy Manager pages.
 * Features green left border accent and subtle green gradient fill.
 * Based on DWx Brand Guide / Contract Manager pattern.
 */
export const PageSubheader: React.FC<IPageSubheaderProps> = ({
  title,
  description,
  iconName,
  icon,
  badgeCount,
  actions,
  children
}) => {
  return (
    <div className={styles.pageSubheader}>
      <div className={styles.pageSubheaderLeft}>
        {(icon || iconName) && (
          <div className={styles.pageSubheaderIcon}>
            {icon || (
              <i className={`ms-Icon ms-Icon--${iconName}`} aria-hidden="true" />
            )}
          </div>
        )}
        <div className={styles.pageSubheaderContent}>
          <div className={styles.pageSubheaderTitleRow}>
            <span className={styles.pageSubheaderTitle}>{title}</span>
            {badgeCount !== undefined && badgeCount > 0 && (
              <span className={styles.pageSubheaderBadge}>{badgeCount}</span>
            )}
          </div>
          {description && (
            <span className={styles.pageSubheaderDescription}>{description}</span>
          )}
          {children}
        </div>
      </div>
      {actions && (
        <div className={styles.pageSubheaderActions}>
          {actions}
        </div>
      )}
    </div>
  );
};

export default PageSubheader;
