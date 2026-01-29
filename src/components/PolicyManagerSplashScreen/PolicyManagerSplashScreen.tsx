// @ts-nocheck
import * as React from 'react';
import styles from './PolicyManagerSplashScreen.module.scss';

export interface IPolicyManagerSplashScreenProps {
  /** Application version to display */
  version?: string;
  /** Callback when user clicks sign in */
  onSignIn?: () => void;
  /** Whether to show loading state */
  isLoading?: boolean;
  /** Custom loading message */
  loadingMessage?: string;
  /** Callback when "Don't show again" preference changes */
  onDontShowAgainChange?: (dontShowAgain: boolean) => void;
  /** Initial state for "Don't show again" checkbox */
  dontShowAgainChecked?: boolean;
}

/**
 * Policy Manager Splash Screen
 * Based on DWx Brand Guide - Forest Teal theme
 *
 * Displays the branded landing page with:
 * - DWx branding and Policy Manager title
 * - Key stats and value proposition
 * - Feature highlights
 * - Microsoft SSO sign-in button
 */
export const PolicyManagerSplashScreen: React.FC<IPolicyManagerSplashScreenProps> = ({
  version = 'v1.2.0',
  onSignIn,
  isLoading = false,
  loadingMessage = 'Loading Policy Manager...',
  onDontShowAgainChange,
  dontShowAgainChecked = false
}) => {
  const [dontShowAgain, setDontShowAgain] = React.useState(dontShowAgainChecked);

  const handleCheckboxChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const checked = e.target.checked;
    setDontShowAgain(checked);
    if (onDontShowAgainChange) {
      onDontShowAgainChange(checked);
    }
  };
  // Policy Manager specific stats
  const stats = [
    { value: '100+', label: 'Policies Managed' },
    { value: '99%', label: 'Compliance Rate' },
    { value: '<1hr', label: 'Acknowledgement Time' }
  ];

  // Policy Manager features
  const features = [
    'Create and manage organizational policies',
    'Track employee acknowledgements in real-time',
    'Built-in quiz assessments for comprehension',
    'Automated compliance reporting and analytics',
    'Version control and approval workflows'
  ];

  if (isLoading) {
    return (
      <div className={styles.loadingOverlay}>
        <div className={styles.spinner} />
        <span className={styles.loadingText}>{loadingMessage}</span>
      </div>
    );
  }

  return (
    <div className={styles.splashContainer}>
      {/* Background decorative elements */}
      <div className={styles.backgroundDecoration} />

      {/* Version badge */}
      <div className={styles.versionBadge}>{version}</div>

      <div className={styles.contentWrapper}>
        {/* Left Panel - Branding */}
        <div className={styles.brandingPanel}>
          {/* Logo */}
          <div className={styles.logoContainer}>
            <div className={styles.logoIcon}>
              <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                <path d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                      strokeLinejoin="round"/>
              </svg>
            </div>
            <div className={styles.logoText}>
              <span className={styles.brandLabel}>First Digital</span>
              <span className={styles.brandName}>
                DW<span className={styles.brandX}>x</span>
              </span>
            </div>
          </div>

          {/* App Title */}
          <h1 className={styles.appTitle}>Policy Manager</h1>
          <p className={styles.appSubtitle}>Policy Governance & Compliance</p>
          <p className={styles.appDescription}>
            Streamline your policy lifecycle with intelligent automation,
            real-time compliance tracking, and seamless employee acknowledgements.
          </p>

          {/* Stats */}
          <div className={styles.statsContainer}>
            {stats.map((stat, index) => (
              <div key={index} className={styles.statItem}>
                <div className={styles.statValue}>{stat.value}</div>
                <div className={styles.statLabel}>{stat.label}</div>
              </div>
            ))}
          </div>

          {/* Testimonial */}
          <div className={styles.testimonial}>
            <p className={styles.testimonialQuote}>
              "Policy Manager has transformed how we handle compliance.
              Our acknowledgement rates went from 60% to 99% within the first month."
            </p>
            <p className={styles.testimonialAuthor}>Sarah Johnson</p>
            <p className={styles.testimonialRole}>Head of Compliance, Acme Corp</p>
          </div>
        </div>

        {/* Right Panel - Welcome Card */}
        <div className={styles.welcomePanel}>
          <div className={styles.welcomeCard}>
            <h2 className={styles.welcomeTitle}>Welcome back!</h2>
            <p className={styles.welcomeSubtitle}>
              Sign in to access your policy management dashboard
            </p>

            {/* Feature List */}
            <ul className={styles.featureList}>
              {features.map((feature, index) => (
                <li key={index} className={styles.featureItem}>
                  <span className={styles.featureIcon}>
                    <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                      <path d="M5 13l4 4L19 7"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            strokeLinecap="round"
                            strokeLinejoin="round"/>
                    </svg>
                  </span>
                  <span className={styles.featureText}>{feature}</span>
                </li>
              ))}
            </ul>

            {/* Don't show again checkbox */}
            <label className={styles.dontShowAgainLabel}>
              <input
                type="checkbox"
                checked={dontShowAgain}
                onChange={handleCheckboxChange}
                className={styles.dontShowAgainCheckbox}
              />
              <span className={styles.dontShowAgainText}>Don't show this screen again</span>
            </label>

            {/* Sign In Button */}
            <button
              className={styles.signInButton}
              onClick={onSignIn}
              type="button"
            >
              <span className={styles.microsoftLogo}>
                <svg viewBox="0 0 21 21" xmlns="http://www.w3.org/2000/svg">
                  <rect x="0" y="0" width="10" height="10" fill="#f25022"/>
                  <rect x="11" y="0" width="10" height="10" fill="#7fba00"/>
                  <rect x="0" y="11" width="10" height="10" fill="#00a4ef"/>
                  <rect x="11" y="11" width="10" height="10" fill="#ffb900"/>
                </svg>
              </span>
              Sign in with Microsoft
            </button>

            {/* SSO Note */}
            <p className={styles.ssoNote}>
              Enterprise SSO with your organization credentials
            </p>

            {/* Trust Badges */}
            <div className={styles.trustBadges}>
              <span className={styles.trustBadge}>
                <span className={styles.trustIcon}>🔒</span>
                <span>Secure</span>
              </span>
              <span className={styles.trustBadge}>
                <span className={styles.trustIcon}>✓</span>
                <span>Compliant</span>
              </span>
              <span className={styles.trustBadge}>
                <span className={styles.trustIcon}>💬</span>
                <span>Support</span>
              </span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default PolicyManagerSplashScreen;
