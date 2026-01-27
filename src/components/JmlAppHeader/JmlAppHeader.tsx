// @ts-nocheck
/**
 * DwxAppHeader - Stub component for standalone Policy Manager
 * Note: File retains Jml naming for import compatibility
 */
import * as React from 'react';

export interface IDwxAppHeaderProps {
  context?: unknown;
}

export const DwxAppHeader: React.FC<IDwxAppHeaderProps> = () => {
  return null;
};

// Export with both names for compatibility
export default DwxAppHeader;
// Legacy alias for backward compatibility
export const JmlAppHeader = DwxAppHeader;
export type IJmlAppHeaderProps = IDwxAppHeaderProps;
