export interface INewHire {
  id: number;
  employeeName: string;
  nickname: string;
  jobTitle: string;
  department: string;
  hireDate: Date;
  profilePhoto: string;
  hobbies: string;
  favoriteWebsites: string; // JSON array or comma-separated
  personalQuote: string;
  skillset: string; // Comma-separated skills
  isActive: boolean;
  displayOrder: number;
}

export enum ViewMode {
  Carousel = 'Carousel',
  Grid = 'Grid',
  Banner = 'Banner',
  Timeline = 'Timeline'
}

export interface INewHireSpotlightProps {
  sp: any;
  siteUrl: string;
  userDisplayName?: string;

  // Display Settings
  viewMode: ViewMode;
  title: string;
  description: string;

  // Layout Settings
  hideInternalHeader?: boolean; // Hide the internal header when used inside JmlAppLayout

  // Content Settings
  maxItems: number;
  showDepartmentFilter: boolean;
  showHobbies: boolean;
  showSkills: boolean;
  showWebsites: boolean;
  showQuote: boolean;
  showHireDate: boolean;

  // Carousel/Banner Settings
  autoRotate: boolean;
  rotationInterval: number; // seconds
  showNavigationDots: boolean;
  showNavigationArrows: boolean;

  // Styling Settings
  themeColor: string;
  cardBackgroundColor: string;
  textColor: string;
  enableAnimations: boolean;
  cardElevation: string; // 'low', 'medium', 'high'

  // Grid Settings
  columnsPerRow: number; // 2, 3, or 4

  // Timeline Settings
  showTimelineConnector: boolean;
  compactTimeline: boolean;
}

export interface IWebsiteLink {
  url: string;
  title: string;
}
