// @ts-nocheck
export interface IFooterLink {
  text: string;
  url: string;
  icon?: string;
}

export interface IFooterLinkGroup {
  title: string;
  links: IFooterLink[];
}

export interface IJmlAppFooterProps {
  context?: unknown;
  version?: string;
  supportUrl?: string;
  supportText?: string;
  linkGroups?: IFooterLinkGroup[];
  compact?: boolean;
  organizationName?: string;
}
