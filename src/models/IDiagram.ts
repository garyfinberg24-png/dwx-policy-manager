/**
 * Diagram Management Models
 * Interfaces for floor plans, emergency diagrams, and annotated images
 */

import { IBaseListItem, IUser } from './ICommon';

// ============================================================================
// ENUMS
// ============================================================================

/**
 * Types of diagrams supported in the system
 */
export enum DiagramType {
  FloorPlan = 'Floor Plan',
  FireDrill = 'Fire Drill',
  EmergencyExit = 'Emergency Exit',
  EvacuationRoute = 'Evacuation Route',
  SafetyZone = 'Safety Zone',
  AccessMap = 'Access Map',
  Seating = 'Seating Chart',
  Infographic = 'Infographic',
  ProcessFlow = 'Process Flow',
  OrgChart = 'Org Chart',
  Custom = 'Custom'
}

/**
 * Status of a diagram
 */
export enum DiagramStatus {
  Draft = 'Draft',
  InReview = 'In Review',
  Approved = 'Approved',
  Published = 'Published',
  Archived = 'Archived'
}

/**
 * Types of annotation markers
 */
export enum MarkerType {
  Exit = 'Exit',
  FireExtinguisher = 'Fire Extinguisher',
  FirstAid = 'First Aid',
  AssemblyPoint = 'Assembly Point',
  Defibrillator = 'Defibrillator',
  EmergencyPhone = 'Emergency Phone',
  Stairs = 'Stairs',
  Elevator = 'Elevator',
  RestRoom = 'Rest Room',
  Entrance = 'Entrance',
  Danger = 'Danger',
  Warning = 'Warning',
  Information = 'Information',
  YouAreHere = 'You Are Here',
  Custom = 'Custom'
}

/**
 * Shape types for drawing on diagrams
 */
export enum ShapeType {
  Rectangle = 'Rectangle',
  Circle = 'Circle',
  Arrow = 'Arrow',
  Line = 'Line',
  Polygon = 'Polygon',
  Text = 'Text'
}

// ============================================================================
// MARKER AND ANNOTATION INTERFACES
// ============================================================================

/**
 * Position on the diagram (percentage-based for responsiveness)
 */
export interface IDiagramPosition {
  x: number; // Percentage from left (0-100)
  y: number; // Percentage from top (0-100)
}

/**
 * A marker/pin on the diagram
 */
export interface IDiagramMarker {
  id: string;
  type: MarkerType;
  position: IDiagramPosition;
  label?: string;
  description?: string;
  color?: string;
  icon?: string;
  size?: 'small' | 'medium' | 'large';
  visible?: boolean;
}

/**
 * A shape drawn on the diagram
 */
export interface IDiagramShape {
  id: string;
  type: ShapeType;
  points: IDiagramPosition[];
  strokeColor?: string;
  strokeWidth?: number;
  fillColor?: string;
  fillOpacity?: number;
  label?: string;
  dashed?: boolean;
}

/**
 * A route/path on the diagram (for evacuation routes, etc.)
 */
export interface IDiagramRoute {
  id: string;
  name: string;
  points: IDiagramPosition[];
  color?: string;
  width?: number;
  dashed?: boolean;
  arrowStart?: boolean;
  arrowEnd?: boolean;
  description?: string;
}

/**
 * A zone/area on the diagram
 */
export interface IDiagramZone {
  id: string;
  name: string;
  points: IDiagramPosition[]; // Polygon points
  fillColor?: string;
  fillOpacity?: number;
  strokeColor?: string;
  strokeWidth?: number;
  description?: string;
  capacity?: number;
}

/**
 * All annotations for a diagram
 */
export interface IDiagramAnnotations {
  markers: IDiagramMarker[];
  shapes: IDiagramShape[];
  routes: IDiagramRoute[];
  zones: IDiagramZone[];
}

// ============================================================================
// DIAGRAM INTERFACES
// ============================================================================

/**
 * Main diagram interface
 */
export interface IDiagram extends IBaseListItem {
  /** Diagram name/title */
  Title: string;
  /** Type of diagram */
  DiagramType: DiagramType;
  /** Current status */
  Status: DiagramStatus;
  /** Description */
  Description?: string;
  /** Building/location this diagram represents */
  Building?: string;
  /** Floor number/name */
  Floor?: string;
  /** Department or area */
  Department?: string;
  /** Base image URL */
  ImageUrl: string;
  /** Thumbnail URL */
  ThumbnailUrl?: string;
  /** JSON string of annotations */
  AnnotationsJson?: string;
  /** Parsed annotations (client-side only, not stored) */
  annotations?: IDiagramAnnotations;
  /** Version number */
  Version: string;
  /** Tags for categorisation */
  Tags?: string[];
  /** Effective date */
  EffectiveDate?: Date;
  /** Expiration date */
  ExpirationDate?: Date;
  /** Is this the current active version */
  IsActive: boolean;
  /** Created by user */
  CreatedBy?: IUser;
  CreatedById?: number;
  /** Last modified by user */
  ModifiedBy?: IUser;
  ModifiedById?: number;
  /** Approved by user */
  ApprovedBy?: IUser;
  ApprovedById?: number;
  /** Approval date */
  ApprovedDate?: Date;
}

/**
 * Options for uploading a diagram
 */
export interface IDiagramUploadOptions {
  title: string;
  diagramType: DiagramType;
  description?: string;
  building?: string;
  floor?: string;
  department?: string;
  tags?: string[];
  effectiveDate?: Date;
  expirationDate?: Date;
}

/**
 * Search/filter options for diagrams
 */
export interface IDiagramSearchFilters {
  searchText?: string;
  diagramTypes?: DiagramType[];
  statuses?: DiagramStatus[];
  buildings?: string[];
  floors?: string[];
  departments?: string[];
  tags?: string[];
  isActive?: boolean;
  effectiveBefore?: Date;
  effectiveAfter?: Date;
}

/**
 * Diagram template for consistent creation
 */
export interface IDiagramTemplate {
  id: string;
  name: string;
  description: string;
  diagramType: DiagramType;
  /** Pre-configured markers for this template */
  defaultMarkers?: Partial<IDiagramMarker>[];
  /** Suggested marker types for this template */
  suggestedMarkerTypes?: MarkerType[];
  /** Default zones */
  defaultZones?: Partial<IDiagramZone>[];
  /** Preview image URL */
  previewUrl?: string;
}

// ============================================================================
// MARKER CONFIGURATION
// ============================================================================

/**
 * Configuration for marker types with icons and colours
 */
export interface IMarkerConfig {
  type: MarkerType;
  icon: string;
  color: string;
  label: string;
  description: string;
}

/**
 * Default marker configurations
 */
export const DEFAULT_MARKER_CONFIGS: IMarkerConfig[] = [
  { type: MarkerType.Exit, icon: 'DoorArrowLeft', color: '#107c10', label: 'Exit', description: 'Emergency exit point' },
  { type: MarkerType.FireExtinguisher, icon: 'Ringer', color: '#d13438', label: 'Fire Extinguisher', description: 'Fire extinguisher location' },
  { type: MarkerType.FirstAid, icon: 'Medical', color: '#107c10', label: 'First Aid', description: 'First aid kit location' },
  { type: MarkerType.AssemblyPoint, icon: 'People', color: '#0078d4', label: 'Assembly Point', description: 'Emergency assembly point' },
  { type: MarkerType.Defibrillator, icon: 'Heart', color: '#d13438', label: 'AED', description: 'Automated External Defibrillator' },
  { type: MarkerType.EmergencyPhone, icon: 'Phone', color: '#ff8c00', label: 'Emergency Phone', description: 'Emergency telephone' },
  { type: MarkerType.Stairs, icon: 'ChevronUp', color: '#605e5c', label: 'Stairs', description: 'Stairway access' },
  { type: MarkerType.Elevator, icon: 'ChevronUpDown', color: '#605e5c', label: 'Elevator', description: 'Elevator (do not use during fire)' },
  { type: MarkerType.RestRoom, icon: 'PersonFeedback', color: '#0078d4', label: 'Rest Room', description: 'Restroom facilities' },
  { type: MarkerType.Entrance, icon: 'Door', color: '#0078d4', label: 'Entrance', description: 'Building entrance' },
  { type: MarkerType.Danger, icon: 'Warning', color: '#d13438', label: 'Danger', description: 'Danger zone' },
  { type: MarkerType.Warning, icon: 'Warning', color: '#ff8c00', label: 'Warning', description: 'Warning area' },
  { type: MarkerType.Information, icon: 'Info', color: '#0078d4', label: 'Information', description: 'Information point' },
  { type: MarkerType.YouAreHere, icon: 'MyLocation', color: '#881798', label: 'You Are Here', description: 'Current location marker' },
  { type: MarkerType.Custom, icon: 'Pinned', color: '#605e5c', label: 'Custom', description: 'Custom marker' }
];

/**
 * Get marker configuration by type
 */
export function getMarkerConfig(type: MarkerType): IMarkerConfig {
  const config = DEFAULT_MARKER_CONFIGS.find(c => c.type === type);
  return config || DEFAULT_MARKER_CONFIGS[DEFAULT_MARKER_CONFIGS.length - 1];
}

// ============================================================================
// DIAGRAM TEMPLATES
// ============================================================================

/**
 * Pre-defined diagram templates
 */
export const DIAGRAM_TEMPLATES: IDiagramTemplate[] = [
  {
    id: 'fire-drill',
    name: 'Fire Drill Floor Plan',
    description: 'Floor plan with fire exits, extinguishers, and evacuation routes',
    diagramType: DiagramType.FireDrill,
    suggestedMarkerTypes: [
      MarkerType.Exit,
      MarkerType.FireExtinguisher,
      MarkerType.AssemblyPoint,
      MarkerType.Stairs,
      MarkerType.YouAreHere
    ]
  },
  {
    id: 'emergency-exit',
    name: 'Emergency Exit Map',
    description: 'Emergency exit locations and evacuation routes',
    diagramType: DiagramType.EmergencyExit,
    suggestedMarkerTypes: [
      MarkerType.Exit,
      MarkerType.AssemblyPoint,
      MarkerType.Stairs,
      MarkerType.Danger,
      MarkerType.YouAreHere
    ]
  },
  {
    id: 'first-aid',
    name: 'First Aid & Safety Map',
    description: 'First aid kits, AEDs, and emergency equipment locations',
    diagramType: DiagramType.SafetyZone,
    suggestedMarkerTypes: [
      MarkerType.FirstAid,
      MarkerType.Defibrillator,
      MarkerType.EmergencyPhone,
      MarkerType.Information
    ]
  },
  {
    id: 'floor-plan',
    name: 'General Floor Plan',
    description: 'General purpose floor plan with common markers',
    diagramType: DiagramType.FloorPlan,
    suggestedMarkerTypes: [
      MarkerType.Entrance,
      MarkerType.Exit,
      MarkerType.Stairs,
      MarkerType.Elevator,
      MarkerType.RestRoom,
      MarkerType.Information
    ]
  },
  {
    id: 'access-map',
    name: 'Building Access Map',
    description: 'Entry points and access control locations',
    diagramType: DiagramType.AccessMap,
    suggestedMarkerTypes: [
      MarkerType.Entrance,
      MarkerType.Exit,
      MarkerType.Information,
      MarkerType.YouAreHere
    ]
  }
];
