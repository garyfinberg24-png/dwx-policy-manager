// @ts-nocheck
/**
 * Image Generation Service
 * Creates images for floorplans, emergency exit routes, and other visual content
 */

import { logger } from './LoggingService';

/**
 * Image generation result
 */
export interface IImageGenerationResult {
  success: boolean;
  blob?: Blob;
  dataUrl?: string;
  fileName?: string;
  error?: string;
}

/**
 * Image document options
 */
export interface IImageDocumentOptions {
  /** Image width in pixels */
  width?: number;
  /** Image height in pixels */
  height?: number;
  /** Output format */
  format?: 'png' | 'jpeg';
  /** JPEG quality (0-1) */
  quality?: number;
  /** Background color */
  backgroundColor?: string;
  /** Title text */
  title?: string;
  /** Footer text */
  footer?: string;
}

/**
 * Floorplan room configuration
 */
export interface IFloorplanRoom {
  /** Room ID */
  id: string;
  /** Room name/label */
  name: string;
  /** X position (percentage 0-100) */
  x: number;
  /** Y position (percentage 0-100) */
  y: number;
  /** Width (percentage 0-100) */
  width: number;
  /** Height (percentage 0-100) */
  height: number;
  /** Room type for styling */
  type?: 'office' | 'meeting' | 'common' | 'restroom' | 'stairs' | 'elevator' | 'entrance' | 'exit' | 'kitchen' | 'storage';
  /** Optional highlight color */
  highlightColor?: string;
  /** Is this an emergency exit */
  isEmergencyExit?: boolean;
}

/**
 * Emergency route configuration
 */
export interface IEmergencyRoute {
  /** Route ID */
  id: string;
  /** Route points as [x, y] percentages */
  points: [number, number][];
  /** Route color */
  color?: string;
  /** Route label */
  label?: string;
  /** Is primary route */
  isPrimary?: boolean;
}

/**
 * Floorplan configuration
 */
export interface IFloorplanConfig {
  /** Floor name/number */
  floorName: string;
  /** Building name */
  buildingName?: string;
  /** Rooms in the floorplan */
  rooms: IFloorplanRoom[];
  /** Emergency routes */
  emergencyRoutes?: IEmergencyRoute[];
  /** Assembly point location */
  assemblyPoint?: { x: number; y: number; label: string };
  /** Fire extinguisher locations */
  fireExtinguishers?: { x: number; y: number }[];
  /** First aid kit locations */
  firstAidKits?: { x: number; y: number }[];
  /** Highlight specific employee location */
  employeeLocation?: { roomId: string; label: string };
}

/**
 * Organizational chart node
 */
export interface IOrgChartNode {
  /** Node ID */
  id: string;
  /** Person name */
  name: string;
  /** Job title */
  title: string;
  /** Department */
  department?: string;
  /** Parent node ID (for hierarchy) */
  parentId?: string | null;
  /** Photo URL or initials */
  photoUrl?: string;
  /** Highlight this node */
  isHighlighted?: boolean;
}

/**
 * Image Generation Service
 */
export class ImageGenerationService {
  private readonly defaultWidth = 1200;
  private readonly defaultHeight = 800;
  private readonly colors = {
    primary: '#0078d4',
    secondary: '#323130',
    success: '#107c10',
    warning: '#ffb900',
    danger: '#d13438',
    background: '#ffffff',
    border: '#e1dfdd',
    text: '#323130',
    lightGray: '#f3f2f1'
  };

  private readonly roomColors: Record<string, string> = {
    office: '#e3f2fd',
    meeting: '#e8f5e9',
    common: '#fff3e0',
    restroom: '#f3e5f5',
    stairs: '#fce4ec',
    elevator: '#e0f2f1',
    entrance: '#e8eaf6',
    exit: '#ffebee',
    kitchen: '#fff8e1',
    storage: '#eceff1'
  };

  /**
   * Generate a floorplan image
   */
  public async generateFloorplan(
    config: IFloorplanConfig,
    options?: IImageDocumentOptions
  ): Promise<IImageGenerationResult> {
    try {
      const width = options?.width || this.defaultWidth;
      const height = options?.height || this.defaultHeight;
      const canvas = this.createCanvas(width, height);
      const ctx = canvas.getContext('2d');

      if (!ctx) {
        throw new Error('Failed to get canvas context');
      }

      // Draw background
      ctx.fillStyle = options?.backgroundColor || this.colors.background;
      ctx.fillRect(0, 0, width, height);

      // Calculate drawing area (leave space for title and footer)
      const padding = 40;
      const titleHeight = options?.title ? 60 : 0;
      const footerHeight = options?.footer ? 40 : 0;
      const drawArea = {
        x: padding,
        y: padding + titleHeight,
        width: width - padding * 2,
        height: height - padding * 2 - titleHeight - footerHeight
      };

      // Draw title
      if (options?.title) {
        ctx.fillStyle = this.colors.primary;
        ctx.font = 'bold 24px Segoe UI, Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(options.title, width / 2, padding + 30);
      }

      // Draw floor label
      ctx.fillStyle = this.colors.secondary;
      ctx.font = '16px Segoe UI, Arial, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText(
        `${config.buildingName || 'Building'} - ${config.floorName}`,
        drawArea.x,
        drawArea.y - 10
      );

      // Draw floor outline
      ctx.strokeStyle = this.colors.border;
      ctx.lineWidth = 2;
      ctx.strokeRect(drawArea.x, drawArea.y, drawArea.width, drawArea.height);

      // Draw rooms
      for (const room of config.rooms) {
        this.drawRoom(ctx, room, drawArea);
      }

      // Draw emergency routes
      if (config.emergencyRoutes) {
        for (const route of config.emergencyRoutes) {
          this.drawEmergencyRoute(ctx, route, drawArea);
        }
      }

      // Draw fire extinguishers
      if (config.fireExtinguishers) {
        for (const fe of config.fireExtinguishers) {
          this.drawFireExtinguisher(ctx, fe, drawArea);
        }
      }

      // Draw first aid kits
      if (config.firstAidKits) {
        for (const fa of config.firstAidKits) {
          this.drawFirstAidKit(ctx, fa, drawArea);
        }
      }

      // Draw assembly point
      if (config.assemblyPoint) {
        this.drawAssemblyPoint(ctx, config.assemblyPoint, drawArea);
      }

      // Highlight employee location
      if (config.employeeLocation) {
        const room = config.rooms.find(r => r.id === config.employeeLocation?.roomId);
        if (room) {
          this.highlightRoom(ctx, room, config.employeeLocation.label, drawArea);
        }
      }

      // Draw legend
      this.drawLegend(ctx, config, { x: drawArea.x + drawArea.width - 200, y: drawArea.y + 10 });

      // Draw footer
      if (options?.footer) {
        ctx.fillStyle = this.colors.secondary;
        ctx.font = '12px Segoe UI, Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(options.footer, width / 2, height - padding + 10);
      }

      return this.canvasToResult(canvas, config.floorName, options);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ImageGenerationService', 'Failed to generate floorplan:', error);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Generate an emergency exit route image
   */
  public async generateEmergencyExitMap(
    config: IFloorplanConfig,
    options?: IImageDocumentOptions
  ): Promise<IImageGenerationResult> {
    // Use floorplan generation with emergency-focused styling
    const emergencyOptions: IImageDocumentOptions = {
      ...options,
      title: options?.title || `Emergency Exit Routes - ${config.floorName}`,
      footer: options?.footer || 'In case of emergency, proceed calmly to the nearest exit. Assembly point marked with green star.',
      backgroundColor: '#fff5f5'
    };

    return this.generateFloorplan(config, emergencyOptions);
  }

  /**
   * Generate an organizational chart image
   */
  public async generateOrgChart(
    nodes: IOrgChartNode[],
    options?: IImageDocumentOptions
  ): Promise<IImageGenerationResult> {
    try {
      const width = options?.width || this.defaultWidth;
      const height = options?.height || this.defaultHeight;
      const canvas = this.createCanvas(width, height);
      const ctx = canvas.getContext('2d');

      if (!ctx) {
        throw new Error('Failed to get canvas context');
      }

      // Draw background
      ctx.fillStyle = options?.backgroundColor || this.colors.background;
      ctx.fillRect(0, 0, width, height);

      // Draw title
      if (options?.title) {
        ctx.fillStyle = this.colors.primary;
        ctx.font = 'bold 24px Segoe UI, Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(options.title, width / 2, 40);
      }

      // Build tree structure
      const rootNodes = nodes.filter(n => !n.parentId);
      const nodeWidth = 160;
      const nodeHeight = 80;
      const verticalGap = 60;
      const horizontalGap = 20;

      // Calculate levels
      const levels = this.calculateOrgLevels(nodes, rootNodes);
      const startY = (options?.title ? 80 : 40);

      // Draw connections first (behind nodes)
      this.drawOrgConnections(ctx, nodes, levels, width, startY, nodeWidth, nodeHeight, verticalGap);

      // Draw nodes
      let levelY = startY;
      for (let i = 0; i < levels.length; i++) {
        const level = levels[i];
        const totalWidth = level.length * nodeWidth + (level.length - 1) * horizontalGap;
        let startX = (width - totalWidth) / 2;

        for (const node of level) {
          this.drawOrgNode(ctx, node, startX, levelY, nodeWidth, nodeHeight);
          startX += nodeWidth + horizontalGap;
        }

        levelY += nodeHeight + verticalGap;
      }

      // Draw footer
      if (options?.footer) {
        ctx.fillStyle = this.colors.secondary;
        ctx.font = '12px Segoe UI, Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(options.footer, width / 2, height - 20);
      }

      const fileName = options?.title?.replace(/\s+/g, '_') || 'OrgChart';
      return this.canvasToResult(canvas, fileName, options);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ImageGenerationService', 'Failed to generate org chart:', error);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Generate a simple location map
   */
  public async generateLocationMap(
    locations: Array<{ name: string; x: number; y: number; type?: string }>,
    options?: IImageDocumentOptions & { mapTitle?: string }
  ): Promise<IImageGenerationResult> {
    try {
      const width = options?.width || 800;
      const height = options?.height || 600;
      const canvas = this.createCanvas(width, height);
      const ctx = canvas.getContext('2d');

      if (!ctx) {
        throw new Error('Failed to get canvas context');
      }

      // Draw background
      ctx.fillStyle = options?.backgroundColor || this.colors.lightGray;
      ctx.fillRect(0, 0, width, height);

      // Draw title
      const title = options?.title || options?.mapTitle || 'Location Map';
      ctx.fillStyle = this.colors.primary;
      ctx.font = 'bold 20px Segoe UI, Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(title, width / 2, 35);

      // Calculate drawing area
      const padding = 50;
      const drawArea = {
        x: padding,
        y: padding + 20,
        width: width - padding * 2,
        height: height - padding * 2 - 40
      };

      // Draw border
      ctx.strokeStyle = this.colors.border;
      ctx.lineWidth = 2;
      ctx.strokeRect(drawArea.x, drawArea.y, drawArea.width, drawArea.height);

      // Draw grid
      ctx.strokeStyle = '#e0e0e0';
      ctx.lineWidth = 0.5;
      const gridSize = 50;
      for (let x = drawArea.x + gridSize; x < drawArea.x + drawArea.width; x += gridSize) {
        ctx.beginPath();
        ctx.moveTo(x, drawArea.y);
        ctx.lineTo(x, drawArea.y + drawArea.height);
        ctx.stroke();
      }
      for (let y = drawArea.y + gridSize; y < drawArea.y + drawArea.height; y += gridSize) {
        ctx.beginPath();
        ctx.moveTo(drawArea.x, y);
        ctx.lineTo(drawArea.x + drawArea.width, y);
        ctx.stroke();
      }

      // Draw locations
      for (const loc of locations) {
        const x = drawArea.x + (loc.x / 100) * drawArea.width;
        const y = drawArea.y + (loc.y / 100) * drawArea.height;

        // Draw marker
        ctx.fillStyle = loc.type === 'exit' ? this.colors.danger : this.colors.primary;
        ctx.beginPath();
        ctx.arc(x, y, 12, 0, Math.PI * 2);
        ctx.fill();

        // Draw label
        ctx.fillStyle = this.colors.text;
        ctx.font = '12px Segoe UI, Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(loc.name, x, y + 28);
      }

      // Draw footer
      if (options?.footer) {
        ctx.fillStyle = this.colors.secondary;
        ctx.font = '11px Segoe UI, Arial, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(options.footer, width / 2, height - 15);
      }

      const fileName = title.replace(/\s+/g, '_');
      return this.canvasToResult(canvas, fileName, options);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ImageGenerationService', 'Failed to generate location map:', error);
      return { success: false, error: errorMessage };
    }
  }

  /**
   * Generate a badge/ID card image
   */
  public async generateBadgeImage(
    employee: {
      name: string;
      title: string;
      department: string;
      employeeId: string;
      photoDataUrl?: string;
    },
    options?: IImageDocumentOptions & { companyName?: string; companyLogo?: string }
  ): Promise<IImageGenerationResult> {
    try {
      const width = options?.width || 350;
      const height = options?.height || 220;
      const canvas = this.createCanvas(width, height);
      const ctx = canvas.getContext('2d');

      if (!ctx) {
        throw new Error('Failed to get canvas context');
      }

      // Draw card background with rounded corners
      ctx.fillStyle = this.colors.background;
      this.roundRect(ctx, 0, 0, width, height, 12);
      ctx.fill();

      // Draw header band
      ctx.fillStyle = this.colors.primary;
      this.roundRect(ctx, 0, 0, width, 50, 12, true, false);
      ctx.fill();

      // Company name in header
      ctx.fillStyle = '#ffffff';
      ctx.font = 'bold 16px Segoe UI, Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(options?.companyName || 'Company Name', width / 2, 32);

      // Photo area
      const photoSize = 80;
      const photoX = 20;
      const photoY = 65;

      ctx.fillStyle = this.colors.lightGray;
      ctx.strokeStyle = this.colors.border;
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.arc(photoX + photoSize / 2, photoY + photoSize / 2, photoSize / 2, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();

      // Initials in photo area (as placeholder)
      const initials = employee.name.split(' ').map(n => n[0]).join('').substring(0, 2);
      ctx.fillStyle = this.colors.secondary;
      ctx.font = 'bold 28px Segoe UI, Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(initials, photoX + photoSize / 2, photoY + photoSize / 2 + 10);

      // Employee info
      const infoX = photoX + photoSize + 20;
      ctx.fillStyle = this.colors.text;
      ctx.textAlign = 'left';

      ctx.font = 'bold 16px Segoe UI, Arial, sans-serif';
      ctx.fillText(employee.name, infoX, 90);

      ctx.font = '13px Segoe UI, Arial, sans-serif';
      ctx.fillText(employee.title, infoX, 110);

      ctx.fillStyle = this.colors.secondary;
      ctx.font = '12px Segoe UI, Arial, sans-serif';
      ctx.fillText(employee.department, infoX, 130);

      // Employee ID at bottom
      ctx.fillStyle = this.colors.lightGray;
      ctx.fillRect(0, height - 45, width, 45);

      ctx.fillStyle = this.colors.secondary;
      ctx.font = '11px Segoe UI, Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText('EMPLOYEE ID', width / 2, height - 28);

      ctx.fillStyle = this.colors.text;
      ctx.font = 'bold 14px Segoe UI, Arial, sans-serif';
      ctx.fillText(employee.employeeId, width / 2, height - 10);

      // Draw border
      ctx.strokeStyle = this.colors.border;
      ctx.lineWidth = 1;
      this.roundRect(ctx, 0, 0, width, height, 12);
      ctx.stroke();

      const fileName = `Badge_${employee.name.replace(/\s+/g, '_')}`;
      return this.canvasToResult(canvas, fileName, options);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ImageGenerationService', 'Failed to generate badge:', error);
      return { success: false, error: errorMessage };
    }
  }

  // Private helper methods

  private createCanvas(width: number, height: number): HTMLCanvasElement {
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    return canvas;
  }

  private async canvasToResult(
    canvas: HTMLCanvasElement,
    baseName: string,
    options?: IImageDocumentOptions
  ): Promise<IImageGenerationResult> {
    const format = options?.format || 'png';
    const quality = options?.quality || 0.92;
    const mimeType = format === 'jpeg' ? 'image/jpeg' : 'image/png';

    return new Promise((resolve) => {
      canvas.toBlob(
        (blob) => {
          if (!blob) {
            resolve({ success: false, error: 'Failed to create image blob' });
            return;
          }

          const dataUrl = canvas.toDataURL(mimeType, quality);
          const timestamp = new Date().toISOString().slice(0, 10);
          const fileName = `${baseName}_${timestamp}.${format}`;

          logger.info('ImageGenerationService', `Image generated: ${fileName}`);

          resolve({
            success: true,
            blob,
            dataUrl,
            fileName
          });
        },
        mimeType,
        quality
      );
    });
  }

  private drawRoom(
    ctx: CanvasRenderingContext2D,
    room: IFloorplanRoom,
    drawArea: { x: number; y: number; width: number; height: number }
  ): void {
    const x = drawArea.x + (room.x / 100) * drawArea.width;
    const y = drawArea.y + (room.y / 100) * drawArea.height;
    const w = (room.width / 100) * drawArea.width;
    const h = (room.height / 100) * drawArea.height;

    // Fill
    ctx.fillStyle = room.highlightColor || this.roomColors[room.type || 'office'] || this.colors.lightGray;
    ctx.fillRect(x, y, w, h);

    // Border
    ctx.strokeStyle = this.colors.border;
    ctx.lineWidth = 1;
    ctx.strokeRect(x, y, w, h);

    // Emergency exit marker
    if (room.isEmergencyExit) {
      ctx.fillStyle = this.colors.danger;
      ctx.font = 'bold 12px Segoe UI, Arial, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText('EXIT', x + w / 2, y + 15);
    }

    // Label
    ctx.fillStyle = this.colors.text;
    ctx.font = '11px Segoe UI, Arial, sans-serif';
    ctx.textAlign = 'center';
    const labelY = room.isEmergencyExit ? y + h / 2 + 5 : y + h / 2 + 4;
    ctx.fillText(room.name, x + w / 2, labelY);
  }

  private drawEmergencyRoute(
    ctx: CanvasRenderingContext2D,
    route: IEmergencyRoute,
    drawArea: { x: number; y: number; width: number; height: number }
  ): void {
    if (route.points.length < 2) return;

    ctx.strokeStyle = route.color || (route.isPrimary ? this.colors.danger : this.colors.warning);
    ctx.lineWidth = route.isPrimary ? 4 : 2;
    ctx.setLineDash(route.isPrimary ? [] : [10, 5]);

    ctx.beginPath();
    const startX = drawArea.x + (route.points[0][0] / 100) * drawArea.width;
    const startY = drawArea.y + (route.points[0][1] / 100) * drawArea.height;
    ctx.moveTo(startX, startY);

    for (let i = 1; i < route.points.length; i++) {
      const x = drawArea.x + (route.points[i][0] / 100) * drawArea.width;
      const y = drawArea.y + (route.points[i][1] / 100) * drawArea.height;
      ctx.lineTo(x, y);
    }

    ctx.stroke();
    ctx.setLineDash([]);

    // Draw arrow at the end
    const lastIdx = route.points.length - 1;
    const endX = drawArea.x + (route.points[lastIdx][0] / 100) * drawArea.width;
    const endY = drawArea.y + (route.points[lastIdx][1] / 100) * drawArea.height;
    const prevX = drawArea.x + (route.points[lastIdx - 1][0] / 100) * drawArea.width;
    const prevY = drawArea.y + (route.points[lastIdx - 1][1] / 100) * drawArea.height;

    const angle = Math.atan2(endY - prevY, endX - prevX);
    const arrowSize = 10;

    ctx.fillStyle = route.color || (route.isPrimary ? this.colors.danger : this.colors.warning);
    ctx.beginPath();
    ctx.moveTo(endX, endY);
    ctx.lineTo(endX - arrowSize * Math.cos(angle - Math.PI / 6), endY - arrowSize * Math.sin(angle - Math.PI / 6));
    ctx.lineTo(endX - arrowSize * Math.cos(angle + Math.PI / 6), endY - arrowSize * Math.sin(angle + Math.PI / 6));
    ctx.closePath();
    ctx.fill();
  }

  private drawFireExtinguisher(
    ctx: CanvasRenderingContext2D,
    pos: { x: number; y: number },
    drawArea: { x: number; y: number; width: number; height: number }
  ): void {
    const x = drawArea.x + (pos.x / 100) * drawArea.width;
    const y = drawArea.y + (pos.y / 100) * drawArea.height;

    ctx.fillStyle = this.colors.danger;
    ctx.beginPath();
    ctx.arc(x, y, 8, 0, Math.PI * 2);
    ctx.fill();

    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 10px Segoe UI, Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('FE', x, y + 3);
  }

  private drawFirstAidKit(
    ctx: CanvasRenderingContext2D,
    pos: { x: number; y: number },
    drawArea: { x: number; y: number; width: number; height: number }
  ): void {
    const x = drawArea.x + (pos.x / 100) * drawArea.width;
    const y = drawArea.y + (pos.y / 100) * drawArea.height;

    ctx.fillStyle = this.colors.success;
    ctx.fillRect(x - 8, y - 8, 16, 16);

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(x - 5, y - 2, 10, 4);
    ctx.fillRect(x - 2, y - 5, 4, 10);
  }

  private drawAssemblyPoint(
    ctx: CanvasRenderingContext2D,
    point: { x: number; y: number; label: string },
    drawArea: { x: number; y: number; width: number; height: number }
  ): void {
    const x = drawArea.x + (point.x / 100) * drawArea.width;
    const y = drawArea.y + (point.y / 100) * drawArea.height;

    // Draw star
    ctx.fillStyle = this.colors.success;
    this.drawStar(ctx, x, y, 5, 15, 7);

    // Draw label
    ctx.fillStyle = this.colors.text;
    ctx.font = 'bold 11px Segoe UI, Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(point.label, x, y + 25);
  }

  private drawStar(
    ctx: CanvasRenderingContext2D,
    cx: number,
    cy: number,
    spikes: number,
    outerRadius: number,
    innerRadius: number
  ): void {
    let rot = (Math.PI / 2) * 3;
    const step = Math.PI / spikes;

    ctx.beginPath();
    ctx.moveTo(cx, cy - outerRadius);

    for (let i = 0; i < spikes; i++) {
      let x = cx + Math.cos(rot) * outerRadius;
      let y = cy + Math.sin(rot) * outerRadius;
      ctx.lineTo(x, y);
      rot += step;

      x = cx + Math.cos(rot) * innerRadius;
      y = cy + Math.sin(rot) * innerRadius;
      ctx.lineTo(x, y);
      rot += step;
    }

    ctx.lineTo(cx, cy - outerRadius);
    ctx.closePath();
    ctx.fill();
  }

  private highlightRoom(
    ctx: CanvasRenderingContext2D,
    room: IFloorplanRoom,
    label: string,
    drawArea: { x: number; y: number; width: number; height: number }
  ): void {
    const x = drawArea.x + (room.x / 100) * drawArea.width;
    const y = drawArea.y + (room.y / 100) * drawArea.height;
    const w = (room.width / 100) * drawArea.width;
    const h = (room.height / 100) * drawArea.height;

    // Highlight border
    ctx.strokeStyle = this.colors.primary;
    ctx.lineWidth = 3;
    ctx.setLineDash([5, 3]);
    ctx.strokeRect(x - 2, y - 2, w + 4, h + 4);
    ctx.setLineDash([]);

    // Label
    ctx.fillStyle = this.colors.primary;
    ctx.font = 'bold 12px Segoe UI, Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(`📍 ${label}`, x + w / 2, y - 8);
  }

  private drawLegend(
    ctx: CanvasRenderingContext2D,
    config: IFloorplanConfig,
    pos: { x: number; y: number }
  ): void {
    const items: Array<{ color: string; label: string }> = [];

    if (config.emergencyRoutes?.some(r => r.isPrimary)) {
      items.push({ color: this.colors.danger, label: 'Primary Exit Route' });
    }
    if (config.emergencyRoutes?.some(r => !r.isPrimary)) {
      items.push({ color: this.colors.warning, label: 'Alt Exit Route' });
    }
    if (config.fireExtinguishers?.length) {
      items.push({ color: this.colors.danger, label: 'Fire Extinguisher' });
    }
    if (config.firstAidKits?.length) {
      items.push({ color: this.colors.success, label: 'First Aid Kit' });
    }
    if (config.assemblyPoint) {
      items.push({ color: this.colors.success, label: 'Assembly Point' });
    }

    if (items.length === 0) return;

    // Draw legend box
    const boxWidth = 150;
    const boxHeight = items.length * 20 + 30;

    ctx.fillStyle = 'rgba(255, 255, 255, 0.9)';
    ctx.strokeStyle = this.colors.border;
    ctx.lineWidth = 1;
    ctx.fillRect(pos.x, pos.y, boxWidth, boxHeight);
    ctx.strokeRect(pos.x, pos.y, boxWidth, boxHeight);

    // Title
    ctx.fillStyle = this.colors.text;
    ctx.font = 'bold 11px Segoe UI, Arial, sans-serif';
    ctx.textAlign = 'left';
    ctx.fillText('Legend', pos.x + 10, pos.y + 18);

    // Items
    ctx.font = '10px Segoe UI, Arial, sans-serif';
    let y = pos.y + 38;

    for (const item of items) {
      ctx.fillStyle = item.color;
      ctx.beginPath();
      ctx.arc(pos.x + 15, y - 4, 5, 0, Math.PI * 2);
      ctx.fill();

      ctx.fillStyle = this.colors.text;
      ctx.fillText(item.label, pos.x + 28, y);
      y += 20;
    }
  }

  private calculateOrgLevels(nodes: IOrgChartNode[], rootNodes: IOrgChartNode[]): IOrgChartNode[][] {
    const levels: IOrgChartNode[][] = [];

    const processLevel = (currentNodes: IOrgChartNode[]): void => {
      if (currentNodes.length === 0) return;

      levels.push(currentNodes);

      const childNodes: IOrgChartNode[] = [];
      for (const node of currentNodes) {
        const children = nodes.filter(n => n.parentId === node.id);
        childNodes.push(...children);
      }

      processLevel(childNodes);
    };

    processLevel(rootNodes);
    return levels;
  }

  private drawOrgConnections(
    ctx: CanvasRenderingContext2D,
    nodes: IOrgChartNode[],
    levels: IOrgChartNode[][],
    canvasWidth: number,
    startY: number,
    nodeWidth: number,
    nodeHeight: number,
    verticalGap: number
  ): void {
    ctx.strokeStyle = this.colors.border;
    ctx.lineWidth = 2;

    const horizontalGap = 20;

    for (let i = 0; i < levels.length - 1; i++) {
      const currentLevel = levels[i];
      const nextLevel = levels[i + 1];

      const currentTotalWidth = currentLevel.length * nodeWidth + (currentLevel.length - 1) * horizontalGap;
      const currentStartX = (canvasWidth - currentTotalWidth) / 2;

      const nextTotalWidth = nextLevel.length * nodeWidth + (nextLevel.length - 1) * horizontalGap;
      const nextStartX = (canvasWidth - nextTotalWidth) / 2;

      const currentY = startY + i * (nodeHeight + verticalGap) + nodeHeight;
      const nextY = startY + (i + 1) * (nodeHeight + verticalGap);

      for (let j = 0; j < currentLevel.length; j++) {
        const parentNode = currentLevel[j];
        const parentX = currentStartX + j * (nodeWidth + horizontalGap) + nodeWidth / 2;

        const children = nextLevel.filter(n => n.parentId === parentNode.id);

        for (const child of children) {
          const childIndex = nextLevel.indexOf(child);
          const childX = nextStartX + childIndex * (nodeWidth + horizontalGap) + nodeWidth / 2;

          ctx.beginPath();
          ctx.moveTo(parentX, currentY);
          ctx.lineTo(parentX, currentY + verticalGap / 2);
          ctx.lineTo(childX, currentY + verticalGap / 2);
          ctx.lineTo(childX, nextY);
          ctx.stroke();
        }
      }
    }
  }

  private drawOrgNode(
    ctx: CanvasRenderingContext2D,
    node: IOrgChartNode,
    x: number,
    y: number,
    width: number,
    height: number
  ): void {
    // Node background
    ctx.fillStyle = node.isHighlighted ? '#e3f2fd' : this.colors.background;
    ctx.strokeStyle = node.isHighlighted ? this.colors.primary : this.colors.border;
    ctx.lineWidth = node.isHighlighted ? 2 : 1;

    this.roundRect(ctx, x, y, width, height, 8);
    ctx.fill();
    ctx.stroke();

    // Name
    ctx.fillStyle = this.colors.text;
    ctx.font = 'bold 12px Segoe UI, Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(this.truncateText(ctx, node.name, width - 20), x + width / 2, y + 28);

    // Title
    ctx.fillStyle = this.colors.secondary;
    ctx.font = '10px Segoe UI, Arial, sans-serif';
    ctx.fillText(this.truncateText(ctx, node.title, width - 20), x + width / 2, y + 45);

    // Department
    if (node.department) {
      ctx.fillStyle = this.colors.primary;
      ctx.font = '9px Segoe UI, Arial, sans-serif';
      ctx.fillText(this.truncateText(ctx, node.department, width - 20), x + width / 2, y + 62);
    }
  }

  private roundRect(
    ctx: CanvasRenderingContext2D,
    x: number,
    y: number,
    width: number,
    height: number,
    radius: number,
    topOnly: boolean = false,
    bottomOnly: boolean = false
  ): void {
    ctx.beginPath();

    if (topOnly) {
      ctx.moveTo(x + radius, y);
      ctx.lineTo(x + width - radius, y);
      ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
      ctx.lineTo(x + width, y + height);
      ctx.lineTo(x, y + height);
      ctx.lineTo(x, y + radius);
      ctx.quadraticCurveTo(x, y, x + radius, y);
    } else if (bottomOnly) {
      ctx.moveTo(x, y);
      ctx.lineTo(x + width, y);
      ctx.lineTo(x + width, y + height - radius);
      ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
      ctx.lineTo(x + radius, y + height);
      ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
      ctx.lineTo(x, y);
    } else {
      ctx.moveTo(x + radius, y);
      ctx.lineTo(x + width - radius, y);
      ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
      ctx.lineTo(x + width, y + height - radius);
      ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
      ctx.lineTo(x + radius, y + height);
      ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
      ctx.lineTo(x, y + radius);
      ctx.quadraticCurveTo(x, y, x + radius, y);
    }

    ctx.closePath();
  }

  private truncateText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string {
    const metrics = ctx.measureText(text);
    if (metrics.width <= maxWidth) {
      return text;
    }

    let truncated = text;
    while (ctx.measureText(truncated + '...').width > maxWidth && truncated.length > 0) {
      truncated = truncated.slice(0, -1);
    }

    return truncated + '...';
  }

  /**
   * Download an image in the browser
   */
  public downloadImage(result: IImageGenerationResult): void {
    if (!result.success || !result.blob) {
      logger.error('ImageGenerationService', 'Cannot download: Image generation failed');
      return;
    }

    const url = URL.createObjectURL(result.blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = result.fileName || 'image.png';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    logger.info('ImageGenerationService', `Image downloaded: ${result.fileName}`);
  }
}

export const imageGenerationService = new ImageGenerationService();
