// @ts-nocheck
// SignalR Service for Real-time Communication
// Note: In a production environment, you would need a SignalR hub hosted on Azure or similar
// For SPFx, we'll use SharePoint's change notification webhooks as an alternative

import { ILiveNotification, IUserPresence, PresenceStatus } from '../models';
import { logger } from './LoggingService';

type NotificationListener = (notification: ILiveNotification) => void;
type PresenceListener = (presences: IUserPresence[]) => void;
type ConnectionStateListener = (connected: boolean) => void;

/**
 * SignalR-like service for real-time updates
 * In SPFx, this simulates real-time updates using polling and local storage events
 * For true real-time, you would need an Azure Function with SignalR Service
 */
class SignalRServiceClass {
  private notificationListeners: NotificationListener[] = [];
  private presenceListeners: PresenceListener[] = [];
  private connectionStateListeners: ConnectionStateListener[] = [];

  private isConnected = false;
  private pollingInterval: any;
  private presenceInterval: any;
  private currentUserId: number = 0;
  private currentUserName: string = '';
  private currentUserEmail: string = '';
  private lastNotificationCheck: Date = new Date();

  // Local storage keys for cross-tab communication
  private readonly PRESENCE_KEY = 'PM_user_presence';
  private readonly NOTIFICATION_KEY = 'PM_notifications';
  private readonly ACTIVITY_KEY = 'PM_activity';

  /**
   * Initialize and connect to real-time service
   */
  public async connect(userId: number, userName: string, userEmail: string): Promise<void> {
    if (this.isConnected) {
      logger.debug('SignalRService', 'SignalR: Already connected');
      return;
    }

    this.currentUserId = userId;
    this.currentUserName = userName;
    this.currentUserEmail = userEmail;

    try {
      // Set up local storage event listener for cross-tab communication
      window.addEventListener('storage', this.handleStorageChange);

      // Update own presence
      this.updatePresence(PresenceStatus.Online);

      // Start polling for notifications (every 30 seconds)
      this.pollingInterval = setInterval(() => {
        this.checkForNotifications();
      }, 30000);

      // Update presence every 2 minutes
      this.presenceInterval = setInterval(() => {
        this.updatePresence(PresenceStatus.Online);
        this.broadcastPresence();
      }, 120000);

      // Handle page visibility changes
      document.addEventListener('visibilitychange', this.handleVisibilityChange);

      // Handle beforeunload to set offline
      window.addEventListener('beforeunload', this.handleBeforeUnload);

      this.isConnected = true;
      this.notifyConnectionState(true);

      logger.debug('SignalRService', 'SignalR: Connected successfully');
    } catch (error) {
      logger.error('SignalRService', 'SignalR: Connection failed', error);
      this.notifyConnectionState(false);
      throw error;
    }
  }

  /**
   * Disconnect from real-time service
   */
  public disconnect(): void {
    if (!this.isConnected) {
      return;
    }

    // Clear intervals
    if (this.pollingInterval) {
      clearInterval(this.pollingInterval);
    }
    if (this.presenceInterval) {
      clearInterval(this.presenceInterval);
    }

    // Remove event listeners
    window.removeEventListener('storage', this.handleStorageChange);
    document.removeEventListener('visibilitychange', this.handleVisibilityChange);
    window.removeEventListener('beforeunload', this.handleBeforeUnload);

    // Set status to offline
    this.updatePresence(PresenceStatus.Offline);

    this.isConnected = false;
    this.notifyConnectionState(false);

    logger.debug('SignalRService', 'SignalR: Disconnected');
  }

  /**
   * Subscribe to notification updates
   */
  public onNotification(listener: NotificationListener): () => void {
    this.notificationListeners.push(listener);
    return () => {
      const index = this.notificationListeners.indexOf(listener);
      if (index > -1) {
        this.notificationListeners.splice(index, 1);
      }
    };
  }

  /**
   * Subscribe to presence updates
   */
  public onPresence(listener: PresenceListener): () => void {
    this.presenceListeners.push(listener);
    return () => {
      const index = this.presenceListeners.indexOf(listener);
      if (index > -1) {
        this.presenceListeners.splice(index, 1);
      }
    };
  }

  /**
   * Subscribe to connection state changes
   */
  public onConnectionState(listener: ConnectionStateListener): () => void {
    this.connectionStateListeners.push(listener);
    return () => {
      const index = this.connectionStateListeners.indexOf(listener);
      if (index > -1) {
        this.connectionStateListeners.splice(index, 1);
      }
    };
  }

  /**
   * Send a notification to the system
   */
  public async sendNotification(notification: Omit<ILiveNotification, 'id' | 'timestamp' | 'read'>): Promise<void> {
    const fullNotification: ILiveNotification = {
      ...notification,
      id: `notif_${Date.now()}_${Math.random()}`,
      timestamp: new Date(),
      read: false
    };

    // Store in local storage for cross-tab communication
    try {
      const existing = this.getStoredNotifications();
      existing.push(fullNotification);

      // Keep only last 100 notifications
      const trimmed = existing.slice(-100);
      localStorage.setItem(this.NOTIFICATION_KEY, JSON.stringify(trimmed));

      // Trigger storage event manually for same tab
      this.notifyNotificationListeners(fullNotification);
    } catch (error) {
      logger.error('SignalRService', 'Failed to send notification:', error);
    }
  }

  /**
   * Update user's current location/activity
   */
  public updateActivity(location: string): void {
    const activity = {
      userId: this.currentUserId,
      location,
      timestamp: new Date().toISOString()
    };

    try {
      localStorage.setItem(this.ACTIVITY_KEY, JSON.stringify(activity));
    } catch (error) {
      logger.error('SignalRService', 'Failed to update activity:', error);
    }
  }

  /**
   * Get all active users' presence
   */
  public getPresence(): IUserPresence[] {
    try {
      const presenceData = localStorage.getItem(this.PRESENCE_KEY);
      if (!presenceData) {
        return [];
      }

      const allPresence: IUserPresence[] = JSON.parse(presenceData);
      const now = new Date();

      // Filter out stale presence (older than 5 minutes) - ES5 compatible
      const filtered = [];
      for (let i = 0; i < allPresence.length; i++) {
        const p = allPresence[i];
        const lastActivity = typeof p.lastActivity === 'string' ? new Date(p.lastActivity) : p.lastActivity;
        const diffMinutes = (now.getTime() - lastActivity.getTime()) / (1000 * 60);
        if (diffMinutes < 5) {
          filtered.push(p);
        }
      }
      return filtered;
    } catch (error) {
      logger.error('SignalRService', 'Failed to get presence:', error);
      return [];
    }
  }

  /**
   * Mark notification as read
   */
  public markNotificationRead(notificationId: string): void {
    try {
      const notifications = this.getStoredNotifications();
      const updated = notifications.map(n =>
        n.id === notificationId ? { ...n, read: true } : n
      );
      localStorage.setItem(this.NOTIFICATION_KEY, JSON.stringify(updated));
    } catch (error) {
      logger.error('SignalRService', 'Failed to mark notification as read:', error);
    }
  }

  /**
   * Get unread notification count
   */
  public getUnreadCount(): number {
    const notifications = this.getStoredNotifications();
    return notifications.filter(n => !n.read && n.userId === this.currentUserId).length;
  }

  /**
   * Get user's notifications
   */
  public getNotifications(limit: number = 50): ILiveNotification[] {
    const notifications = this.getStoredNotifications();
    return notifications
      .filter(n => n.userId === this.currentUserId)
      .sort((a, b) => {
        const aTime = typeof a.timestamp === 'string' ? new Date(a.timestamp).getTime() : a.timestamp.getTime();
        const bTime = typeof b.timestamp === 'string' ? new Date(b.timestamp).getTime() : b.timestamp.getTime();
        return bTime - aTime;
      })
      .slice(0, limit);
  }

  // Private helper methods

  private handleStorageChange = (event: StorageEvent): void => {
    if (event.key === this.NOTIFICATION_KEY && event.newValue) {
      try {
        const notifications: ILiveNotification[] = JSON.parse(event.newValue);
        // Notify about new notifications
        const newNotifications = notifications.filter(n => {
          if (n.userId !== this.currentUserId || n.read) {
            return false;
          }
          const timestamp = typeof n.timestamp === 'string' ? new Date(n.timestamp) : n.timestamp;
          return timestamp > this.lastNotificationCheck;
        });

        newNotifications.forEach(n => this.notifyNotificationListeners(n));
        this.lastNotificationCheck = new Date();
      } catch (error) {
        logger.error('SignalRService', 'Error processing storage change:', error);
      }
    } else if (event.key === this.PRESENCE_KEY && event.newValue) {
      try {
        const presence: IUserPresence[] = JSON.parse(event.newValue);
        this.notifyPresenceListeners(presence);
      } catch (error) {
        logger.error('SignalRService', 'Error processing presence change:', error);
      }
    }
  };

  private handleVisibilityChange = (): void => {
    if (document.hidden) {
      this.updatePresence(PresenceStatus.Away);
    } else {
      this.updatePresence(PresenceStatus.Online);
    }
  };

  private handleBeforeUnload = (): void => {
    this.updatePresence(PresenceStatus.Offline);
  };

  private updatePresence(status: PresenceStatus): void {
    try {
      const allPresence = this.getPresence();

      // Update or add current user's presence - ES5 compatible
      let existingIndex = -1;
      for (let i = 0; i < allPresence.length; i++) {
        if (allPresence[i].userId === this.currentUserId) {
          existingIndex = i;
          break;
        }
      }

      const userPresence: IUserPresence = {
        userId: this.currentUserId,
        userName: this.currentUserName,
        email: this.currentUserEmail,
        status,
        lastActivity: new Date()
      };

      if (existingIndex > -1) {
        allPresence[existingIndex] = userPresence;
      } else {
        allPresence.push(userPresence);
      }

      // Remove offline users - ES5 compatible
      const activePresence = [];
      for (let i = 0; i < allPresence.length; i++) {
        if (allPresence[i].status !== PresenceStatus.Offline) {
          activePresence.push(allPresence[i]);
        }
      }

      localStorage.setItem(this.PRESENCE_KEY, JSON.stringify(activePresence));
    } catch (error) {
      logger.error('SignalRService', 'Failed to update presence:', error);
    }
  }

  private broadcastPresence(): void {
    const presence = this.getPresence();
    this.notifyPresenceListeners(presence);
  }

  private async checkForNotifications(): Promise<void> {
    // In a real implementation, this would query SharePoint for new notifications
    // For now, we'll rely on local storage events
    const notifications = this.getStoredNotifications();
    const newNotifications = notifications.filter(n => {
      if (n.userId !== this.currentUserId || n.read) {
        return false;
      }
      const timestamp = typeof n.timestamp === 'string' ? new Date(n.timestamp) : n.timestamp;
      return timestamp > this.lastNotificationCheck;
    });

    newNotifications.forEach(n => this.notifyNotificationListeners(n));
    this.lastNotificationCheck = new Date();
  }

  private getStoredNotifications(): ILiveNotification[] {
    try {
      const data = localStorage.getItem(this.NOTIFICATION_KEY);
      if (!data) {
        return [];
      }
      return JSON.parse(data);
    } catch (error) {
      logger.error('SignalRService', 'Failed to get stored notifications:', error);
      return [];
    }
  }

  private notifyNotificationListeners(notification: ILiveNotification): void {
    this.notificationListeners.forEach(listener => {
      try {
        listener(notification);
      } catch (error) {
        logger.error('SignalRService', 'Error in notification listener:', error);
      }
    });
  }

  private notifyPresenceListeners(presence: IUserPresence[]): void {
    this.presenceListeners.forEach(listener => {
      try {
        listener(presence);
      } catch (error) {
        logger.error('SignalRService', 'Error in presence listener:', error);
      }
    });
  }

  private notifyConnectionState(connected: boolean): void {
    this.connectionStateListeners.forEach(listener => {
      try {
        listener(connected);
      } catch (error) {
        logger.error('SignalRService', 'Error in connection state listener:', error);
      }
    });
  }

  public isServiceConnected(): boolean {
    return this.isConnected;
  }
}

export const SignalRService = new SignalRServiceClass();
