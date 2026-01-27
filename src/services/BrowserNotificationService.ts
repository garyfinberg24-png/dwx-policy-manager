// @ts-nocheck
// Browser Notification Service
// Handles native browser notifications with permission management

import { ILiveNotification } from '../models';
import { logger } from './LoggingService';

class BrowserNotificationServiceClass {
  private permission: NotificationPermission = 'default';
  private enabled = false;
  private soundEnabled = true;

  constructor() {
    if ('Notification' in window) {
      this.permission = Notification.permission;
    }
  }

  /**
   * Request permission to show browser notifications
   */
  public async requestPermission(): Promise<boolean> {
    if (!('Notification' in window)) {
      logger.warn('BrowserNotificationService', 'Browser notifications not supported');
      return false;
    }

    if (this.permission === 'granted') {
      this.enabled = true;
      return true;
    }

    if (this.permission === 'denied') {
      logger.warn('BrowserNotificationService', 'Browser notifications denied by user');
      return false;
    }

    try {
      const result = await Notification.requestPermission();
      this.permission = result;
      this.enabled = result === 'granted';
      return this.enabled;
    } catch (error) {
      logger.error('BrowserNotificationService', 'Error requesting notification permission:', error);
      return false;
    }
  }

  /**
   * Show a browser notification
   */
  public async show(notification: ILiveNotification): Promise<void> {
    if (!this.enabled || this.permission !== 'granted') {
      logger.debug('BrowserNotificationService', 'Browser notifications not enabled');
      return;
    }

    try {
      const options: NotificationOptions = {
        body: notification.message,
        icon: this.getIconForNotification(notification),
        badge: '/assets/badge-icon.png',
        tag: notification.id,
        requireInteraction: notification.priority === 'critical',
        timestamp: notification.timestamp.getTime(),
        data: {
          processId: notification.processId,
          taskId: notification.taskId,
          actionUrl: notification.actionUrl
        }
      };

      const browserNotification = new Notification(notification.title, options);

      // Handle notification click
      browserNotification.onclick = () => {
        window.focus();
        if (notification.actionUrl) {
          window.location.href = notification.actionUrl;
        }
        browserNotification.close();
      };

      // Auto-close after duration based on priority
      const duration = this.getDurationForPriority(notification.priority);
      setTimeout(() => {
        browserNotification.close();
      }, duration);

      // Play sound if enabled
      if (this.soundEnabled) {
        this.playNotificationSound(notification.priority);
      }
    } catch (error) {
      logger.error('BrowserNotificationService', 'Error showing browser notification:', error);
    }
  }

  /**
   * Show a simple notification with title and message
   */
  public async showSimple(title: string, message: string, priority: 'low' | 'normal' | 'high' | 'critical' = 'normal'): Promise<void> {
    const notification: ILiveNotification = {
      id: `simple_${Date.now()}`,
      type: 'SystemAlert' as any,
      title,
      message,
      timestamp: new Date(),
      read: false,
      userId: 0,
      priority
    };

    await this.show(notification);
  }

  /**
   * Enable or disable browser notifications
   */
  public setEnabled(enabled: boolean): void {
    this.enabled = enabled && this.permission === 'granted';
  }

  /**
   * Enable or disable notification sounds
   */
  public setSoundEnabled(enabled: boolean): void {
    this.soundEnabled = enabled;
  }

  /**
   * Check if browser notifications are supported
   */
  public isSupported(): boolean {
    return 'Notification' in window;
  }

  /**
   * Check if browser notifications are enabled
   */
  public isEnabled(): boolean {
    return this.enabled && this.permission === 'granted';
  }

  /**
   * Get current permission status
   */
  public getPermission(): NotificationPermission {
    return this.permission;
  }

  /**
   * Check if in quiet hours
   */
  public isQuietHours(quietHoursStart?: string, quietHoursEnd?: string): boolean {
    if (!quietHoursStart || !quietHoursEnd) {
      return false;
    }

    try {
      const now = new Date();
      const currentHour = now.getHours();
      const currentMinute = now.getMinutes();
      const currentTime = currentHour * 60 + currentMinute;

      const startParts = quietHoursStart.split(':');
      const startTime = parseInt(startParts[0], 10) * 60 + parseInt(startParts[1], 10);

      const endParts = quietHoursEnd.split(':');
      const endTime = parseInt(endParts[0], 10) * 60 + parseInt(endParts[1], 10);

      // Handle overnight quiet hours (e.g., 22:00 to 08:00)
      if (startTime > endTime) {
        return currentTime >= startTime || currentTime <= endTime;
      } else {
        return currentTime >= startTime && currentTime <= endTime;
      }
    } catch (error) {
      logger.error('BrowserNotificationService', 'Error checking quiet hours:', error);
      return false;
    }
  }

  // Private helper methods

  private getIconForNotification(notification: ILiveNotification): string {
    // Map notification types to icons
    // In production, use actual icon URLs
    switch (notification.type) {
      case 'TaskAssigned':
        return '/assets/task-assigned-icon.png';
      case 'TaskCompleted':
        return '/assets/task-complete-icon.png';
      case 'DueDateApproaching':
      case 'DueDatePassed':
        return '/assets/calendar-icon.png';
      case 'ApprovalRequired':
        return '/assets/approval-icon.png';
      case 'Mention':
        return '/assets/mention-icon.png';
      default:
        return '/assets/notification-icon.png';
    }
  }

  private getDurationForPriority(priority: 'low' | 'normal' | 'high' | 'critical'): number {
    switch (priority) {
      case 'low':
        return 3000; // 3 seconds
      case 'normal':
        return 5000; // 5 seconds
      case 'high':
        return 8000; // 8 seconds
      case 'critical':
        return 0; // Requires manual dismissal
      default:
        return 5000;
    }
  }

  private playNotificationSound(priority: 'low' | 'normal' | 'high' | 'critical'): void {
    try {
      // Create audio context for notification sounds
      const audioContext = new (window.AudioContext || (window as any).webkitAudioContext)();
      const oscillator = audioContext.createOscillator();
      const gainNode = audioContext.createGain();

      oscillator.connect(gainNode);
      gainNode.connect(audioContext.destination);

      // Different frequencies for different priorities
      let frequency = 440; // A4
      let duration = 0.1;

      switch (priority) {
        case 'low':
          frequency = 330; // E4
          duration = 0.05;
          break;
        case 'normal':
          frequency = 440; // A4
          duration = 0.1;
          break;
        case 'high':
          frequency = 554; // C#5
          duration = 0.15;
          break;
        case 'critical':
          frequency = 659; // E5
          duration = 0.2;
          break;
      }

      oscillator.frequency.setValueAtTime(frequency, audioContext.currentTime);
      gainNode.gain.setValueAtTime(0.3, audioContext.currentTime);
      gainNode.gain.exponentialRampToValueAtTime(0.01, audioContext.currentTime + duration);

      oscillator.start(audioContext.currentTime);
      oscillator.stop(audioContext.currentTime + duration);
    } catch (error) {
      // Silently fail if audio context not supported
      logger.debug('BrowserNotificationService', 'Could not play notification sound', { error });
    }
  }
}

export const BrowserNotificationService = new BrowserNotificationServiceClass();
