// @ts-nocheck
// ToastService - Centralized toast notification service
// Provides consistent user feedback across the application

import { logger } from './LoggingService';

export enum ToastType {
  Success = 'success',
  Error = 'error',
  Warning = 'warning',
  Info = 'info'
}

export interface IToastMessage {
  id: string;
  type: ToastType;
  title: string;
  message?: string;
  duration?: number;
  action?: {
    label: string;
    onClick: () => void;
  };
}

type ToastListener = (toast: IToastMessage) => void;

class ToastServiceClass {
  private listeners: ToastListener[] = [];
  private toastCounter = 0;

  /**
   * Subscribe to toast notifications
   */
  public subscribe(listener: ToastListener): () => void {
    this.listeners.push(listener);

    // Return unsubscribe function
    return () => {
      const index = this.listeners.indexOf(listener);
      if (index > -1) {
        this.listeners.splice(index, 1);
      }
    };
  }

  /**
   * Show a toast notification
   */
  private showToast(toast: Omit<IToastMessage, 'id'>): void {
    const toastWithId: IToastMessage = {
      ...toast,
      id: `toast-${++this.toastCounter}-${Date.now()}`,
      duration: toast.duration || this.getDefaultDuration(toast.type)
    };

    this.listeners.forEach(listener => listener(toastWithId));
  }

  /**
   * Get default duration based on toast type
   */
  private getDefaultDuration(type: ToastType): number {
    switch (type) {
      case ToastType.Error:
        return 8000; // Errors stay longer
      case ToastType.Warning:
        return 6000;
      case ToastType.Success:
        return 4000;
      case ToastType.Info:
        return 5000;
      default:
        return 5000;
    }
  }

  /**
   * Show success toast
   */
  public success(title: string, message?: string, duration?: number): void {
    this.showToast({
      type: ToastType.Success,
      title,
      message,
      duration
    });
  }

  /**
   * Show error toast
   */
  public error(title: string, message?: string, duration?: number, action?: IToastMessage['action']): void {
    this.showToast({
      type: ToastType.Error,
      title,
      message,
      duration,
      action
    });
  }

  /**
   * Show warning toast
   */
  public warning(title: string, message?: string, duration?: number): void {
    this.showToast({
      type: ToastType.Warning,
      title,
      message,
      duration
    });
  }

  /**
   * Show info toast
   */
  public info(title: string, message?: string, duration?: number): void {
    this.showToast({
      type: ToastType.Info,
      title,
      message,
      duration
    });
  }

  /**
   * Show error toast with user-friendly message based on error type
   */
  public handleError(error: any, context?: string): void {
    logger.error('ToastService', `Error in ${context || 'application'}:`, error);

    let title = 'Something went wrong';
    let message = 'Please try again later';

    // Parse SharePoint/PnP errors
    if (error?.status) {
      switch (error.status) {
        case 400:
          title = 'Invalid Request';
          message = 'The request was not formatted correctly. Please check your input.';
          break;
        case 401:
          title = 'Authentication Required';
          message = 'Your session may have expired. Please refresh the page.';
          break;
        case 403:
          title = 'Access Denied';
          message = 'You do not have permission to perform this action.';
          break;
        case 404:
          title = 'Not Found';
          message = 'The requested item could not be found.';
          break;
        case 409:
          title = 'Conflict';
          message = 'This item has been modified by someone else. Please refresh and try again.';
          break;
        case 500:
        case 502:
        case 503:
          title = 'Server Error';
          message = 'The server encountered an error. Please try again later.';
          break;
        case 429:
          title = 'Too Many Requests';
          message = 'Please wait a moment before trying again.';
          break;
      }
    } else if (error?.message) {
      // Use error message if available
      message = error.message;
    }

    // Add context if provided
    if (context) {
      title = `${title} - ${context}`;
    }

    this.error(title, message, undefined, {
      label: 'Retry',
      onClick: () => {
        window.location.reload();
      }
    });
  }
}

// Singleton instance
export const ToastService = new ToastServiceClass();
