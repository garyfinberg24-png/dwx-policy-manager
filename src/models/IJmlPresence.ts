// Real-time Presence and Collaboration Models

export enum PresenceStatus {
  Online = 'Online',
  Away = 'Away',
  Busy = 'Busy',
  Offline = 'Offline'
}

export interface IUserPresence {
  userId: number;
  userName: string;
  email: string;
  status: PresenceStatus;
  lastActivity: Date;
  currentLocation?: string; // Which page/process they're viewing
  isTyping?: boolean;
}

export interface IProcessPresence {
  processId: number;
  viewers: IUserPresence[];
  editors: IUserPresence[];
  lastUpdate: Date;
}

export interface ITaskPresence {
  taskId: number;
  assignedTo?: IUserPresence;
  viewers: IUserPresence[];
  lastUpdate: Date;
}

export enum LiveNotificationType {
  TaskAssigned = 'TaskAssigned',
  TaskCompleted = 'TaskCompleted',
  TaskUpdated = 'TaskUpdated',
  TaskCommented = 'TaskCommented',
  ProcessCreated = 'ProcessCreated',
  ProcessUpdated = 'ProcessUpdated',
  ProcessCompleted = 'ProcessCompleted',
  ProcessDeleted = 'ProcessDeleted',
  Mention = 'Mention',
  DueDateApproaching = 'DueDateApproaching',
  DueDatePassed = 'DueDatePassed',
  ApprovalRequired = 'ApprovalRequired',
  ApprovalApproved = 'ApprovalApproved',
  ApprovalRejected = 'ApprovalRejected',
  SystemAlert = 'SystemAlert'
}

export interface ILiveNotification {
  id: string;
  type: LiveNotificationType;
  title: string;
  message: string;
  timestamp: Date;
  read: boolean;
  userId: number;
  processId?: number;
  taskId?: number;
  actionUrl?: string;
  actionLabel?: string;
  priority: 'low' | 'normal' | 'high' | 'critical';
  sender?: {
    id: number;
    name: string;
    email: string;
  };
  metadata?: Record<string, any>;
}

export interface INotificationPreferences {
  userId: number;
  emailNotifications: boolean;
  browserNotifications: boolean;
  soundEnabled: boolean;
  notifyOnTaskAssignment: boolean;
  notifyOnTaskCompletion: boolean;
  notifyOnMention: boolean;
  notifyOnDueDate: boolean;
  notifyOnApproval: boolean;
  quietHoursEnabled: boolean;
  quietHoursStart?: string; // "22:00"
  quietHoursEnd?: string; // "08:00"
}
