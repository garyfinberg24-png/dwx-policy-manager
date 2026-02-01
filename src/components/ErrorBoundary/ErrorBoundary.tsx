import * as React from 'react';
import { MessageBar, MessageBarType, PrimaryButton, Stack, Text } from '@fluentui/react';

interface IErrorBoundaryProps {
  children: React.ReactNode;
  fallbackMessage?: string;
}

interface IErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

/**
 * React Error Boundary component for catching render errors.
 * Wraps child components and displays a user-friendly error message
 * instead of crashing the entire application.
 */
export class ErrorBoundary extends React.Component<IErrorBoundaryProps, IErrorBoundaryState> {
  constructor(props: IErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  public static getDerivedStateFromError(error: Error): Partial<IErrorBoundaryState> {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error('ErrorBoundary caught an error:', error, errorInfo);
  }

  private handleRetry = (): void => {
    this.setState({ hasError: false, error: null });
  };

  public render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <Stack tokens={{ childrenGap: 16, padding: 24 }}>
          <MessageBar messageBarType={MessageBarType.error}>
            {this.props.fallbackMessage || 'Something went wrong while loading this section.'}
          </MessageBar>
          {this.state.error && (
            <Text variant="small" style={{ color: '#605e5c', fontFamily: 'monospace' }}>
              {this.state.error.message}
            </Text>
          )}
          <div>
            <PrimaryButton
              text="Try Again"
              iconProps={{ iconName: 'Refresh' }}
              onClick={this.handleRetry}
            />
          </div>
        </Stack>
      );
    }

    return this.props.children;
  }
}
