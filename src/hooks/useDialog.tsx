// @ts-nocheck
/**
 * useDialog Hook
 * Provides Fluent UI dialog replacements for native alert, confirm, and prompt dialogs
 * Offers better UX and consistent styling across the application
 */

import * as React from 'react';
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  TextField,
  IDialogContentProps,
  IModalProps,
  MessageBar,
  MessageBarType,
  Stack,
  Text
} from '@fluentui/react';

// ============================================================================
// INTERFACES
// ============================================================================

export type DialogVariant = 'info' | 'success' | 'warning' | 'error';

export interface IAlertOptions {
  title?: string;
  variant?: DialogVariant;
  confirmText?: string;
}

export interface IConfirmOptions {
  title?: string;
  confirmText?: string;
  cancelText?: string;
  variant?: DialogVariant;
  isDanger?: boolean;
}

export interface IPromptOptions {
  title?: string;
  label?: string;
  defaultValue?: string;
  placeholder?: string;
  confirmText?: string;
  cancelText?: string;
  required?: boolean;
  multiline?: boolean;
  rows?: number;
}

interface IDialogState {
  isOpen: boolean;
  type: 'alert' | 'confirm' | 'prompt';
  message: string;
  options: IAlertOptions | IConfirmOptions | IPromptOptions;
  resolve: ((value: boolean | string | null) => void) | null;
  inputValue: string;
}

interface IDialogContextValue {
  showAlert: (message: string, options?: IAlertOptions) => Promise<void>;
  showConfirm: (message: string, options?: IConfirmOptions) => Promise<boolean>;
  showPrompt: (message: string, options?: IPromptOptions) => Promise<string | null>;
}

// ============================================================================
// CONTEXT
// ============================================================================

const DialogContext = React.createContext<IDialogContextValue | undefined>(undefined);

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function getMessageBarType(variant: DialogVariant): MessageBarType {
  switch (variant) {
    case 'success':
      return MessageBarType.success;
    case 'warning':
      return MessageBarType.warning;
    case 'error':
      return MessageBarType.error;
    default:
      return MessageBarType.info;
  }
}

function getDialogIcon(variant: DialogVariant): string {
  switch (variant) {
    case 'success':
      return 'CheckMark';
    case 'warning':
      return 'Warning';
    case 'error':
      return 'ErrorBadge';
    default:
      return 'Info';
  }
}

// ============================================================================
// DIALOG PROVIDER COMPONENT
// ============================================================================

export const DialogProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [dialogState, setDialogState] = React.useState<IDialogState>({
    isOpen: false,
    type: 'alert',
    message: '',
    options: {},
    resolve: null,
    inputValue: ''
  });

  const showAlert = React.useCallback((message: string, options: IAlertOptions = {}): Promise<void> => {
    return new Promise((resolve) => {
      setDialogState({
        isOpen: true,
        type: 'alert',
        message,
        options,
        resolve: () => resolve(),
        inputValue: ''
      });
    });
  }, []);

  const showConfirm = React.useCallback((message: string, options: IConfirmOptions = {}): Promise<boolean> => {
    return new Promise((resolve) => {
      setDialogState({
        isOpen: true,
        type: 'confirm',
        message,
        options,
        resolve: (value) => resolve(value as boolean),
        inputValue: ''
      });
    });
  }, []);

  const showPrompt = React.useCallback((message: string, options: IPromptOptions = {}): Promise<string | null> => {
    return new Promise((resolve) => {
      setDialogState({
        isOpen: true,
        type: 'prompt',
        message,
        options,
        resolve: (value) => resolve(value as string | null),
        inputValue: options.defaultValue || ''
      });
    });
  }, []);

  const handleDismiss = React.useCallback((): void => {
    const { resolve, type } = dialogState;
    if (resolve) {
      if (type === 'confirm') {
        resolve(false);
      } else if (type === 'prompt') {
        resolve(null);
      } else {
        resolve(true);
      }
    }
    setDialogState(prev => ({ ...prev, isOpen: false, resolve: null }));
  }, [dialogState]);

  const handleConfirm = React.useCallback((): void => {
    const { resolve, type, inputValue, options } = dialogState;
    if (resolve) {
      if (type === 'prompt') {
        const promptOptions = options as IPromptOptions;
        if (promptOptions.required && !inputValue.trim()) {
          return; // Don't close if required and empty
        }
        resolve(inputValue);
      } else {
        resolve(true);
      }
    }
    setDialogState(prev => ({ ...prev, isOpen: false, resolve: null }));
  }, [dialogState]);

  const handleInputChange = React.useCallback((
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void => {
    setDialogState(prev => ({ ...prev, inputValue: newValue || '' }));
  }, []);

  const contextValue = React.useMemo(() => ({
    showAlert,
    showConfirm,
    showPrompt
  }), [showAlert, showConfirm, showPrompt]);

  const renderDialogContent = (): JSX.Element => {
    const { type, message, options, inputValue } = dialogState;

    if (type === 'alert') {
      const alertOptions = options as IAlertOptions;
      const variant = alertOptions.variant || 'info';

      return (
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={getMessageBarType(variant)}>
            {message}
          </MessageBar>
        </Stack>
      );
    }

    if (type === 'confirm') {
      const confirmOptions = options as IConfirmOptions;
      const variant = confirmOptions.variant;

      return (
        <Stack tokens={{ childrenGap: 16 }}>
          {variant ? (
            <MessageBar messageBarType={getMessageBarType(variant)}>
              {message}
            </MessageBar>
          ) : (
            <Text>{message}</Text>
          )}
        </Stack>
      );
    }

    if (type === 'prompt') {
      const promptOptions = options as IPromptOptions;

      return (
        <Stack tokens={{ childrenGap: 16 }}>
          <Text>{message}</Text>
          <TextField
            label={promptOptions.label}
            value={inputValue}
            onChange={handleInputChange}
            placeholder={promptOptions.placeholder}
            required={promptOptions.required}
            multiline={promptOptions.multiline}
            rows={promptOptions.rows || 3}
            autoFocus
          />
        </Stack>
      );
    }

    return <Text>{message}</Text>;
  };

  const renderDialogFooter = (): JSX.Element => {
    const { type, options, inputValue } = dialogState;

    if (type === 'alert') {
      const alertOptions = options as IAlertOptions;
      return (
        <DialogFooter>
          <PrimaryButton onClick={handleConfirm} text={alertOptions.confirmText || 'OK'} />
        </DialogFooter>
      );
    }

    if (type === 'confirm') {
      const confirmOptions = options as IConfirmOptions;
      const confirmButtonStyles = confirmOptions.isDanger
        ? { root: { backgroundColor: '#a80000', borderColor: '#a80000' }, rootHovered: { backgroundColor: '#8b0000' } }
        : undefined;

      return (
        <DialogFooter>
          <PrimaryButton
            onClick={handleConfirm}
            text={confirmOptions.confirmText || 'Confirm'}
            styles={confirmButtonStyles}
          />
          <DefaultButton onClick={handleDismiss} text={confirmOptions.cancelText || 'Cancel'} />
        </DialogFooter>
      );
    }

    if (type === 'prompt') {
      const promptOptions = options as IPromptOptions;
      const isDisabled = promptOptions.required && !inputValue.trim();

      return (
        <DialogFooter>
          <PrimaryButton
            onClick={handleConfirm}
            text={promptOptions.confirmText || 'Submit'}
            disabled={isDisabled}
          />
          <DefaultButton onClick={handleDismiss} text={promptOptions.cancelText || 'Cancel'} />
        </DialogFooter>
      );
    }

    return <DialogFooter><PrimaryButton onClick={handleConfirm} text="OK" /></DialogFooter>;
  };

  const getDialogTitle = (): string => {
    const { type, options } = dialogState;

    if (type === 'alert') {
      const alertOptions = options as IAlertOptions;
      if (alertOptions.title) return alertOptions.title;
      switch (alertOptions.variant) {
        case 'success': return 'Success';
        case 'warning': return 'Warning';
        case 'error': return 'Error';
        default: return 'Information';
      }
    }

    if (type === 'confirm') {
      return (options as IConfirmOptions).title || 'Confirm';
    }

    if (type === 'prompt') {
      return (options as IPromptOptions).title || 'Input Required';
    }

    return 'Dialog';
  };

  const dialogContentProps: IDialogContentProps = {
    type: DialogType.normal,
    title: getDialogTitle(),
    closeButtonAriaLabel: 'Close'
  };

  const modalProps: IModalProps = {
    isBlocking: true,
    styles: { main: { maxWidth: 500 } }
  };

  return (
    <DialogContext.Provider value={contextValue}>
      {children}
      <Dialog
        hidden={!dialogState.isOpen}
        onDismiss={handleDismiss}
        dialogContentProps={dialogContentProps}
        modalProps={modalProps}
      >
        {dialogState.isOpen && renderDialogContent()}
        {dialogState.isOpen && renderDialogFooter()}
      </Dialog>
    </DialogContext.Provider>
  );
};

// ============================================================================
// HOOK
// ============================================================================

export function useDialog(): IDialogContextValue {
  const context = React.useContext(DialogContext);
  if (!context) {
    throw new Error('useDialog must be used within a DialogProvider');
  }
  return context;
}

// ============================================================================
// STANDALONE FUNCTIONS (for class components)
// ============================================================================

/**
 * Creates a standalone dialog manager for use in class components
 * Returns a render function that must be included in the component's render method
 */
export function createDialogManager(): {
  showAlert: (message: string, options?: IAlertOptions) => Promise<void>;
  showConfirm: (message: string, options?: IConfirmOptions) => Promise<boolean>;
  showPrompt: (message: string, options?: IPromptOptions) => Promise<string | null>;
  DialogComponent: React.FC;
} {
  let resolveRef: ((value: boolean | string | null) => void) | null = null;
  let setStateRef: React.Dispatch<React.SetStateAction<IDialogState>> | null = null;

  const DialogComponent: React.FC = () => {
    const [state, setState] = React.useState<IDialogState>({
      isOpen: false,
      type: 'alert',
      message: '',
      options: {},
      resolve: null,
      inputValue: ''
    });

    React.useEffect(() => {
      setStateRef = setState;
      return () => { setStateRef = null; };
    }, []);

    const handleDismiss = (): void => {
      if (resolveRef) {
        if (state.type === 'confirm') {
          resolveRef(false);
        } else if (state.type === 'prompt') {
          resolveRef(null);
        } else {
          resolveRef(true);
        }
        resolveRef = null;
      }
      setState(prev => ({ ...prev, isOpen: false }));
    };

    const handleConfirm = (): void => {
      if (resolveRef) {
        if (state.type === 'prompt') {
          const promptOptions = state.options as IPromptOptions;
          if (promptOptions.required && !state.inputValue.trim()) {
            return;
          }
          resolveRef(state.inputValue);
        } else {
          resolveRef(true);
        }
        resolveRef = null;
      }
      setState(prev => ({ ...prev, isOpen: false }));
    };

    const handleInputChange = (
      _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
      newValue?: string
    ): void => {
      setState(prev => ({ ...prev, inputValue: newValue || '' }));
    };

    const renderContent = (): JSX.Element => {
      if (state.type === 'alert') {
        const alertOptions = state.options as IAlertOptions;
        const variant = alertOptions.variant || 'info';
        return (
          <Stack tokens={{ childrenGap: 16 }}>
            <MessageBar messageBarType={getMessageBarType(variant)}>
              {state.message}
            </MessageBar>
          </Stack>
        );
      }

      if (state.type === 'confirm') {
        const confirmOptions = state.options as IConfirmOptions;
        const variant = confirmOptions.variant;
        return (
          <Stack tokens={{ childrenGap: 16 }}>
            {variant ? (
              <MessageBar messageBarType={getMessageBarType(variant)}>
                {state.message}
              </MessageBar>
            ) : (
              <Text>{state.message}</Text>
            )}
          </Stack>
        );
      }

      if (state.type === 'prompt') {
        const promptOptions = state.options as IPromptOptions;
        return (
          <Stack tokens={{ childrenGap: 16 }}>
            <Text>{state.message}</Text>
            <TextField
              label={promptOptions.label}
              value={state.inputValue}
              onChange={handleInputChange}
              placeholder={promptOptions.placeholder}
              required={promptOptions.required}
              multiline={promptOptions.multiline}
              rows={promptOptions.rows || 3}
              autoFocus
            />
          </Stack>
        );
      }

      return <Text>{state.message}</Text>;
    };

    const renderFooter = (): JSX.Element => {
      if (state.type === 'alert') {
        const alertOptions = state.options as IAlertOptions;
        return (
          <DialogFooter>
            <PrimaryButton onClick={handleConfirm} text={alertOptions.confirmText || 'OK'} />
          </DialogFooter>
        );
      }

      if (state.type === 'confirm') {
        const confirmOptions = state.options as IConfirmOptions;
        const confirmButtonStyles = confirmOptions.isDanger
          ? { root: { backgroundColor: '#a80000', borderColor: '#a80000' }, rootHovered: { backgroundColor: '#8b0000' } }
          : undefined;
        return (
          <DialogFooter>
            <PrimaryButton
              onClick={handleConfirm}
              text={confirmOptions.confirmText || 'Confirm'}
              styles={confirmButtonStyles}
            />
            <DefaultButton onClick={handleDismiss} text={confirmOptions.cancelText || 'Cancel'} />
          </DialogFooter>
        );
      }

      if (state.type === 'prompt') {
        const promptOptions = state.options as IPromptOptions;
        const isDisabled = promptOptions.required && !state.inputValue.trim();
        return (
          <DialogFooter>
            <PrimaryButton
              onClick={handleConfirm}
              text={promptOptions.confirmText || 'Submit'}
              disabled={isDisabled}
            />
            <DefaultButton onClick={handleDismiss} text={promptOptions.cancelText || 'Cancel'} />
          </DialogFooter>
        );
      }

      return <DialogFooter><PrimaryButton onClick={handleConfirm} text="OK" /></DialogFooter>;
    };

    const getTitle = (): string => {
      if (state.type === 'alert') {
        const alertOptions = state.options as IAlertOptions;
        if (alertOptions.title) return alertOptions.title;
        switch (alertOptions.variant) {
          case 'success': return 'Success';
          case 'warning': return 'Warning';
          case 'error': return 'Error';
          default: return 'Information';
        }
      }
      if (state.type === 'confirm') {
        return (state.options as IConfirmOptions).title || 'Confirm';
      }
      if (state.type === 'prompt') {
        return (state.options as IPromptOptions).title || 'Input Required';
      }
      return 'Dialog';
    };

    return (
      <Dialog
        hidden={!state.isOpen}
        onDismiss={handleDismiss}
        dialogContentProps={{
          type: DialogType.normal,
          title: getTitle(),
          closeButtonAriaLabel: 'Close'
        }}
        modalProps={{
          isBlocking: true,
          styles: { main: { maxWidth: 500 } }
        }}
      >
        {state.isOpen && renderContent()}
        {state.isOpen && renderFooter()}
      </Dialog>
    );
  };

  const showAlert = (message: string, options: IAlertOptions = {}): Promise<void> => {
    return new Promise((resolve) => {
      resolveRef = () => resolve();
      if (setStateRef) {
        setStateRef({
          isOpen: true,
          type: 'alert',
          message,
          options,
          resolve: null,
          inputValue: ''
        });
      }
    });
  };

  const showConfirm = (message: string, options: IConfirmOptions = {}): Promise<boolean> => {
    return new Promise((resolve) => {
      resolveRef = (value) => resolve(value as boolean);
      if (setStateRef) {
        setStateRef({
          isOpen: true,
          type: 'confirm',
          message,
          options,
          resolve: null,
          inputValue: ''
        });
      }
    });
  };

  const showPrompt = (message: string, options: IPromptOptions = {}): Promise<string | null> => {
    return new Promise((resolve) => {
      resolveRef = (value) => resolve(value as string | null);
      if (setStateRef) {
        setStateRef({
          isOpen: true,
          type: 'prompt',
          message,
          options,
          resolve: null,
          inputValue: options.defaultValue || ''
        });
      }
    });
  };

  return {
    showAlert,
    showConfirm,
    showPrompt,
    DialogComponent
  };
}

export default useDialog;
