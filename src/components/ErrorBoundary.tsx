/**
 * Catches render-time errors in any subtree so a single bad tab doesn't
 * blank-screen the whole renderer. Shows the message + stack so users can
 * report exactly what broke.
 */

import { Component, type ErrorInfo, type ReactNode } from 'react';
import { AlertOctagon } from 'lucide-react';
import { tImp } from '../lib/i18n';

interface Props {
  children: ReactNode;
  /** Optional label shown above the error (e.g. tab name). */
  label?: string;
}

interface State {
  error: Error | null;
  info: ErrorInfo | null;
  copied: boolean;
}

export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null, info: null, copied: false };
  private copyTimer: ReturnType<typeof setTimeout> | null = null;

  static getDerivedStateFromError(error: Error): State {
    return { error, info: null, copied: false };
  }

  componentDidCatch(error: Error, info: ErrorInfo): void {
    this.setState({ error, info });
    // Surface to console so DevTools shows it.
    // eslint-disable-next-line no-console
    console.error('[ErrorBoundary]', error, info);
  }

  componentWillUnmount(): void {
    if (this.copyTimer) clearTimeout(this.copyTimer);
  }

  reset = () => {
    if (this.copyTimer) {
      clearTimeout(this.copyTimer);
      this.copyTimer = null;
    }
    this.setState({ error: null, info: null, copied: false });
  };

  copyDiagnostics = async () => {
    const { error, info } = this.state;
    if (!error) return;
    // Bundle everything a developer would otherwise have to expand and select
    // by hand: label, message, JS stack, React component stack.
    const lines = [
      this.props.label ? `[Context] ${this.props.label}` : null,
      `[Error] ${error.message}`,
      error.stack ? `\n[Stack]\n${error.stack}` : null,
      info?.componentStack ? `\n[Component stack]${info.componentStack}` : null,
    ].filter(Boolean);
    const text = lines.join('\n');
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      // Clipboard can fail in older Electron contexts or when the document is
      // not focused. Fall back to a transient textarea + execCommand so the
      // user still gets the copy rather than a silent no-op.
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try {
        document.execCommand('copy');
      } catch {
        /* give up — at least the panel still shows the stack for manual copy */
      }
      document.body.removeChild(ta);
    }
    this.setState({ copied: true });
    if (this.copyTimer) clearTimeout(this.copyTimer);
    this.copyTimer = setTimeout(() => {
      this.copyTimer = null;
      this.setState({ copied: false });
    }, 1500);
  };

  render(): ReactNode {
    const { error, info } = this.state;
    if (!error) return this.props.children;
    return (
      <div className="h-full w-full overflow-auto bg-background p-6">
        <div className="max-w-2xl mx-auto space-y-4">
          <div className="flex items-center gap-2 text-destructive">
            <AlertOctagon className="h-5 w-5" />
            <h2 className="text-base font-semibold">{tImp('編輯器發生錯誤', 'Editor encountered an error')}</h2>
          </div>
          {this.props.label ? (
            <div className="text-xs text-muted-foreground">{this.props.label}</div>
          ) : null}
          <div className="text-sm font-mono bg-secondary/40 border border-border rounded p-3 whitespace-pre-wrap">
            {error.message}
          </div>
          {error.stack ? (
            <details className="text-xs">
              <summary className="cursor-pointer text-muted-foreground">{tImp('堆疊', 'Stack')}</summary>
              <pre className="mt-2 text-[11px] font-mono bg-secondary/40 border border-border rounded p-3 overflow-auto whitespace-pre-wrap">
                {error.stack}
              </pre>
            </details>
          ) : null}
          {info?.componentStack ? (
            <details className="text-xs">
              <summary className="cursor-pointer text-muted-foreground">React component stack</summary>
              <pre className="mt-2 text-[11px] font-mono bg-secondary/40 border border-border rounded p-3 overflow-auto whitespace-pre-wrap">
                {info.componentStack}
              </pre>
            </details>
          ) : null}
          <div className="flex gap-2 pt-2">
            <button
              type="button"
              onClick={this.reset}
              className="text-xs px-3 py-1.5 rounded border border-border hover:bg-secondary"
            >
              {tImp('重試', 'Retry')}
            </button>
            <button
              type="button"
              onClick={this.copyDiagnostics}
              className="text-xs px-3 py-1.5 rounded border border-border hover:bg-secondary"
            >
              {this.state.copied ? tImp('已複製', 'Copied') : tImp('複製錯誤資訊', 'Copy error details')}
            </button>
          </div>
        </div>
      </div>
    );
  }
}
