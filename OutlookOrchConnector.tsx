/**
 * OutlookOrchConnector — port of MiBuddy's OutlookConnector.
 *
 * Keeps MiBuddy's UX intact (email icon, bullet-list of capabilities,
 * privacy note, Connect/Cancel buttons, Success screen with example
 * prompts) but rewritten in Tailwind instead of FluentUI so it matches
 * the rest of the orchestrator styling.
 *
 * Talks to the orchestrator-only endpoints under `/api/outlook-orch/...`
 * — NOT the existing `outlook-chat` routes, which are a different
 * use-case and must not be touched.
 */
import { useEffect, useRef, useState } from "react";
import { X, Loader2, Mail, CheckCircle2 } from "lucide-react";

interface OutlookOrchConnectorProps {
  isOpen: boolean;
  onDismiss: () => void;
  onConnected: () => void;
}

export default function OutlookOrchConnector({
  isOpen,
  onDismiss,
  onConnected,
}: OutlookOrchConnectorProps) {
  const [loading, setLoading] = useState(false);
  const [isConnected, setIsConnected] = useState(false);
  const connectingRef = useRef(false);
  const authWindowRef = useRef<Window | null>(null);

  // Reset state when the modal opens / closes
  useEffect(() => {
    if (isOpen) {
      setIsConnected(false);
      return;
    }
    setLoading(false);
    connectingRef.current = false;
    authWindowRef.current = null;
  }, [isOpen]);

  const handleConnect = () => {
    if (connectingRef.current) return;
    connectingRef.current = true;
    setLoading(true);

    const authWindow = window.open(
      "/api/outlook-orch/auth/login",
      "outlookAuth",
      "width=600,height=700",
    );
    authWindowRef.current = authWindow;
    if (!authWindow) {
      connectingRef.current = false;
      setLoading(false);
      alert("Popup blocked. Please allow popups and try again.");
    }
  };

  // Listen for the postMessage from /auth/callback, and also detect if the
  // user closed the popup themselves without completing.
  useEffect(() => {
    const onMessage = (event: MessageEvent) => {
      // The callback posts with target "*", so we can't strictly enforce
      // origin here — but the only thing we accept is our two known types.
      if (event.data?.type === "OUTLOOK_AUTH_SUCCESS") {
        setIsConnected(true);
        onConnected();
        connectingRef.current = false;
        setLoading(false);
        setTimeout(() => onDismiss(), 1200);
      }
      if (event.data?.type === "OUTLOOK_AUTH_ERROR") {
        connectingRef.current = false;
        setLoading(false);
        alert("Failed to connect to Outlook. Please try again.");
      }
    };

    const timer = window.setInterval(() => {
      if (
        connectingRef.current &&
        authWindowRef.current &&
        authWindowRef.current.closed
      ) {
        connectingRef.current = false;
        setLoading(false);
        authWindowRef.current = null;
      }
    }, 300);

    window.addEventListener("message", onMessage);
    return () => {
      window.removeEventListener("message", onMessage);
      window.clearInterval(timer);
    };
  }, [onConnected, onDismiss]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/50">
      <div className="flex max-h-[90vh] w-full max-w-[600px] flex-col rounded-2xl border border-border bg-popover shadow-2xl">
        {/* Header */}
        <div className="flex items-center justify-between border-b border-border px-5 py-4">
          <h2 className="text-base font-semibold text-foreground">
            {isConnected ? "Outlook Connected!" : "Connect to Outlook"}
          </h2>
          <button
            onClick={onDismiss}
            className="rounded-md p-1.5 text-muted-foreground hover:bg-accent hover:text-foreground"
          >
            <X size={18} />
          </button>
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-6">
          {loading ? (
            /* ── Loading screen ── */
            <div className="flex flex-col items-center gap-4 py-10">
              <Loader2 size={36} className="animate-spin text-blue-600" />
              <p className="text-sm text-muted-foreground">
                Connecting to your Outlook account...
              </p>
            </div>
          ) : isConnected ? (
            /* ── Connected screen ── */
            <div className="flex flex-col items-center gap-4 px-4 py-4 text-center">
              <CheckCircle2 size={56} className="text-green-600" />
              <div className="text-xl font-semibold text-green-600">
                Successfully Connected!
              </div>
              <div className="text-sm text-muted-foreground">
                You can now ask questions about your emails and calendar.
              </div>

              <div className="mt-2 w-full rounded-lg border border-border bg-muted/50 p-4 text-left">
                <p className="mb-2 text-sm font-semibold text-foreground">
                  Try asking:
                </p>
                <ul className="list-disc pl-5 text-sm text-muted-foreground">
                  <li>"Show me my recent emails"</li>
                  <li>"Do I have any meetings today?"</li>
                  <li>"Search for emails from John about the project"</li>
                  <li>"What's on my calendar this week?"</li>
                </ul>
              </div>
            </div>
          ) : (
            /* ── Consent screen ── */
            <div className="flex flex-col items-center gap-5 px-4 text-center">
              <div className="text-5xl">
                <Mail size={48} className="text-blue-600" />
              </div>
              <div>
                <p className="text-lg font-semibold text-foreground">
                  Connect MiBuddy to your Outlook
                </p>
                <p className="mt-1 text-sm text-muted-foreground">
                  Access your emails and calendar to get intelligent assistance
                </p>
              </div>

              {/* Capabilities list */}
              <div className="w-full rounded-lg border border-border bg-muted/30 p-4 text-left">
                <p className="mb-3 text-sm font-semibold text-foreground">
                  This will allow MiBuddy to:
                </p>
                <ul className="flex flex-col gap-2">
                  <li className="flex items-center gap-3 text-sm text-foreground">
                    <span className="font-bold text-green-600">✓</span>
                    Read your emails and search your mailbox
                  </li>
                  <li className="flex items-center gap-3 text-sm text-foreground">
                    <span className="font-bold text-green-600">✓</span>
                    View your calendar events and meetings
                  </li>
                  <li className="flex items-center gap-3 text-sm text-foreground">
                    <span className="font-bold text-green-600">✓</span>
                    Access your profile information
                  </li>
                </ul>
              </div>

              {/* Privacy note (MiBuddy's yellow card) */}
              <div className="w-full rounded-md border border-yellow-400 bg-yellow-50 p-3 text-left text-sm text-yellow-900 dark:border-yellow-700 dark:bg-yellow-950/30 dark:text-yellow-200">
                <strong>🔒 Privacy Note:</strong> Your emails and calendar data
                are only accessed when you ask questions. We never store your
                email content or share it with third parties.
              </div>
            </div>
          )}
        </div>

        {/* Footer (only on consent screen) */}
        {!loading && !isConnected && (
          <div className="flex items-center justify-center gap-3 border-t border-border px-5 py-3">
            <button
              onClick={handleConnect}
              disabled={loading}
              className="rounded-lg bg-blue-600 px-6 py-2.5 text-sm font-semibold text-white hover:bg-blue-700 disabled:opacity-60"
            >
              Connect to Outlook
            </button>
            <button
              onClick={onDismiss}
              className="rounded-lg border border-border px-5 py-2.5 text-sm font-medium text-foreground hover:bg-accent"
            >
              Cancel
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

/**
 * Auto-reconnect hook — port of MiBuddy's `useOutlookAutoReconnect`.
 * Polls /api/outlook-orch/status once on mount to restore "connected"
 * state after a page reload.
 */
export function useOutlookOrchStatus() {
  const [isOutlookConnected, setIsOutlookConnected] = useState(false);

  useEffect(() => {
    const check = async () => {
      try {
        const res = await fetch("/api/outlook-orch/status", {
          method: "GET",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
        });
        if (res.ok) {
          const data = await res.json();
          setIsOutlookConnected(Boolean(data.connected));
        }
      } catch {
        // ignore — treat as not connected
      }
    };
    check();
  }, []);

  return { isOutlookConnected, setIsOutlookConnected };
}

/** Disconnect helper (parallel to MiBuddy's inline usage). */
export async function disconnectOutlookOrch(): Promise<void> {
  await fetch("/api/outlook-orch/disconnect", {
    method: "POST",
    credentials: "include",
    headers: { "Content-Type": "application/json" },
  });
}
