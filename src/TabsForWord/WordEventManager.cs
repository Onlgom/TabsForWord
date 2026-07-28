using System;
using Word = Microsoft.Office.Interop.Word;

namespace TabsForWord
{
    /// <summary>
    /// Subscribes to Word Application events and translates them into Reconcile calls.
    ///
    /// IMPORTANT: every delegate is stored in a field - subscribing with a rootless
    /// lambda lets the GC collect the delegate, after which COM events silently stop
    /// arriving (a documented problem, see docs/RESEARCH.md).
    ///
    /// Every handler is wrapped in try/catch: no exception may escape into Word.
    /// </summary>
    public sealed class WordEventManager : IDisposable
    {
        private readonly Word.Application _word;
        private readonly DocumentWindowManager _manager;
        private bool _subscribed;

        // --- delegates (GC roots) ---
        private Word.ApplicationEvents4_DocumentOpenEventHandler _onDocumentOpen;
        private Word.ApplicationEvents4_NewDocumentEventHandler _onNewDocument;
        private Word.ApplicationEvents4_DocumentChangeEventHandler _onDocumentChange;
        private Word.ApplicationEvents4_DocumentBeforeCloseEventHandler _onDocumentBeforeClose;
        private Word.ApplicationEvents4_DocumentBeforeSaveEventHandler _onDocumentBeforeSave;
        private Word.ApplicationEvents4_WindowActivateEventHandler _onWindowActivate;
        private Word.ApplicationEvents4_WindowDeactivateEventHandler _onWindowDeactivate;
        private Word.ApplicationEvents4_ProtectedViewWindowOpenEventHandler _onPvOpen;
        private Word.ApplicationEvents4_ProtectedViewWindowActivateEventHandler _onPvActivate;
        private Word.ApplicationEvents4_ProtectedViewWindowDeactivateEventHandler _onPvDeactivate;
        private Word.ApplicationEvents4_ProtectedViewWindowBeforeCloseEventHandler _onPvBeforeClose;
        private Word.ApplicationEvents4_QuitEventHandler _onQuit;

        public WordEventManager(Word.Application word, DocumentWindowManager manager)
        {
            _word = word;
            _manager = manager;
        }

        public void Subscribe()
        {
            if (_subscribed) return;

            _onDocumentOpen = doc => Safely("DocumentOpen", () => _manager.RequestReconcile("DocumentOpen"));
            _onNewDocument = doc => Safely("NewDocument", () => _manager.RequestReconcile("NewDocument"));
            _onDocumentChange = () => Safely("DocumentChange", () => _manager.RequestReconcile("DocumentChange"));
            _onWindowActivate = (doc, wn) => Safely("WindowActivate", () => _manager.RequestReconcile("WindowActivate"));
            _onWindowDeactivate = (doc, wn) => Safely("WindowDeactivate", () => _manager.RequestDeferredReconcile("WindowDeactivate"));

            // BeforeClose is only a hint: the user can still cancel the close in the
            // save dialog AFTER the event (see ADR-004).
            _onDocumentBeforeClose = new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(OnDocumentBeforeClose);

            // BeforeSave: the outcome (a new name after Save As) becomes known later -
            // the deferred reconcile and the background poll pick it up.
            _onDocumentBeforeSave = new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(OnDocumentBeforeSave);

            _onPvOpen = pv => Safely("ProtectedViewWindowOpen", () => _manager.RequestReconcile("PVOpen"));
            _onPvActivate = pv => Safely("ProtectedViewWindowActivate", () => _manager.RequestReconcile("PVActivate"));
            _onPvDeactivate = pv => Safely("ProtectedViewWindowDeactivate", () => _manager.RequestDeferredReconcile("PVDeactivate"));
            _onPvBeforeClose = new Word.ApplicationEvents4_ProtectedViewWindowBeforeCloseEventHandler(OnPvBeforeClose);

            _onQuit = () => Safely("Quit", () => LoggingService.Info("Word Quit event"));

            try
            {
                _word.DocumentOpen += _onDocumentOpen;
                // NewDocument clashes with the _Application.NewDocument member - subscribe through a cast.
                ((Word.ApplicationEvents4_Event)_word).NewDocument += _onNewDocument;
                _word.DocumentChange += _onDocumentChange;
                _word.DocumentBeforeClose += _onDocumentBeforeClose;
                _word.DocumentBeforeSave += _onDocumentBeforeSave;
                _word.WindowActivate += _onWindowActivate;
                _word.WindowDeactivate += _onWindowDeactivate;
                _word.ProtectedViewWindowOpen += _onPvOpen;
                _word.ProtectedViewWindowActivate += _onPvActivate;
                _word.ProtectedViewWindowDeactivate += _onPvDeactivate;
                _word.ProtectedViewWindowBeforeClose += _onPvBeforeClose;
                // Quit clashes with the _Application.Quit method - subscribe through a cast to the events interface.
                ((Word.ApplicationEvents4_Event)_word).Quit += _onQuit;
            }
            catch (Exception ex)
            {
                // A partial subscription is worse than none: remove the handlers already
                // added (Dispose would not touch them, since _subscribed=false) and let
                // the error propagate - initialisation then rolls back as a whole.
                LoggingService.Error("Event subscribe failed; removing partial handlers", ex);
                UnhookAll();
                throw;
            }

            _subscribed = true;
            LoggingService.Info("Word events subscribed");
        }

        public void Dispose()
        {
            if (!_subscribed) return;
            UnhookAll();
            _subscribed = false;
            LoggingService.Info("Word events unsubscribed");
        }

        /// <summary>
        /// Every unsubscribe sits in its own try/catch: a failure on one "-=" must not
        /// skip the rest (that could leave duplicated events and calls into already
        /// disposed managers). Unsubscribing a handler that was never added is a safe
        /// no-op, which is why the same code also serves to roll back a partial
        /// subscription.
        /// </summary>
        private void UnhookAll()
        {
            Unhook("DocumentOpen", () => _word.DocumentOpen -= _onDocumentOpen);
            Unhook("NewDocument", () => ((Word.ApplicationEvents4_Event)_word).NewDocument -= _onNewDocument);
            Unhook("DocumentChange", () => _word.DocumentChange -= _onDocumentChange);
            Unhook("DocumentBeforeClose", () => _word.DocumentBeforeClose -= _onDocumentBeforeClose);
            Unhook("DocumentBeforeSave", () => _word.DocumentBeforeSave -= _onDocumentBeforeSave);
            Unhook("WindowActivate", () => _word.WindowActivate -= _onWindowActivate);
            Unhook("WindowDeactivate", () => _word.WindowDeactivate -= _onWindowDeactivate);
            Unhook("ProtectedViewWindowOpen", () => _word.ProtectedViewWindowOpen -= _onPvOpen);
            Unhook("ProtectedViewWindowActivate", () => _word.ProtectedViewWindowActivate -= _onPvActivate);
            Unhook("ProtectedViewWindowDeactivate", () => _word.ProtectedViewWindowDeactivate -= _onPvDeactivate);
            Unhook("ProtectedViewWindowBeforeClose", () => _word.ProtectedViewWindowBeforeClose -= _onPvBeforeClose);
            Unhook("Quit", () => ((Word.ApplicationEvents4_Event)_word).Quit -= _onQuit);
        }

        private static void Unhook(string name, Action remove)
        {
            try { remove(); }
            catch (Exception ex) { LoggingService.Error("Event unsubscribe failed: " + name, ex); }
        }

        private void OnDocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            // Nothing is cancelled here (cancel is left alone) - we only record the intent.
            try
            {
                string name = "<unknown>";
                try { name = doc.Name; } catch { }
                LoggingService.Info("Document closing: " + name);
                _manager.RequestDeferredReconcile("DocumentBeforeClose");
            }
            catch (Exception ex)
            {
                LoggingService.Error("DocumentBeforeClose handler failed", ex);
            }
        }

        private void OnDocumentBeforeSave(Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            // Nothing is cancelled - we only schedule a reconcile after the save.
            try
            {
                _manager.RequestDeferredReconcile(saveAsUi ? "DocumentBeforeSave(SaveAs)" : "DocumentBeforeSave");
            }
            catch (Exception ex)
            {
                LoggingService.Error("DocumentBeforeSave handler failed", ex);
            }
        }

        private void OnPvBeforeClose(Word.ProtectedViewWindow pv, int closeReason, ref bool cancel)
        {
            try
            {
                // closeReason=1 (wdProtectedViewCloseEdit) => the user pressed
                // "Enable Editing": DocumentOpen follows and the tab turns into an
                // ordinary one.
                LoggingService.Info("PV window closing (reason=" + closeReason + ")");
                _manager.RequestDeferredReconcile("PVBeforeClose");
            }
            catch (Exception ex)
            {
                LoggingService.Error("PVBeforeClose handler failed", ex);
            }
        }

        private static void Safely(string eventName, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                LoggingService.Error(eventName + " handler failed", ex);
            }
        }
    }
}
