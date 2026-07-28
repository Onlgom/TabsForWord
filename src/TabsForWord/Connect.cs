using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

[assembly: ComVisible(false)]

namespace TabsForWord
{
    /// <summary>
    /// COM add-in entry point. Word creates this class by the ProgId "TabsForWord.Connect"
    /// (HKCU\Software\Microsoft\Office\Word\Addins\TabsForWord.Connect, LoadBehavior=3).
    ///
    /// Word calls in this order: OnConnection -> QI ICustomTaskPaneConsumer ->
    /// CTPFactoryAvailable -> OnStartupComplete (the last one only when loaded at startup).
    /// When the add-in is enabled by hand (ext_cm_AfterStartup) OnStartupComplete never
    /// arrives - initialisation is started from CTPFactoryAvailable instead.
    ///
    /// No exception may leave the methods of this class: otherwise Word flips
    /// LoadBehavior 3 -> 2 or puts the add-in into DisabledItems.
    /// </summary>
    [ComVisible(true)]
    [Guid("A3B7C9D1-5E2F-4A6B-8C0D-1F2E3D4C5B6A")]
    [ProgId("TabsForWord.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2, Office.ICustomTaskPaneConsumer
    {
        private Word.Application _word;
        private Office.ICTPFactory _ctpFactory;
        private ext_ConnectMode _connectMode;
        private bool _initialized;
        private bool _initializing;

        private TabPaneManager _paneManager;
        private NativeHost.NativeTabHostManager _nativeHostManager;   // null in the classic CTP mode
        private DocumentWindowManager _docManager;
        private WordEventManager _eventManager;

        /// <summary>
        /// A beacon for diagnosing other machines: this line appears as soon as the CLR
        /// has created the object by CLSID - BEFORE any call into Word. If the log has
        /// the beacon but no "Add-in connecting", the IDTExtensibility2 chain broke; if
        /// there is no beacon either, Word never reached our DLL at all
        /// (registration / CLR / policies).
        /// </summary>
        public Connect()
        {
            try
            {
                LoggingService.Info("Connect instantiated: v" + typeof(Connect).Assembly.GetName().Version +
                    ", process64=" + Environment.Is64BitProcess +
                    ", CLR=" + Environment.Version);
            }
            catch { /* logging must not get in the way of creating the COM object */ }
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _word = (Word.Application)Application;
                _connectMode = ConnectMode;
                LoggingService.Info("Add-in connecting (mode=" + ConnectMode + ", version=" +
                    typeof(Connect).Assembly.GetName().Version + ")");
                LogEnvironment();
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnConnection failed", ex);
            }
        }

        /// <summary>Environment details - the basis for diagnosing "it does not work for someone".</summary>
        private void LogEnvironment()
        {
            try
            {
                string wordVer = "?", wordBuild = "?";
                try { wordVer = _word.Version; } catch { }
                try { wordBuild = _word.Build; } catch { }
                LoggingService.Info("Environment: Word " + wordVer + " (build " + wordBuild +
                    "), OS " + Environment.OSVersion.Version +
                    ", os64=" + Environment.Is64BitOperatingSystem +
                    ", process64=" + Environment.Is64BitProcess +
                    ", culture=" + System.Globalization.CultureInfo.CurrentUICulture.Name +
                    ", dpiAware=" + (NativeMethods.IsProcessDpiAwareSafe() ? "1" : "0"));
            }
            catch (Exception ex)
            {
                LoggingService.Warn("LogEnvironment failed: " + ex.Message);
            }
        }

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            try
            {
                _ctpFactory = CTPFactoryInst;
                LoggingService.Info("CTP factory available");

                // The add-in was enabled by hand after Word started: OnStartupComplete will not come.
                if (_connectMode == ext_ConnectMode.ext_cm_AfterStartup)
                    TryInitialize();
            }
            catch (Exception ex)
            {
                LoggingService.Error("CTPFactoryAvailable failed", ex);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                TryInitialize();
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnStartupComplete failed", ex);
            }
        }

        private void TryInitialize()
        {
            if (_initialized || _initializing) return;
            if (_word == null) return;
            if (_ctpFactory == null)
            {
                // The factory has not arrived yet (it follows OnConnection) - wait for it.
                LoggingService.Warn("Initialize deferred: CTP factory not yet available");
                return;
            }

            _initializing = true;
            try
            {
                // Language first, before the strips are created: menu captions are built on
                // first show, but headers and tooltips are read straight away.
                UiLanguage.Init(_word);

                _paneManager = new TabPaneManager();

                // The main mode is the tab host inside the Word window (ADR-015).
                // The classic CustomTaskPane stays as the fallback: it is switched on by
                // config (mode=ctp) and it also backs up any failure to initialise the
                // in-window host.
                ITabPaneSync paneSync = _paneManager;
                if (NativeHost.TabHostSettings.Mode == NativeHost.TabHostMode.Native)
                {
                    try
                    {
                        _nativeHostManager = new NativeHost.NativeTabHostManager(_paneManager);
                        paneSync = _nativeHostManager;
                    }
                    catch (Exception ex)
                    {
                        LoggingService.Error("Native host init failed; using CustomTaskPane", ex);
                        _nativeHostManager = null;
                        paneSync = _paneManager;
                    }
                }

                _docManager = new DocumentWindowManager(_word, paneSync);
                _paneManager.Init(_ctpFactory, _docManager);
                if (_nativeHostManager != null) _nativeHostManager.Init(_docManager);
                _eventManager = new WordEventManager(_word, _docManager);

                _eventManager.Subscribe();
                _docManager.Start(); // initial scan of already open documents + timers
                                     // (this also installs the Ctrl+Tab hook if enabled in settings)

                // The success flag is set strictly AFTER full initialisation: otherwise a
                // failure half-way would block every retry for this Word session and leave
                // half-built managers and hooks behind.
                _initialized = true;
                LoggingService.Info("Add-in started");
            }
            catch (Exception ex)
            {
                LoggingService.Error("Initialize failed; rolling back partial state", ex);
                RollbackPartialInit();
            }
            finally
            {
                _initializing = false;
            }
        }

        /// <summary>
        /// Rolls back a partial initialisation: the next attempt (if Word calls again)
        /// starts from a clean state and no resources leak.
        /// </summary>
        private void RollbackPartialInit()
        {
            try { if (_eventManager != null) _eventManager.Dispose(); }
            catch (Exception ex) { LoggingService.Error("Rollback: event manager", ex); }
            _eventManager = null;

            try { if (_docManager != null) _docManager.Dispose(); }
            catch (Exception ex) { LoggingService.Error("Rollback: document manager", ex); }
            _docManager = null;

            try { if (_nativeHostManager != null) _nativeHostManager.Dispose(); }
            catch (Exception ex) { LoggingService.Error("Rollback: native host manager", ex); }
            _nativeHostManager = null;

            try { if (_paneManager != null) _paneManager.Dispose(); }
            catch (Exception ex) { LoggingService.Error("Rollback: pane manager", ex); }
            _paneManager = null;
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                LoggingService.Info("Word shutting down");
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnBeginShutdown failed", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (_eventManager != null) { _eventManager.Dispose(); _eventManager = null; }
                if (_docManager != null) { _docManager.Dispose(); _docManager = null; }
                if (_nativeHostManager != null) { _nativeHostManager.Dispose(); _nativeHostManager = null; }
                if (_paneManager != null) { _paneManager.Dispose(); _paneManager = null; }
                UiLanguage.Shutdown();
                _ctpFactory = null;
                _word = null;
                _initialized = false;
                LoggingService.Info("Add-in stopped (mode=" + RemoveMode + ")");
            }
            catch (Exception ex)
            {
                LoggingService.Error("OnDisconnection failed", ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }
    }
}
