using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using KiCadSync.UI;

namespace KiCadSync
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567891")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class SwAddin : ISwAddin
    {
        private static readonly string LogFile = Path.Combine(
            Path.GetDirectoryName(typeof(SwAddin).Assembly.Location) ?? "C:\\temp",
            "KiCadSync.log");

        public static ISldWorks SwApp { get; private set; } = null!;
        private int _addinCookie;
        private ITaskpaneView? _taskpaneView;
        private SyncTaskPane? _taskpane;

        public static void Log(string msg)
        {
            try { File.AppendAllText(LogFile, $"[{DateTime.Now:HH:mm:ss}] {msg}\n"); }
            catch { }
        }

        #region ISwAddin

        public bool ConnectToSW(object thisSW, int cookie)
        {
            try
            {
                Log("ConnectToSW starting...");
                SwApp = (ISldWorks)thisSW;
                _addinCookie = cookie;
                SwApp.SetAddinCallbackInfo2(0, this, cookie);
                Log("Callback info set.");

                var resDir = Path.GetDirectoryName(typeof(SwAddin).Assembly.Location)!;
                var imageList = new string[]
                {
                    Path.Combine(resDir, "Resources", "icon_20.png"),
                    Path.Combine(resDir, "Resources", "icon_32.png"),
                    Path.Combine(resDir, "Resources", "icon_40.png"),
                    Path.Combine(resDir, "Resources", "icon_64.png"),
                    Path.Combine(resDir, "Resources", "icon_96.png"),
                    Path.Combine(resDir, "Resources", "icon_128.png"),
                };
                _taskpaneView = SwApp.CreateTaskpaneView3(imageList, "KiCad Sync");
                Log($"CreateTaskpaneView3 result: {_taskpaneView != null}");

                if (_taskpaneView == null)
                {
                    Log("Falling back to CreateTaskpaneView2");
                    _taskpaneView = SwApp.CreateTaskpaneView2("", "KiCad Sync");
                }

                _taskpane = new SyncTaskPane();
                _taskpane.CreateControl();
                Log($"Control created, handle: {_taskpane.Handle}");

                bool attached = _taskpaneView.DisplayWindowFromHandlex64(_taskpane.Handle.ToInt64());
                Log($"DisplayWindowFromHandlex64 result: {attached}");

                (_taskpaneView as DTaskpaneViewEvents_Event)!.TaskPaneActivateNotify += OnTaskPaneActivated;

                _taskpane.Init(SwApp);
                Log("Taskpane init complete. ConnectToSW complete.");

                return true;
            }
            catch (Exception ex)
            {
                Log($"ConnectToSW FAILED: {ex}");
                return false;
            }
        }

        public bool DisconnectFromSW()
        {
            try
            {
                _taskpane?.Dispose();
                if (_taskpaneView != null)
                {
                    if (_taskpaneView is DTaskpaneViewEvents_Event tpEvents)
                        tpEvents.TaskPaneActivateNotify -= OnTaskPaneActivated;
                    _taskpaneView.DeleteView();
                    Marshal.ReleaseComObject(_taskpaneView);
                }
                Marshal.ReleaseComObject(SwApp);
            }
            catch (Exception ex)
            {
                Log($"DisconnectFromSW error: {ex}");
            }
            return true;
        }

        #endregion

        private int OnTaskPaneActivated()
        {
            _taskpane?.BeginInvoke(new Action(() => _taskpane.SyncSizeToParent()));
            return 0;
        }

        #region COM Registration

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            var hklm = Microsoft.Win32.Registry.LocalMachine;
            var addinKey = hklm.CreateSubKey(
                $@"SOFTWARE\SolidWorks\AddIns\{{{t.GUID}}}");
            addinKey.SetValue(null, 1);   // load on startup
            addinKey.SetValue("Description", "KiCad ↔ SolidWorks ECAD/MCAD Sync");
            addinKey.SetValue("Title", "KiCad Sync");
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKey(
                $@"SOFTWARE\SolidWorks\AddIns\{{{t.GUID}}}", throwOnMissingSubKey: false);
        }

        #endregion
    }
}
