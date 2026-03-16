using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;

namespace KiCadSync
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567891")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SwAddin : ISwAddin
    {
        public static ISldWorks SwApp { get; private set; } = null!;
        private ICommandManager _cmdMgr = null!;
        private int _addinCookie;
        private CommandPanel _panel = null!;

        #region ISwAddin

        public bool ConnectToSW(object thisSW, int cookie)
        {
            SwApp = (ISldWorks)thisSW;
            _addinCookie = cookie;
            SwApp.SetAddinCallbackInfo2(0, this, cookie);

            _cmdMgr = SwApp.GetCommandManager(cookie);
            _panel = new CommandPanel(SwApp, _cmdMgr, cookie);
            _panel.Register();

            return true;
        }

        public bool DisconnectFromSW()
        {
            _panel?.Unregister();
            Marshal.ReleaseComObject(_cmdMgr);
            Marshal.ReleaseComObject(SwApp);
            return true;
        }

        #endregion

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
