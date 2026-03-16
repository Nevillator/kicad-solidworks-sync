using System;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace KiCadSync.UI
{
    /// <summary>
    /// Registers Push/Pull toolbar commands and handles their callbacks.
    /// </summary>
    public class CommandPanel
    {
        private readonly ISldWorks _swApp;
        private readonly ICommandManager _cmdMgr;
        private readonly int _cookie;
        private ICommandGroup? _cmdGroup;

        private const int CmdIdPull = 1;
        private const int CmdIdPush = 2;

        public CommandPanel(ISldWorks swApp, ICommandManager cmdMgr, int cookie)
        {
            _swApp = swApp;
            _cmdMgr = cmdMgr;
            _cookie = cookie;
        }

        public void Register()
        {
            int cmdGroupErr = 0;
            _cmdGroup = _cmdMgr.CreateCommandGroup2(
                _cookie, "KiCad Sync", "KiCad ↔ SolidWorks Sync",
                "KiCad ↔ SolidWorks Sync", -1, false, ref cmdGroupErr);

            _cmdGroup.AddCommandItem2("Pull from KiCad", -1,
                "Import board and components from KiCad",
                "Pull from KiCad", 0,
                nameof(OnPull), "", CmdIdPull,
                (int)swCommandItemType_e.swMenuItem | (int)swCommandItemType_e.swToolbarItem);

            _cmdGroup.AddCommandItem2("Push to KiCad", -1,
                "Export board outline and component positions to KiCad",
                "Push to KiCad", 1,
                nameof(OnPush), "", CmdIdPush,
                (int)swCommandItemType_e.swMenuItem | (int)swCommandItemType_e.swToolbarItem);

            _cmdGroup.HasToolbar = true;
            _cmdGroup.HasMenu = true;
            _cmdGroup.Activate();
        }

        public void Unregister()
        {
            if (_cmdGroup != null)
            {
                _cmdMgr.RemoveCommandGroup(_cookie);
                _cmdGroup = null;
            }
        }

        // ── Callbacks (called by SolidWorks via add-in callback mechanism) ────

        public void OnPull()
        {
            var syncDir = GetSyncDir();
            if (syncDir == null) return;

            try
            {
                var mgr = new SyncManager(_swApp, syncDir);
                var doc = mgr.PullFromKiCad(out var changes);

                if (doc != null)
                    MessageBox.Show($"Pulled {changes.Count} item(s) from KiCad.",
                        "Pull Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Pull failed:\n{ex.Message}",
                    "Pull Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnPush()
        {
            var syncDir = GetSyncDir();
            if (syncDir == null) return;

            var comment = Microsoft.VisualBasic.Interaction.InputBox(
                "Describe your changes (optional):", "Push to KiCad", "");
            if (comment == null) return; // cancelled

            try
            {
                var mgr = new SyncManager(_swApp, syncDir);
                var changes = mgr.PushToKiCad(comment);

                MessageBox.Show($"Pushed {changes.Count} change(s) to KiCad sync directory.",
                    "Push Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Push failed:\n{ex.Message}",
                    "Push Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string? GetSyncDir()
        {
            var dir = System.Environment.GetEnvironmentVariable("KICAD_SW_SYNC_DIR");
            if (!string.IsNullOrEmpty(dir)) return dir;

            // Prompt user to select sync directory
            using var dlg = new FolderBrowserDialog
            {
                Description = "Select the KiCad↔SolidWorks sync directory"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
                return dlg.SelectedPath;

            return null;
        }
    }
}
