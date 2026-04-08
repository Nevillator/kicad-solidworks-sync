using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using SolidWorks.Interop.sldworks;

namespace KiCadSync.UI
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567892")]
    [ProgId("KiCadSync.SyncTaskPane")]
    public class SyncTaskPane : UserControl
    {
        private ISldWorks _swApp = null!;
        private Button _btnPull = null!;
        private Button _btnPush = null!;
        private Label _lblSyncDir = null!;
        private Button _btnBrowse = null!;
        private TextBox _txtSyncDir = null!;
        private CheckBox _chkHideImport = null!;
        private Label _lblStatus = null!;

        private string? _syncDir;

        private static readonly string SettingsPath = Path.Combine(
            System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
            "KiCadSync", "settings.json");

        public SyncTaskPane()
        {
            InitializeControls();
        }

        public void Init(ISldWorks swApp)
        {
            _swApp = swApp;

            // Load saved settings (env var overrides sync_dir)
            var (savedDir, hideImport) = LoadSettings();
            _syncDir = System.Environment.GetEnvironmentVariable("KICAD_SW_SYNC_DIR");
            if (string.IsNullOrEmpty(_syncDir))
                _syncDir = savedDir;
            _chkHideImport.Checked = hideImport;

            if (!string.IsNullOrEmpty(_syncDir))
                _txtSyncDir.Text = _syncDir;
        }

        private (string? syncDir, bool hideImport) LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var json = JObject.Parse(File.ReadAllText(SettingsPath));
                    return (json["sync_dir"]?.ToString(),
                            json["hide_during_import"]?.ToObject<bool>() ?? false);
                }
            }
            catch { }
            return (null, false);
        }

        private void SaveSettings()
        {
            try
            {
                var folder = Path.GetDirectoryName(SettingsPath)!;
                Directory.CreateDirectory(folder);
                var json = new JObject
                {
                    ["sync_dir"]           = _syncDir,
                    ["hide_during_import"] = _chkHideImport.Checked
                };
                File.WriteAllText(SettingsPath, json.ToString());
            }
            catch { }
        }

        private void InitializeControls()
        {
            BackColor = Color.White;
            Dock = DockStyle.Fill;
            Padding = new Padding(8);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                Padding = new Padding(0)
            };

            // Title
            var title = new Label
            {
                Text = "KiCad Sync",
                Font = new Font(Font.FontFamily, 12f, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 12)
            };
            layout.Controls.Add(title);

            // Sync directory section
            _lblSyncDir = new Label
            {
                Text = "Sync Directory:",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4)
            };
            layout.Controls.Add(_lblSyncDir);

            var dirPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0, 0, 0, 12)
            };

            _txtSyncDir = new TextBox
            {
                Width = 160,
                ReadOnly = true,
                Text = "(not set)"
            };
            dirPanel.Controls.Add(_txtSyncDir);

            _btnBrowse = new Button
            {
                Text = "...",
                Width = 30,
                Height = _txtSyncDir.Height
            };
            _btnBrowse.Click += OnBrowseClick;
            dirPanel.Controls.Add(_btnBrowse);

            layout.Controls.Add(dirPanel);

            // Hide during import checkbox
            _chkHideImport = new CheckBox
            {
                Text = "Hide SW during import",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8)
            };
            _chkHideImport.CheckedChanged += (_, _) => SaveSettings();
            layout.Controls.Add(_chkHideImport);

            // Pull button
            _btnPull = new Button
            {
                Text = "Pull from KiCad",
                Width = 200,
                Height = 36,
                Margin = new Padding(0, 0, 0, 6),
                FlatStyle = FlatStyle.System
            };
            _btnPull.Click += OnPullClick;
            layout.Controls.Add(_btnPull);

            // Push button
            _btnPush = new Button
            {
                Text = "Push to KiCad",
                Width = 200,
                Height = 36,
                Margin = new Padding(0, 0, 0, 12),
                FlatStyle = FlatStyle.System
            };
            _btnPush.Click += OnPushClick;
            layout.Controls.Add(_btnPush);

            // Status label
            _lblStatus = new Label
            {
                Text = "",
                AutoSize = true,
                ForeColor = Color.Gray
            };
            layout.Controls.Add(_lblStatus);

            Controls.Add(layout);
        }

        private void OnBrowseClick(object? sender, EventArgs e)
        {
            var path = BrowseForFolder("Select the KiCad/SolidWorks sync directory");
            if (path != null)
            {
                _syncDir = path;
                _txtSyncDir.Text = _syncDir;
                SaveSettings();
            }
        }

        /// <summary>
        /// Use the Vista-style IFileOpenDialog for folder picking (full Explorer window).
        /// Falls back to FolderBrowserDialog if COM fails.
        /// </summary>
        private string? BrowseForFolder(string title)
        {
            try
            {
                var dialog = (IFileOpenDialog)new FileOpenDialog();
                dialog.SetOptions(FOS.FOS_PICKFOLDERS | FOS.FOS_FORCEFILESYSTEM);
                dialog.SetTitle(title);

                // Set initial folder if we have one
                if (!string.IsNullOrEmpty(_syncDir) && Directory.Exists(_syncDir))
                {
                    SHCreateItemFromParsingName(_syncDir, IntPtr.Zero,
                        typeof(IShellItem).GUID, out var folder);
                    if (folder != null)
                        dialog.SetFolder(folder);
                }

                if (dialog.Show(Handle) == 0) // S_OK
                {
                    dialog.GetResult(out var item);
                    item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out var path);
                    return path;
                }
            }
            catch
            {
                // Fallback to classic dialog
                using var dlg = new FolderBrowserDialog { Description = title };
                if (dlg.ShowDialog() == DialogResult.OK)
                    return dlg.SelectedPath;
            }
            return null;
        }

        private string? GetSyncDir()
        {
            if (!string.IsNullOrEmpty(_syncDir))
                return _syncDir;

            OnBrowseClick(null, EventArgs.Empty);
            return _syncDir;
        }

        private void OnPullClick(object? sender, EventArgs e)
        {
            var syncDir = GetSyncDir();
            if (syncDir == null) return;

            _lblStatus.ForeColor = Color.Gray;
            _lblStatus.Text = "Starting pull...";
            _lblStatus.Update();
            _btnPull.Enabled = false;
            _btnPush.Enabled = false;

            var progress = new Progress<string>(msg =>
            {
                _lblStatus.Text = msg;
                _lblStatus.Update();
            });

            try
            {
                var mgr = new SyncManager(_swApp, syncDir);
                var doc = mgr.PullFromKiCad(out var changes, progress, _chkHideImport.Checked);

                _lblStatus.ForeColor = Color.Green;
                _lblStatus.Text = $"Pulled {changes.Count} item(s).";
            }
            catch (Exception ex)
            {
                _lblStatus.ForeColor = Color.Red;
                _lblStatus.Text = "Pull failed.";
                MessageBox.Show($"Pull failed:\n{ex.Message}",
                    "Pull Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnPull.Enabled = true;
                _btnPush.Enabled = true;
            }
        }

        private void OnPushClick(object? sender, EventArgs e)
        {
            var syncDir = GetSyncDir();
            if (syncDir == null) return;

            var comment = Microsoft.VisualBasic.Interaction.InputBox(
                "Describe your changes (optional):", "Push to KiCad", "");
            if (comment == null) return;

            _lblStatus.ForeColor = Color.Gray;
            _lblStatus.Text = "Pushing...";
            _btnPush.Enabled = false;

            try
            {
                var mgr = new SyncManager(_swApp, syncDir);
                var changes = mgr.PushToKiCad(comment);

                _lblStatus.ForeColor = Color.Green;
                _lblStatus.Text = $"Pushed {changes.Count} change(s).";
            }
            catch (Exception ex)
            {
                _lblStatus.ForeColor = Color.Red;
                _lblStatus.Text = "Push failed.";
                MessageBox.Show($"Push failed:\n{ex.Message}",
                    "Push Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnPush.Enabled = true;
            }
        }

        // ── COM interop for Vista IFileOpenDialog (folder picker) ────────

        [ComImport, Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
        private class FileOpenDialog { }

        [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog
        {
            [PreserveSig] int Show(IntPtr hwndOwner);
            void SetFileTypes();
            void SetFileTypeIndex();
            void GetFileTypeIndex();
            void Advise();
            void Unadvise();
            void SetOptions(FOS fos);
            void GetOptions();
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder();
            void GetCurrentSelection();
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName();
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel();
            void SetFileNameLabel();
            void GetResult(out IShellItem ppsi);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler();
            void GetParent();
            void GetDisplayName(SIGDN sigdnName,
                [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        }

        [Flags]
        private enum FOS : uint
        {
            FOS_PICKFOLDERS = 0x00000020,
            FOS_FORCEFILESYSTEM = 0x00000040,
        }

        private enum SIGDN : uint
        {
            SIGDN_FILESYSPATH = 0x80058000,
        }

        /// <summary>
        /// Sync our WinForms size to the Win32 parent (task pane container) client rect,
        /// then do a full repaint. Call this when the task pane tab is activated.
        /// </summary>
        public void SyncSizeToParent()
        {
            var parent = _GetParent(Handle);
            if (parent == IntPtr.Zero) { Refresh(); return; }
            _GetClientRect(parent, out var r);
            var sz = new Size(r.right - r.left, r.bottom - r.top);
            if (sz.Width > 0 && sz.Height > 0 && Size != sz)
            {
                Size = sz;
                PerformLayout();
            }
            Refresh();
        }

        [DllImport("user32.dll", EntryPoint = "GetParent")]
        private static extern IntPtr _GetParent(IntPtr hwnd);

        [DllImport("user32.dll", EntryPoint = "GetClientRect")]
        private static extern bool _GetClientRect(IntPtr hwnd, out _RECT rect);

        [StructLayout(LayoutKind.Sequential)]
        private struct _RECT { public int left, top, right, bottom; }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            string pszPath, IntPtr pbc, [In] Guid riid, out IShellItem ppv);

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // SyncTaskPane
            // 
            this.Name = "SyncTaskPane";
            this.Size = new System.Drawing.Size(807, 771);
            this.ResumeLayout(false);

        }
    }
}
