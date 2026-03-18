"""KiCad action plugin — single button opens KiCad↔SolidWorks sync panel."""

import json
import os
import wx
import pcbnew
from pathlib import Path

from .sync import manifest, writer, reader


# ── Config persistence ────────────────────────────────────────────────

def _config_path(board: pcbnew.BOARD) -> Path | None:
    board_path = board.GetFileName()
    if not board_path:
        return None
    return Path(board_path).parent / ".kicad_sw_sync.json"


def load_config(board: pcbnew.BOARD) -> dict:
    p = _config_path(board)
    if p and p.exists():
        try:
            return json.loads(p.read_text())
        except Exception:
            pass
    return {}


def save_config(board: pcbnew.BOARD, cfg: dict):
    p = _config_path(board)
    if p:
        p.write_text(json.dumps(cfg, indent=2))


# ── Modeless sync panel ──────────────────────────────────────────────

_sync_frame = None  # singleton instance


class SyncFrame(wx.Frame):
    def __init__(self, parent, board: pcbnew.BOARD):
        super().__init__(parent, title="KiCad \u2194 SolidWorks Sync",
                         size=(340, 360),
                         style=wx.DEFAULT_FRAME_STYLE & ~wx.MAXIMIZE_BOX)
        self._board = board
        self._cfg = load_config(board)
        self._sync_dir = self._cfg.get("sync_dir", "")

        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # ── Sync directory ────────────────────────────────────────
        dir_box = wx.StaticBoxSizer(wx.VERTICAL, panel, "Sync Directory")

        dir_row = wx.BoxSizer(wx.HORIZONTAL)
        self._txt_dir = wx.TextCtrl(panel, value=self._sync_dir,
                                    style=wx.TE_READONLY)
        dir_row.Add(self._txt_dir, 1, wx.EXPAND | wx.RIGHT, 4)

        btn_browse = wx.Button(panel, label="...", size=(30, -1))
        btn_browse.SetToolTip("Browse for existing sync directory")
        btn_browse.Bind(wx.EVT_BUTTON, self._on_browse)
        dir_row.Add(btn_browse, 0)

        btn_new = wx.Button(panel, label="New", size=(50, -1))
        btn_new.SetToolTip("Create a new sync directory")
        btn_new.Bind(wx.EVT_BUTTON, self._on_new_dir)
        dir_row.Add(btn_new, 0, wx.LEFT, 4)

        dir_box.Add(dir_row, 0, wx.EXPAND | wx.ALL, 8)
        vbox.Add(dir_box, 0, wx.EXPAND | wx.ALL, 10)

        # ── Actions ───────────────────────────────────────────────
        act_box = wx.StaticBoxSizer(wx.VERTICAL, panel, "Actions")

        btn_push = wx.Button(panel, label="Push to SolidWorks", size=(-1, 36))
        btn_push.Bind(wx.EVT_BUTTON, self._on_push)
        act_box.Add(btn_push, 0, wx.EXPAND | wx.ALL, 8)

        btn_pull = wx.Button(panel, label="Pull from SolidWorks", size=(-1, 36))
        btn_pull.Bind(wx.EVT_BUTTON, self._on_pull)
        act_box.Add(btn_pull, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 8)

        vbox.Add(act_box, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

        # ── Status ────────────────────────────────────────────────
        self._status = wx.StaticText(panel, label="")
        vbox.Add(self._status, 0, wx.ALL, 10)

        panel.SetSizer(vbox)

        # Hide instead of destroy on close so we can reuse it
        self.Bind(wx.EVT_CLOSE, self._on_close)

        if not self._sync_dir:
            self._status.SetLabel("Set a sync directory to get started.")

    def _on_close(self, event):
        global _sync_frame
        _sync_frame = None
        self.Destroy()

    def _set_sync_dir(self, path: str):
        self._sync_dir = path
        self._txt_dir.SetValue(path)
        self._cfg["sync_dir"] = path
        save_config(self._board, self._cfg)
        self._status.SetLabel("Sync directory set.")

    def _on_browse(self, event):
        dlg = wx.DirDialog(self, "Select sync directory",
                           defaultPath=self._sync_dir or "")
        if dlg.ShowModal() == wx.ID_OK:
            self._set_sync_dir(dlg.GetPath())
        dlg.Destroy()

    def _on_new_dir(self, event):
        board_path = self._board.GetFileName()
        default = str(Path(board_path).parent) if board_path else ""
        dlg = wx.DirDialog(self, "Select parent folder for new sync directory",
                           defaultPath=default)
        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return
        parent = Path(dlg.GetPath())
        dlg.Destroy()

        name_dlg = wx.TextEntryDialog(self, "Sync folder name:", "New Sync Directory",
                                      "sw-sync")
        if name_dlg.ShowModal() != wx.ID_OK:
            name_dlg.Destroy()
            return
        name = name_dlg.GetValue().strip()
        name_dlg.Destroy()

        if not name:
            return

        sync_path = parent / name
        sync_path.mkdir(parents=True, exist_ok=True)
        (sync_path / "ecad_to_mcad").mkdir(exist_ok=True)
        (sync_path / "mcad_to_ecad").mkdir(exist_ok=True)
        self._set_sync_dir(str(sync_path))

    def _get_sync_dir(self) -> Path | None:
        if not self._sync_dir:
            wx.MessageBox("Set a sync directory first.",
                          "KiCad Sync", wx.OK | wx.ICON_WARNING)
            return None
        return Path(self._sync_dir)

    def _on_push(self, event):
        sync_dir = self._get_sync_dir()
        if not sync_dir:
            return

        board = self._board

        # Check drill origin
        origin_warning = writer.check_drill_origin(board)
        if origin_warning:
            result = wx.MessageBox(
                origin_warning + "\n\nContinue anyway?",
                "Drill Origin Warning",
                wx.YES_NO | wx.NO_DEFAULT | wx.ICON_WARNING)
            if result != wx.YES:
                return

        # Prompt for comment
        dlg = wx.TextEntryDialog(self, "Describe your changes (optional):",
                                 "Push to SolidWorks")
        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return
        comment = dlg.GetValue().strip()
        dlg.Destroy()

        try:
            changes = writer.push(board, sync_dir)
            mf = manifest.load(sync_dir)
            mf = manifest.record_push(mf, "ecad_to_mcad", "KiCad", comment, changes)
            manifest.save(sync_dir, mf)
            self._status.SetLabel(f"Pushed {len(changes)} change(s).")
        except Exception as e:
            self._status.SetLabel("Push failed.")
            wx.MessageBox(f"Push failed:\n{e}", "Push Error", wx.OK | wx.ICON_ERROR)

    def _on_pull(self, event):
        sync_dir = self._get_sync_dir()
        if not sync_dir:
            return

        board = self._board

        pending = reader.get_pending_changes(sync_dir)
        if not pending:
            wx.MessageBox("No changes from SolidWorks.",
                          "Pull from SolidWorks", wx.OK | wx.ICON_INFORMATION)
            return

        dlg = ChangeSummaryDialog(self, pending)
        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return
        selected = dlg.get_selected()
        dlg.Destroy()

        if not selected:
            return

        try:
            applied = reader.apply(board, sync_dir, selected)
            mf = manifest.load(sync_dir)
            mf = manifest.record_push(mf, "mcad_to_ecad", "SolidWorks", "", applied)
            manifest.save(sync_dir, mf)
            self._status.SetLabel(f"Applied {len(applied)} change(s).")
        except Exception as e:
            self._status.SetLabel("Pull failed.")
            wx.MessageBox(f"Pull failed:\n{e}", "Pull Error", wx.OK | wx.ICON_ERROR)


class ChangeSummaryDialog(wx.Dialog):
    """Shows a checklist of incoming changes so the user can accept/reject individually."""

    def __init__(self, parent, changes: list):
        super().__init__(parent, title="Pull from SolidWorks", size=(420, 360))
        self._changes = changes

        vbox = wx.BoxSizer(wx.VERTICAL)
        vbox.Add(wx.StaticText(self, label="Select changes to apply:"), 0, wx.ALL, 10)

        self._checklist = wx.CheckListBox(self, choices=[_describe(c) for c in changes])
        for i in range(len(changes)):
            self._checklist.Check(i, True)
        vbox.Add(self._checklist, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

        btns = self.CreateButtonSizer(wx.OK | wx.CANCEL)
        vbox.Add(btns, 0, wx.ALL | wx.ALIGN_RIGHT, 10)
        self.SetSizer(vbox)

    def get_selected(self) -> list:
        return [self._changes[i] for i in range(len(self._changes))
                if self._checklist.IsChecked(i)]


def _describe(change: dict) -> str:
    t = change["type"]
    if t == "board_outline_updated":
        return "Board outline updated"
    if t == "component_moved":
        pos = change.get("to", {})
        return f"{change['ref']} moved to ({pos.get('x_mm', '?'):.2f}, {pos.get('y_mm', '?'):.2f})"
    return t


# ── Plugin entry point ────────────────────────────────────────────────

class SyncPlugin(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "KiCad Sync"
        self.category = "ECAD/MCAD Sync"
        self.description = "Synchronize board and components with SolidWorks"
        self.icon_file_name = ""
        self.show_toolbar_button = True

    def Run(self):
        global _sync_frame

        board = pcbnew.GetBoard()
        if not board.GetFileName():
            wx.MessageBox("Save the board first so settings can be stored.",
                          "KiCad Sync", wx.OK | wx.ICON_WARNING)
            return

        # If already open, just bring it to front
        if _sync_frame is not None:
            _sync_frame.Raise()
            _sync_frame.Show()
            return

        _sync_frame = SyncFrame(None, board)
        _sync_frame.Show()


SyncPlugin().register()
