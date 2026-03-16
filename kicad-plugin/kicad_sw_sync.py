"""KiCad action plugin — Push/Pull buttons for SolidWorks sync."""

import os
import wx
import pcbnew
from pathlib import Path

from sync import manifest, writer, reader


def get_sync_dir(board: pcbnew.BOARD) -> Path | None:
    """Return sync dir from environment variable or prompt user."""
    env = os.environ.get("KICAD_SW_SYNC_DIR")
    if env:
        return Path(env)
    # Fall back to a sibling directory of the board file
    board_path = board.GetFileName()
    if board_path:
        return Path(board_path).parent / "sw-sync"
    return None


class PushPlugin(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "Push to SolidWorks"
        self.category = "ECAD/MCAD Sync"
        self.description = "Export board and component data to the SolidWorks sync directory"
        self.icon_file_name = ""
        self.show_toolbar_button = True

    def Run(self):
        board = pcbnew.GetBoard()
        sync_dir = get_sync_dir(board)
        if not sync_dir:
            wx.MessageBox("Sync directory not set. Set KICAD_SW_SYNC_DIR or open a saved board.",
                          "KiCad↔SW Sync", wx.OK | wx.ICON_ERROR)
            return

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
        dlg = wx.TextEntryDialog(None, "Describe your changes (optional):", "Push to SolidWorks")
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
            wx.MessageBox(f"Pushed to SolidWorks sync directory.\n{len(changes)} change(s) recorded.",
                          "Push Complete", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
            wx.MessageBox(f"Push failed:\n{e}", "Push Error", wx.OK | wx.ICON_ERROR)


class PullPlugin(pcbnew.ActionPlugin):
    def defaults(self):
        self.name = "Pull from SolidWorks"
        self.category = "ECAD/MCAD Sync"
        self.description = "Apply board outline and component position updates from SolidWorks"
        self.icon_file_name = ""
        self.show_toolbar_button = True

    def Run(self):
        board = pcbnew.GetBoard()
        sync_dir = get_sync_dir(board)
        if not sync_dir:
            wx.MessageBox("Sync directory not set.", "KiCad↔SW Sync", wx.OK | wx.ICON_ERROR)
            return

        pending = reader.get_pending_changes(sync_dir)
        if not pending:
            wx.MessageBox("No changes from SolidWorks.", "Pull from SolidWorks", wx.OK | wx.ICON_INFORMATION)
            return

        # Show change summary dialog
        dlg = ChangeSummaryDialog(None, pending)
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
            wx.MessageBox(f"Applied {len(applied)} change(s) from SolidWorks.",
                          "Pull Complete", wx.OK | wx.ICON_INFORMATION)
        except Exception as e:
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
        return [self._changes[i] for i in range(len(self._changes)) if self._checklist.IsChecked(i)]


def _describe(change: dict) -> str:
    t = change["type"]
    if t == "board_outline_updated":
        return "Board outline updated"
    if t == "component_moved":
        pos = change.get("to", {})
        return f"{change['ref']} moved to ({pos.get('x_mm', '?'):.2f}, {pos.get('y_mm', '?'):.2f})"
    return t


PushPlugin().register()
PullPlugin().register()
