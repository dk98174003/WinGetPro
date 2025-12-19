#!/usr/bin/env python3
"""
Winget GUI (Tkinter) - single-page app for:
- Search packages (winget search)
- View upgrades (winget upgrade)
- View installed packages (winget list)
- Install / Upgrade / Uninstall selected
- Pin / Unpin selected (mark "do not update" via winget pin)

Tested conceptually against WinGet CLI table output. If Microsoft changes the output format,
you may need to adjust the table parser.

Requires:
- Windows 10/11 with WinGet (App Installer)
- Python 3.9+ with tkinter
"""

from __future__ import annotations

import subprocess
import os
import threading
import queue
import sys
import shutil
import re
import urllib.parse
import webbrowser
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

import tkinter as tk
from tkinter import ttk, messagebox


# -------------------------
# WinGet helpers + parsing
# -------------------------

@dataclass
class PackageRow:
    name: str = ""
    id: str = ""
    version: str = ""
    available: str = ""
    source: str = ""
    pinned: str = ""   # "Yes"/"No"
    pin_type: str = "" # Pinning/Blocking/Gating (if known)
    pin_version: str = ""  # gating version/range (if any)


def _create_no_window_flag() -> int:
    # Hide console window on Windows
    if sys.platform.startswith("win"):
        return getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return 0


def run_winget(args: List[str], timeout: Optional[int] = None) -> Tuple[int, str]:
    """
    Runs winget and returns (returncode, combined_output).
    """
    exe = shutil.which("winget")
    if not exe:
        return 127, "winget not found in PATH. Install 'App Installer' from Microsoft Store."
    cmd = [exe] + args

    try:
        p = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            creationflags=_create_no_window_flag(),
        )
        out = (p.stdout or "") + (("\n" + p.stderr) if p.stderr else "")
        return p.returncode, out.strip()
    except subprocess.TimeoutExpired:
        return 124, f"Timeout running: {' '.join(cmd)}"
    except Exception as e:
        return 1, f"Error running: {' '.join(cmd)}\n{e}"


def parse_winget_table(output: str) -> List[Dict[str, str]]:
    """
    Parse typical WinGet table output like:

    Name                      Id                          Version     Available  Source
    ------------------------------------------------------------------------------------
    ... rows ...

    Returns list of dict rows keyed by column header.
    """
    lines = [ln.rstrip("\n") for ln in output.splitlines() if ln.strip() != ""]
    if not lines:
        return []

    # Find header line with common columns
    header_idx = -1
    for i, ln in enumerate(lines):
        if ("Id" in ln) and (("Name" in ln) or ("Package" in ln)) and ("---" not in ln):
            header_idx = i
            break
    if header_idx == -1:
        return []

    # Usually the delimiter row is right after header
    delim_idx = header_idx + 1
    while delim_idx < len(lines) and not re.match(r"^-{3,}\s*$", lines[delim_idx].strip()):
        # some builds use a long dash line with spaces
        if re.match(r"^-{3,}", lines[delim_idx].strip()):
            break
        delim_idx += 1
    if delim_idx >= len(lines):
        return []

    header_line = lines[header_idx]
    # Column starts by finding occurrences of header names in the header line.
    # We detect words separated by 2+ spaces in the header.
    col_names = [c.strip() for c in re.split(r"\s{2,}", header_line.strip()) if c.strip()]
    if not col_names:
        return []

    # Determine start indices for each column name by searching in header_line.
    starts = []
    search_from = 0
    for name in col_names:
        pos = header_line.find(name, search_from)
        if pos == -1:
            # fallback: best-effort scan
            pos = header_line.find(name)
            if pos == -1:
                pos = search_from
        starts.append(pos)
        search_from = pos + len(name)

    # Build slices from starts to next start
    slices: List[Tuple[str, int, int]] = []
    for i, name in enumerate(col_names):
        start = starts[i]
        end = starts[i + 1] if i + 1 < len(starts) else None
        slices.append((name, start, end if end is not None else 10_000))

    rows: List[Dict[str, str]] = []
    for ln in lines[delim_idx + 1 :]:
        # Stop if winget prints a summary section
        if ln.strip().lower().startswith("no installed package"):
            break
        if ln.strip().startswith("---"):
            continue
        row: Dict[str, str] = {}
        for name, start, end in slices:
            row[name] = ln[start:end].strip()
        # drop empty junk rows
        if any(v for v in row.values()):
            rows.append(row)
    return rows


def get_pins() -> Dict[str, Dict[str, str]]:
    """
    Returns dict keyed by package Id -> { "Pin type": "...", "Version": "...", "Source": "..." }
    (columns depend on winget version; we normalize what we can)
    """
    rc, out = run_winget(["pin", "list"])
    if rc != 0:
        return {}
    table = parse_winget_table(out)
    pins: Dict[str, Dict[str, str]] = {}
    # Typical columns: Id, Source, Version, Pin type (but order can vary by version)
    for r in table:
        pid = r.get("Id", "") or r.get("ID", "") or ""
        if not pid:
            continue
        pins[pid] = {
            "Source": r.get("Source", ""),
            "Version": r.get("Version", ""),
            "Pin type": r.get("Pin type", "") or r.get("Pin", "") or r.get("Type", ""),
        }
    return pins


# -------------------------
# Tkinter App
# -------------------------

class WingetGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("WinGetPro")
        self.geometry("1150x700")
        self.minsize(1000, 600)

        self._set_window_icon()

        self.work_q: "queue.Queue[tuple]" = queue.Queue()
        self.current_mode = tk.StringVar(value="Upgrades")  # Search / Upgrades / Installed
        self.search_text = tk.StringVar(value="")
        self.silent_var = tk.BooleanVar(value=False)
        self.accept_var = tk.BooleanVar(value=True)
        self.include_pinned_var = tk.BooleanVar(value=False)
        self.include_unknown_var = tk.BooleanVar(value=False)
        self.uninstall_source_winget_var = tk.BooleanVar(value=True)

        self.pins: Dict[str, Dict[str, str]] = {}

        # PanedWindow (table/log) and default sash init
        self._paned: Optional[ttk.PanedWindow] = None
        self._sash_inited: bool = False

        # Sorting state
        self._sort_col: str = "Name"
        self._sort_reverse: bool = False
        self._sort_reverse_by_col: Dict[str, bool] = {}
        # Last loaded rows from winget (before local filter)
        self._all_rows: List[PackageRow] = []

        self._build_ui()
        self._check_winget()
        self._poll_queue()
        # Make the table take most vertical space by default (user can drag sash).
        self.after(200, self._init_default_sash)

        # Initial load: upgrades
        self.refresh()

    # ---- UI ----

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        # Let the main area (row=2) take extra vertical space, not the options row.
        self.rowconfigure(1, weight=0)
        self.rowconfigure(2, weight=1)
        self.rowconfigure(3, weight=0)

        top = ttk.Frame(self, padding=8)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(3, weight=1)

        ttk.Label(top, text="Mode:").grid(row=0, column=0, sticky="w")
        self.mode_combo = ttk.Combobox(top, textvariable=self.current_mode, state="readonly",
                            values=["Upgrades", "Search", "Installed"], width=12)
        self.mode_combo.grid(row=0, column=1, padx=(6, 12), sticky="w")
        self.mode_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh())

        ttk.Label(top, text="Search:").grid(row=0, column=2, sticky="w")
        self.search_entry = ttk.Entry(top, textvariable=self.search_text)
        self.search_entry.grid(row=0, column=3, sticky="ew", padx=(6, 12))
        self.search_entry.bind("<Return>", lambda e: self.refresh())
        self.search_entry.bind("<KeyRelease>", lambda e: self._on_filter_change())

        self.btn_refresh = ttk.Button(top, text="Refresh", command=self.refresh)
        self.btn_refresh.grid(row=0, column=4, padx=6)

        self.btn_install = ttk.Button(top, text="Install", command=self.install_selected)
        self.btn_install.grid(row=0, column=5, padx=6)

        self.btn_upgrade = ttk.Button(top, text="Upgrade", command=self.upgrade_selected)
        self.btn_upgrade.grid(row=0, column=6, padx=6)

        self.btn_upgrade_all = ttk.Button(top, text="Upgrade All", command=self.upgrade_all)
        self.btn_upgrade_all.grid(row=0, column=7, padx=6)

        self.btn_uninstall = ttk.Button(top, text="Uninstall", command=self.uninstall_selected)
        self.btn_uninstall.grid(row=0, column=8, padx=6)

        # Options row
        opt = ttk.Frame(self, padding=(8, 0, 8, 6))
        opt.grid(row=1, column=0, sticky="ew")
        opt.columnconfigure(6, weight=1)

        ttk.Checkbutton(opt, text="Silent (no UI)", variable=self.silent_var).grid(row=0, column=0, padx=(0, 10), sticky="w")
        ttk.Checkbutton(opt, text="Auto-accept agreements", variable=self.accept_var).grid(row=0, column=1, padx=(0, 10), sticky="w")
        ttk.Checkbutton(opt, text="Include pinned (upgrades)", variable=self.include_pinned_var).grid(row=0, column=2, padx=(0, 10), sticky="w")
        ttk.Checkbutton(opt, text="Include unknown (upgrades)", variable=self.include_unknown_var).grid(row=0, column=3, padx=(0, 10), sticky="w")
        ttk.Checkbutton(opt, text="Uninstall: force source=winget", variable=self.uninstall_source_winget_var).grid(row=0, column=4, padx=(0, 10), sticky="w")

        self.btn_pin = ttk.Button(opt, text="Pin (do not update)", command=self.pin_selected)
        self.btn_pin.grid(row=0, column=5, padx=6, sticky="e")

        self.btn_unpin = ttk.Button(opt, text="Unpin", command=self.unpin_selected)
        self.btn_unpin.grid(row=0, column=6, padx=6, sticky="e")

        # Main split: table + log
        main = ttk.PanedWindow(self, orient="vertical")
        self._paned = main
        main.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))

        table_frame = ttk.Frame(main)
        log_frame = ttk.Frame(main)
        # Give the table most of the resize weight.
        main.add(table_frame, weight=8)
        main.add(log_frame, weight=1)

        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        cols = ("Name", "Id", "Version", "Available", "Source", "Pinned", "PinType", "PinVersion")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="extended")
        for c in cols:
            # Click headers to sort
            self.tree.heading(c, text=c, command=lambda col=c: self._on_sort(col))
            # reasonable default widths
            w = 220 if c == "Name" else 170
            if c in ("Version", "Available", "Pinned"):
                w = 90
            if c in ("PinType", "PinVersion"):
                w = 110
            self.tree.column(c, width=w, anchor="w", stretch=(c == "Name"))
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        ysb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")

        # Context menu
        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="Refresh", command=self.refresh)
        self.menu.add_separator()
        self.menu.add_command(label="Install selected", command=self.install_selected)
        self.menu.add_command(label="Upgrade selected", command=self.upgrade_selected)
        self.menu.add_command(label="Upgrade all", command=self.upgrade_all)
        self.menu.add_command(label="Uninstall selected", command=self.uninstall_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Pin selected (do not update)", command=self.pin_selected)
        self.menu.add_command(label="Unpin selected", command=self.unpin_selected)
        self.menu.add_separator()
        self.menu.add_command(label="Copy Id", command=self.copy_selected_id)
        self.menu.add_command(label="Open Winstall page", command=self.open_winstall_selected)
        self.menu.add_command(label="Open winget.run page", command=self.open_wingetrun_selected)
        self.tree.bind("<Button-3>", self._popup_menu)

        # Log
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=10, wrap="word")
        self.log.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        log_scroll.grid(row=0, column=1, sticky="ns")

        # Status bar (text + progress)
        self.status = tk.StringVar(value="Ready")

        sb = ttk.Frame(self, padding=6)
        sb.grid(row=3, column=0, sticky="ew")
        sb.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(sb, textvariable=self.status, anchor="w")
        self.status_label.grid(row=0, column=0, sticky="ew")

        self.progress = ttk.Progressbar(sb, mode="indeterminate", length=180)
        self.progress.grid(row=0, column=1, padx=(10, 0), sticky="e")

    # ---- basics ----

    def _check_winget(self) -> None:
        rc, out = run_winget(["--version"])
        if rc != 0:
            messagebox.showerror(
                "WinGet not available",
                "WinGet (winget.exe) not found.\n\n"
                "Install 'App Installer' from Microsoft Store, then reopen this app.\n\n"
                f"Details:\n{out}"
            )

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.work_q.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "status":
                    self.status.set(payload)
                elif kind == "all_rows":
                    self._all_rows = payload
                elif kind == "table":
                    self._fill_table(payload)
                elif kind == "done":
                    self._set_busy(False)
                    # Keep "Loaded X rows." if that was the last status.
                    if not (self.status.get() or "").startswith("Loaded"):
                        self.status.set("Ready")
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _append_log(self, text: str) -> None:
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def _set_busy(self, busy: bool, text: Optional[str] = None) -> None:
        """Show a simple busy indicator and disable actions while WinGet runs."""
        if text is not None:
            self.status.set(text)

        # progressbar
        try:
            if busy:
                self.progress.start(10)
            else:
                self.progress.stop()
        except Exception:
            pass

        state = "disabled" if busy else "normal"
        widgets = [
            getattr(self, "mode_combo", None),
            getattr(self, "search_entry", None),
            getattr(self, "btn_refresh", None),
            getattr(self, "btn_install", None),
            getattr(self, "btn_upgrade", None),
            getattr(self, "btn_upgrade_all", None),
            getattr(self, "btn_uninstall", None),
            getattr(self, "btn_pin", None),
            getattr(self, "btn_unpin", None),
        ]
        for w in widgets:
            if w is None:
                continue
            try:
                w.configure(state=state)
            except Exception:
                pass


    def _popup_menu(self, event) -> None:
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def copy_selected_id(self) -> None:
        items = self.tree.selection()
        if not items:
            return
        vals = self.tree.item(items[0], "values")
        pid = vals[1] if len(vals) > 1 else ""
        if not pid:
            return
        self.clipboard_clear()
        self.clipboard_append(pid)
        self.status.set(f"Copied Id: {pid}")

    # ---- window icon ----


    def _winstall_url_for(self, pkg_id: str, pkg_name: str = "") -> str:
        """Return a winstall.app URL for a package id/name.

        If pkg_id looks like a normal WinGet id (no backslashes), open the app page.
        Otherwise, fall back to a winget.run search page (works for ARP/MSIX entries).
        """
        pkg_id = (pkg_id or "").strip()
        pkg_name = (pkg_name or "").strip()

        if pkg_id and "\\" not in pkg_id and not pkg_id.lower().startswith(("arp\\", "msix\\")):
            return "https://winstall.app/apps/" + urllib.parse.quote(pkg_id, safe="")

        q = pkg_name or pkg_id
        return "https://winget.run/search?query=" + urllib.parse.quote(q, safe="")

    def _wingetrun_url_for(self, pkg_id: str, pkg_name: str = "") -> str:
        """Return a winget.run URL for a package id/name.

        For real WinGet IDs (e.g., Microsoft.AppInstaller), open the pkg page.
        For ARP/MSIX-style entries, fall back to a winget.run search.
        """
        pkg_id = (pkg_id or "").strip()
        pkg_name = (pkg_name or "").strip()

        if pkg_id and "\\" not in pkg_id and not pkg_id.lower().startswith(("arp\\", "msix\\")):
            # winget.run uses /pkg/<publisher>/<app> style for dot-separated IDs
            return "https://winget.run/pkg/" + "/".join(urllib.parse.quote(p, safe="") for p in pkg_id.split("."))

        q = pkg_name or pkg_id
        return "https://winget.run/search?query=" + urllib.parse.quote(q, safe="")


    def open_package_page_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values") or ()
        if len(vals) < 2:
            return
        name = vals[0]
        pkg_id = vals[1]
        webbrowser.open(self._winstall_url_for(pkg_id, name))

    def open_winstall_selected(self) -> None:
        """Open the selected package on winstall.app."""
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values") or ()
        if len(vals) < 2:
            return
        name = vals[0]
        pkg_id = vals[1]
        webbrowser.open(self._winstall_url_for(pkg_id, name))

    def open_wingetrun_selected(self) -> None:
        """Open the selected package on winget.run."""
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values") or ()
        if len(vals) < 2:
            return
        name = vals[0]
        pkg_id = vals[1]
        webbrowser.open(self._wingetrun_url_for(pkg_id, name))

    def _on_tree_double_click(self, event) -> None:
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        vals = self.tree.item(row_id, "values") or ()
        if len(vals) < 2:
            return
        name = vals[0]
        pkg_id = vals[1]
        # Open package info page (winstall has a nice description)
        webbrowser.open(self._winstall_url_for(pkg_id, name))

        # Double-click on the "Id" column opens the package page
        if col_index != 1:
            return

        vals = self.tree.item(row_id, "values") or ()
        if len(vals) < 2:
            return
        name = vals[0]
        pkg_id = vals[1]
        webbrowser.open(self._winstall_url_for(pkg_id, name))


    def _set_window_icon(self) -> None:
        """
        Set the main window icon from an .ico file (Windows).
        Looks for wingetgetpro.ico (preferred) or wingetpro.ico next to the script.
        """
        if not sys.platform.startswith("win"):
            return
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            base = os.getcwd()

        candidates = [
            os.path.join(base, "wingetgetpro.ico"),
            os.path.join(base, "wingetpro.ico"),
            os.path.join(os.getcwd(), "wingetgetpro.ico"),
            os.path.join(os.getcwd(), "wingetpro.ico"),
        ]

        for p in candidates:
            if os.path.exists(p):
                try:
                    self.iconbitmap(p)
                except Exception:
                    pass
                break


    # ---- layout helpers ----

    def _init_default_sash(self) -> None:
        """
        Set a sensible default sash so the table fills the window and the log is smaller.
        Users can still drag the sash.
        """
        if self._sash_inited:
            return
        if not self._paned:
            return
        try:
            # Ensure geometry has settled
            self.update_idletasks()
            h = self._paned.winfo_height()
            if h <= 0:
                return
            log_target = 170  # px (default log area height)
            # Place sash so bottom pane ~log_target tall
            sash_y = max(200, h - log_target)
            self._paned.sashpos(0, sash_y)
            self._sash_inited = True
        except Exception:
            # Non-fatal; just skip
            pass


    # ---- sorting ----

    def _on_sort(self, col: str) -> None:
        """
        Toggle sort order for a column when the header is clicked.
        """
        reverse = self._sort_reverse_by_col.get(col, False)
        self._sort_col = col
        self._sort_reverse = reverse
        self._sort_reverse_by_col[col] = not reverse
        self._sort_tree(col, reverse)

    def _version_key(self, s: str):
        """
        Best-effort version key for sorting.
        Handles dotted versions like '1.2.10' and ignores non-numeric suffixes.
        Falls back to lowercase string.
        """
        if not s:
            return (0, ())
        s = s.strip()
        # Extract numeric groups; keep some text as a tiebreaker
        nums = re.findall(r"\d+", s)
        if nums:
            return (1, tuple(int(n) for n in nums), s.lower())
        return (0, (), s.lower())

    def _cell_key(self, col: str, value: str):
        v = (value or "").strip()

        if col in ("Version", "Available", "PinVersion"):
            return self._version_key(v)

        if col == "Pinned":
            # Yes first when ascending
            return 0 if v.lower() == "yes" else 1

        # Default: case-insensitive text
        return v.lower()

    def _sort_tree(self, col: str, reverse: bool) -> None:
        """
        Sort currently displayed rows in the treeview by column.
        """
        col_index = {"Name": 0, "Id": 1, "Version": 2, "Available": 3, "Source": 4,
                     "Pinned": 5, "PinType": 6, "PinVersion": 7}.get(col, 0)

        items = list(self.tree.get_children(""))
        sortable = []
        for iid in items:
            vals = self.tree.item(iid, "values")
            cell = vals[col_index] if col_index < len(vals) else ""
            sortable.append((self._cell_key(col, str(cell)), iid))

        sortable.sort(key=lambda t: t[0], reverse=reverse)

        for i, (_, iid) in enumerate(sortable):
            self.tree.move(iid, "", i)

    def _apply_last_sort(self) -> None:
        if self.tree.get_children(""):
            self._sort_tree(self._sort_col, self._sort_reverse)


    # ---- local filter (Installed/Upgrades) ----

    def _filter_rows(self, rows: List[PackageRow], text: str) -> List[PackageRow]:
        t = (text or "").strip().lower()
        if not t:
            return rows
        out: List[PackageRow] = []
        for r in rows:
            hay = f"{r.name} {r.id} {r.version} {r.available} {r.source} {r.pinned} {r.pin_type} {r.pin_version}".lower()
            if t in hay:
                out.append(r)
        return out

    def _apply_filter_only(self) -> None:
        """
        Apply filter to the *already loaded* rows (no winget call).
        Only used for Installed/Upgrades.
        """
        mode = self.current_mode.get()
        if mode == "Search":
            return
        filtered = self._filter_rows(self._all_rows, self.search_text.get())
        self._fill_table(filtered)

    def _on_filter_change(self) -> None:
        # Live filter while typing, but only for Installed/Upgrades.
        self._apply_filter_only()


    # ---- data loading ----

    def refresh(self) -> None:
        mode = self.current_mode.get()
        query = self.search_text.get().strip()

        # Search box is also used as a local filter in Installed/Upgrades.
        self.search_entry.state(["!disabled"]) 

        self._append_log(f"\n=== Refresh: {mode} ===")
        self._set_busy(True, f"Loading {mode}...")

        def worker() -> None:
            self.work_q.put(("status", f"Reading pins..."))
            self.pins = get_pins()

            if mode == "Search":
                if not query:
                    self.work_q.put(("log", "Search mode: enter a search term."))
                    self.work_q.put(("table", []))
                    self.work_q.put(("done", None))
                    return
                rc, out = run_winget(["search", query])
                self.work_q.put(("log", f"$ winget search {query}\n{out}\n"))
                rows = self._to_rows(parse_winget_table(out))
                self.work_q.put(("all_rows", rows))
                self.work_q.put(("table", self._filter_rows(rows, query)))
                self.work_q.put(("done", None))
                return

            if mode == "Installed":
                rc, out = run_winget(["list"])
                self.work_q.put(("log", f"$ winget list\n{out}\n"))
                rows = self._to_rows(parse_winget_table(out))
                self.work_q.put(("all_rows", rows))
                self.work_q.put(("table", self._filter_rows(rows, query)))
                self.work_q.put(("done", None))
                return

            # Upgrades
            cmd = ["upgrade"]
            if self.include_unknown_var.get():
                cmd.append("--include-unknown")
            if self.include_pinned_var.get():
                cmd.append("--include-pinned")
            rc, out = run_winget(cmd)
            self.work_q.put(("log", f"$ winget {' '.join(cmd)}\n{out}\n"))
            rows = self._to_rows(parse_winget_table(out))
            self.work_q.put(("all_rows", rows))
            self.work_q.put(("table", self._filter_rows(rows, query)))
            self.work_q.put(("done", None))

        threading.Thread(target=worker, daemon=True).start()

    def _to_rows(self, table_rows: List[Dict[str, str]]) -> List[PackageRow]:
        # Normalize dict rows to PackageRow
        rows: List[PackageRow] = []
        for r in table_rows:
            name = r.get("Name", "") or r.get("Package", "")
            pid = r.get("Id", "") or r.get("ID", "")
            version = r.get("Version", "")
            available = r.get("Available", "")
            source = r.get("Source", "")

            pin = self.pins.get(pid, {})
            pinned_yes = "Yes" if pid in self.pins else "No"
            pin_type = pin.get("Pin type", "")
            pin_version = pin.get("Version", "")

            rows.append(PackageRow(
                name=name, id=pid, version=version,
                available=available, source=source,
                pinned=pinned_yes, pin_type=pin_type, pin_version=pin_version
            ))
        return rows

    def _fill_table(self, rows: List[PackageRow]) -> None:
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=(
                r.name, r.id, r.version, r.available, r.source, r.pinned, r.pin_type, r.pin_version
            ))
        self._apply_last_sort()
        self.status.set(f"Loaded {len(rows)} rows.")

    # ---- Actions ----

    def _selected_ids(self) -> List[str]:
        ids: List[str] = []
        for item in self.tree.selection():
            vals = self.tree.item(item, "values")
            if len(vals) >= 2 and vals[1]:
                ids.append(vals[1])
        # unique preserve order
        seen = set()
        out = []
        for x in ids:
            if x not in seen:
                out.append(x)
                seen.add(x)
        return out

    def _build_common_flags(self, exact: bool = True) -> List[str]:
        flags: List[str] = []
        if exact:
            flags.append("--exact")

        # Show progress/UI by default. Silent suppresses all UI.
        if self.silent_var.get():
            flags.append("--silent")
        else:
            flags.append("--interactive")

        if self.accept_var.get():
            flags.extend(["--accept-package-agreements", "--accept-source-agreements"])
        return flags


    def install_selected(self) -> None:
        ids = self._selected_ids()
        if not ids:
            return
        self._run_many("Install", [["install", "--id", pid] + self._build_common_flags() for pid in ids])

    def upgrade_selected(self) -> None:
        ids = self._selected_ids()
        if not ids:
            return
        self._run_many("Upgrade", [["upgrade", "--id", pid] + self._build_common_flags() for pid in ids])


    def upgrade_all(self) -> None:
        if not messagebox.askyesno("Confirm upgrade all", "Upgrade all packages with available updates?"):
            return

        cmd = ["upgrade", "--all"]
        if self.include_unknown_var.get():
            cmd.append("--include-unknown")
        if self.include_pinned_var.get():
            cmd.append("--include-pinned")

        cmd += self._build_common_flags(exact=False)
        self._run_many("Upgrade All", [cmd])

    def uninstall_selected(self) -> None:
        ids = self._selected_ids()
        if not ids:
            return

        if not messagebox.askyesno("Confirm uninstall", f"Uninstall {len(ids)} selected package(s)?"):
            return

        def build(pid: str) -> List[str]:
            cmd = ["uninstall", "--id", pid, "--exact"]
            if self.uninstall_source_winget_var.get():
                cmd.extend(["--source", "winget"])  # reduce MS Store agreement prompts
            if self.silent_var.get():
                cmd.append("--silent")
            else:
                cmd.append("--interactive")
            if self.accept_var.get():
                cmd.append("--accept-source-agreements")
            cmd.append("--force")
            return cmd

        self._run_many("Uninstall", [build(pid) for pid in ids])

    def pin_selected(self) -> None:
        ids = self._selected_ids()
        if not ids:
            return
        # Default (non-blocking) pin: excluded from `winget upgrade --all`, but allows explicit upgrade.
        # See Microsoft docs: winget pin add --id <ID> (Pinning)
        cmds = [["pin", "add", "--id", pid, "--exact"] + (["--accept-source-agreements"] if self.accept_var.get() else [])
                for pid in ids]
        self._run_many("Pin", cmds, refresh_after=True)

    def unpin_selected(self) -> None:
        ids = self._selected_ids()
        if not ids:
            return
        cmds = [["pin", "remove", "--id", pid, "--exact"] + (["--accept-source-agreements"] if self.accept_var.get() else [])
                for pid in ids]
        self._run_many("Unpin", cmds, refresh_after=True)

    def _run_many(self, title: str, cmd_lists: List[List[str]], refresh_after: bool = False) -> None:
        self._set_busy(True, f"{title} running...")
        self._append_log(f"\n=== {title} ===")

        def worker() -> None:
            total = max(1, len(cmd_lists))
            for i, cmd in enumerate(cmd_lists, start=1):
                self.work_q.put(("status", f"{title} ({i}/{total}): winget {' '.join(cmd)}"))
                rc, out = run_winget(cmd)
                prefix = "$ winget " + " ".join(cmd)
                self.work_q.put(("log", f"{prefix}\n{out}\n"))
            if refresh_after:
                self.work_q.put(("status", "Refreshing..."))
            self.work_q.put(("done", None))
            # Refresh after actions so the list stays correct
            self.after(0, self.refresh)

        threading.Thread(target=worker, daemon=True).start()



def main() -> None:
    app = WingetGui()
    app.mainloop()


if __name__ == "__main__":
    main()
