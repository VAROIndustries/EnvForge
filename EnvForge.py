"""
Windows Environment Variable Manager
Manage System and User environment variables with copy/move between scopes.
Auto-elevates to Administrator for full System variable editing.
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import winreg
import ctypes
import sys
import os

# Registry paths
USER_ENV_KEY = r"Environment"
SYSTEM_ENV_KEY = r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"

SCOPE_USER = "User"
SCOPE_SYSTEM = "System"


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def broadcast_env_change():
    """Notify Windows that environment variables changed."""
    HWND_BROADCAST = 0xFFFF
    WM_SETTINGCHANGE = 0x001A
    SMTO_ABORTIFHUNG = 0x0002
    result = ctypes.c_long()
    ctypes.windll.user32.SendMessageTimeoutW(
        HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment",
        SMTO_ABORTIFHUNG, 5000, ctypes.byref(result)
    )


def read_env_vars(scope):
    """Read all environment variables for a scope."""
    vars_dict = {}
    try:
        if scope == SCOPE_USER:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, USER_ENV_KEY, 0, winreg.KEY_READ)
        else:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, SYSTEM_ENV_KEY, 0, winreg.KEY_READ)

        i = 0
        while True:
            try:
                name, value, reg_type = winreg.EnumValue(key, i)
                vars_dict[name] = (value, reg_type)
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
    except PermissionError:
        messagebox.showerror("Permission Denied",
                             f"Cannot read {scope} variables. Try running as Administrator.")
    except Exception as e:
        messagebox.showerror("Error", f"Error reading {scope} variables:\n{e}")
    return vars_dict


def write_env_var(scope, name, value, reg_type=None):
    """Write an environment variable. Returns True on success."""
    if reg_type is None:
        reg_type = winreg.REG_EXPAND_SZ if "%" in value else winreg.REG_SZ
    try:
        if scope == SCOPE_USER:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, USER_ENV_KEY, 0,
                                 winreg.KEY_SET_VALUE)
        else:
            if not is_admin():
                messagebox.showerror("Permission Denied",
                                     "Editing System variables requires Administrator privileges.\n"
                                     "Restart the app as Administrator.")
                return False
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, SYSTEM_ENV_KEY, 0,
                                 winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, name, 0, reg_type, value)
        winreg.CloseKey(key)
        broadcast_env_change()
        return True
    except PermissionError:
        messagebox.showerror("Permission Denied",
                             f"Cannot write to {scope} variables. Run as Administrator.")
        return False
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write variable:\n{e}")
        return False


def delete_env_var(scope, name):
    """Delete an environment variable. Returns True on success."""
    try:
        if scope == SCOPE_USER:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, USER_ENV_KEY, 0,
                                 winreg.KEY_SET_VALUE)
        else:
            if not is_admin():
                messagebox.showerror("Permission Denied",
                                     "Deleting System variables requires Administrator privileges.")
                return False
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, SYSTEM_ENV_KEY, 0,
                                 winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, name)
        winreg.CloseKey(key)
        broadcast_env_change()
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to delete variable:\n{e}")
        return False


def is_multivalue(value):
    """Return True if the value looks like a semicolon-separated list."""
    return ";" in value


class VarEditDialog(tk.Toplevel):
    """Dialog for editing or adding a single environment variable."""

    def __init__(self, parent, name="", value="", on_save=None, title="Edit Variable"):
        super().__init__(parent)
        self.title(title)
        self.on_save = on_save
        self.resizable(True, False)
        self.geometry("600x140")

        tk.Label(self, text="Name:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=8)
        self.name_var = tk.StringVar(value=name)
        self.name_entry = tk.Entry(self, textvariable=self.name_var, font=("Consolas", 10), width=50)
        self.name_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0, 10), pady=8)
        if name:
            self.name_entry.config(state="readonly")

        tk.Label(self, text="Value:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=4)
        self.value_var = tk.StringVar(value=value)
        self.value_entry = tk.Entry(self, textvariable=self.value_var, font=("Consolas", 10), width=50)
        self.value_entry.grid(row=1, column=1, sticky=tk.EW, padx=(0, 10), pady=4)

        btn_frame = tk.Frame(self)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=12)
        tk.Button(btn_frame, text="Save", command=self.save, bg="#0078d4", fg="white",
                  relief=tk.FLAT, padx=16).pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Cancel", command=self.destroy,
                  relief=tk.FLAT, padx=16).pack(side=tk.LEFT, padx=6)

        self.columnconfigure(1, weight=1)
        self.grab_set()
        self.transient(parent)

    def save(self):
        name = self.name_var.get().strip()
        value = self.value_var.get()
        if not name:
            messagebox.showwarning("Missing Name", "Variable name cannot be empty.", parent=self)
            return
        if self.on_save:
            self.on_save(name, value)
        self.destroy()


class EnvPanel(tk.Frame):
    """Panel displaying one scope's environment variables, with drill-down for multi-value vars."""

    def __init__(self, parent, scope, app, **kwargs):
        super().__init__(parent, **kwargs)
        self.scope = scope
        self.app = app
        self.vars_data = {}  # name -> (value, reg_type)
        self._sort_col = "Name"
        self._sort_rev = False
        self._drill_var_name = None

        header_color = "#0078d4" if scope == SCOPE_SYSTEM else "#107c10"

        # Header (always visible — content swaps between normal and drill mode)
        self.header = tk.Frame(self, bg=header_color)
        self.header.pack(fill=tk.X)

        self.back_btn = tk.Button(
            self.header, text="← Back", command=self._request_exit_drill,
            bg=header_color, fg="white", relief=tk.FLAT, padx=8, pady=4,
            activebackground=header_color, activeforeground="#ccffcc", cursor="hand2"
        )
        # Packed only in drill mode

        self.header_title = tk.Label(
            self.header, text=f"{scope} Variables",
            font=("Segoe UI", 11, "bold"),
            bg=header_color, fg="white", pady=6
        )
        self.header_title.pack(side=tk.LEFT, padx=10)

        self.count_label = tk.Label(self.header, text="", bg=header_color, fg="white")
        self.count_label.pack(side=tk.RIGHT, padx=10)

        # Main variables view
        self.main_frame = tk.Frame(self)
        self._build_main_view()
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Drill-down view (hidden until activated)
        self.drill_frame = tk.Frame(self)
        self._build_drill_view()

    # -------------------------------------------------------------------------
    # Main view
    # -------------------------------------------------------------------------

    def _build_main_view(self):
        search_frame = tk.Frame(self.main_frame, pady=4)
        search_frame.pack(fill=tk.X, padx=6)
        tk.Label(search_frame, text="Filter:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.apply_filter())
        tk.Entry(search_frame, textvariable=self.search_var, width=30).pack(side=tk.LEFT, padx=4)
        tk.Button(search_frame, text="x", command=self.clear_filter, width=2).pack(side=tk.LEFT)

        cols = ("Name", "Value")
        tree_frame = tk.Frame(self.main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 4))

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                 yscrollcommand=vsb.set, xscrollcommand=hsb.set,
                                 selectmode="browse")
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.heading("Name", text="Name", command=lambda: self.sort_by("Name"))
        self.tree.heading("Value", text="Value", command=lambda: self.sort_by("Value"))
        self.tree.column("Name", width=180, minwidth=100)
        self.tree.column("Value", width=340, minwidth=150)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.bind("<Double-Button-1>", self._on_main_double_click)
        self.tree.bind("<Delete>", self.delete_selected)

        btn_frame = tk.Frame(self.main_frame)
        btn_frame.pack(fill=tk.X, padx=6, pady=(0, 6))

        other = SCOPE_SYSTEM if self.scope == SCOPE_USER else SCOPE_USER
        tk.Button(btn_frame, text="+ New", command=self.new_var,
                  bg="#5c5c5c", fg="white", relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Edit", command=self.edit_selected,
                  relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Delete", command=self.delete_selected,
                  relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=2)

        tk.Button(btn_frame, text=f"Copy \u2192 {other}",
                  command=lambda: self.transfer(other, delete_source=False),
                  relief=tk.FLAT, padx=8, bg="#e8f0fe").pack(side=tk.RIGHT, padx=2)
        tk.Button(btn_frame, text=f"Move \u2192 {other}",
                  command=lambda: self.transfer(other, delete_source=True),
                  relief=tk.FLAT, padx=8, bg="#fce8e6").pack(side=tk.RIGHT, padx=2)

    def _on_main_double_click(self, event=None):
        name = self.selected_name()
        if not name:
            return
        value, reg_type = self.vars_data.get(name, ("", winreg.REG_SZ))
        if is_multivalue(value):
            self.app.try_drill_down(name, self.scope)
        else:
            VarEditDialog(self, name=name, value=value,
                          title=f"Edit {self.scope} Variable",
                          on_save=lambda n, v: self._save_var(n, v, reg_type))

    def reload(self):
        self.vars_data = read_env_vars(self.scope)
        self.apply_filter()
        self.count_label.config(text=f"{len(self.vars_data)} vars")

    def apply_filter(self):
        query = self.search_var.get().lower()
        self.tree.delete(*self.tree.get_children())
        items = sorted(self.vars_data.items(),
                       key=lambda x: x[0].lower() if self._sort_col == "Name" else x[1][0].lower(),
                       reverse=self._sort_rev)
        for name, (value, _) in items:
            if query in name.lower() or query in value.lower():
                tag = "multivalue" if is_multivalue(value) else ""
                self.tree.insert("", tk.END, iid=name, values=(name, value), tags=(tag,))
        self.tree.tag_configure("multivalue", foreground="#0055aa")

    def clear_filter(self):
        self.search_var.set("")

    def sort_by(self, col):
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = False
        self.apply_filter()

    def selected_name(self):
        sel = self.tree.selection()
        return sel[0] if sel else None

    def new_var(self):
        VarEditDialog(self, title=f"New {self.scope} Variable",
                      on_save=lambda n, v: self._save_var(n, v))

    def edit_selected(self, event=None):
        name = self.selected_name()
        if not name:
            return
        value, reg_type = self.vars_data.get(name, ("", winreg.REG_SZ))
        VarEditDialog(self, name=name, value=value, title=f"Edit {self.scope} Variable",
                      on_save=lambda n, v: self._save_var(n, v, reg_type))

    def _save_var(self, name, value, reg_type=None):
        if write_env_var(self.scope, name, value, reg_type):
            self.reload()

    def delete_selected(self, event=None):
        name = self.selected_name()
        if not name:
            return
        if messagebox.askyesno("Confirm Delete",
                               f"Delete {self.scope} variable '{name}'?",
                               parent=self):
            if delete_env_var(self.scope, name):
                self.reload()

    def transfer(self, target_scope, delete_source=False):
        name = self.selected_name()
        if not name:
            messagebox.showinfo("No Selection", "Select a variable first.", parent=self)
            return
        value, reg_type = self.vars_data[name]
        action = "Move" if delete_source else "Copy"
        if not messagebox.askyesno("Confirm",
                                   f"{action} '{name}' from {self.scope} to {target_scope}?",
                                   parent=self):
            return
        if write_env_var(target_scope, name, value, reg_type):
            if delete_source:
                delete_env_var(self.scope, name)
            self.app.reload_all()

    # -------------------------------------------------------------------------
    # Drill-down view
    # -------------------------------------------------------------------------

    def _build_drill_view(self):
        tree_frame = tk.Frame(self.drill_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(4, 0))

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        self.drill_tree = ttk.Treeview(
            tree_frame, columns=("#", "Entry"), show="headings",
            yscrollcommand=vsb.set, selectmode="browse"
        )
        vsb.config(command=self.drill_tree.yview)
        self.drill_tree.heading("#", text="#")
        self.drill_tree.heading("Entry", text="Entry")
        self.drill_tree.column("#", width=38, minwidth=30, stretch=False)
        self.drill_tree.column("Entry", width=420, minwidth=150)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.drill_tree.pack(fill=tk.BOTH, expand=True)
        self.drill_tree.bind("<Double-Button-1>", self._drill_edit_selected)

        btn_frame = tk.Frame(self.drill_frame)
        btn_frame.pack(fill=tk.X, padx=6, pady=4)

        other = SCOPE_SYSTEM if self.scope == SCOPE_USER else SCOPE_USER
        tk.Button(btn_frame, text="+ Add", command=self._drill_add,
                  bg="#5c5c5c", fg="white", relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Edit", command=self._drill_edit_selected,
                  relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="Delete", command=self._drill_delete,
                  relief=tk.FLAT, padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="\u25b2", command=self._drill_move_up,
                  relief=tk.FLAT, padx=6).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame, text="\u25bc", command=self._drill_move_down,
                  relief=tk.FLAT, padx=6).pack(side=tk.LEFT, padx=2)

        tk.Button(btn_frame, text=f"Copy \u2192 {other}",
                  command=lambda: self.app.drill_transfer_entry(
                      self._drill_selected_entry(), self._drill_var_name,
                      self.scope, other, delete_source=False),
                  relief=tk.FLAT, padx=8, bg="#e8f0fe").pack(side=tk.RIGHT, padx=2)
        tk.Button(btn_frame, text=f"Move \u2192 {other}",
                  command=lambda: self.app.drill_transfer_entry(
                      self._drill_selected_entry(), self._drill_var_name,
                      self.scope, other, delete_source=True),
                  relief=tk.FLAT, padx=8, bg="#fce8e6").pack(side=tk.RIGHT, padx=2)

    def enter_drill(self, var_name):
        """Switch to drill-down view for the given variable."""
        self._drill_var_name = var_name
        self.back_btn.pack(side=tk.LEFT, padx=(4, 0), before=self.header_title)
        self.header_title.config(text=f"  {var_name}")
        self.main_frame.pack_forget()
        self._refresh_drill()
        self.drill_frame.pack(fill=tk.BOTH, expand=True)

    def exit_drill(self):
        """Return to main variable list."""
        self._drill_var_name = None
        self.back_btn.pack_forget()
        self.header_title.config(text=f"{self.scope} Variables")
        self.drill_frame.pack_forget()
        self.reload()
        self.main_frame.pack(fill=tk.BOTH, expand=True)

    def _request_exit_drill(self):
        self.app.exit_drill_down()

    def _get_drill_entries(self):
        """Return (list_of_entries, reg_type) for the currently drilled variable."""
        if not self._drill_var_name:
            return [], winreg.REG_EXPAND_SZ
        value, reg_type = self.vars_data.get(self._drill_var_name, ("", winreg.REG_EXPAND_SZ))
        entries = [e.strip() for e in value.split(";") if e.strip()]
        return entries, reg_type

    def _save_drill_entries(self, entries):
        """Write entries to registry and refresh the drill view. Returns True on success."""
        value = ";".join(entries)
        _, reg_type = self.vars_data.get(self._drill_var_name, ("", winreg.REG_EXPAND_SZ))
        if write_env_var(self.scope, self._drill_var_name, value, reg_type):
            self.vars_data[self._drill_var_name] = (value, reg_type)
            self._refresh_drill()
            return True
        return False

    def refresh_drill_from_data(self):
        """Re-read vars_data from registry and refresh drill view (used after external changes)."""
        self.vars_data = read_env_vars(self.scope)
        self._refresh_drill()

    def _refresh_drill(self):
        entries, _ = self._get_drill_entries()
        self.drill_tree.delete(*self.drill_tree.get_children())
        for i, entry in enumerate(entries):
            self.drill_tree.insert("", tk.END, iid=str(i), values=(i + 1, entry))
        self.count_label.config(text=f"{len(entries)} entries")

    def _drill_selected_idx(self):
        sel = self.drill_tree.selection()
        return int(sel[0]) if sel else None

    def _drill_selected_entry(self):
        idx = self._drill_selected_idx()
        if idx is None:
            return None
        entries, _ = self._get_drill_entries()
        return entries[idx] if idx < len(entries) else None

    def _drill_add(self):
        val = simpledialog.askstring("Add Entry", "Enter value:", parent=self)
        if val and val.strip():
            entries, _ = self._get_drill_entries()
            entries.append(val.strip())
            self._save_drill_entries(entries)

    def _drill_edit_selected(self, event=None):
        idx = self._drill_selected_idx()
        if idx is None:
            return
        entries, _ = self._get_drill_entries()
        old = entries[idx]
        val = simpledialog.askstring("Edit Entry", "Edit value:", initialvalue=old, parent=self)
        if val is not None:
            entries[idx] = val.strip()
            self._save_drill_entries(entries)
            if self.drill_tree.exists(str(idx)):
                self.drill_tree.selection_set(str(idx))

    def _drill_delete(self):
        idx = self._drill_selected_idx()
        if idx is None:
            return
        entries, _ = self._get_drill_entries()
        if messagebox.askyesno("Confirm Delete",
                               f"Delete entry:\n{entries[idx]}", parent=self):
            entries.pop(idx)
            self._save_drill_entries(entries)

    def _drill_move_up(self):
        idx = self._drill_selected_idx()
        if idx is None or idx == 0:
            return
        entries, _ = self._get_drill_entries()
        entries[idx - 1], entries[idx] = entries[idx], entries[idx - 1]
        self._save_drill_entries(entries)
        new_iid = str(idx - 1)
        if self.drill_tree.exists(new_iid):
            self.drill_tree.selection_set(new_iid)

    def _drill_move_down(self):
        idx = self._drill_selected_idx()
        if idx is None:
            return
        entries, _ = self._get_drill_entries()
        if idx >= len(entries) - 1:
            return
        entries[idx], entries[idx + 1] = entries[idx + 1], entries[idx]
        self._save_drill_entries(entries)
        new_iid = str(idx + 1)
        if self.drill_tree.exists(new_iid):
            self.drill_tree.selection_set(new_iid)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EnvForge — Environment Variable Manager [Administrator]")
        self.geometry("1200x700")
        self.minsize(800, 500)

        menubar = tk.Menu(self)
        self.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Refresh All", accelerator="F5", command=self.reload_all)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)

        self.bind("<F5>", lambda e: self.reload_all())

        self.status_var = tk.StringVar(value="Ready")
        status_bar = tk.Label(self, textvariable=self.status_var, bd=1,
                              relief=tk.SUNKEN, anchor=tk.W, font=("Segoe UI", 9))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self.user_panel = EnvPanel(paned, SCOPE_USER, self, relief=tk.GROOVE, bd=1)
        self.sys_panel = EnvPanel(paned, SCOPE_SYSTEM, self, relief=tk.GROOVE, bd=1)
        paned.add(self.user_panel, weight=1)
        paned.add(self.sys_panel, weight=1)

        self.reload_all()

    def _panel_for(self, scope):
        return self.user_panel if scope == SCOPE_USER else self.sys_panel

    def _other_scope(self, scope):
        return SCOPE_SYSTEM if scope == SCOPE_USER else SCOPE_USER

    # -------------------------------------------------------------------------
    # Drill-down coordination
    # -------------------------------------------------------------------------

    def try_drill_down(self, var_name, source_scope):
        """
        Attempt to drill down into var_name from source_scope.
        Checks that the variable exists in both scopes; prompts to create if missing.
        """
        other_scope = self._other_scope(source_scope)
        other_panel = self._panel_for(other_scope)

        # Case-insensitive lookup in other scope
        canonical = next(
            (k for k in other_panel.vars_data if k.lower() == var_name.lower()), None
        )

        if canonical is None:
            # Variable doesn't exist in the other scope
            ans = messagebox.askyesno(
                "Variable Not Found",
                f"'{var_name}' does not exist in {other_scope} variables.\n\n"
                f"Create it as empty to enable drill-down?",
                parent=self
            )
            if not ans:
                return
            if not write_env_var(other_scope, var_name, ""):
                return
            other_panel.reload()
            canonical = var_name
        else:
            var_name = canonical  # Normalise to canonical casing

        self.user_panel.enter_drill(var_name)
        self.sys_panel.enter_drill(var_name)
        self.status_var.set(
            f"Drill-down: {var_name}  |  \u2190 Back to return  |  Double-click to edit entry"
        )

    def exit_drill_down(self):
        self.user_panel.exit_drill()
        self.sys_panel.exit_drill()
        self.status_var.set(
            "Ready  |  F5 to refresh  |  Double-click to edit  |  Double-click PATH-style vars to drill down"
        )

    def drill_transfer_entry(self, entry, var_name, src_scope, tgt_scope, delete_source=False):
        """Copy or move a single entry from src_scope's var to tgt_scope's var."""
        if entry is None:
            messagebox.showinfo("No Selection", "Select an entry first.", parent=self)
            return

        tgt_panel = self._panel_for(tgt_scope)
        t_value, t_reg_type = tgt_panel.vars_data.get(var_name, ("", winreg.REG_EXPAND_SZ))
        t_entries = [e.strip() for e in t_value.split(";") if e.strip()]

        if entry in t_entries:
            messagebox.showinfo(
                "Already Exists",
                f"This entry already exists in {tgt_scope}:\n\n{entry}",
                parent=self
            )
            return

        action = "Move" if delete_source else "Copy"
        if not messagebox.askyesno(
            "Confirm",
            f"{action} entry to {tgt_scope}?\n\n{entry}",
            parent=self
        ):
            return

        t_entries.append(entry)
        new_tgt_value = ";".join(t_entries)
        if not write_env_var(tgt_scope, var_name, new_tgt_value, t_reg_type):
            return
        tgt_panel.vars_data[var_name] = (new_tgt_value, t_reg_type)
        tgt_panel._refresh_drill()

        if delete_source:
            src_panel = self._panel_for(src_scope)
            s_value, s_reg_type = src_panel.vars_data.get(var_name, ("", winreg.REG_EXPAND_SZ))
            s_entries = [e.strip() for e in s_value.split(";") if e.strip()]
            s_entries = [e for e in s_entries if e != entry]
            new_src_value = ";".join(s_entries)
            if write_env_var(src_scope, var_name, new_src_value, s_reg_type):
                src_panel.vars_data[var_name] = (new_src_value, s_reg_type)
                src_panel._refresh_drill()

    # -------------------------------------------------------------------------

    def reload_all(self):
        self.status_var.set("Loading...")
        self.update_idletasks()
        self.user_panel.reload()
        self.sys_panel.reload()
        self.status_var.set(
            "Ready  |  F5 to refresh  |  Double-click to edit  |  Double-click PATH-style vars to drill down"
        )


if __name__ == "__main__":
    if not is_admin():
        # Re-launch with UAC elevation
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, f'"{os.path.abspath(__file__)}"', None, 1
        )
        sys.exit(0)
    app = App()
    app.mainloop()
