from tkinter import *
from tkinter import filedialog, simpledialog
from tkinter import messagebox
from tkinter import ttk
import sys
import threading
import re
from pathlib import Path
import yaml
from update_live_reqs import update_live_requests
import pandas as pd


def load_conf(conf_file):
    with open(conf_file, "r") as cf:
        conf = yaml.load(cf, Loader=yaml.FullLoader)
        return conf


def write_conf(cfg, conf_file):
    with open(conf_file, "w") as cf:
        yaml.dump(cfg, cf)

conf = {}
selected_file = None


class CredentialsDialog(Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Prijava na LiveSD")
        self.transient(parent)
        self.resizable(False, False)
        self.result = None

        body = ttk.Frame(self, padding=16)
        body.pack(fill=BOTH, expand=True)

        lbl_u = ttk.Label(body, text="Username (email):")
        lbl_u.grid(row=0, column=0, sticky=W, padx=(0, 8), pady=(0, 8))
        self.ent_user = ttk.Entry(body, width=42)
        self.ent_user.grid(row=0, column=1, sticky=EW, pady=(0, 8))

        lbl_p = ttk.Label(body, text="Password:")
        lbl_p.grid(row=1, column=0, sticky=W, padx=(0, 8))
        self.ent_pass = ttk.Entry(body, width=42, show='*')
        self.ent_pass.grid(row=1, column=1, sticky=EW)

        btns = ttk.Frame(body)
        btns.grid(row=2, column=0, columnspan=2, sticky=E, pady=(16, 0))
        btn_ok = ttk.Button(btns, text="Prijavi se", command=self.on_ok)
        btn_ok.pack(side=RIGHT, padx=(8, 0))
        btn_cancel = ttk.Button(btns, text="Otkaži", command=self.on_cancel)
        btn_cancel.pack(side=RIGHT)

        body.columnconfigure(1, weight=1)

        self.bind('<Return>', lambda e: self.on_ok())
        self.bind('<Escape>', lambda e: self.on_cancel())

        self.after(10, self._center_on_parent)
        self.grab_set()
        self.ent_user.focus_set()

    def _center_on_parent(self):
        try:
            self.update_idletasks()
            pw = self.master.winfo_width()
            ph = self.master.winfo_height()
            px = self.master.winfo_rootx()
            py = self.master.winfo_rooty()
            w = max(420, self.winfo_width())
            h = max(180, self.winfo_height())
            x = px + (pw // 2) - (w // 2)
            y = py + (ph // 2) - (h // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def on_ok(self):
        user = self.ent_user.get().strip()
        pwd = self.ent_pass.get()
        if not user:
            messagebox.showerror("Greška u logovanju", "Morate uneti username")
            return
        self.result = (user, pwd)
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()


def run_update_async(username, pwd, mapping):
    total = len(mapping)
    progress = {"done": 0, "ok": 0}

    def progress_cb(rid, year, ok):
        progress["done"] += 1
        if ok:
            progress["ok"] += 1
        pct = int(progress["done"] * 100 / total) if total else 0
        # Update UI safely from main thread
        window.after(0, lambda: update_progress_ui(progress["done"], total, pct))

    def worker():
        try:
            res = update_live_requests(conf, username, pwd, mapping, progress_callback=progress_cb)
            if isinstance(res, tuple) and len(res) == 2:
                ok_count, total_count = res
            else:
                ok_count, total_count = progress["ok"], total
        finally:
            window.after(0, on_update_finished)

    threading.Thread(target=worker, daemon=True).start()


def update_progress_ui(done, total, pct):
    progress_var.set(pct)
    lbl_status.configure(text=f"Obrađeno: {done}/{total}")


def on_update_finished():
    btn_start.configure(state=NORMAL)
    btn_browse.configure(state=NORMAL)
    try:
        pbar.stop()
        pbar.configure(mode='determinate')
        progress_var.set(100)
    except Exception:
        pass
    lbl_status.configure(text="Završeno ažuriranje")
    messagebox.showinfo("Obaveštenje", "Završeno ažuriranje.")


def choosefile():
    global selected_file
    window.filename = filedialog.askopenfilename(initialdir="/", title="Odaberite Excel fajl",
                                                 filetypes=(("xlsx files", "*.xlsx"), ("xls files", "*.xls")))
    excelPath = str(window.filename or "")
    if not (excelPath.endswith(".xlsx") or excelPath.endswith(".xls")):
        messagebox.showerror("Greška", "Morate izabrati Excel fajl")
        return
    selected_file = excelPath
    lbl_file.configure(text=excelPath)
    lbl_status.configure(text="Spremno")


def start_update():
    if not selected_file:
        messagebox.showerror("Greška", "Najpre izaberite Excel fajl.")
        return
    dlg = CredentialsDialog(window)
    window.wait_window(dlg)
    if not dlg.result:
        return
    username, pwd = dlg.result
    try:
        df = pd.read_excel(
            selected_file,
            usecols=["tp_ID", "Godina"],
            dtype={"tp_ID": str, "Godina": str},
            engine="openpyxl"
        )
    except Exception:
        messagebox.showerror("Greška", "Excel mora imati kolone 'tp_ID' i 'Godina'.")
        return

    df = df.dropna(subset=["tp_ID", "Godina"]).copy()
    df["tp_ID"] = df["tp_ID"].astype(str).str.replace(r"\D+", "", regex=True).str.strip()
    df["Godina"] = df["Godina"].astype(str).str.strip()
    df = df[(df["tp_ID"] != "") & (df["Godina"].str.isdigit())]
    m = dict(zip(df["tp_ID"], df["Godina"]))

    if not m:
        messagebox.showerror("Greška", "Nijedan ID i godina nisu pronađeni u fajlu.")
        return

    # UI state + progress start
    btn_start.configure(state=DISABLED)
    btn_browse.configure(state=DISABLED)
    progress_var.set(0)
    try:
        pbar.configure(mode='indeterminate')
        pbar.start(12)
    except Exception:
        pass
    lbl_status.configure(text="Pokrećem ažuriranje...")
    run_update_async(username, pwd, m)


if __name__ == '__main__':
    base_dir = Path(getattr(sys, '_MEIPASS', Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path.cwd()))
    conf_name = base_dir / 'updatecustomfields.yaml'

    if conf_name.is_file():
        conf = load_conf(conf_name)
    else:
        print('Conf file {} not exists!'.format(conf_name))
        try:
            messagebox.showerror("Greška", f"Nije pronađen konfiguracioni fajl: {conf_name}")
        except Exception:
            pass
        sys.exit(1)

    window = Tk()
    window.title("Update customfields")
    window.geometry("520x260")
    try:
        window.resizable(True, False)
    except Exception:
        pass
    try:
        window.iconbitmap(default='')
    except Exception:
        pass

    # Main frame
    frm = ttk.Frame(window, padding=16)
    frm.pack(fill=BOTH, expand=True)

    # File choose row
    row_file = ttk.Frame(frm)
    row_file.pack(fill=X, pady=8)
    ttk.Label(row_file, text="Excel fajl:").pack(side=LEFT)
    lbl_file = ttk.Label(row_file, text="(Nije izabran)", foreground="#666")
    lbl_file.pack(side=LEFT, padx=8)
    btn_browse = ttk.Button(row_file, text="Izaberi", command=choosefile)
    btn_browse.pack(side=RIGHT)

    # Start button
    btn_start = ttk.Button(frm, text="Pokreni ažuriranje", command=start_update)
    btn_start.pack(anchor=E, pady=(16, 24), padx=(0, 8))
    try:
        btn_start.configure(width=24)
    except Exception:
        pass

    # Progress bar and status
    progress_var = IntVar(value=0)
    pbar = ttk.Progressbar(frm, orient=HORIZONTAL, length=480, mode='determinate', maximum=100, variable=progress_var)
    pbar.pack(pady=(8, 16))
    lbl_status = ttk.Label(frm, text="")
    lbl_status.pack(anchor=W, pady=(8, 0))

    window.mainloop()
