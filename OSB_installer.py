import os
import threading
import subprocess
import win32gui
import win32con
import shutil
import zipfile
import re
import winreg
import tkinter as tk
import string
from ctypes import windll
from tkinter import ttk, filedialog, messagebox
from urllib.request import urlopen
from io import BytesIO

# Windows registry access for Steam path detection
try:
    import winreg
except ImportError:
    winreg = None

# Constants
STEAMCMD_URL      = "https://steamcdn-a.akamaihd.net/client/installer/steamcmd.zip"
STARBOUND_APP_ID  = "211820"
WORKSHOP_MOD_IDS  = ["3534616750"]
NIGHTLY_URL       = (
    "https://nightly.link/OpenStarbound/OpenStarbound/"
    "workflows/build/main/OpenStarbound-Windows-Client.zip"
)

def list_drives():
    """Return a list of all mounted drive letters (e.g. ['C:\\','D:\\',…])."""
    drives = []
    mask = windll.kernel32.GetLogicalDrives()
    for i, letter in enumerate(string.ascii_uppercase):
        if mask & (1 << i):
            drives.append(f"{letter}:\\")
    return drives

def get_steam_libraries():
    """
    Extract valid Steam library folders that are specifically named 'SteamLibrary'
    and contain 'steamapps/common'.
    """
    found_paths = set()

    def read_path(hive, flag):
        try:
            key = winreg.OpenKey(hive, r"Software\Valve\Steam", 0,
                                winreg.KEY_READ | flag)
            path, _ = winreg.QueryValueEx(key, "SteamPath")
            return path
        except OSError:
            return None

    roots = set()
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            path = read_path(hive, view)
            if path:
                roots.add(path)

    roots.add(r"C:\Program Files (x86)\Steam")  # fallback

    for root in roots:
        vdf_path = os.path.join(root, "steamapps", "libraryfolders.vdf")
        if not os.path.isfile(vdf_path):
            continue

        with open(vdf_path, encoding="utf-8") as f:
            text = f.read()

        paths = re.findall(r'"path"\s*"([^"]+)"', text)
        for path in paths:
            norm = os.path.normpath(path.replace("\\\\", "\\")).strip()
            if os.path.isdir(norm):
                # Only keep paths inside a folder named 'SteamLibrary'
                if os.path.basename(norm).lower() == "steamlibrary":
                    found_paths.add(norm)

    print("→ Found Steam libraries named 'SteamLibrary':")
    for path in found_paths:
        print("   ", path)

    return sorted(found_paths)

def detect_starbound_install():
    """
    Return the path to the Starbound install if found (including starbound.exe),
    otherwise return empty string.
    """
    paths = []

    for lib in get_steam_libraries():
        candidate = os.path.join(lib, "steamapps", "common", "Starbound")
        exe_path = os.path.join(candidate, "starbound.exe")
        if os.path.isfile(exe_path):
            paths.append(candidate)

    if paths:
        print(f"→ Found Starbound installs: {paths}")
        return paths[0]  # prioritize first one with exe present

    # fallback scan for other drives
    print("→ No SB in registered libraries, scanning all drives for SteamLibrary…")
    for drive in list_drives():
        candidate = os.path.join(drive, "SteamLibrary", "steamapps", "common", "Starbound")
        exe_path = os.path.join(candidate, "starbound.exe")
        if os.path.isfile(exe_path):
            print("→ Found via fallback scan:", candidate)
            return candidate

    print("→ No existing Starbound install detected.")
    return ""

def bring_to_front(root):
    """Bring the Tk window back on top."""
    root.deiconify()
    root.lift()
    root.focus_force()

class InstallerWizard(tk.Tk):
    def __init__(self, libraries, existing_sb):
        super().__init__()
        self.title("Starbound + OpenStarbound Installer")
        self.resizable(False, False)

        # Store detection results
        self.libraries        = libraries
        self.steam_installed  = bool(existing_sb)
        self.steam_dir        = tk.StringVar(value=existing_sb)
        # Default install_dir for SteamCMD (first library + starbound path)
        default_install = ""
        if not self.steam_installed and libraries:
            default_install = os.path.join(
                libraries[0], "steamapps", "common", "Starbound"
            )
        self.install_dir      = tk.StringVar(value=default_install)
        self.osb_dir          = tk.StringVar(
            value=os.path.join(os.getcwd(), "OpenStarbound")
        )
        self.run_when_done    = tk.BooleanVar(value=True)

        # Prepare wizard frames
        self.frames = {}
        for Frame in (StepPaths, StepInstall, StepFinish):
            page = Frame(self)
            self.frames[Frame] = page
            page.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StepPaths)

    def show_frame(self, frame_cls):
        self.frames[frame_cls].tkraise()

class StepPaths(tk.Frame):
    def __init__(self, master):
        super().__init__(master, padx=10, pady=10)
        tk.Label(self,
                text="Step 1: Locate or Install Starbound",
                font=("Segoe UI", 12, "bold"))\
        .grid(columnspan=3, pady=(0,10))

        if master.steam_installed:
            # Only show existing install path
            tk.Label(self, text="Existing Starbound Install:")\
            .grid(row=1, column=0, sticky="e")
            tk.Entry(self, width=40, textvariable=master.steam_dir)\
            .grid(row=1, column=1)
            tk.Button(self, text="Browse…",
                    command=lambda: self.browse(master.steam_dir))\
            .grid(row=1, column=2)
        else:
            # Only show install location field
            tk.Label(self, text="Install Starbound Here:")\
            .grid(row=1, column=0, sticky="e")
            tk.Entry(self, width=40, textvariable=master.install_dir)\
            .grid(row=1, column=1)
            tk.Button(self, text="Browse…",
                    command=lambda: self.browse(master.install_dir))\
            .grid(row=1, column=2)

        # Always show OSB install path
        tk.Label(self, text="OpenStarbound Install Dir:")\
        .grid(row=2, column=0, sticky="e", pady=10)
        tk.Entry(self, width=40, textvariable=master.osb_dir)\
        .grid(row=2, column=1)
        tk.Button(self, text="Browse…",
                command=lambda: self.browse(master.osb_dir))\
        .grid(row=2, column=2)

        tk.Button(self, text="Next →", width=10,
                command=self.validate)\
        .grid(row=3, column=2, pady=15)

    def browse(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def validate(self):
        m = self.master
        # Validate required fields
        if m.steam_installed:
            if not m.steam_dir.get().strip():
                return messagebox.showerror(
                    "Error", "Please specify existing Starbound path."
                )
        else:
            if not m.install_dir.get().strip():
                return messagebox.showerror(
                    "Error", "Please specify where to install Starbound."
                )
        if not m.osb_dir.get().strip():
            return messagebox.showerror(
                "Error", "Please specify OpenStarbound install directory."
            )
        m.show_frame(StepInstall)
        m.frames[StepInstall].start_install()

def minimize_steam_window():
    def enum_windows_callback(hwnd, result):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "Steam" in title:
                result.append(hwnd)
    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    for hwnd in windows:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        except Exception as e:
            print(f"Could not minimize window {hwnd}: {e}")


class StepInstall(tk.Frame):
    def __init__(self, master):
        super().__init__(master, padx=10, pady=10)
        tk.Label(self,
                text="Step 2: Installing…",
                font=("Segoe UI", 12, "bold"))\
        .pack(anchor="w")

        self.progress = ttk.Progressbar(self, length=500, mode="determinate")
        self.progress.pack(pady=(5,10))

        self.log = tk.Text(self, width=70, height=15,
                        bg="#f9f9f9", state="disabled")
        self.log.pack()

        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", pady=10)
        tk.Button(btn_frame, text="← Back",
                command=lambda: master.show_frame(StepPaths))\
        .pack(side="left")
        self.next_btn = tk.Button(btn_frame, text="Next →",
                                state="disabled",
                                command=lambda: master.show_frame(StepFinish))
        self.next_btn.pack(side="right")

    def log_write(self, txt):
        self.log.config(state="normal")
        self.log.insert("end", txt + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def start_install(self):
        threading.Thread(target=self._install, daemon=True).start()

    def _install(self):
        steps = [
            ("Ensure Steam is running", self._step_steam),
            ("Minimize Steam window", lambda: minimize_steam_window()),
            ("Download SteamCMD",       self._step_steamcmd),
            ("Install Starbound",       self._step_starbound),
            ("Download and run OSB installer", self._step_installer_release),
            ("Move OSB files to install folder", self._step_merge_osb_output),
            ("Copy Steam assets", self._step_assets),
            ("Download mods", self._step_mods),
            ("Final OSB file copy & cleanup", self._step_final_osb_copy),
        ]
        self.progress["maximum"] = len(steps)

        for i, (desc, func) in enumerate(steps, start=1):
            self.log_write(f"→ {desc}…")
            try:
                func()
                self.log_write(f"✔ {desc}")
            except Exception as e:
                self.log_write(f"✘ {desc}: {e}")
                messagebox.showerror("Install Error", f"{desc} failed:\n{e}")
                return
            self.progress["value"] = i

        self.next_btn.config(state="normal")

    def _step_steam(self):
        # Check if Steam.exe is running
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq Steam.exe"],
            stderr=subprocess.DEVNULL, text=True
        )
        if "Steam.exe" not in out:
            # Launch Steam
            steam_root = ""
            if winreg:
                try:
                    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                        r"Software\Valve\Steam")
                    steam_root, _ = winreg.QueryValueEx(key, "SteamPath")
                except Exception:
                    pass
            steam_exe = os.path.join(steam_root, "Steam.exe")
            if os.path.isfile(steam_exe):
                subprocess.Popen([steam_exe],
                                stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL)
                # Brief pause to let Steam start
                import time; time.sleep(2)
            else:
                raise FileNotFoundError("Steam.exe not found.")
        # Bring installer back to front
        bring_to_front(self.master)

    def _step_steamcmd(self):
        dest = os.path.join(os.getcwd(), "steamcmd")
        exe  = os.path.join(dest, "steamcmd.exe")
        if not os.path.isfile(exe):
            os.makedirs(dest, exist_ok=True)
            data = urlopen(STEAMCMD_URL).read()
            z = zipfile.ZipFile(BytesIO(data))
            z.extractall(dest)

    def _step_starbound(self):
        m = self.master
        sb_dir = m.steam_dir.get().strip()
        exe = os.path.join(sb_dir, "starbound.exe")

        if os.path.isfile(exe):
            self.log_write(f"  → Found existing Starbound at: {sb_dir}")
            return  # skip install

        # Otherwise, install using SteamCMD
        sb_dir = m.install_dir.get().strip()
        m.steam_dir.set(sb_dir)  # update in GUI too

        self.log_write(f"  → Installing Starbound to: {sb_dir}")
        cmd = [
            os.path.join(os.getcwd(), "steamcmd", "steamcmd.exe"),
            "+force_install_dir", sb_dir,
            "+login", "anonymous",
            "+app_update", STARBOUND_APP_ID, "validate",
            "+quit"
        ]
        subprocess.check_call(cmd, stdout=subprocess.DEVNULL)

    def _step_installer_release(self):
        import urllib.request
        from urllib.error import URLError

        # Step 1: Resolve tag like v0.1.14
        latest_url = "https://github.com/OpenStarbound/OpenStarbound/releases/latest"
        req = urllib.request.Request(latest_url, method="HEAD")
        with urllib.request.urlopen(req) as resp:
            final_url = resp.geturl()

        match = re.search(r"/tag/(v[\d\.]+)$", final_url)
        if not match:
            raise ValueError(f"Could not extract release tag from: {final_url}")
        tag = match.group(1)
        self.log_write(f"→ Latest OSB release: {tag}")

        # Step 2: Build installer zip URL
        installer_url = (
            f"https://github.com/OpenStarbound/OpenStarbound/releases/download/"
            f"{tag}/OpenStarbound-Windows-Installer.zip"
        )
        self.log_write("→ Downloading OSB installer ZIP…")

        # Step 3: Download and unzip
        data = urlopen(installer_url).read()
        temp_dir = os.path.join(os.getcwd(), "osb_installer_temp")
        if os.path.isdir(temp_dir):
            shutil.rmtree(temp_dir)
        os.makedirs(temp_dir, exist_ok=True)

        with zipfile.ZipFile(BytesIO(data)) as z:
            z.extractall(temp_dir)

        # Step 4: Run installer (find .exe)
        exe = None
        for file in os.listdir(temp_dir):
            if file.endswith(".exe"):
                exe = os.path.join(temp_dir, file)
                break

        if not exe or not os.path.isfile(exe):
            raise FileNotFoundError("Installer .exe not found in ZIP")

        self.log_write(f"→ Running installer: {exe}")

        # Step 5: Run with /NOICONS to avoid desktop shortcut
        subprocess.Popen([exe, "/VERYSILENT", "/NOICONS", "/NORESTART"])

    def _step_merge_osb_output(self):
        import time

        m = self.master
        osb_src = r"C:\Program Files\OpenStarbound"
        osb_dst = m.osb_dir.get()

        self.log_write("→ Waiting for OSB installer to finish...")
        timeout = 30
        while not os.path.isdir(osb_src) and timeout > 0:
            time.sleep(1)
            timeout -= 1

        if not os.path.isdir(osb_src):
            raise FileNotFoundError("OSB output not found at C:\\Program Files\\OpenStarbound")

        # Wait a little more to let all file ops finish
        time.sleep(2)

        self.log_write(f"→ Merging all files from {osb_src} → {osb_dst} (overwrite enabled)")
        failed = []
        for root, dirs, files in os.walk(osb_src):
            rel_path = os.path.relpath(root, osb_src)
            dst_dir = os.path.join(osb_dst, rel_path) if rel_path != "." else osb_dst

            os.makedirs(dst_dir, exist_ok=True)

            for f in files:
                # Skip known temporary files that may disappear
                if f.startswith("is-") and f.endswith(".tmp"):
                    continue

                src_file = os.path.join(root, f)
                dst_file = os.path.join(dst_dir, f)

                try:
                    # Always replace the file, even if it exists
                    if os.path.exists(dst_file):
                        os.remove(dst_file)
                    shutil.copy2(src_file, dst_file)
                except Exception as e:
                    failed.append((src_file, dst_file, str(e)))

        if failed:
            self.log_write("⚠ Some files could not be copied:")
            for src, dst, err in failed:
                self.log_write(f"   {src} → {dst} | {err}")

        # Attempt to delete installer folder
        self.log_write(f"→ Cleaning up {osb_src}")
        try:
            shutil.rmtree(osb_src)
        except Exception as e:
            self.log_write(f"⚠ Could not delete installer output: {e}")

    def _step_assets(self):
        m   = self.master
        src = os.path.join(m.steam_dir.get(), "assets")
        dst = os.path.join(m.osb_dir.get(), "assets")
        if os.path.isdir(dst):
            shutil.rmtree(dst)
        shutil.copytree(src, dst)

    def _step_final_osb_copy(self):
        osb_src = r"C:\Program Files\OpenStarbound"
        osb_dst = self.master.osb_dir.get()

        if not os.path.isdir(osb_src):
            return  # nothing left

        self.log_write("→ Final pass: copying any remaining OSB files…")
        for root, dirs, files in os.walk(osb_src):
            rel_path = os.path.relpath(root, osb_src)
            dst_dir = os.path.join(osb_dst, rel_path) if rel_path != "." else osb_dst
            os.makedirs(dst_dir, exist_ok=True)

            for f in files:
                # Avoid temp/locked files but copy everything else
                if f.startswith("is-") and f.endswith(".tmp"):
                    continue

                src_file = os.path.join(root, f)
                dst_file = os.path.join(dst_dir, f)

                try:
                    if os.path.exists(dst_file):
                        os.remove(dst_file)
                    shutil.copy2(src_file, dst_file)
                except Exception as e:
                    self.log_write(f"⚠ Could not copy {f}: {e}")

        # Try to delete the leftover Program Files folder one last time
        try:
            shutil.rmtree(osb_src)
            self.log_write("→ Final cleanup: Program Files folder removed.")
        except Exception as e:
            self.log_write(f"⚠ Could not delete Program Files folder: {e}")

    def _step_mods(self):
        exe = os.path.join(os.getcwd(), "steamcmd", "steamcmd.exe")
        for mod in WORKSHOP_MOD_IDS:
            subprocess.check_call([
                exe,
                "+login", "anonymous",
                "+workshop_download_item", STARBOUND_APP_ID, mod,
                "+quit"
            ], stdout=subprocess.DEVNULL)

class StepFinish(tk.Frame):
    def __init__(self, master):
        super().__init__(master, padx=10, pady=10)
        tk.Label(self, text="Step 3: Done!",
                font=("Segoe UI", 12, "bold"))\
        .pack(pady=(0,10))

        tk.Checkbutton(self,
                    text="Run Starbound now",
                    variable=master.run_when_done)\
        .pack()

        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", pady=15)
        tk.Button(btn_frame, text="← Back",
                command=lambda: master.show_frame(StepInstall))\
        .pack(side="left")
        tk.Button(btn_frame, text="Finish", width=10,
                command=self.finish)\
        .pack(side="right")

    def finish(self):
        if self.master.run_when_done.get():
            exe = os.path.join(self.master.osb_dir.get(),
                            "win", "starbound.exe")
            if os.path.isfile(exe):
                subprocess.Popen([exe])
            else:
                messagebox.showwarning(
                    "Warning", f"Could not find:\n{exe}"
                )
        self.master.destroy()

if __name__ == "__main__":
    libs     = get_steam_libraries()
    existing = detect_starbound_install()
    app      = InstallerWizard(libs, existing)
    app.mainloop()
