import argparse
import ctypes
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import threading
import time
from ctypes import wintypes
from dataclasses import asdict, dataclass
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Callable, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

import keyring
import tkinter as tk
import winreg
from tkinter import messagebox, scrolledtext, ttk

from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService

try:
    import msvcrt
except Exception:
    msvcrt = None

try:
    import pystray
except Exception:
    pystray = None

try:
    from PIL import Image, ImageDraw, ImageTk
except Exception:
    Image = None
    ImageDraw = None
    ImageTk = None


APP_NAME = "NWAFUAutoLogin"
CREDENTIAL_SERVICE = "NWAFUAutoLogin"
RUN_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
PORTAL_URL = "https://portal.nwafu.edu.cn/srun_portal_success?ac_id=1&theme=pro"
TARGET_SSID = "NWAFU"
CREATE_NO_WINDOW = 0x08000000


class SingleInstanceGuard:
    MUTEX_NAME = f"Local\\{APP_NAME}_SingleInstance"
    ERROR_ALREADY_EXISTS = 183
    LOCKFILE_NAME = "instance.lock"

    def __init__(self) -> None:
        self._handle = None
        self._kernel32 = None
        self.already_running = False
        self._lock_file = None
        self._lock_len = 1
        if os.name != "nt":
            return
        self._init_mutex()
        self._init_file_lock()

    def _init_mutex(self) -> None:
        self._kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        self._kernel32.CreateMutexW.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
        self._kernel32.CreateMutexW.restype = wintypes.HANDLE
        self._kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        self._kernel32.CloseHandle.restype = wintypes.BOOL
        self._handle = self._kernel32.CreateMutexW(None, False, self.MUTEX_NAME)
        if not self._handle:
            raise ctypes.WinError(ctypes.get_last_error())
        if ctypes.get_last_error() == self.ERROR_ALREADY_EXISTS:
            self.already_running = True

    def _init_file_lock(self) -> None:
        if msvcrt is None:
            return
        lock_file = _local_app_data_dir() / APP_NAME / self.LOCKFILE_NAME
        lock_file.parent.mkdir(parents=True, exist_ok=True)
        self._lock_file = open(lock_file, "a+b")
        self._lock_file.seek(0)
        try:
            msvcrt.locking(self._lock_file.fileno(), msvcrt.LK_NBLCK, self._lock_len)
        except OSError:
            self.already_running = True

    def release(self) -> None:
        if self._lock_file:
            try:
                if msvcrt is not None:
                    self._lock_file.seek(0)
                    msvcrt.locking(self._lock_file.fileno(), msvcrt.LK_UNLCK, self._lock_len)
            except OSError:
                pass
            try:
                self._lock_file.close()
            except OSError:
                pass
            self._lock_file = None
        if self._handle and self._kernel32:
            self._kernel32.CloseHandle(self._handle)
            self._handle = None


def _local_app_data_dir() -> Path:
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data)
    return Path.home() / "AppData" / "Local"


def _roaming_app_data_dir() -> Path:
    app_data = os.getenv("APPDATA")
    if app_data:
        return Path(app_data)
    return Path.home() / "AppData" / "Roaming"


def _startup_folder_dir() -> Path:
    return (
        _roaming_app_data_dir()
        / "Microsoft"
        / "Windows"
        / "Start Menu"
        / "Programs"
        / "Startup"
    )


def _hidden_subprocess_kwargs() -> dict:
    kwargs = {}
    if os.name == "nt":
        startup_info = subprocess.STARTUPINFO()
        startup_info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        kwargs["startupinfo"] = startup_info
        kwargs["creationflags"] = CREATE_NO_WINDOW
    return kwargs


def _find_edge_executable() -> Optional[str]:
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    found = shutil.which("msedge.exe")
    return found


def _default_profile_dir() -> str:
    return str(_local_app_data_dir() / APP_NAME / "edge-profile")


@dataclass
class AppConfig:
    target_ssid: str = TARGET_SSID
    portal_url: str = PORTAL_URL
    poll_interval_sec: int = 5
    auth_probe_interval_sec: int = 60
    login_timeout_sec: int = 25
    edge_profile_dir: str = _default_profile_dir()
    keep_browser_on_failure: bool = True
    autostart_enabled: bool = False
    exit_on_close: bool = True
    auto_hide_to_tray_on_success: bool = True


class ConfigManager:
    def __init__(self) -> None:
        self.base_dir = _local_app_data_dir() / APP_NAME
        self.config_file = self.base_dir / "config.json"
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def load(self) -> AppConfig:
        config = AppConfig()
        if not self.config_file.exists():
            return config
        try:
            raw = json.loads(self.config_file.read_text(encoding="utf-8"))
        except Exception:
            return config
        for key, value in raw.items():
            if hasattr(config, key):
                setattr(config, key, value)
        # Lock core routing values to fixed product requirements.
        config.target_ssid = TARGET_SSID
        config.portal_url = PORTAL_URL
        config.poll_interval_sec = max(1, int(config.poll_interval_sec))
        config.auth_probe_interval_sec = max(10, int(config.auth_probe_interval_sec))
        config.login_timeout_sec = max(5, int(config.login_timeout_sec))
        return config

    def save(self, config: AppConfig) -> None:
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self.config_file.write_text(
            json.dumps(asdict(config), ensure_ascii=False, indent=2), encoding="utf-8"
        )


class CredentialStore:
    USERNAME_KEY = "username"
    PASSWORD_KEY = "password"

    def save(self, username: str, password: str) -> None:
        keyring.set_password(CREDENTIAL_SERVICE, self.USERNAME_KEY, username)
        keyring.set_password(CREDENTIAL_SERVICE, self.PASSWORD_KEY, password)

    def load(self) -> Optional[Tuple[str, str]]:
        username = keyring.get_password(CREDENTIAL_SERVICE, self.USERNAME_KEY)
        password = keyring.get_password(CREDENTIAL_SERVICE, self.PASSWORD_KEY)
        if not username or not password:
            return None
        return username, password

    def clear(self) -> None:
        for key in (self.USERNAME_KEY, self.PASSWORD_KEY):
            try:
                keyring.delete_password(CREDENTIAL_SERVICE, key)
            except keyring.errors.PasswordDeleteError:
                pass


class WifiService:
    def get_current_ssid(self) -> Optional[str]:
        try:
            result = subprocess.run(
                ["netsh", "wlan", "show", "interfaces"],
                capture_output=True,
                text=True,
                errors="ignore",
                check=False,
                **_hidden_subprocess_kwargs(),
            )
        except Exception:
            return None

        output = (result.stdout or "") + "\n" + (result.stderr or "")
        for line in output.splitlines():
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            key = key.strip().lower()
            if key == "ssid":
                ssid = value.strip()
                if ssid and ssid.upper() != "N/A":
                    return ssid

        # Fallback for localized labels that still include "SSID" text.
        for line in output.splitlines():
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            key = key.strip().lower()
            if "ssid" in key and "bssid" not in key:
                ssid = value.strip()
                if ssid and ssid.upper() != "N/A":
                    return ssid
        return None


class LoginResult(str, Enum):
    ALREADY_LOGGED_IN = "already_logged_in"
    LOGGED_IN_NOW = "logged_in_now"
    FAILED = "failed"


class PortalAutomator:
    LOGOUT_XPATHS = [
        "//button[contains(normalize-space(.),'注销')]",
        "//a[contains(normalize-space(.),'注销')]",
        "//input[contains(@value,'注销')]",
    ]
    LOGIN_XPATHS = [
        "//button[contains(normalize-space(.),'登录')]",
        "//a[contains(normalize-space(.),'登录')]",
        "//input[contains(@value,'登录')]",
    ]
    PROBE_URL = "http://www.msftconnecttest.com/redirect"
    JS_NAVIGATION_BUDGET_SEC = 15
    LOCAL_EDGE_FALLBACK_COOLDOWN_SEC = 300

    def __init__(self) -> None:
        self._retained_drivers = []
        self._last_local_edge_fallback_at = 0.0

    def ensure_logged_in(
        self,
        config: AppConfig,
        creds: Tuple[str, str],
        logger: Callable[[str], None],
        source: str = "auto",
    ) -> LoginResult:
        username, password = creds
        driver: Optional[webdriver.Edge] = None
        keep_browser_open = False
        try:
            driver = self._create_driver(config)
            page_timeout = max(10, int(config.login_timeout_sec))
            logger("已启动 Edge 自动化窗口。")
            if not self._navigate_to_portal(driver, config.portal_url, page_timeout, logger):
                logger("JS 导航失败，触发本地 Edge 回退。")
                self.open_portal_in_local_edge(
                    config,
                    logger,
                    source=source,
                    reason="selenium_navigation_failed",
                    force=False,
                )
                return LoginResult.FAILED

            if self._wait_for_any_xpath(driver, self.LOGOUT_XPATHS, timeout=4):
                logger("页面已显示“注销”元素，当前已登录。")
                return LoginResult.ALREADY_LOGGED_IN

            login_btn = self._wait_for_any_xpath(driver, self.LOGIN_XPATHS, timeout=8)
            if not login_btn:
                logger("未找到“登录”按钮，无法继续。")
                keep_browser_open = config.keep_browser_on_failure
                return LoginResult.FAILED

            username_input = self._find_username_input(driver)
            password_input = self._find_password_input(driver)
            if not username_input or not password_input:
                logger("未找到账号或密码输入框，无法继续。")
                keep_browser_open = config.keep_browser_on_failure
                return LoginResult.FAILED

            self._set_input(username_input, username)
            self._set_input(password_input, password)
            self._safe_click(driver, login_btn)
            logger("已点击“登录”，等待认证结果。")

            if self._wait_for_any_xpath(
                driver, self.LOGOUT_XPATHS, timeout=config.login_timeout_sec
            ):
                logger("检测到“注销”元素，认证成功。")
                return LoginResult.LOGGED_IN_NOW

            logger("超时未检测到认证成功信号，认证失败。")
            keep_browser_open = config.keep_browser_on_failure
            return LoginResult.FAILED
        except WebDriverException as exc:
            short_msg = str(exc).splitlines()[0] if str(exc) else repr(exc)
            logger(f"Selenium/Edge 运行异常: {short_msg}")
            keep_browser_open = config.keep_browser_on_failure
            return LoginResult.FAILED
        except Exception as exc:
            logger(f"认证流程发生异常: {exc}")
            keep_browser_open = config.keep_browser_on_failure
            return LoginResult.FAILED
        finally:
            if driver:
                if keep_browser_open:
                    self._retained_drivers.append(driver)
                else:
                    try:
                        driver.quit()
                    except Exception:
                        pass

    def cleanup(self) -> None:
        for driver in self._retained_drivers:
            try:
                driver.quit()
            except Exception:
                pass
        self._retained_drivers.clear()

    def _create_driver(self, config: AppConfig) -> webdriver.Edge:
        Path(config.edge_profile_dir).mkdir(parents=True, exist_ok=True)

        options = EdgeOptions()
        options.page_load_strategy = "eager"
        options.add_argument(f"--user-data-dir={config.edge_profile_dir}")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        edge_binary = _find_edge_executable()
        if edge_binary:
            options.binary_location = edge_binary

        service_kwargs = {
            "creation_flags": CREATE_NO_WINDOW,
            "log_output": subprocess.DEVNULL,
        }
        try:
            service = EdgeService(**service_kwargs)
        except TypeError:
            service_kwargs.pop("creation_flags", None)
            try:
                service = EdgeService(**service_kwargs)
            except TypeError:
                service = EdgeService()
        return webdriver.Edge(service=service, options=options)

    @staticmethod
    def _wait_for_any_xpath(driver: webdriver.Edge, xpaths: list, timeout: int):
        deadline = time.time() + max(1, timeout)
        while time.time() < deadline:
            for xpath in xpaths:
                elements = driver.find_elements(By.XPATH, xpath)
                for element in elements:
                    if element.is_displayed():
                        return element
            time.sleep(0.3)
        return None

    @staticmethod
    def _safe_click(driver: webdriver.Edge, element) -> None:
        try:
            element.click()
            return
        except Exception:
            pass
        driver.execute_script("arguments[0].click();", element)

    @staticmethod
    def _set_input(element, value: str) -> None:
        element.clear()
        element.send_keys(value)

    @staticmethod
    def _is_about_blank(url: str) -> bool:
        if not url:
            return True
        return url.strip().lower().startswith("about:blank")

    @staticmethod
    def _safe_current_url(driver: webdriver.Edge) -> str:
        try:
            return str(driver.current_url or "")
        except Exception:
            return ""

    def _wait_until_not_about_blank(
        self, driver: webdriver.Edge, timeout_sec: int
    ) -> Tuple[bool, str]:
        deadline = time.time() + max(1, int(timeout_sec))
        while time.time() < deadline:
            current = self._safe_current_url(driver)
            if not self._is_about_blank(current):
                return True, current
            time.sleep(0.25)
        return False, self._safe_current_url(driver)

    def _navigate_to_portal(
        self,
        driver: webdriver.Edge,
        portal_url: str,
        timeout_sec: int,
        logger: Callable[[str], None],
    ) -> bool:
        nav_budget = max(5, min(int(timeout_sec), self.JS_NAVIGATION_BUDGET_SEC))
        first_stage_timeout = max(2, min(4, nav_budget // 2))
        second_stage_timeout = max(2, nav_budget - first_stage_timeout)

        logger("导航策略: JS 跳转门户。")
        try:
            driver.execute_script("window.location.replace(arguments[0]);", portal_url)
        except Exception as exc:
            logger(f"JS 跳转门户失败: {exc}")
        ok, current = self._wait_until_not_about_blank(driver, first_stage_timeout)
        logger(f"当前页面: {current or '<空>'}")
        if ok:
            logger("门户导航成功: JS 跳转。")
            return True

        logger("JS 跳转后仍为 about:blank，尝试跳转探测页。")
        try:
            driver.execute_script("window.location.replace(arguments[0]);", self.PROBE_URL)
        except Exception as exc:
            logger(f"JS 跳转探测页失败: {exc}")
        ok, current = self._wait_until_not_about_blank(driver, second_stage_timeout)
        logger(f"当前页面: {current or '<空>'}")
        if ok:
            logger("门户导航成功: 探测页跳转。")
            return True

        logger("门户导航失败: 仍为 about:blank。")
        return False

    def open_portal_in_local_edge(
        self,
        config: AppConfig,
        logger: Callable[[str], None],
        source: str = "manual",
        reason: str = "manual_open",
        force: bool = False,
    ) -> bool:
        source_normalized = (source or "auto").strip().lower()
        now = time.time()
        if source_normalized == "auto" and not force:
            cooldown_left = int(
                self.LOCAL_EDGE_FALLBACK_COOLDOWN_SEC
                - (now - self._last_local_edge_fallback_at)
            )
            if cooldown_left > 0:
                logger(f"本地 Edge 回退跳过: 冷却中，剩余 {cooldown_left} 秒。")
                return False

        edge_path = _find_edge_executable()
        if not edge_path:
            logger("本地 Edge 回退失败: 未找到 Edge 可执行文件。")
            return False

        Path(config.edge_profile_dir).mkdir(parents=True, exist_ok=True)
        target_url = self.PROBE_URL if source_normalized == "auto" else config.portal_url
        command = [
            edge_path,
            f"--user-data-dir={config.edge_profile_dir}",
            target_url,
        ]
        try:
            subprocess.Popen(command, **_hidden_subprocess_kwargs())
            if source_normalized == "auto":
                self._last_local_edge_fallback_at = now
            logger(
                f"已触发本地 Edge 回退: source={source_normalized}, reason={reason}, url={target_url}。"
            )
            return True
        except Exception as exc:
            logger(f"本地 Edge 回退失败: {exc}")
            return False

    @staticmethod
    def _field_score(element, user_field: bool) -> int:
        attrs = " ".join(
            [
                element.get_attribute("id") or "",
                element.get_attribute("name") or "",
                element.get_attribute("placeholder") or "",
                element.get_attribute("class") or "",
            ]
        ).lower()
        field_type = (element.get_attribute("type") or "").lower()
        score = 0
        if user_field:
            keywords = ["user", "username", "account", "name", "账号", "学号"]
            if any(token in attrs for token in keywords):
                score += 3
            if field_type in ("text", "email", "tel", "number", ""):
                score += 1
            if field_type == "password":
                score -= 10
        else:
            keywords = ["pass", "password", "pwd", "密码"]
            if any(token in attrs for token in keywords):
                score += 3
            if field_type == "password":
                score += 2
        return score

    def _find_username_input(self, driver: webdriver.Edge):
        candidates = driver.find_elements(By.XPATH, "//input[not(@disabled)]")
        best = None
        best_score = -999
        for element in candidates:
            score = self._field_score(element, user_field=True)
            if score > best_score:
                best = element
                best_score = score
        return best if best_score >= 1 else None

    def _find_password_input(self, driver: webdriver.Edge):
        candidates = driver.find_elements(By.XPATH, "//input[not(@disabled)]")
        best = None
        best_score = -999
        for element in candidates:
            score = self._field_score(element, user_field=False)
            if score > best_score:
                best = element
                best_score = score
        return best if best_score >= 1 else None


class AutostartService:
    VBS_FILENAME = f"{APP_NAME}.vbs"
    RUN_PATTERN = re.compile(r'WshShell\.Run\s+"((?:""|[^"])*)"\s*,\s*0\s*,\s*False', re.IGNORECASE)

    def __init__(self) -> None:
        self.startup_script = _startup_folder_dir() / self.VBS_FILENAME

    @staticmethod
    def _normalize_command(value: Optional[str]) -> Optional[str]:
        command = str(value or "").strip()
        return command or None

    def get_run_command(self) -> Optional[str]:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_REG_KEY, 0, winreg.KEY_READ) as key:
                value, _ = winreg.QueryValueEx(key, APP_NAME)
                return self._normalize_command(value)
        except FileNotFoundError:
            return None
        except OSError:
            return None

    def get_startup_folder_command(self) -> Optional[str]:
        if not self.startup_script.exists():
            return None
        content = None
        for encoding in ("utf-16", "utf-8", "mbcs"):
            try:
                content = self.startup_script.read_text(encoding=encoding, errors="ignore")
                break
            except OSError:
                return None
            except UnicodeError:
                continue
        if content is None:
            return None
        match = self.RUN_PATTERN.search(content)
        if not match:
            return None
        command = match.group(1).replace('""', '"').strip()
        return command or None

    def get_command(self) -> Optional[str]:
        return self.get_run_command() or self.get_startup_folder_command()

    def get_channel_commands(self) -> Tuple[Optional[str], Optional[str]]:
        return self.get_run_command(), self.get_startup_folder_command()

    def is_enabled(self) -> bool:
        run_command, startup_command = self.get_channel_commands()
        return bool(run_command or startup_command)

    @staticmethod
    def _build_vbs_content(command: str) -> str:
        escaped_command = command.replace('"', '""')
        return (
            'Set WshShell = CreateObject("WScript.Shell")\n'
            f'WshShell.Run "{escaped_command}", 0, False\n'
        )

    def set_run_command(self, command: str) -> None:
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, RUN_REG_KEY) as key:
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, command)
        except OSError as exc:
            raise OSError(f"写入注册表启动项失败: {exc}") from exc

    def set_startup_folder_command(self, command: str) -> None:
        try:
            self.startup_script.parent.mkdir(parents=True, exist_ok=True)
            # Use UTF-16 so Windows Script Host reliably handles non-ASCII paths.
            self.startup_script.write_text(self._build_vbs_content(command), encoding="utf-16")
        except OSError as exc:
            raise OSError(f"写入启动文件夹脚本失败: {exc}") from exc

    def _enable_run_preferred(self, command: str) -> str:
        run_error = None
        try:
            self.set_run_command(command)
            # Keep a single startup channel to avoid duplicate launches on boot.
            self.disable_startup_folder()
            return "run"
        except OSError as exc:
            run_error = exc

        try:
            self.set_startup_folder_command(command)
            self.disable_run()
            return "startup_folder"
        except OSError as exc:
            raise OSError(
                f"写入开机启动项失败: run={run_error}; startup_folder={exc}"
            ) from exc

    def enable(self, command: str) -> str:
        return self._enable_run_preferred(command)

    def disable_run(self) -> None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_REG_KEY, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, APP_NAME)
        except FileNotFoundError:
            pass
        except OSError:
            pass

    def disable_startup_folder(self) -> None:
        try:
            self.startup_script.unlink()
        except FileNotFoundError:
            pass
        except OSError:
            pass

    def disable(self) -> None:
        self.disable_run()
        self.disable_startup_folder()


def build_autostart_command() -> str:
    if getattr(sys, "frozen", False):
        exe_path = Path(os.path.abspath(sys.executable))
        # Prefer nearby onedir executable when both onefile and onedir artifacts exist.
        candidates = [
            exe_path.with_suffix("") / exe_path.name,
            exe_path.parent / APP_NAME / exe_path.name,
        ]
        for candidate in candidates:
            try:
                if candidate != exe_path and candidate.is_file():
                    exe_path = candidate
                    break
            except OSError:
                continue
        return f'"{exe_path}" --autostart --minimized'

    def is_usable_file(path: Optional[str]) -> bool:
        if not path:
            return False
        try:
            return os.path.isfile(path)
        except OSError:
            return False

    python_exec = str(sys.executable or "")
    pythonw_exec = ""
    lower_exec = python_exec.lower()
    if lower_exec.endswith("pythonw.exe") and is_usable_file(python_exec):
        pythonw_exec = python_exec
    elif lower_exec.endswith("python.exe"):
        candidate = python_exec[:-10] + "pythonw.exe"
        if is_usable_file(candidate):
            pythonw_exec = candidate

    if not pythonw_exec:
        from_path = shutil.which("pythonw.exe")
        pythonw_exec = from_path if from_path else "pythonw"

    script_path = os.path.abspath(__file__)
    return f'"{pythonw_exec}" "{script_path}" --autostart --minimized'


class AppIconFactory:
    @staticmethod
    def available() -> bool:
        return Image is not None and ImageDraw is not None

    @classmethod
    def build(cls, size: int = 64):
        if not cls.available():
            return None
        side = max(16, int(size))
        image = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)

        pad = max(1, int(side * 0.15))
        radius = max(2, int(side * 0.24))
        bg_color = (22, 128, 112, 255)
        border_color = (255, 255, 255, 72)
        white = (255, 255, 255, 255)
        accent = (255, 196, 66, 255)

        left = pad
        top = pad
        right = side - pad - 1
        bottom = side - pad - 1
        draw.rounded_rectangle((left, top, right, bottom), radius=radius, fill=bg_color)
        draw.rounded_rectangle(
            (left, top, right, bottom),
            radius=radius,
            outline=border_color,
            width=max(1, side // 24),
        )

        stroke = max(2, int(side * 0.15))
        n_left = int(side * 0.33)
        n_right = int(side * 0.67)
        n_top = int(side * 0.29)
        n_bottom = int(side * 0.71)
        draw.line(
            [(n_left, n_bottom), (n_left, n_top), (n_right, n_bottom), (n_right, n_top)],
            fill=white,
            width=stroke,
            joint="curve",
        )

        dot_r = max(2, int(side * 0.08))
        dot_cx = int(side * 0.72)
        dot_cy = int(side * 0.27)
        draw.ellipse(
            (dot_cx - dot_r, dot_cy - dot_r, dot_cx + dot_r, dot_cy + dot_r),
            fill=accent,
        )
        return image


class TrayController:
    def __init__(
        self,
        on_show: Callable[[], None],
        on_exit: Callable[[], None],
        logger: Callable[[str], None],
    ) -> None:
        self._on_show = on_show
        self._on_exit = on_exit
        self._logger = logger
        self._icon = None
        self._thread: Optional[threading.Thread] = None

    @property
    def available(self) -> bool:
        return pystray is not None and Image is not None and ImageDraw is not None

    def start(self) -> bool:
        if not self.available:
            return False
        if self._icon:
            return True
        try:
            icon_image = AppIconFactory.build(64)
            if icon_image is None:
                return False
            menu = pystray.Menu(
                pystray.MenuItem("显示主窗口", self._menu_show),
                pystray.MenuItem("退出程序", self._menu_exit),
            )
            self._icon = pystray.Icon(
                APP_NAME,
                icon_image,
                "NWAFU 自动认证",
                menu,
            )
            self._thread = threading.Thread(target=self._icon.run, daemon=True)
            self._thread.start()
            return True
        except Exception as exc:
            self._logger(f"创建托盘图标失败: {exc}")
            self._icon = None
            self._thread = None
            return False

    def stop(self) -> None:
        icon = self._icon
        self._icon = None
        if not icon:
            return
        try:
            icon.stop()
        except Exception as exc:
            self._logger(f"关闭托盘图标失败: {exc}")
        self._thread = None

    def _menu_show(self, icon, item) -> None:
        self._on_show()

    def _menu_exit(self, icon, item) -> None:
        self._on_exit()


class NWAFUGuiApp:
    AUTO_RETRY_SECONDS = 60
    CRED_ALERT_COOLDOWN_SECONDS = 300
    AUTH_PROBE_TIMEOUT_SECONDS = 8
    AUTH_PROBE_HTTP_URL = "http://www.msftconnecttest.com/connecttest.txt"
    AUTH_PROBE_HTTPS_URL = "https://www.msftconnecttest.com/connecttest.txt"
    AUTH_PROBE_EXPECTED_TEXT = "Microsoft Connect Test"
    AUTO_START_FAST_DELAY_MS = 500
    AUTO_START_BOOT_DELAY_MS = 30000
    STARTUP_LOG_MAX_BYTES = 1024 * 1024

    def __init__(self, root: tk.Tk, launch_autostart: bool = False, minimized: bool = False):
        self.root = root
        self.root.title("NWAFU 自动认证")
        self.root.geometry("860x670")
        self.root.minsize(760, 600)
        self._window_icon_refs = []

        self.config_manager = ConfigManager()
        self.config = self.config_manager.load()
        self.startup_log_path = self.config_manager.base_dir / "startup.log"
        self.startup_log_lock = threading.Lock()
        self.credential_store = CredentialStore()
        self.wifi_service = WifiService()
        self.portal_automator = PortalAutomator()
        self.autostart_service = AutostartService()
        self.tray_controller = TrayController(
            on_show=lambda: self._run_on_ui(self._restore_from_tray),
            on_exit=lambda: self._run_on_ui(self.exit_application),
            logger=lambda message: self._run_on_ui(self._log, message),
        )
        self.window_hidden_to_tray = False
        self.tray_fallback_logged = False

        self.monitor_thread: Optional[threading.Thread] = None
        self.monitor_stop_event = threading.Event()
        self.monitor_running = False
        self.check_lock = threading.Lock()

        self.last_seen_target = False
        self.session_authenticated = False
        self.next_auto_retry_at = 0.0
        self.last_cred_alert_at = 0.0
        self.last_auth_probe_at = 0.0

        self.status_var = tk.StringVar(value="未运行")
        self.current_ssid_var = tk.StringVar(value="未知")
        self.last_result_var = tk.StringVar(value="暂无")
        self.last_run_time_var = tk.StringVar(value="暂无")
        self.poll_interval_var = tk.StringVar(value=str(self.config.poll_interval_sec))
        self.auth_probe_interval_var = tk.StringVar(value=str(self.config.auth_probe_interval_sec))
        autostart_enabled = bool(self.config.autostart_enabled or self.autostart_service.is_enabled())
        self.autostart_var = tk.BooleanVar(value=autostart_enabled)
        self.exit_on_close_var = tk.BooleanVar(value=bool(self.config.exit_on_close))
        self.auto_hide_to_tray_var = tk.BooleanVar(
            value=bool(self.config.auto_hide_to_tray_on_success)
        )

        self.username_var = tk.StringVar(value="")
        self.password_var = tk.StringVar(value="")

        self._build_ui()
        self._apply_window_icon()
        self._load_credentials_to_form()
        self._log(f"启动参数: --autostart={launch_autostart}, --minimized={minimized}")
        before_state = self._collect_autostart_channel_state()
        self._startup_log(
            f"startup_args: autostart={launch_autostart}, minimized={minimized}, "
            f"runtime_mode={'frozen_exe' if getattr(sys, 'frozen', False) else 'script'}, "
            f"executable={os.path.abspath(sys.executable)}"
        )
        self._record_autostart_channel_state("before_integrity_check", before_state)
        repaired = self._ensure_autostart_command_integrity()
        after_state = self._collect_autostart_channel_state()
        self._record_autostart_channel_state("after_integrity_check", after_state)
        self._startup_log(f"autostart_repair_applied={repaired}")
        self._log("程序已启动。")

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        if minimized:
            self.root.after(300, lambda: self._hide_to_tray("启动参数"))
        auto_by_config = bool(self.autostart_var.get() or self.config.autostart_enabled)
        should_auto_start = launch_autostart or auto_by_config
        if should_auto_start:
            reasons = []
            if launch_autostart:
                reasons.append("参数触发")
            if auto_by_config:
                reasons.append("配置触发")
            self._log(f"自动启动监控触发来源: {' + '.join(reasons)}")
            delay_ms = self.AUTO_START_BOOT_DELAY_MS if launch_autostart else self.AUTO_START_FAST_DELAY_MS
            self._log(f"自动监控将在 {delay_ms}ms 后启动。")
            self._startup_log(
                f"monitor_schedule: should_auto_start={should_auto_start}, launch_autostart={launch_autostart}, "
                f"auto_by_config={auto_by_config}, delay_ms={delay_ms}, reasons={'+'.join(reasons)}"
            )
            self.root.after(delay_ms, self.start_monitoring)
        else:
            self._log("自动启动监控触发来源: 无")
            self._startup_log("monitor_schedule: should_auto_start=False")

    def _apply_window_icon(self) -> None:
        if ImageTk is None:
            return
        try:
            icon_refs = []
            for size in (16, 24, 32, 48, 64):
                icon = AppIconFactory.build(size)
                if icon is None:
                    return
                icon_refs.append(ImageTk.PhotoImage(icon))
            if icon_refs:
                self._window_icon_refs = icon_refs
                self.root.iconphoto(True, *icon_refs)
        except Exception:
            pass

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)

        status_frame = ttk.LabelFrame(outer, text="状态")
        status_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(status_frame, text="监控状态:").grid(row=0, column=0, sticky=tk.W, padx=8, pady=6)
        ttk.Label(status_frame, textvariable=self.status_var).grid(row=0, column=1, sticky=tk.W, padx=8, pady=6)
        ttk.Label(status_frame, text="当前 SSID:").grid(row=0, column=2, sticky=tk.W, padx=8, pady=6)
        ttk.Label(status_frame, textvariable=self.current_ssid_var).grid(row=0, column=3, sticky=tk.W, padx=8, pady=6)

        ttk.Label(status_frame, text="最近结果:").grid(row=1, column=0, sticky=tk.W, padx=8, pady=6)
        ttk.Label(status_frame, textvariable=self.last_result_var).grid(row=1, column=1, sticky=tk.W, padx=8, pady=6)
        ttk.Label(status_frame, text="最近执行:").grid(row=1, column=2, sticky=tk.W, padx=8, pady=6)
        ttk.Label(status_frame, textvariable=self.last_run_time_var).grid(row=1, column=3, sticky=tk.W, padx=8, pady=6)

        control_frame = ttk.LabelFrame(outer, text="控制")
        control_frame.pack(fill=tk.X, pady=(0, 8))

        self.start_btn = ttk.Button(control_frame, text="开始监控", command=self.start_monitoring)
        self.stop_btn = ttk.Button(control_frame, text="停止监控", command=self.stop_monitoring)
        self.check_btn = ttk.Button(control_frame, text="立即检测", command=self.manual_check)
        self.open_btn = ttk.Button(control_frame, text="打开认证页", command=self.open_portal_page)
        self.to_tray_btn = ttk.Button(control_frame, text="收至托盘", command=self.hide_to_tray_manually)
        self.exit_btn = ttk.Button(control_frame, text="退出程序", command=self.exit_application)

        self.start_btn.grid(row=0, column=0, padx=8, pady=8)
        self.stop_btn.grid(row=0, column=1, padx=8, pady=8)
        self.check_btn.grid(row=0, column=2, padx=8, pady=8)
        self.open_btn.grid(row=0, column=3, padx=8, pady=8)
        self.to_tray_btn.grid(row=0, column=4, padx=8, pady=8)
        self.exit_btn.grid(row=0, column=5, padx=8, pady=8)
        self.stop_btn.configure(state=tk.DISABLED)

        cred_frame = ttk.LabelFrame(outer, text="凭据")
        cred_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(cred_frame, text="账号:").grid(row=0, column=0, padx=8, pady=8, sticky=tk.W)
        ttk.Entry(cred_frame, textvariable=self.username_var, width=30).grid(
            row=0, column=1, padx=8, pady=8, sticky=tk.W
        )
        ttk.Label(cred_frame, text="密码:").grid(row=0, column=2, padx=8, pady=8, sticky=tk.W)
        ttk.Entry(cred_frame, textvariable=self.password_var, width=30, show="*").grid(
            row=0, column=3, padx=8, pady=8, sticky=tk.W
        )

        ttk.Button(cred_frame, text="保存凭据", command=self.save_credentials).grid(
            row=1, column=1, padx=8, pady=8, sticky=tk.W
        )
        ttk.Button(cred_frame, text="清除凭据", command=self.clear_credentials).grid(
            row=1, column=3, padx=8, pady=8, sticky=tk.W
        )

        setting_frame = ttk.LabelFrame(outer, text="设置")
        setting_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(setting_frame, text="轮询间隔(秒):").grid(row=0, column=0, padx=8, pady=8, sticky=tk.W)
        ttk.Spinbox(setting_frame, from_=1, to=120, textvariable=self.poll_interval_var, width=8).grid(
            row=0, column=1, padx=8, pady=8, sticky=tk.W
        )
        ttk.Label(setting_frame, text="认证保活检测间隔(秒):").grid(
            row=1, column=0, padx=8, pady=8, sticky=tk.W
        )
        ttk.Spinbox(
            setting_frame,
            from_=10,
            to=600,
            textvariable=self.auth_probe_interval_var,
            width=8,
        ).grid(row=1, column=1, padx=8, pady=8, sticky=tk.W)
        ttk.Button(setting_frame, text="保存设置", command=self.save_settings).grid(
            row=0, column=2, padx=8, pady=8, sticky=tk.W
        )

        ttk.Checkbutton(
            setting_frame,
            text="开机自启",
            variable=self.autostart_var,
            command=self.apply_autostart,
        ).grid(row=0, column=3, padx=12, pady=8, sticky=tk.W)
        ttk.Checkbutton(
            setting_frame,
            text="点 X 时退出程序",
            variable=self.exit_on_close_var,
            command=self.apply_close_behavior,
        ).grid(row=1, column=3, padx=12, pady=8, sticky=tk.W)
        ttk.Checkbutton(
            setting_frame,
            text="认证成功后自动收至托盘",
            variable=self.auto_hide_to_tray_var,
            command=self.apply_auto_hide_to_tray,
        ).grid(row=2, column=3, padx=12, pady=8, sticky=tk.W)

        hint_color = "#666666"
        ttk.Label(
            setting_frame,
            text="轮询间隔用于检测是否连接到 NWAFU，以及是否需要触发登录。",
            foreground=hint_color,
        ).grid(row=3, column=0, columnspan=4, padx=8, pady=(2, 0), sticky=tk.W)
        ttk.Label(
            setting_frame,
            text="建议 3~10 秒：越小响应越快，但检测更频繁。",
            foreground=hint_color,
        ).grid(row=4, column=0, columnspan=4, padx=8, pady=(0, 2), sticky=tk.W)
        ttk.Label(
            setting_frame,
            text="认证保活检测间隔仅在已登录后生效，用于确认认证状态是否仍有效。",
            foreground=hint_color,
        ).grid(row=5, column=0, columnspan=4, padx=8, pady=(2, 0), sticky=tk.W)
        ttk.Label(
            setting_frame,
            text="建议 30~120 秒：越小恢复越快，但网络检测更频繁。",
            foreground=hint_color,
        ).grid(row=6, column=0, columnspan=4, padx=8, pady=(0, 6), sticky=tk.W)

        log_frame = ttk.LabelFrame(outer, text="日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=16, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    def _on_close(self) -> None:
        if self.exit_on_close_var.get():
            self.exit_application()
            return
        self._hide_to_tray("点击关闭按钮")

    def exit_application(self) -> None:
        self.stop_monitoring()
        self.portal_automator.cleanup()
        self.tray_controller.stop()
        self.root.destroy()

    def _run_on_ui(self, callback: Callable, *args) -> None:
        self.root.after(0, lambda: callback(*args))

    def hide_to_tray_manually(self) -> None:
        self._hide_to_tray("手动操作")

    def _ensure_tray_icon(self) -> bool:
        if self.tray_controller.start():
            return True
        if not self.tray_fallback_logged:
            self.tray_fallback_logged = True
            self._log("托盘功能不可用（缺少 pystray/Pillow 或系统不支持），回退为最小化到任务栏。")
        return False

    def _hide_to_tray(self, reason: str) -> None:
        if self.window_hidden_to_tray:
            return
        if self._ensure_tray_icon():
            self.root.withdraw()
            self.window_hidden_to_tray = True
            self._log(f"已收至托盘（{reason}），监控继续运行。")
            return
        self.root.iconify()
        self._log(f"已最小化到任务栏（{reason}），监控继续运行。")

    def _restore_from_tray(self) -> None:
        if self.window_hidden_to_tray:
            self.window_hidden_to_tray = False
        self.root.deiconify()
        self.root.lift()
        try:
            self.root.focus_force()
        except Exception:
            pass
        self._log("已从托盘恢复主窗口。")

    def _startup_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        try:
            with self.startup_log_lock:
                self.startup_log_path.parent.mkdir(parents=True, exist_ok=True)
                if self.startup_log_path.exists():
                    current_size = self.startup_log_path.stat().st_size
                    if current_size > self.STARTUP_LOG_MAX_BYTES:
                        self.startup_log_path.write_text("", encoding="utf-8")
                with self.startup_log_path.open("a", encoding="utf-8") as fh:
                    fh.write(line)
        except OSError:
            pass

    @staticmethod
    def _format_command_for_log(command: Optional[str]) -> str:
        if not command:
            return "<missing>"
        compact = " ".join(command.split())
        if len(compact) > 240:
            return compact[:237] + "..."
        return compact

    def _collect_autostart_channel_state(self) -> dict:
        run_command = self.autostart_service.get_run_command()
        startup_command = self.autostart_service.get_startup_folder_command()
        return {
            "run_command": run_command,
            "run_present": bool(run_command),
            "run_valid": self._is_autostart_command_valid(run_command),
            "startup_command": startup_command,
            "startup_present": bool(startup_command),
            "startup_valid": self._is_autostart_command_valid(startup_command),
        }

    def _record_autostart_channel_state(self, stage: str, state: dict) -> None:
        self._startup_log(
            f"{stage}: "
            f"run_present={state['run_present']}, run_valid={state['run_valid']}, "
            f"run_command={self._format_command_for_log(state['run_command'])}; "
            f"startup_present={state['startup_present']}, startup_valid={state['startup_valid']}, "
            f"startup_command={self._format_command_for_log(state['startup_command'])}"
        )

    def _log(self, message: str) -> None:
        if threading.current_thread() is not threading.main_thread():
            self._run_on_ui(self._log, message)
            return
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{now}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _set_last_run_time_now(self) -> None:
        self.last_run_time_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def _load_credentials_to_form(self) -> None:
        creds = self.credential_store.load()
        if creds:
            self.username_var.set(creds[0])
            self._log("已读取已保存账号。")

    def save_credentials(self) -> None:
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        if not username or not password:
            messagebox.showwarning("提示", "账号和密码都不能为空。")
            return
        try:
            self.credential_store.save(username, password)
            self.password_var.set("")
            self._log("凭据已保存到 Windows 凭据管理器。")
            messagebox.showinfo("完成", "凭据保存成功。")
        except Exception as exc:
            self._log(f"凭据保存失败: {exc}")
            messagebox.showerror("错误", f"凭据保存失败: {exc}")

    def clear_credentials(self) -> None:
        try:
            self.credential_store.clear()
            self.password_var.set("")
            self._log("凭据已清除。")
            messagebox.showinfo("完成", "凭据已清除。")
        except Exception as exc:
            self._log(f"清除凭据失败: {exc}")
            messagebox.showerror("错误", f"清除凭据失败: {exc}")

    def _parse_interval_settings(self) -> Optional[Tuple[int, int]]:
        try:
            poll_interval = int(self.poll_interval_var.get().strip())
        except ValueError:
            messagebox.showwarning("提示", "轮询间隔必须是整数。")
            return None

        try:
            auth_probe_interval = int(self.auth_probe_interval_var.get().strip())
        except ValueError:
            messagebox.showwarning("提示", "认证保活检测间隔必须是整数。")
            return None

        if poll_interval < 1:
            messagebox.showwarning("提示", "轮询间隔不能小于 1 秒。")
            return None

        if auth_probe_interval < 10:
            messagebox.showwarning("提示", "认证保活检测间隔不能小于 10 秒。")
            return None

        return poll_interval, auth_probe_interval

    def save_settings(self) -> None:
        parsed = self._parse_interval_settings()
        if not parsed:
            return
        poll_interval, auth_probe_interval = parsed

        self.config.poll_interval_sec = poll_interval
        self.config.auth_probe_interval_sec = auth_probe_interval
        self.config.autostart_enabled = bool(self.autostart_var.get())
        self.config.exit_on_close = bool(self.exit_on_close_var.get())
        self.config.auto_hide_to_tray_on_success = bool(self.auto_hide_to_tray_var.get())
        self.config_manager.save(self.config)
        self._log(
            f"设置已保存，轮询间隔={poll_interval}s，认证保活检测间隔={auth_probe_interval}s。"
        )
        messagebox.showinfo("完成", "设置已保存。")

    def apply_autostart(self) -> None:
        desired = bool(self.autostart_var.get())
        try:
            if desired:
                cmd = build_autostart_command()
                channel = self.autostart_service.enable(cmd)
                channel_label = "注册表" if channel == "run" else "启动文件夹"
                self._log(f"已启用开机自启（通道: {channel_label}）。")
                self._startup_log(
                    f"autostart_toggle: enabled=True, channel={channel}, "
                    f"command={self._format_command_for_log(cmd)}"
                )
            else:
                self.autostart_service.disable()
                self._log("已关闭开机自启。")
                self._startup_log("autostart_toggle: enabled=False")
            self.config.autostart_enabled = desired
            self.config_manager.save(self.config)
        except Exception as exc:
            actual_enabled = self.autostart_service.is_enabled()
            self.autostart_var.set(actual_enabled)
            self._log(f"更新开机自启失败: {exc}")
            self._startup_log(f"autostart_toggle: failed={exc}")
            messagebox.showerror("错误", f"更新开机自启失败: {exc}")

    def apply_close_behavior(self) -> None:
        previous = bool(self.config.exit_on_close)
        desired = bool(self.exit_on_close_var.get())
        try:
            self.config.exit_on_close = desired
            self.config_manager.save(self.config)
            action = "退出程序" if desired else "收至托盘"
            self._log(f"关闭窗口行为已更新：点击 X 时{action}。")
        except Exception as exc:
            self.config.exit_on_close = previous
            self.exit_on_close_var.set(previous)
            self._log(f"更新关闭窗口行为失败: {exc}")
            messagebox.showerror("错误", f"更新关闭窗口行为失败: {exc}")

    def apply_auto_hide_to_tray(self) -> None:
        previous = bool(self.config.auto_hide_to_tray_on_success)
        desired = bool(self.auto_hide_to_tray_var.get())
        try:
            self.config.auto_hide_to_tray_on_success = desired
            self.config_manager.save(self.config)
            status_text = "已启用" if desired else "已关闭"
            self._log(f"{status_text}认证成功后自动收至托盘。")
        except Exception as exc:
            self.config.auto_hide_to_tray_on_success = previous
            self.auto_hide_to_tray_var.set(previous)
            self._log(f"更新自动收托盘设置失败: {exc}")
            messagebox.showerror("错误", f"更新自动收托盘设置失败: {exc}")

    @staticmethod
    def _split_command(command: str) -> list:
        try:
            return shlex.split(command, posix=False)
        except ValueError:
            return []

    @staticmethod
    def _is_accessible_command_path(path: str) -> bool:
        cleaned = str(path or "").strip().strip('"')
        if not cleaned:
            return False
        if os.path.isfile(cleaned):
            return True
        return shutil.which(cleaned) is not None

    def _is_autostart_command_valid(self, command: Optional[str]) -> bool:
        if not command:
            return False
        lower_command = command.lower()
        if "--autostart" not in lower_command or "--minimized" not in lower_command:
            return False

        parts = self._split_command(command)
        if not parts:
            return False
        if not self._is_accessible_command_path(parts[0]):
            return False

        exe_name = os.path.basename(parts[0]).strip('"').lower()
        if exe_name.startswith("python"):
            if len(parts) < 2:
                return False
            script_path = parts[1].strip('"')
            if not os.path.isfile(script_path):
                return False
        return True

    def _ensure_autostart_command_integrity(self) -> bool:
        autostart_enabled = bool(self.autostart_var.get() or self.config.autostart_enabled)
        if not autostart_enabled:
            self._log("启动项校验: 开机自启未启用，跳过。")
            self._startup_log("autostart_integrity: skipped_disabled")
            return False

        run_command = self.autostart_service.get_run_command()
        startup_command = self.autostart_service.get_startup_folder_command()
        run_present = bool(run_command)
        startup_present = bool(startup_command)
        run_valid = self._is_autostart_command_valid(run_command)
        startup_valid = self._is_autostart_command_valid(startup_command)

        target_command = build_autostart_command()

        def same_command(left: Optional[str], right: Optional[str]) -> bool:
            if not left or not right:
                return False
            return " ".join(left.split()) == " ".join(right.split())

        repaired_actions = []
        try:
            if run_present and startup_present:
                self.autostart_service.disable_startup_folder()
                repaired_actions.append("startup_folder_removed")
                startup_command = self.autostart_service.get_startup_folder_command()
                startup_valid = self._is_autostart_command_valid(startup_command)

            if run_valid:
                if not same_command(run_command, target_command):
                    self.autostart_service.set_run_command(target_command)
                    repaired_actions.append("run_updated")
            elif startup_valid:
                channel = self.autostart_service.enable(target_command)
                repaired_actions.append(f"migrated_to_{channel}")
            else:
                channel = self.autostart_service.enable(target_command)
                repaired_actions.append(f"recreated_{channel}")

            if repaired_actions:
                self.autostart_var.set(True)
                self.config.autostart_enabled = True
                self.config_manager.save(self.config)
                self._log(f"已修复开机启动项: {' + '.join(repaired_actions)}")
                self._startup_log(
                    f"autostart_integrity: repaired_actions={'+'.join(repaired_actions)}, "
                    f"target_command={self._format_command_for_log(target_command)}"
                )
                return True

            self._log("启动项校验: 注册表启动项有效，且未发现重复启动通道。")
            self._startup_log("autostart_integrity: healthy_single_channel=True")
            return False
        except Exception as exc:
            self._log(f"修复开机启动项失败: {exc}")
            self._startup_log(f"autostart_integrity: repair_failed={exc}")
            return False

    @staticmethod
    def _is_expected_probe_response(status_code: int, final_url: str, body_sample: str) -> bool:
        if status_code != 200:
            return False
        final_host = urlparse(final_url).netloc.lower()
        if final_host not in ("www.msftconnecttest.com", "msftconnecttest.com"):
            return False
        normalized_body = " ".join((body_sample or "").split()).lower()
        return NWAFUGuiApp.AUTH_PROBE_EXPECTED_TEXT.lower() in normalized_body

    def _probe_auth_session(self) -> bool:
        portal_host = urlparse(self.config.portal_url).netloc.lower()
        probe_urls = (self.AUTH_PROBE_HTTP_URL, self.AUTH_PROBE_HTTPS_URL)

        for probe_url in probe_urls:
            request = Request(
                probe_url,
                headers={
                    "User-Agent": "Mozilla/5.0",
                    "Cache-Control": "no-cache",
                    "Pragma": "no-cache",
                },
            )
            try:
                with urlopen(request, timeout=self.AUTH_PROBE_TIMEOUT_SECONDS) as response:
                    status_code = int(response.getcode() or 0)
                    final_url = str(response.geturl() or probe_url)
                    body_sample = response.read(200).decode("utf-8", errors="ignore")
            except (HTTPError, URLError, TimeoutError, OSError) as exc:
                self._log(f"认证保活检测失败 ({probe_url}): {exc}")
                continue

            final_url_lower = final_url.lower()
            if portal_host and portal_host in final_url_lower:
                self._log(f"认证保活检测结果: 已重定向到认证页 ({final_url})。")
                return False

            if self._is_expected_probe_response(status_code, final_url, body_sample):
                self._log(f"认证保活检测结果: 在线 (status={status_code}, url={final_url})。")
                return True

            body_preview = " ".join(body_sample.split())
            if len(body_preview) > 80:
                body_preview = body_preview[:80] + "..."
            if not body_preview:
                body_preview = "<空>"
            self._log(
                f"认证保活检测异常响应 ({probe_url}): status={status_code}, url={final_url}, body={body_preview}。"
            )

        self._log("认证保活检测结果: 未通过连通性校验，判定离线。")
        return False

    def open_portal_page(self) -> None:
        opened = self.portal_automator.open_portal_in_local_edge(
            self.config,
            self._log,
            source="manual",
            reason="manual_open_button",
            force=True,
        )
        if not opened:
            messagebox.showerror("错误", "打开认证页面失败，请查看日志。")

    def start_monitoring(self) -> None:
        if self.monitor_running:
            return
        parsed = self._parse_interval_settings()
        if not parsed:
            return
        poll_interval, auth_probe_interval = parsed
        self.config.poll_interval_sec = poll_interval
        self.config.auth_probe_interval_sec = auth_probe_interval
        self.config_manager.save(self.config)

        self.monitor_stop_event.clear()
        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()
        self.monitor_running = True
        self.status_var.set("运行中")
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self._log("监控已启动。")
        started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._startup_log(
            f"monitor_started_at={started_at}, poll_interval={poll_interval}, "
            f"auth_probe_interval={auth_probe_interval}"
        )

    def stop_monitoring(self) -> None:
        if not self.monitor_running:
            return
        self.monitor_stop_event.set()
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=2)
        self.monitor_running = False
        self.status_var.set("已停止")
        self.start_btn.configure(state=tk.NORMAL)
        self.stop_btn.configure(state=tk.DISABLED)
        self._log("监控已停止。")

    def manual_check(self) -> None:
        self._start_check_worker(source="manual", force=True)

    def _monitor_loop(self) -> None:
        while not self.monitor_stop_event.is_set():
            now = time.time()
            ssid = self.wifi_service.get_current_ssid()
            self._run_on_ui(self.current_ssid_var.set, ssid or "未连接")

            is_target = ssid == self.config.target_ssid
            if is_target and not self.last_seen_target:
                self.session_authenticated = False
                self.next_auto_retry_at = 0
                self.last_auth_probe_at = 0
                self._log(f"检测到连接 {self.config.target_ssid}。")

            if is_target and self.session_authenticated:
                elapsed = now - self.last_auth_probe_at
                if elapsed >= self.config.auth_probe_interval_sec:
                    self.last_auth_probe_at = now
                    if self._probe_auth_session():
                        self._log(
                            f"认证保活检测: 正常，下一次检测间隔 {self.config.auth_probe_interval_sec} 秒。"
                        )
                    else:
                        self.session_authenticated = False
                        self.next_auto_retry_at = 0
                        reconnect_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        self._log(f"认证状态失效，{reconnect_at} 立即触发自动重连。")

            if is_target and not self.session_authenticated:
                if now >= self.next_auto_retry_at:
                    self._start_check_worker(source="auto", force=False)

            if not is_target and self.last_seen_target:
                self.session_authenticated = False
                self.next_auto_retry_at = 0
                self.last_auth_probe_at = 0
                self._log(f"已离开 {self.config.target_ssid}，下次连接将重新认证。")

            self.last_seen_target = is_target
            self.monitor_stop_event.wait(self.config.poll_interval_sec)

    def _start_check_worker(self, source: str, force: bool) -> None:
        if self.check_lock.locked():
            return
        worker = threading.Thread(target=self._check_once, args=(source, force), daemon=True)
        worker.start()

    def _show_missing_credential_warning(self) -> None:
        messagebox.showwarning("缺少凭据", "请先保存账号和密码。")

    def _check_once(self, source: str, force: bool) -> None:
        with self.check_lock:
            self._run_on_ui(self._set_last_run_time_now)
            self._run_on_ui(self.last_result_var.set, "执行中")
            ssid = self.wifi_service.get_current_ssid()
            self._run_on_ui(self.current_ssid_var.set, ssid or "未连接")

            if not force and ssid != self.config.target_ssid:
                self._run_on_ui(self.last_result_var.set, "跳过")
                self._log(f"自动检测跳过：当前 SSID 不是 {self.config.target_ssid}。")
                return

            creds = self.credential_store.load()
            if not creds:
                self._run_on_ui(self.last_result_var.set, "缺少凭据")
                self._log("未找到凭据，无法执行登录。")
                now = time.time()
                should_alert = source == "manual" or (
                    now - self.last_cred_alert_at >= self.CRED_ALERT_COOLDOWN_SECONDS
                )
                if should_alert:
                    self.last_cred_alert_at = now
                    self._run_on_ui(self._show_missing_credential_warning)
                self.next_auto_retry_at = now + self.AUTO_RETRY_SECONDS
                return

            source_text = {"auto": "自动", "manual": "手动"}.get(source, source)
            self._log(f"开始执行认证（来源: {source_text}）。")
            result = self.portal_automator.ensure_logged_in(
                self.config, creds, self._log, source=source
            )
            should_auto_hide = source == "auto" and bool(self.auto_hide_to_tray_var.get())

            if result == LoginResult.ALREADY_LOGGED_IN:
                self.session_authenticated = True
                self.next_auto_retry_at = 0
                self.last_auth_probe_at = time.time()
                self._run_on_ui(self.last_result_var.set, "已登录")
                if should_auto_hide:
                    self._run_on_ui(self._hide_to_tray, "认证已在线")
            elif result == LoginResult.LOGGED_IN_NOW:
                self.session_authenticated = True
                self.next_auto_retry_at = 0
                self.last_auth_probe_at = time.time()
                self._run_on_ui(self.last_result_var.set, "登录成功")
                if should_auto_hide:
                    self._run_on_ui(self._hide_to_tray, "认证成功")
            else:
                self.session_authenticated = False
                retry_at_epoch = time.time() + self.AUTO_RETRY_SECONDS
                self.next_auto_retry_at = retry_at_epoch
                retry_at_text = datetime.fromtimestamp(retry_at_epoch).strftime("%Y-%m-%d %H:%M:%S")
                self._log(
                    f"登录失败，{self.AUTO_RETRY_SECONDS} 秒后自动重试（预计时间: {retry_at_text}）。"
                )
                self._run_on_ui(self.last_result_var.set, "登录失败")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--autostart", action="store_true")
    parser.add_argument("--minimized", action="store_true")
    return parser.parse_args()


def _set_windows_app_user_model_id() -> None:
    if os.name != "nt":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_NAME)
    except Exception:
        pass


def main() -> None:
    _set_windows_app_user_model_id()
    args = parse_args()
    try:
        instance_guard = SingleInstanceGuard()
    except Exception:
        instance_guard = None

    if instance_guard and instance_guard.already_running:
        if not args.autostart:
            try:
                prompt_root = tk.Tk()
                prompt_root.withdraw()
                messagebox.showinfo("提示", "程序已在运行，无需重复启动。")
                prompt_root.destroy()
            except Exception:
                pass
        return

    try:
        root = tk.Tk()
        NWAFUGuiApp(root, launch_autostart=args.autostart, minimized=args.minimized)
        root.mainloop()
    finally:
        if instance_guard:
            instance_guard.release()


if __name__ == "__main__":
    main()
