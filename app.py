# -*- coding: utf-8 -*-
"""업무 도우미 — 데스크톱 앱 진입점.

pywebview 창 안에 web/index.html 을 띄우고,
JS 에서 window.pywebview.api.<메서드>() 로 파이썬 백엔드를 호출한다.
파일/폴더 선택은 pywebview 내장 네이티브 창을 쓴다.
"""

import os
import sys
import time
import tempfile
import threading
import traceback
import winreg
from datetime import datetime

import webview

# WebView2 Evergreen 런타임 감지용
WEBVIEW2_GUID = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
WEBVIEW2_URL = "https://developer.microsoft.com/microsoft-edge/webview2/"

from backend import mail_recipients as mr
from backend import xls_merge as xm
from backend import folders as fld

EXCEL_TYPES = ("Excel 파일 (*.xlsx;*.xlsm)", "모든 파일 (*.*)")

# ── 로그 ────────────────────────────────────────────────────
LOG_PATH = os.path.join(tempfile.gettempdir(), "WorkHelper.log")
_log_lock = threading.Lock()


def log(msg):
    line = f"[{datetime.now().strftime('%H:%M:%S.%f')[:-3]}] {msg}"
    try:
        with _log_lock:
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(line + "\n")
    except Exception:
        pass
    try:
        print(line, flush=True)
    except Exception:
        pass


def resource_path(rel):
    """개발 실행 / PyInstaller(onefile) 실행 모두에서 리소스 경로를 찾는다."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def build_html():
    """web/index.html 을 자체 완결형 HTML 문자열로 만든다.
    - 외부 폰트(googleapis) 링크 제거: 사내/오프라인 PC 에서 로딩 지연·멈춤 방지
    → create_window(html=...) 로 직접 주입하면 내부 HTTP 서버/파일 접근이 필요 없어
      frozen(exe) 환경에서도 즉시 로드된다.
    """
    import re
    web_dir = resource_path("web")
    with open(os.path.join(web_dir, "index.html"), encoding="utf-8") as f:
        html = f.read()
    # 외부 폰트 링크 / @font-face 제거 (시스템 폰트로 자연 폴백)
    html = re.sub(r'<link[^>]*fonts\.googleapis\.com[^>]*>', '', html)
    html = re.sub(r'<link[^>]*rel="preconnect"[^>]*>', '', html)
    html = re.sub(r"@font-face\{[^}]*\}", '', html)
    return html


def webview2_installed():
    """Edge WebView2 런타임 설치 여부를 레지스트리로 확인. (설치여부, 버전)"""
    locations = [
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\%s" % WEBVIEW2_GUID),
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\Microsoft\EdgeUpdate\Clients\%s" % WEBVIEW2_GUID),
        (winreg.HKEY_CURRENT_USER,
         r"Software\Microsoft\EdgeUpdate\Clients\%s" % WEBVIEW2_GUID),
    ]
    for hive, path in locations:
        try:
            with winreg.OpenKey(hive, path) as k:
                pv, _ = winreg.QueryValueEx(k, "pv")
                if pv and pv not in ("", "0.0.0.0"):
                    return True, pv
        except OSError:
            continue
    return False, None


def show_message(title, text, icon=0x40):
    """네이티브 메시지 박스 (GUI 프레임워크 없이 동작). 0x40=정보 아이콘."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, title, icon)
    except Exception:
        pass


def ask_yesno(title, text):
    """예/아니오 메시지 박스. 예=True. (MB_YESNO|MB_ICONQUESTION)"""
    try:
        import ctypes
        return ctypes.windll.user32.MessageBoxW(0, text, title, 0x04 | 0x20) == 6  # IDYES
    except Exception:
        return False


def app_dir():
    """exe(또는 스크립트)가 놓인 실제 폴더. (onefile 압축해제 폴더가 아닌 배포 위치)"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ── 첫 실행 로딩 스플래시 ──────────────────────────────────────
# WebView2 는 첫 실행 때 (보안 스캔·프로필 최초 생성으로) 화면이 뜨기까지 수 초~수십 초
# 걸릴 수 있고, 그동안 메인 창이 '응답없음'처럼 보인다. 이를 위해 WebView2 와 별개인
# 가벼운 네이티브 창(tkinter)을 '별도 스레드'에서 띄운다. 메인 UI 스레드가 바빠도
# 이 창은 자기 스레드에서 독립적으로 움직여 "준비 중"임을 계속 보여준다.
_splash_stop = threading.Event()


def _run_splash():
    root = None
    try:
        import tkinter as tk
        root = tk.Tk()
        root.overrideredirect(True)          # 제목표시줄 없는 심플 창
        root.attributes("-topmost", True)
        w, h = 380, 165
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")
        root.configure(bg="#E5E3DC")
        inner = tk.Frame(root, bg="#FAFAF7")
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text="업무 도우미", bg="#FAFAF7", fg="#0E0D0B",
                 font=("Malgun Gothic", 14, "bold")).pack(pady=(30, 8))
        tk.Label(inner,
                 text="프로그램을 준비하고 있습니다.\n첫 실행은 시간이 걸릴 수 있어요. 잠시만 기다려 주세요…",
                 bg="#FAFAF7", fg="#6E6A60", font=("Malgun Gothic", 10),
                 justify="center").pack()
        dot = tk.Label(inner, text="", bg="#FAFAF7", fg="#8A1538",
                       font=("Malgun Gothic", 13, "bold"))
        dot.pack(pady=(12, 0))
        st = {"n": 0, "c": 0}

        def tick():
            st["c"] += 1
            # loaded 신호가 오거나, 안전장치로 약 120초가 지나면 닫는다.
            if _splash_stop.is_set() or st["c"] > 400:
                root.quit()
                return
            st["n"] = (st["n"] + 1) % 4
            dot.config(text="●" * st["n"] + "○" * (3 - st["n"]))
            root.after(300, tick)

        root.after(10, tick)
        root.mainloop()
    except Exception as e:
        log(f"스플래시 표시 실패(무시): {e!r}")
    finally:
        # 반드시 이 스레드에서 정리(다른 스레드 GC 로 인한 Tcl 오류 방지)
        try:
            if root is not None:
                root.destroy()
        except Exception:
            pass
        root = None


def guard(fn):
    """API 메서드: 진입/종료/예외를 로그로 남기고, 예외는 {'ok':False,...} 로 반환."""
    def wrapper(*args, **kwargs):
        t0 = time.time()
        # 개인정보/조직명 보호: 인자 '값'은 로그에 남기지 않고 개수만 기록
        log(f"→ API 호출: {fn.__name__} (인자 {max(0, len(args) - 1)}개)")
        try:
            result = fn(*args, **kwargs)
            dt = (time.time() - t0) * 1000
            ok = result.get("ok") if isinstance(result, dict) else "?"
            log(f"← API 완료: {fn.__name__}  ok={ok}  ({dt:.0f}ms)")
            return result
        except Exception as e:
            log(f"✗ API 예외: {fn.__name__}  {e!r}\n{traceback.format_exc()}")
            return {"ok": False, "error": f"오류가 발생했습니다: {e}"}
    wrapper.__name__ = fn.__name__
    return wrapper


class Api:
    def __init__(self):
        self.window = None
        self._roster = None          # 주소록 파싱 상태
        self._last_recipients = None  # 마지막 수신자 계산 결과
        self._last_roles = ([], [])
        self._folder_excel = None     # 폴더 생성용 엑셀 경로

    # ── 공용 다이얼로그 ──────────────────────────────────
    def _open_file(self, directory=""):
        res = self.window.create_file_dialog(
            webview.FileDialog.OPEN, directory=directory,
            allow_multiple=False, file_types=EXCEL_TYPES)
        if not res:
            return None
        return res[0] if isinstance(res, (list, tuple)) else res

    def _open_folder(self, directory=""):
        res = self.window.create_file_dialog(
            webview.FileDialog.FOLDER, directory=directory)
        if not res:
            return None
        return res[0] if isinstance(res, (list, tuple)) else res

    def log_js(self, msg):
        """JS 측에서 보내는 로그/오류를 파일 로그에 기록 (guard 미적용: 빠르게)."""
        log(f"[JS] {msg}")
        return {"ok": True}

    @guard
    def open_path(self, path):
        """탐색기에서 파일/폴더 위치 열기."""
        if path and os.path.exists(path):
            os.startfile(os.path.dirname(path) if os.path.isfile(path) else path)
            return {"ok": True}
        return {"ok": False, "error": "경로를 찾을 수 없습니다."}

    # ── 1) 주소록 ────────────────────────────────────────
    @guard
    def pick_roster(self):
        """주소록 파일 선택 → 분석 → 역할/팀/컬럼 정보 반환."""
        path = self._open_file()
        if not path:
            return {"ok": False, "cancelled": True}
        res = mr.analyze(path)
        if "meta" not in res:
            return res  # {'ok':False,'error':...}
        self._roster = res["_state"]
        return res["meta"]

    @guard
    def run_recipients(self, teams_text, to_roles, cc_roles):
        """수신/참조 이메일 계산."""
        if not self._roster:
            return {"ok": False, "error": "먼저 주소록 파일을 선택하세요."}
        result = mr.compute(self._roster, teams_text, to_roles or [], cc_roles or [])
        if result.get("ok"):
            self._last_recipients = result
            self._last_roles = (to_roles or [], cc_roles or [])
        return result

    @guard
    def save_recipients_txt(self):
        """마지막 계산 결과를 txt 로 저장."""
        if not (self._roster and self._last_recipients):
            return {"ok": False, "error": "저장할 결과가 없습니다."}
        to_roles, cc_roles = self._last_roles
        path = mr.save_txt(self._roster, self._last_recipients, to_roles, cc_roles)
        return {"ok": True, "path": path}

    # ── 2) 엑셀 취합 ─────────────────────────────────────
    @guard
    def pick_xls_folder(self):
        path = self._open_folder()
        if not path:
            return {"ok": False, "cancelled": True}
        # 미리 대상 파일 개수만 훑어본다 (접두어 미정이라 전체 xlsx 개수)
        return {"ok": True, "folder": path}

    @guard
    def run_xls_merge(self, folder, prefix, sheets_text, headers_text, recursive):
        if not folder:
            return {"ok": False, "error": "취합 대상 폴더를 선택하세요."}
        return xm.run(folder, prefix, sheets_text, headers_text, bool(recursive))

    # ── 3) 폴더 생성 ─────────────────────────────────────
    @guard
    def pick_folder_excel(self):
        path = self._open_file()
        if not path:
            return {"ok": False, "cancelled": True}
        res = fld.load_excel(path)
        if not res.get("ok"):
            return res
        self._folder_excel = path
        res["file_name"] = os.path.basename(path)
        return res

    @guard
    def preview_folder_tree(self, sheet, upper_col, lower_col, number_mode):
        if not self._folder_excel:
            return {"ok": False, "error": "먼저 엑셀 파일을 선택하세요."}
        return fld.preview(self._folder_excel, sheet, upper_col, lower_col, number_mode)

    @guard
    def pick_output_dir(self):
        base = os.path.dirname(self._folder_excel) if self._folder_excel else ""
        path = self._open_folder(directory=base)
        if not path:
            return {"ok": False, "cancelled": True}
        return {"ok": True, "path": path}

    @guard
    def run_create_folders(self, sheet, upper_col, lower_col, number_mode, output_dir):
        if not self._folder_excel:
            return {"ok": False, "error": "먼저 엑셀 파일을 선택하세요."}
        if not output_dir:
            return {"ok": False, "error": "생성 위치를 선택하세요."}
        return fld.run(self._folder_excel, sheet, upper_col, lower_col,
                       number_mode, output_dir)


def main():
    # 새 실행마다 로그 초기화
    try:
        open(LOG_PATH, "w", encoding="utf-8").close()
    except Exception:
        pass
    log("=" * 50)
    log(f"업무 도우미 시작  (frozen={getattr(sys,'frozen',False)})")
    log(f"index.html: {resource_path(os.path.join('web','index.html'))}")
    log(f"로그 파일: {LOG_PATH}")

    # WebView2 런타임 확인 (없으면 '응답없음' 대신 안내창)
    ok, pv = webview2_installed()
    log(f"WebView2 런타임: 설치={ok} 버전={pv}")
    if not ok:
        log("WebView2 런타임 없음")
        setup = os.path.join(app_dir(), "MicrosoftEdgeWebview2Setup.exe")
        log(f"함께 배포된 설치파일 존재: {os.path.isfile(setup)} ({setup})")
        if os.path.isfile(setup):
            # 옆에 설치파일이 있으면 바로 설치를 제안
            if ask_yesno(
                "실행에 필요한 구성요소 설치",
                "이 프로그램을 실행하려면 'Microsoft Edge WebView2 런타임'이 필요합니다.\n\n"
                "함께 배포된 설치 파일이 있습니다. 지금 설치하시겠습니까?\n"
                "([예] → 설치 진행 · 설치 후 이 프로그램을 다시 실행해 주세요.)\n\n"
                "※ 설치에는 인터넷 연결이 필요합니다."):
                try:
                    log("설치파일 실행")
                    os.startfile(setup)
                    show_message(
                        "설치를 시작했습니다",
                        "WebView2 런타임 설치가 진행됩니다.\n"
                        "설치가 끝난 뒤 이 프로그램을 다시 실행해 주세요.")
                except Exception as e:
                    log(f"설치파일 실행 실패: {e!r}")
                    show_message("설치 실행 실패",
                                 f"설치 파일을 실행하지 못했습니다.\n{e}\n\n"
                                 f"직접 실행해 주세요:\n{setup}", icon=0x10)
            return
        # 설치파일이 없으면 다운로드 페이지 안내
        log("설치파일 없음 → 다운로드 페이지 안내")
        show_message(
            "실행에 필요한 구성요소가 없습니다",
            "이 프로그램을 실행하려면 'Microsoft Edge WebView2 런타임'이 필요합니다.\n\n"
            "[확인]을 누르면 Microsoft 공식 다운로드 페이지가 열립니다.\n"
            "페이지에서 'Evergreen Bootstrapper'(상시 최신) 를 내려받아 설치한 뒤,\n"
            "이 프로그램을 다시 실행해 주세요.\n\n"
            "※ Windows 11 및 최신 Windows 10 에는 대부분 기본 설치되어 있습니다.")
        try:
            import webbrowser
            webbrowser.open(WEBVIEW2_URL)
        except Exception:
            pass
        return

    # 첫 실행 로딩 안내창을 먼저 띄운다 (WebView2 준비 대기 동안 '응답없음'처럼 보이는 것 방지)
    _splash_stop.clear()
    threading.Thread(target=_run_splash, daemon=True).start()

    api = Api()
    html = build_html()
    log(f"HTML 준비 완료 (길이={len(html)}, 외부 링크 제거·SVG 인라인)")
    window = webview.create_window(
        "업무 도우미",
        html=html,
        js_api=api,
        width=1280,
        height=880,
        min_size=(1000, 700),
        background_color="#F2F1EC",
    )
    api.window = window

    # 라이프사이클 이벤트 로깅 (어느 단계까지 진행되는지 추적)
    def _ev(name):
        def handler(*a, **k):
            log(f"[EVENT] {name}")
        return handler
    for _name in ("initialized", "before_load", "loaded", "shown", "closed"):
        try:
            getattr(window.events, _name).__iadd__(_ev(_name))
        except Exception as _e:
            log(f"이벤트 등록 실패 {_name}: {_e!r}")

    # 메인 화면 로드가 끝나면 스플래시를 닫는다.
    try:
        window.events.loaded.__iadd__(lambda *a, **k: _splash_stop.set())
    except Exception as _e:
        log(f"스플래시 종료 핸들러 등록 실패: {_e!r}")

    # ★ WebView2 사용자 데이터 폴더(UDF)를 '실재하는 쓰기가능 경로'로 고정.
    #   pywebview 기본값(private_mode)은 UDF 를 tempfile.TemporaryDirectory().name 으로
    #   잡는데, 그 임시폴더가 즉시 GC 되며 삭제되어 → PC/타이밍에 따라 WebView2 가
    #   프로필을 못 만들고 로딩이 멈추는(before_load 후 정지) 문제가 있다.
    #   고정 경로를 주면 이 문제가 사라진다.
    storage = os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(),
                           "WorkHelper", "webview2")
    try:
        os.makedirs(storage, exist_ok=True)
    except Exception as e:
        log(f"storage_path 생성 실패({storage}): {e!r}")
        storage = None
    log(f"storage_path={storage}")

    debug = os.environ.get("APP_DEBUG") == "1"
    log("webview.start() 진입")
    try:
        webview.start(debug=debug, storage_path=storage, private_mode=False)
        log("webview.start() 반환 (창 종료)")
    except Exception as e:
        log(f"webview.start() 예외: {e!r}\n{traceback.format_exc()}")
        show_message(
            "화면을 열지 못했습니다",
            "앱 창을 생성하는 중 오류가 발생했습니다.\n\n"
            f"{e}\n\n"
            "Microsoft Edge WebView2 런타임을 (재)설치한 뒤 다시 시도해 주세요.\n"
            f"로그: {LOG_PATH}",
            icon=0x10)  # 오류 아이콘
        raise


if __name__ == "__main__":
    main()
