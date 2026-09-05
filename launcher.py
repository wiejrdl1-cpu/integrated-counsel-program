import json
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from pastel_theme import PALETTE, apply_button_palette, apply_pastel_theme


_NativeButton = tk.Button


class RaisedButton(_NativeButton):
    """샘플 UI와 같은 얇은 테두리의 평면형 버튼입니다."""
    def __init__(self, master=None, **kwargs):
        kwargs.pop("style", None)
        kwargs["relief"] = "flat"
        kwargs["bd"] = 0
        kwargs["highlightthickness"] = 1
        kwargs.setdefault("font", ("Malgun Gothic", 10, "bold"))
        kwargs.setdefault("padx", 12)
        kwargs.setdefault("pady", 7)
        kwargs.setdefault("cursor", "hand2")
        apply_button_palette(kwargs)
        super().__init__(master, **kwargs)


# 기존 ttk.Button과 tk.Button 호출도 모두 같은 모양으로 표시합니다.
tk.Button = RaisedButton
ttk.Button = RaisedButton


APP_TITLE = "통합조사 상담 프로그램 v1"
CONFIG_DIR = Path.home() / ".integrated_investigation_launcher"
RECENT_FILE = CONFIG_DIR / "recent_team.json"


def load_recent_team():
    try:
        data = json.loads(RECENT_FILE.read_text(encoding="utf-8"))
        return data.get("team", "")
    except Exception:
        return ""


def save_recent_team(team):
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        RECENT_FILE.write_text(
            json.dumps({"team": team}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


class Launcher:
    def __init__(self, root):
        self.root = root
        self.selected_team = None
        root.title(APP_TITLE)
        root.geometry("720x330")
        root.minsize(640, 310)
        root.resizable(True, True)
        root.protocol("WM_DELETE_WINDOW", root.destroy)

        self.setup_style()
        self.build_ui()
        self.center_window()

    def setup_style(self):
        style = apply_pastel_theme(self.root, 10)
        style.configure("Title.TLabel", font=("Malgun Gothic", 20, "bold"), foreground=PALETTE['primary'])
        style.configure("Guide.TLabel", font=("Malgun Gothic", 11), foreground=PALETTE['muted'])
        style.configure("Team.TLabel", font=("Malgun Gothic", 14, "bold"), foreground=PALETTE['text'])
        style.configure("Use.TButton", font=("Malgun Gothic", 11, "bold"), padding=(16, 10))
        style.configure("Exit.TButton", font=("Malgun Gothic", 10), padding=(12, 7))

    def build_ui(self):
        outer = tk.Frame(self.root, padx=42, pady=28, background=PALETTE['window'])
        outer.pack(fill="both", expand=True)

        tk.Label(outer, text=APP_TITLE, background=PALETTE['window'], foreground=PALETTE['primary'], font=("Malgun Gothic", 20, "bold")).pack()

        cards = tk.Frame(outer, background=PALETTE['window'])
        cards.pack(fill="both", expand=True, pady=(24, 0))
        cards.columnconfigure(0, weight=1, uniform="team")
        cards.columnconfigure(1, weight=1, uniform="team")

        self.add_team_card(cards, 0, "1팀(국기초)", "team1", PALETTE['surface'])
        self.add_team_card(cards, 1, "2팀(차상위, 기초연금 등)", "team2", PALETTE['surface'])

        footer = tk.Frame(outer, background=PALETTE['window'])
        footer.pack(fill="x", pady=(20, 0))
        recent = load_recent_team()
        recent_text = {"team1": "1팀(국기초)", "team2": "2팀(차상위, 기초연금 등)"}.get(recent, "선택 기록 없음")
        tk.Label(footer, text=f"최근 사용: {recent_text}", background=PALETTE['window'], foreground=PALETTE['muted'], font=("Malgun Gothic", 10)).pack(side="left")
        ttk.Button(footer, text="프로그램 종료", style="Exit.TButton", command=self.root.destroy).pack(side="right")

    def add_team_card(self, parent, column, title, team, color):
        card = tk.Frame(parent, padx=24, pady=22, background=color, highlightbackground=PALETTE['border'], highlightthickness=1)
        card.grid(row=0, column=column, sticky="nsew", padx=(0, 8) if column == 0 else (8, 0))
        tk.Label(card, text=title, background=PALETTE['mint'], foreground=PALETTE['primary'], font=("Malgun Gothic", 14, "bold"), anchor="center", pady=10).pack(fill='x', pady=(0, 20))
        ttk.Button(
            card,
            text="사용하기",
            style="Use.TButton",
            command=lambda: self.select_team(team),
        ).pack(fill="x")

    def select_team(self, team):
        self.selected_team = team
        save_recent_team(team)
        self.root.destroy()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = max(0, (self.root.winfo_screenwidth() - width) // 2)
        y = max(0, (self.root.winfo_screenheight() - height) // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")


def run_selected_app(team):
    try:
        if team == "team1":
            from team1_app import App
        elif team == "team2":
            from team2_app import App
        else:
            return

        root = tk.Tk()
        app = App(root)
        root.mainloop()
        return bool(getattr(app, "return_to_launcher", False))
    except Exception as exc:
        error_root = tk.Tk()
        error_root.withdraw()
        messagebox.showerror("실행 오류", f"프로그램을 여는 중 오류가 발생했습니다.\n\n{exc}")
        error_root.destroy()
        return False


def main():
    while True:
        root = tk.Tk()
        launcher = Launcher(root)
        root.mainloop()
        if not launcher.selected_team:
            break
        if not run_selected_app(launcher.selected_team):
            break


if __name__ == "__main__":
    main()
