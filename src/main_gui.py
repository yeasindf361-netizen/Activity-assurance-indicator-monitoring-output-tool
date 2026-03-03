# main_gui.py - 极简现代化界面
import sys
import os
import re
import threading
import traceback
import inspect
from tkinter import filedialog
from tkinter import messagebox as tkmsg
import customtkinter as ctk
from customtkinter import CTk, CTkFrame, CTkLabel, CTkEntry, CTkButton, CTkTextbox, CTkScrollableFrame, CTkCheckBox
from CTkMessagebox import CTkMessagebox
from kpi_tool.config.constants import APP_VERSION

THEME = {
    'primary':       '#3B82F6',
    'primary_hover': '#2563EB',
    'primary_light': '#EFF6FF',
}

# --- message box helper ---
def _safe_msg(title: str, message: str, icon: str | None = None):
    try:
        CTkMessagebox(title=title, message=message, icon=icon or "info").get()
        return
    except Exception:
        pass
    try:
        if icon in ("cancel", "error"):
            tkmsg.showerror(title, message)
        elif icon in ("warning",):
            tkmsg.showwarning(title, message)
        else:
            tkmsg.showinfo(title, message)
    except Exception:
        try:
            print(f"[{title}] {message}")
        except Exception:
            pass

import backend_logic
import license_manager

try:
    print(f"[SELFCHK] backend_logic_file={getattr(backend_logic, '__file__', None)}")
    print(f"[SELFCHK] backend_logic_main_sig={inspect.signature(backend_logic.main) if hasattr(backend_logic,'main') else 'NO main'}")
    print(f"[SELFCHK] has_get_activity_list={hasattr(backend_logic, 'get_activity_list')}")
except Exception as _e:
    print(f"[SELFCHK] failed: {_e}")

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

def get_app_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# ==========================================
# 后台线程工作函数（保留原有结构）
# ==========================================
def run_task(paths, target_activities=None, log_callback=None, progress_callback=None, finished_callback=None):
    try:
        log_callback(">>> 正在初始化后台逻辑...")

        backend_logic.LIST_FILE_PATH = paths['list']
        backend_logic.RAW_DIR_4G = paths['4g']
        backend_logic.RAW_DIR_5G = paths['5g']

        app_dir = get_app_dir()
        out_dir_default = os.path.join(app_dir, "输出结果")
        log_dir_default = os.path.join(app_dir, "logs")
        backend_logic.OUTPUT_DIR = out_dir_default
        backend_logic.LOG_DIR = log_dir_default

        os.makedirs(out_dir_default, exist_ok=True)
        os.makedirs(log_dir_default, exist_ok=True)

        def gui_log(msg):
            if log_callback: log_callback(msg)

        def gui_progress(percent):
            if progress_callback: progress_callback(percent)
            gui_log(f"[{percent}%] 处理中...")

        backend_logic.log_info = gui_log
        backend_logic.print_progress = gui_progress

        log_callback(f">>> 任务开始运行... (目标活动: {target_activities if target_activities else '全部'})")

        main_fn = getattr(backend_logic, "main", None)
        if main_fn is None:
            raise AttributeError("backend_logic.main 不存在")

        called = False
        try:
            sig = inspect.signature(main_fn)
            params = sig.parameters

            if ("target_activities" in params) or any(
                p.kind == inspect.Parameter.VAR_KEYWORD for p in params.values()
            ):
                try:
                    main_fn(target_activities=target_activities)
                    called = True
                except TypeError:
                    called = False

            if not called:
                if len(params) >= 1:
                    main_fn(target_activities)
                    called = True
                else:
                    main_fn()
                    called = True

        except Exception:
            try:
                main_fn(target_activities=target_activities)
                called = True
            except TypeError:
                try:
                    main_fn(target_activities)
                    called = True
                except TypeError:
                    main_fn()
                    called = True

        if not called:
            raise RuntimeError("调用 backend_logic.main 失败")

        progress_callback(100)
        finished_callback(True, "处理完成")

    except Exception as e:
        import traceback
        err_msg = traceback.format_exc()
        log_callback(f"发生错误:\n{err_msg}")
        finished_callback(False, str(e))

# ==========================================
# 主窗口
# ==========================================
class App(CTk):
    def __init__(self):
        super().__init__()

        # 授权验证
        is_valid, machine_code = license_manager.check_license()
        if not is_valid:
            CTkMessagebox(title="软件未授权", message="您的软件授权已过期或未激活。",
                          icon="cancel", option_1="退出").get()
            sys.exit()

        # 图标
        icon_path = os.path.join(get_app_dir(), "logo.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        # 加载火箭图标
        rocket_icon_path = os.path.join(get_app_dir(), "rocket_icon.png")
        self.rocket_icon = None
        if os.path.exists(rocket_icon_path):
            try:
                from PIL import Image
                self.rocket_icon = ctk.CTkImage(
                    light_image=Image.open(rocket_icon_path),
                    dark_image=Image.open(rocket_icon_path),
                    size=(24, 24)
                )
            except Exception:
                pass

        self.title(f"通信指标自动化通报工具 {APP_VERSION}")
        self.geometry("900x620")
        self.minsize(900, 520)
        self.configure(fg_color="#F9F9FA")

        # 白色圆角卡片主框架
        main_card = CTkFrame(self, fg_color="#FFFFFF", corner_radius=12)
        main_card.pack(fill="both", expand=True, padx=24, pady=24)

        content = CTkScrollableFrame(main_card, fg_color="#FFFFFF")
        content.pack(fill="both", expand=True, padx=16, pady=(16, 8))

        # 1. 数据源配置
        group1 = CTkFrame(content, fg_color="transparent")
        group1.pack(fill="x", pady=(0, 20))
        CTkLabel(group1, text="数据源配置", font=ctk.CTkFont(size=15, weight="bold"),
                 text_color="#1E293B").pack(anchor="w", pady=(4, 10))

        self.inputs = {}
        self.add_path_selector(group1, "保障清单文件:", "list", False)
        self.add_path_selector(group1, "4G指标文件夹:", "4g", True)
        self.add_path_selector(group1, "5G指标文件夹:", "5g", True)

        app_dir = get_app_dir()
        default_list = os.path.join(app_dir, '保障小区清单', '保障小区清单.xlsx')
        default_4g = os.path.join(app_dir, '网管指标', '4G指标')
        default_5g = os.path.join(app_dir, '网管指标', '5G指标')
        if not self.inputs['list'].get(): self.inputs['list'].insert(0, default_list)
        if not self.inputs['4g'].get(): self.inputs['4g'].insert(0, default_4g)
        if not self.inputs['5g'].get(): self.inputs['5g'].insert(0, default_5g)

        # 2. 活动筛选
        group2 = CTkFrame(content, fg_color="transparent")
        group2.pack(fill="x", pady=(0, 20))
        CTkLabel(group2, text="活动筛选", font=ctk.CTkFont(size=15, weight="bold"),
                 text_color="#1E293B").pack(anchor="w", pady=(4, 10))

        # 按钮组
        btn_group = CTkFrame(group2, fg_color="transparent")
        btn_group.pack(anchor="w", pady=(0, 8))

        btn_load = CTkButton(btn_group, text="读取清单中的活动", width=140,
                             command=self.load_activities,
                             fg_color=THEME['primary'], hover_color=THEME['primary_hover'])
        btn_load.pack(side="left", padx=(0, 8))

        btn_select_all = CTkButton(btn_group, text="全选", width=70,
                                   command=self.select_all_activities,
                                   fg_color=THEME['primary_light'], text_color=THEME['primary'],
                                   hover_color="#DBEAFE")
        btn_select_all.pack(side="left", padx=(0, 8))

        btn_deselect_all = CTkButton(btn_group, text="取消全选", width=80,
                                     command=self.deselect_all_activities,
                                     fg_color="#F1F5F9", text_color="#64748B",
                                     hover_color="#E2E8F0")
        btn_deselect_all.pack(side="left")


        self.checkboxes = {}
        self.act_frame = CTkScrollableFrame(group2, height=80, fg_color="#F9F9FA", corner_radius=8)
        self.act_frame.pack(fill="x")

        # 3. 开始按钮（幽灵风格）
        btn_frame = CTkFrame(content, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 16))

        self.btn_run = CTkButton(
            btn_frame, text="开始生成通报",
            image=self.rocket_icon if self.rocket_icon else None,
            compound="left",
            font=ctk.CTkFont(size=14, weight="bold"), height=42,
            fg_color="transparent", border_width=2,
            border_color=THEME['primary'], text_color=THEME['primary'],
            hover_color=THEME['primary_light'], command=self.start_task,
        )
        self.btn_run.pack(fill="x")

        # 4. 进度条 + 日志（弱化）
        self.progress = ctk.CTkProgressBar(content, progress_color=THEME['primary'])
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(0, 8))

        self.log_box = CTkTextbox(content, height=100,
                                  font=ctk.CTkFont(size=11), text_color="#94A3B8",
                                  fg_color="#F9F9FA", corner_radius=8)
        self.log_box.pack(fill="both", expand=True)
        self.log_box.insert("0.0", "运行日志:\n")

    def add_path_selector(self, parent, label, key, is_folder):
        frame = CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=4)
        CTkLabel(frame, text=label, width=120, text_color="#475569").pack(side="left")
        entry = CTkEntry(frame, fg_color="#F9F9FA", border_color="#E2E8F0")
        entry.pack(side="left", fill="x", expand=True, padx=8)
        btn = CTkButton(frame, text="浏览", width=72,
                        fg_color=THEME['primary_light'], text_color=THEME['primary'],
                        hover_color="#DBEAFE",
                        command=lambda: self.select_path(entry, is_folder))
        btn.pack(side="right")
        self.inputs[key] = entry

    def select_path(self, entry, is_folder):
        if is_folder:
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    def load_activities(self):
        list_path = self.inputs['list'].get()
        if not list_path or not os.path.exists(list_path):
            CTkMessagebox(title="提示", message="请先选择有效的保障清单文件！", icon="warning").get()
            return

        try:
            if hasattr(backend_logic, "get_activity_list"):
                acts = backend_logic.get_activity_list(list_path)
            else:
                import pandas as pd
                df = pd.read_excel(list_path, sheet_name=0, dtype=str)
                col = None
                for c in df.columns:
                    if isinstance(c, str) and ("活动" in c):
                        col = c
                        break
                if col is None:
                    acts = []
                else:
                    s = df[col].fillna("").astype(str).str.strip()
                    acts = [x for x in s.tolist() if x]
                    seen=set(); acts=[x for x in acts if not (x in seen or seen.add(x))]

            for widget in self.act_frame.winfo_children():
                widget.destroy()
            self.checkboxes.clear()

            if acts:
                cols = 4
                for i, act in enumerate(acts):
                    var = ctk.BooleanVar()
                    cb = CTkCheckBox(self.act_frame, text=act, variable=var,
                                     checkbox_width=18, checkbox_height=18)
                    cb.grid(row=i // cols, column=i % cols, sticky="w", padx=4, pady=2)
                    self.checkboxes[act] = var
                self.log_box.insert("end", f"成功读取 {len(acts)} 个活动\n")
            else:
                self.log_box.insert("end", "未在清单中找到活动列\n")
        except Exception as e:
            _safe_msg("错误", f"读取失败: {str(e)}", icon="cancel")

    def select_all_activities(self):
        """全选所有活动"""
        if not self.checkboxes:
            CTkMessagebox(title="提示", message="请先读取活动清单！", icon="warning").get()
            return

        for var in self.checkboxes.values():
            var.set(True)
        self.log_box.insert("end", f"已全选 {len(self.checkboxes)} 个活动\n")

    def deselect_all_activities(self):
        """取消全选所有活动"""
        if not self.checkboxes:
            return

        for var in self.checkboxes.values():
            var.set(False)
        self.log_box.insert("end", "已取消全部活动选择\n")

    def start_task(self):
        paths = {k: v.get() for k, v in self.inputs.items()}
        if not all(paths.values()):
            CTkMessagebox(title="提示", message="请先选择所有必要的文件路径！", icon="warning").get()
            return

        target_acts = [act for act, var in self.checkboxes.items() if var.get()]
        if not target_acts:
            target_acts = None

        self.btn_run.configure(state="disabled", text="处理中...",
                               fg_color=THEME['primary_light'], text_color="#94A3B8")
        self.log_box.delete("1.0", "end")
        self.progress.set(0)

        threading.Thread(target=run_task, args=(
            paths, target_acts,
            lambda msg: self.after(0, lambda: self.log_box.insert("end", msg + "\n")),
            lambda p: self.after(0, lambda: self.progress.set(p/100)),
            lambda success, msg: self.after(0, lambda: self.on_finished(success, msg))
        ), daemon=True).start()

    def on_finished(self, success, msg):
        self.btn_run.configure(state="normal", text="开始生成通报",
                               fg_color="transparent", text_color=THEME['primary'])
        if success:
            CTkMessagebox(title="完成", message="通报生成完毕！", icon="info").get()
        else:
            _safe_msg("失败", f"运行出错: {msg}", icon="cancel")

if __name__ == "__main__":
    app = App()
    app.mainloop()
