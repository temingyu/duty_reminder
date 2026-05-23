# 此软件使用AI编写
# ai写的可能会有亿点点问题（ 
# 作者本人并不会写代码，全由AI生成，当然，也欢迎大家一起参与到此软件的维护当中，此软件将会在GitHub上共享源码
# 如果你发现了问题，可以发邮件或者提交issue给我，虽然不一定会修 嘻嘻（学业繁忙）
# 先这样吧，也没啥好说上的了（
# 哦对，我的邮箱:temingmail@163.com
# 就这样   
import ctypes
import tkinter as tk
import tkinter.colorchooser as colorchooser      # 系统颜色选择器
import pyttsx3.drivers
import pyttsx3.drivers.sapi5
import pystray._win32
from tkinter import ttk, messagebox, filedialog
import random
import pyttsx3                                   # 文字转语音引擎
import json
import os
import sys
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler   # 定时任务调度
import threading
from PIL import Image, ImageDraw                # 用于生成托盘图标
import pystray                                  # 系统托盘
from pystray import MenuItem as item
import atexit
import pandas as pd                             # Excel 导入
import winreg                                   # Windows 注册表操作（开机自启）
import logging                                  # 日志模块
from logging.handlers import RotatingFileHandler

# ========== 全局日志配置：捕获所有控制台输出 ==========
class StreamToLogger:
    def __init__(self, logger, log_level=logging.INFO):
        self.logger = logger
        self.log_level = log_level
        self.linebuf = ''

    def write(self, buf):
        try:
            # 防止递归调用：如果 logger 的 handler 正在处理本条记录，直接返回
            if getattr(self.logger, '_in_emit', False):
                return
            for line in buf.rstrip().splitlines():
                if line.strip():
                    self.logger._in_emit = True   # 标记正在处理
                    self.logger.log(self.log_level, line.rstrip())
                    self.logger._in_emit = False
        except AttributeError:
            # 如果 self.logger 为 None 或没有 _in_emit 属性，忽略
            pass
        except Exception:
            pass

    def flush(self):
        pass
    
def setup_global_logging():
    """全局日志配置：捕获所有print和异常输出"""
    # 确定日志文件路径（程序同目录的 logs 目录下）
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
    
    log_dir = os.path.join(app_dir, "logs")
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, "duty_reminder.log")
    
    # 创建日志记录器
    logger = logging.getLogger("DutyReminder")
    logger.setLevel(logging.DEBUG)
    
    # 避免重复添加处理器
    if logger.hasHandlers():
        logger.handlers.clear()
    
    # 创建文件处理器（支持日志轮转，最大10MB，保留5个备份）
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=5,
        encoding='utf-8'
    )
    
    # 创建控制台处理器（保留控制台输出）
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG)
    
    # 创建格式器
    formatter = logging.Formatter(
        '[%(asctime)s] %(levelname)-8s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 设置处理器格式
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # 添加处理器到记录器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    # 重定向标准输出和错误到日志
    sys.stdout = StreamToLogger(logger, logging.INFO)
    sys.stderr = StreamToLogger(logger, logging.ERROR)
    
    # 捕获未处理的异常
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            # 调用默认处理程序处理Ctrl+C
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logger.error("【未捕获异常】程序发生未处理的异常:", exc_info=(exc_type, exc_value, exc_traceback))
    
    sys.excepthook = handle_exception
    
    logger.info("=" * 80)
    logger.info("【日志系统】日志系统初始化完成")
    logger.info(f"【日志系统】日志文件路径: {log_file}")
    logger.info(f"【日志系统】所有控制台输出和异常将被记录到日志文件")
    logger.info("=" * 80)
    
    return logger

# 初始化全局日志（在类定义之前）
global_logger = setup_global_logging()

class DutyReminderApp:
    """值日提醒应用程序主类"""

    def __init__(self):
        self.logger = global_logger
        self.logger.info("=" * 80)
        self.logger.info("【程序启动】值日提醒系统开始初始化...")
        self.logger.info("=" * 80)
        
        # 设置 Windows 任务栏独立图标（避免和 Python 主程序图标重叠）
        app_id = 'mycompany.myapp.subproduct.version'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

        # 语音引擎状态
        self.tts_initialized = False
        self.tts_lock = threading.Lock()         # 防止并发播放语音

        # 加载配置文件（若无则创建默认值）
        self.logger.info("【配置加载】开始加载配置文件...")
        self.load_data()
        self.logger.info("【配置加载】配置文件加载完成")

        # 初始化 TTS 引擎（目前被注释掉，如需语音请取消注释）
        self.tts_engine = None
        # self.setup_tts()

        # 创建隐藏的 Tk 根窗口（用于管理子窗口）
        self.root = tk.Tk()
        self.root.withdraw()                     # 不显示主窗口
        ico = self.get_icon_path()
        try:
            self.root.iconbitmap(ico)            # 设置根窗口图标（影响任务栏）
            self.logger.info("【窗口图标】成功设置窗口图标")
        except Exception as e:
            self.logger.warning(f"【窗口图标】无法设置窗口图标，原因: {e}")
            print(f"警告：无法设置窗口图标，原因: {e}")

        self.main_window = None                  # 主设置窗口对象

        # 初始化系统托盘图标
        self.logger.info("【系统托盘】开始创建系统托盘图标...")
        self.create_system_tray()
        self.logger.info("【系统托盘】系统托盘图标创建成功")

        # 根据配置创建桌面浮窗
        self.floating_widgets = {}
        self.logger.info("【浮窗创建】开始创建桌面浮窗...")
        self.create_floating_widgets()
        self.logger.info(f"【浮窗创建】成功创建 {len(self.floating_widgets)} 个浮窗")

        # 启动每日定时提醒（基于第一个任务的设置）
        self.logger.info("【定时任务】开始启动每日定时提醒...")
        self.start_scheduler()
        self.logger.info("【定时任务】每日定时提醒启动成功")

        # 程序退出时自动保存数据
        atexit.register(self.cleanup)
        
        self.logger.info("=" * 80)
        self.logger.info("【程序启动】值日提醒系统初始化完成，进入主循环")
        self.logger.info("=" * 80)

    # ========== 配置文件加载与保存（保持原有代码）==========
    def load_data(self):
        """从 JSON 文件加载任务配置，如果不存在或损坏则生成默认配置"""
        # 确定配置文件路径（程序同目录的 duty_config.json）
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_file = os.path.join(app_dir, "duty_config.json")
        
        self.logger.info(f"【配置加载】配置文件路径: {self.config_file}")

        # 单个任务的默认字段模板（新增 floating_locked 控制浮窗锁定）
        default_task = {
            'name': '',
            'duty_list': [],
            'starting_duty_index': 0,
            'reminder_hour': 8,
            'reminder_minute': 0,
            'always_on_top': True,
            'floating_x': None,
            'floating_y': None,
            'voice_enabled': True,
            'window_scale_factor': 1.0,
            'font_size_factor': 1.0,
            'floating_enabled': True,
            'custom_voice_template': '现在是%H:%M，明天是%Y年%m月%d日，请%DUTY%同学记得完成明天的%TASK%任务！',
            'override_person': None,            # 临时覆盖的人员
            'override_until': None,              # 覆盖截止日期 (YYYY-MM-DD)
            'floating_locked': False,            # 浮窗位置锁定（新增）
            'duty_label_color': 'blue'            # 浮窗人员字体颜色
        }

        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.task_groups = data.get('task_groups', {})
                self.logger.info(f"【配置加载】成功从文件加载配置: {len(self.task_groups)} 个任务")
            except (json.JSONDecodeError, KeyError) as e:
                self.logger.error(f"【配置加载】配置文件解析失败: {e}，将使用默认配置")
                self.task_groups = {}
        else:
            self.logger.warning(f"【配置加载】配置文件不存在: {self.config_file}，将创建默认配置")
            self.task_groups = {}

        # 如果配置为空，生成三个默认任务
        if not self.task_groups:
            self.logger.info("【配置加载】配置为空，生成默认任务配置")
            self.task_groups = {
                'task1': {**default_task, 'name': '刷勺', 'duty_list': ['张三', '李四', '王五', '赵六']},
                'task2': {**default_task, 'name': '打扫', 'duty_list': ['小明', '小红', '小刚', '小美']},
                'task3': {**default_task, 'name': '黑板', 'duty_list': ['阿强', '阿华', '阿丽', '阿军']}
            }
            self.save_data()
        else:
            # 补充缺失的字段（兼容旧版本配置文件）
            updated = False
            for task_key, task_data in self.task_groups.items():
                if 'name' not in task_data:
                    task_data['name'] = f'任务{task_key}'
                    updated = True
                if 'duty_list' not in task_data:
                    task_data['duty_list'] = []
                    updated = True
                for key, val in default_task.items():
                    if key not in task_data:
                        task_data[key] = val
                        updated = True
                # 确保起始索引合法
                if task_data['duty_list'] and (task_data['starting_duty_index'] >= len(task_data['duty_list']) or task_data['starting_duty_index'] < 0):
                    task_data['starting_duty_index'] = 0
                    updated = True
            if updated:
                self.logger.info("【配置加载】检测到旧版本配置，已自动补充缺失字段")
                self.save_data()

    def save_data(self):
        """将当前任务配置序列化保存到 JSON 文件（原子写入）"""
        import copy, tempfile, shutil
        data = {'task_groups': copy.deepcopy(self.task_groups)}
        try:
            # 先写入临时文件，再重命名，避免写入中断导致文件损坏
            temp_fd, temp_path = tempfile.mkstemp(suffix='.json', dir=os.path.dirname(self.config_file))
            with os.fdopen(temp_fd, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            shutil.move(temp_path, self.config_file)
            self.logger.info(f"【配置保存】配置保存成功: {len(self.task_groups)} 个任务已保存到 {self.config_file}")
            return True
        except Exception as e:
            self.logger.error(f"【配置保存】保存配置失败: {e}")
            print(f"保存配置失败: {e}")
            return False

    # ========== 工具方法（保持原有代码）==========
    def get_icon_path(self, filename='reminder_icon.ico'):
        """获取图标文件的路径（支持 PyInstaller 打包）"""
        if getattr(sys, 'frozen', False):
            base_dir = sys._MEIPASS          # 打包后的临时资源目录
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, filename)

    # def setup_tts(self):
    #     """初始化 TTS 语音引擎（未启用）"""
    #     try:
    #         self.tts_engine = pyttsx3.init()
    #         voices = self.tts_engine.getProperty('voices')
    #         if voices:
    #             for voice in voices:
    #                 if 'Chinese' in voice.name or 'zh-CN' in voice.id or 'Microsoft' in voice.name:
    #                     self.tts_engine.setProperty('voice', voice.id)
    #                     break
    #         self.tts_engine.setProperty('rate', 180)
    #         self.tts_engine.setProperty('volume', 0.9)
    #         self.tts_initialized = True
    #     except Exception as e:
    #         print(f"TTS初始化失败: {e}")
    #         self.tts_initialized = False

    def get_current_day_index(self):
        """计算从基准日期 2024-01-01 到今天的绝对天数差，用于值日轮换计算"""
        base_date = datetime(2024, 1, 1).date()
        return (datetime.now().date() - base_date).days

    def is_override_active(self, task_key):
        """检查当前任务是否处于"覆盖值日"状态（且未过期）"""
        task_data = self.task_groups[task_key]
        if task_data.get('override_person') and task_data.get('override_until'):
            try:
                until = datetime.strptime(task_data['override_until'], "%Y-%m-%d").date()
                return datetime.now().date() <= until
            except:
                pass
        return False

    # ========== 值日人员计算（保持原有代码）==========
    def get_current_duty_person(self, task_key):
        """获取当前应该显示的值日人员（考虑提醒时间节点）"""
        if self.is_override_active(task_key):
            return self.task_groups[task_key]['override_person']   # 覆盖模式直接返回

        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            return "无值日人员"
        now = datetime.now()
        # 今天的提醒时间点
        reminder_time = now.replace(hour=task_data['reminder_hour'],
                                    minute=task_data['reminder_minute'],
                                    second=0, microsecond=0)
        current_day_index = self.get_current_day_index()
        actual_index = (task_data['starting_duty_index'] + current_day_index) % len(task_data['duty_list'])
        # 如果当前时间已经过了提醒时间，则显示下一个人（即明天的值日者）
        if now >= reminder_time:
            next_index = (actual_index + 1) % len(task_data['duty_list'])
            return task_data['duty_list'][next_index] if task_data['duty_list'] and next_index < len(task_data['duty_list']) else "无值日人员"
        else:
            return task_data['duty_list'][actual_index] if task_data['duty_list'] and actual_index < len(task_data['duty_list']) else "无值日人员"

    def get_tomorrow_duty_person(self, task_key):
        """获取明天的值日人员（不受提醒时间影响）"""
        if self.is_override_active(task_key):
            return self.task_groups[task_key]['override_person']
        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            return "无值日人员"
        tomorrow = datetime.now() + timedelta(days=1)
        base_date = datetime(2024, 1, 1).date()
        tomorrow_day_index = (tomorrow.date() - base_date).days
        actual_index = (task_data['starting_duty_index'] + tomorrow_day_index) % len(task_data['duty_list'])
        return task_data['duty_list'][actual_index] if actual_index < len(task_data['duty_list']) else "无值日人员"

    def get_current_or_tomorrow_label(self, task_key):
        """根据提醒时间返回标签文字："今日某某"或"明日某某"""
        task_data = self.task_groups[task_key]
        now = datetime.now()
        reminder_time = now.replace(hour=task_data['reminder_hour'],
                                    minute=task_data['reminder_minute'],
                                    second=0, microsecond=0)
        return f"明日{task_data['name']}" if now >= reminder_time else f"今日{task_data['name']}"

    # ========== 桌面浮窗管理（保持原有代码，增加锁定判断）==========
    def create_floating_widgets(self):
        """为所有启用浮窗的任务创建桌面浮窗"""
        self.logger.info("【浮窗管理】开始创建桌面浮窗...")
        created_count = 0
        for task_key, task_data in self.task_groups.items():
            if task_data['floating_enabled']:
                self.create_single_floating_widget(task_key, task_data)
                created_count += 1
                self.logger.debug(f"【浮窗管理】已创建浮窗: {task_data['name']}")
        self.logger.info(f"【浮窗管理】桌面浮窗创建完成，共创建 {created_count} 个浮窗")

    def create_single_floating_widget(self, task_key, task_data):
        """创建一个无边框、可拖拽的浮动小窗口"""
        floating_widget = tk.Toplevel(self.root)
        floating_widget.title(f"当前{task_data['name']}")
        floating_widget.geometry("220x100")                # 基础尺寸
        floating_widget.overrideredirect(True)             # 无标题栏
        floating_widget.attributes('-topmost', task_data['always_on_top'])  # 置顶
        try:
            floating_widget.attributes('-alpha', 0.9)      # 半透明
        except tk.TclError:
            pass
        # 绑定拖拽事件
        floating_widget.bind('<Button-1>', lambda e, tk=task_key: self.start_drag(e, tk))
        floating_widget.bind('<B1-Motion>', lambda e, tk=task_key: self.drag_window(e, tk))
        floating_widget.bind('<ButtonRelease-1>', lambda e, tk=task_key: self.save_position_on_release(e, tk))

        floating_widget.task_key = task_key
        self.create_widget_ui(floating_widget, task_key)   # 创建内部显示内容
        self.floating_widgets[task_key] = floating_widget
        self.update_floating_size_and_font(task_key)        # 应用缩放设置
        self.update_floating_display(task_key)              # 立即更新显示
        floating_widget.after(60000, lambda tk=task_key: self.periodic_update(tk))  # 每分钟更新一次
        self.set_initial_position(task_key)                 # 设置初始位置

        def bind_drag_events(self, floating_widget, task_key):
            """根据锁定状态绑定或解绑拖拽事件"""
            task_data = self.task_groups.get(task_key)
            if task_data and task_data.get('floating_locked', False):
                # 锁定：解绑所有拖拽事件
                floating_widget.unbind('<Button-1>')
                floating_widget.unbind('<B1-Motion>')
                floating_widget.unbind('<ButtonRelease-1>')
            else:
                # 未锁定：绑定拖拽事件
                floating_widget.bind('<Button-1>', lambda e, tk=task_key: self.start_drag(e, tk))
                floating_widget.bind('<B1-Motion>', lambda e, tk=task_key: self.drag_window(e, tk))
            floating_widget.bind('<ButtonRelease-1>', lambda e, tk=task_key: self.save_position_on_release(e, tk))

    def set_initial_position(self, task_key):
        """设置浮窗初始位置：默认右上角，若已保存位置则恢复"""
        floating_widget = self.floating_widgets[task_key]
        task_data = self.task_groups[task_key]
        floating_widget.update_idletasks()  # 确保已获取正确尺寸
        sw = floating_widget.winfo_screenwidth()
        sh = floating_widget.winfo_screenheight()
        w = int(220 * task_data['window_scale_factor'])
        h = int(100 * task_data['window_scale_factor'])
        # 有保存位置且仍在屏幕内则恢复
        if task_data['floating_x'] is not None and task_data['floating_y'] is not None:
            if 0 <= task_data['floating_x'] <= sw - w and 0 <= task_data['floating_y'] <= sh - h:
                floating_widget.geometry(f"{w}x{h}+{task_data['floating_x']}+{task_data['floating_y']}")
                return
        # 没有保存位置时，按任务编号错开排列在右上角
        offset_x = (ord(task_key[-1]) - ord('1')) * 230
        x = sw - w - 20 - offset_x
        y = 20
        floating_widget.geometry(f"{w}x{h}+{x}+{y}")

    def create_widget_ui(self, floating_widget, task_key):
        """构建浮窗内部的 UI：标题标签 + 值日人员标签"""
        task_data = self.task_groups[task_key]
        color = task_data.get('duty_label_color', 'blue')  # 自定义字体颜色
        main_frame = tk.Frame(floating_widget, bg='#f0f0f0', bd=2, relief='solid')
        main_frame.pack(fill='both', expand=True, padx=1, pady=1)
        # 上方标签：今日/明日任务名
        label_text = tk.StringVar()
        label_text.set(self.get_current_or_tomorrow_label(task_key))
        title_label = tk.Label(main_frame, textvariable=label_text, bg='#f0f0f0',
                              font=("微软雅黑", 10, "bold"), fg='gray')
        title_label.pack(pady=(5, 0))
        # 下方标签：值日人员名字
        duty_label = tk.Label(main_frame, text="", bg='#f0f0f0',
                             font=("微软雅黑", 12, "bold"), fg=color)
        duty_label.pack(pady=(0, 5))
        # 双击浮窗打开主设置窗口
        main_frame.bind('<Double-Button-1>', lambda e: self.show_main_window())
        # 保存引用以便后续更新
        floating_widget.label_text = label_text
        floating_widget.duty_label = duty_label
        floating_widget.title_label = title_label

    def start_drag(self, event, task_key):
        """记录拖拽起始位置（锁定状态下不执行）"""
        task_data = self.task_groups.get(task_key)
        if task_data and task_data.get('floating_locked', False):
            return   # 锁定状态下不记录起始位置
        floating_widget = self.floating_widgets[task_key]
        floating_widget.x = event.x
        floating_widget.y = event.y

    def drag_window(self, event, task_key):
        """根据鼠标移动更新浮窗位置（锁定状态下不移动）"""
        task_data = self.task_groups.get(task_key)
        if task_data and task_data.get('floating_locked', False):
            return   # 锁定状态下不移动
        floating_widget = self.floating_widgets[task_key]
        task_data = self.task_groups[task_key]  # 这里重复了，但为了清晰保留
        x = floating_widget.winfo_x() + event.x - floating_widget.x
        y = floating_widget.winfo_y() + event.y - floating_widget.y
        floating_widget.geometry(f"+{x}+{y}")
        task_data['floating_x'] = x
        task_data['floating_y'] = y

    def save_position_on_release(self, event, task_key):
        """鼠标释放时自动保存位置到配置文件（防止异常退出丢失）"""
        self.logger.debug(f"【浮窗管理】保存浮窗位置: {task_key} -> ({self.task_groups[task_key]['floating_x']}, {self.task_groups[task_key]['floating_y']})")
        self.save_data()

    def update_floating_display(self, task_key):
        """更新浮窗内显示的今日/明日人员和标签"""
        if task_key not in self.floating_widgets:
            return
        floating_widget = self.floating_widgets[task_key]
        task_data = self.task_groups[task_key]
        color = task_data.get('duty_label_color', 'blue')
        floating_widget.duty_label.config(text=self.get_current_duty_person(task_key), fg=color)
        floating_widget.label_text.set(self.get_current_or_tomorrow_label(task_key))

    def periodic_update(self, task_key):
        """定时更新浮窗显示（每分钟一次）"""
        self.update_floating_display(task_key)
        if task_key in self.floating_widgets:
            self.floating_widgets[task_key].after(60000, lambda tk=task_key: self.periodic_update(tk))

    def update_floating_size_and_font(self, task_key):
        """根据缩放因子调整浮窗尺寸和内部字体大小"""
        if task_key not in self.floating_widgets:
            return
        floating_widget = self.floating_widgets[task_key]
        task_data = self.task_groups[task_key]
        w = int(220 * task_data['window_scale_factor'])
        h = int(100 * task_data['window_scale_factor'])
        title_fs = max(8, int(10 * task_data['window_scale_factor'] * task_data['font_size_factor']))
        duty_fs = max(10, int(12 * task_data['window_scale_factor'] * task_data['font_size_factor']))
        floating_widget.geometry(f"{w}x{h}")
        if hasattr(floating_widget, 'title_label'):
            floating_widget.title_label.config(font=("微软雅黑", title_fs, "bold"))
        if hasattr(floating_widget, 'duty_label'):
            floating_widget.duty_label.config(font=("微软雅黑", duty_fs, "bold"))

    # ========== 系统托盘（保持原有代码）==========
    def create_system_tray(self):
        """创建 Windows 系统托盘图标和右键菜单"""
        ico_path = self.get_icon_path('reminder_icon.ico')
        try:
            image = Image.open(ico_path)
        except:
            # 如果找不到图标，生成一个简易彩色块
            image = Image.new('RGB', (64, 64), color='lightblue')
            draw = ImageDraw.Draw(image)
            draw.rectangle([10, 10, 54, 54], outline='black', width=2)
            draw.text((20, 20), '值', fill='black', font_size=30)
        menu = (item('主页面', self.show_main_window),
                item('退出', self.quit_app))
        self.icon = pystray.Icon("值日提醒", image, "值日提醒", menu=menu)
        self.icon.run_detached()  # 在独立线程运行

    # ========== 主设置窗口（保持原有代码，增加锁定按钮）==========
    def show_main_window(self, icon=None, item=None):
        """显示或恢复主设置窗口"""
        if self.main_window and self.main_window.winfo_exists():
            self.main_window.deiconify()
            self.main_window.lift()
            self.main_window.focus_force()
            self.update_main_window_display()
            self.logger.info("【窗口管理】主设置窗口已恢复显示")
        else:
            self.main_window = self.create_main_window()
            self.main_window.deiconify()
            self.main_window.lift()
            self.main_window.focus_force()
            self.logger.info("【窗口管理】主设置窗口已创建并显示")

    def create_main_window(self):
        """构建主设置窗口界面（紧凑布局，适配低分辨率）"""
        window = tk.Toplevel(self.root)
        window.title("值日提醒")
        window.geometry("640x680")          # 优化尺寸，适配大部分屏幕
        window.resizable(True, True)
        # 设置窗口图标
        ico = self.get_icon_path()
        try:
            window.iconbitmap(ico)
        except Exception as e:
            self.logger.warning(f"【窗口管理】主窗口图标失败: {e}")
            print(f"主窗口图标失败: {e}")

        first_task = next(iter(self.task_groups.values()))
        window.attributes('-topmost', first_task['always_on_top'])  # 初始置顶状态

        main_frame = ttk.Frame(window, padding="8")
        main_frame.pack(fill='both', expand=True)

        # 标题
        ttk.Label(main_frame, text="值日提醒", font=("微软雅黑", 16, "bold")).grid(row=0, column=0, columnspan=4, pady=(0, 10))

        # 任务管理栏（选择任务 + 操作按钮）
        task_select_frame = ttk.LabelFrame(main_frame, text="任务管理", padding="8")
        task_select_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 8))
        self.selected_task_var = tk.StringVar()
        task_names = [data['name'] for data in self.task_groups.values()]
        self.task_combo = ttk.Combobox(task_select_frame, textvariable=self.selected_task_var, values=task_names, state="readonly", width=18)
        self.task_combo.grid(row=0, column=0, padx=(0, 8), pady=5)
        self.task_combo.set(task_names[0])
        self.task_combo.bind('<<ComboboxSelected>>', self.on_task_selection_changed)

        btn_frame = ttk.Frame(task_select_frame)
        btn_frame.grid(row=0, column=1, sticky='e')
        ttk.Button(btn_frame, text="刷新", command=self.update_main_window_display).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="重命名", command=lambda w=window: self.rename_task(w)).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="添加任务", command=lambda w=window: self.add_new_task(w)).pack(side='left', padx=3)
        ttk.Button(btn_frame, text="删除", command=lambda w=window: self.delete_current_task(w)).pack(side='left', padx=3)

        # 左右分栏
        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 4))
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(4, 0))
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # 左栏：值日信息 & 值日顺序
        current_frame = ttk.LabelFrame(left_frame, text="值日信息", padding="8")
        current_frame.pack(fill='both', expand=True, pady=(0, 8))
        self.current_duty_var = tk.StringVar()
        ttk.Label(current_frame, text="当前值日:", font=("微软雅黑", 10, "bold")).pack(anchor='w', pady=(0, 3))
        ttk.Label(current_frame, textvariable=self.current_duty_var, font=("微软雅黑", 14, "bold"), foreground="red").pack(pady=3)
        self.tomorrow_duty_var = tk.StringVar()
        ttk.Label(current_frame, text="明天值日:", font=("微软雅黑", 10, "bold")).pack(anchor='w', pady=(8, 3))
        ttk.Label(current_frame, textvariable=self.tomorrow_duty_var, font=("微软雅黑", 14, "bold"), foreground="blue").pack(pady=3)

        list_frame = ttk.LabelFrame(left_frame, text="值日顺序", padding="8")
        list_frame.pack(fill='both', expand=True)
        self.listbox = tk.Listbox(list_frame, height=8, font=("微软雅黑", 9))
        self.listbox.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.listbox.configure(yscrollcommand=scrollbar.set)

        # 右栏：提醒时间 & 功能控制
        time_frame = ttk.LabelFrame(right_frame, text="提醒时间设置", padding="8")
        time_frame.pack(fill='x', pady=(0, 8))
        time_row = ttk.Frame(time_frame)
        time_row.pack(fill='x', pady=3)
        ttk.Label(time_row, text="提醒时间:").pack(side='left', padx=(0, 5))
        self.hour_var = tk.StringVar(value="8")
        ttk.Spinbox(time_row, from_=0, to=23, width=4, textvariable=self.hour_var).pack(side='left', padx=(0, 3))
        ttk.Label(time_row, text=":").pack(side='left')
        self.minute_var = tk.StringVar(value="00")
        ttk.Spinbox(time_row, from_=0, to=59, width=4, textvariable=self.minute_var).pack(side='left', padx=(3, 8))
        ttk.Button(time_row, text="保存", command=lambda w=window: self.save_reminder_time(w)).pack(side='left')
        self.info_label = ttk.Label(time_frame, text="每天 08:00 提醒", font=("微软雅黑", 8))
        self.info_label.pack(pady=(8, 0))

        control_frame = ttk.LabelFrame(right_frame, text="功能控制", padding="8")
        control_frame.pack(fill='both', expand=True)
        # 采用2列网格排列功能按钮，节省垂直空间
        ctrl_inner = ttk.Frame(control_frame)
        ctrl_inner.pack(fill='both', expand=True)
        buttons = [
            ("随机打乱顺序", lambda w=window: self.shuffle_order(w)),
            ("保存当前顺序", lambda w=window: self.save_current_order(w)),
            ("导入Excel", lambda w=window: self.import_from_excel(w)),
            ("更改当前值日人员", lambda w=window: self.change_current_duty(w)),
            ("添加值日人员", lambda w=window: self.add_duty_person(w)),
            ("移除值日人员", lambda w=window: self.remove_duty_person(w)),
            ("设置覆盖值日", lambda w=window: self.set_override_duty(w)),
            ("取消覆盖", self.cancel_override),
        ]
        for i, (text, cmd) in enumerate(buttons):
            ttk.Button(ctrl_inner, text=text, command=cmd).grid(row=i//2, column=i%2, sticky='ew', padx=2, pady=2)
        ctrl_inner.columnconfigure(0, weight=1)
        ctrl_inner.columnconfigure(1, weight=1)

        # 显示控制开关（增加锁定浮窗按钮）
        switch_frame = ttk.LabelFrame(control_frame, text="显示控制", padding="8")
        switch_frame.pack(fill='x', pady=(8, 0))
        top_status = "开" if first_task['always_on_top'] else "关"
        self.top_btn = ttk.Button(switch_frame, text=f"置顶: {top_status}", command=lambda w=window: self.toggle_always_on_top(w))
        self.top_btn.pack(fill='x', pady=2)
        floating_status = "开" if first_task['floating_enabled'] else "关"
        self.floating_btn = ttk.Button(switch_frame, text=f"浮窗: {floating_status}", command=lambda w=window: self.toggle_floating(w))
        self.floating_btn.pack(fill='x', pady=2)
        # 新增锁定浮窗按钮
        lock_status = "开" if first_task.get('floating_locked', False) else "关"
        self.lock_btn = ttk.Button(switch_frame, text=f"锁定浮窗: {lock_status}",
                                   command=lambda w=window: self.toggle_floating_locked(w))
        self.lock_btn.pack(fill='x', pady=2)
        self.resize_btn = ttk.Button(switch_frame, text="调整浮窗大小", command=lambda w=window: self.open_resize_window(w))
        self.resize_btn.pack(fill='x', pady=2)
        self.font_resize_btn = ttk.Button(switch_frame, text="调整字体大小", command=lambda w=window: self.open_font_resize_window(w))
        self.font_resize_btn.pack(fill='x', pady=2)
        self.color_btn = ttk.Button(switch_frame, text="更改浮窗颜色", command=lambda w=window: self.change_floating_font_color(w))
        self.color_btn.pack(fill='x', pady=2)
        autostart_status = "开" if self.check_autostart() else "关"
        self.autostart_btn = ttk.Button(switch_frame, text=f"开机自启: {autostart_status}", command=self.toggle_autostart)
        self.autostart_btn.pack(fill='x', pady=2)

        # 底部说明
        ttk.Label(control_frame, text="Excel导入：第一列姓名", font=("微软雅黑", 8), foreground="gray").pack(pady=(8, 0))

        # "关于"按钮（右上角）
        about_btn = ttk.Button(window, text="关于", command=self.show_about)
        about_btn.place(relx=1.0, rely=0.0, anchor='ne', x=-8, y=8)

        self.update_main_window_display()
        # 关闭窗口时隐藏而不是销毁
        window.protocol("WM_DELETE_WINDOW", lambda: self.hide_main_window(window))
        return window

    # ========== 浮窗锁定切换 ==========
    def toggle_floating_locked(self, window):
        """切换当前任务浮窗的位置锁定状态"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        task_data['floating_locked'] = not task_data.get('floating_locked', False)
        self.logger.info(f"【窗口管理】任务 {task_data['name']} 浮窗锁定状态已切换为: {'开' if task_data['floating_locked'] else '关'}")
        self.save_data()
        status = "开" if task_data['floating_locked'] else "关"
        self.lock_btn.config(text=f"锁定浮窗: {status}")

    # ========== 浮窗字体颜色设置（保持原有代码）==========
    def change_floating_font_color(self, window):
        """弹出系统颜色选择器，修改当前任务浮窗的字体颜色"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        current_color = task_data.get('duty_label_color', 'blue')
        color_code = colorchooser.askcolor(color=current_color, title="选择浮窗字体颜色")
        if color_code and color_code[1]:
            new_color = color_code[1]
            task_data['duty_label_color'] = new_color
            self.logger.info(f"【颜色设置】任务 {task_data['name']} 的浮窗字体颜色已更改为 {new_color}")
            self.save_data()
            if task_key in self.floating_widgets:
                self.floating_widgets[task_key].duty_label.config(fg=new_color)
            messagebox.showinfo("成功", f"颜色已更改为 {new_color}")

    # ========== 开机自启动（注册表操作）（保持原有代码）==========
    def _get_exe_path(self):
        """获取当前可执行文件路径（打包后返回 exe 路径，否则返回脚本路径）"""
        if getattr(sys, 'frozen', False):
            return sys.executable
        return os.path.abspath(sys.argv[0])

    def check_autostart(self):
        """检查注册表 Run 键中是否已设置本程序开机自启"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows\CurrentVersion\Run",
                                 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, "值日提醒")
            winreg.CloseKey(key)
            return value == self._get_exe_path()
        except FileNotFoundError:
            return False
        except Exception as e:
            self.logger.error(f"【开机自启】检查开机自启状态失败: {e}")
            print(f"检查开机自启状态失败: {e}")
            return False

    def toggle_autostart(self):
        """切换开机自启状态并刷新按钮文字"""
        if self.check_autostart():
            self._remove_autostart()
        else:
            self._add_autostart()
        status = "开" if self.check_autostart() else "关"
        self.autostart_btn.config(text=f"开机自启: {status}")

    def _add_autostart(self):
        """向注册表添加开机自启项"""
        try:
            exe_path = self._get_exe_path()
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows\CurrentVersion\Run",
                                 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, "值日提醒", 0, winreg.REG_SZ, exe_path)
            winreg.CloseKey(key)
            self.logger.info(f"【开机自启】已设置开机自启: {exe_path}")
            messagebox.showinfo("成功", "已开启开机自启")
        except Exception as e:
            self.logger.error(f"【开机自启】设置开机自启失败: {e}")
            messagebox.showerror("错误", f"设置开机自启失败: {e}")

    def _remove_autostart(self):
        """从注册表删除开机自启项"""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r"Software\Microsoft\Windows\CurrentVersion\Run",
                                 0, winreg.KEY_SET_VALUE)
            winreg.DeleteValue(key, "值日提醒")
            winreg.CloseKey(key)
            self.logger.info("【开机自启】已关闭开机自启")
            messagebox.showinfo("成功", "已关闭开机自启")
        except FileNotFoundError:
            pass
        except Exception as e:
            self.logger.error(f"【开机自启】取消开机自启失败: {e}")
            messagebox.showerror("错误", f"取消开机自启失败: {e}")

    # ========== 覆盖值日（保持原有代码）==========
    def set_override_duty(self, window):
        """设置临时覆盖值日人员（若干天内显示指定人）"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            messagebox.showwarning("警告", "值日列表为空，无法设置覆盖")
            return
        override_win = tk.Toplevel(window)
        override_win.title("设置覆盖值日")
        override_win.geometry("320x200")
        override_win.transient(window)
        override_win.grab_set()
        win_x = window.winfo_rootx() + (window.winfo_width() // 2) - 160
        win_y = window.winfo_rooty() + (window.winfo_height() // 2) - 100
        override_win.geometry(f"320x200+{win_x}+{win_y}")
        ttk.Label(override_win, text="选择覆盖人员:").pack(pady=(10,0))
        person_var = tk.StringVar()
        person_combo = ttk.Combobox(override_win, textvariable=person_var,
                                    values=task_data['duty_list'], state="readonly")
        person_combo.pack(pady=5)
        ttk.Label(override_win, text="覆盖天数:").pack(pady=(5,0))
        days_var = tk.IntVar(value=1)
        ttk.Spinbox(override_win, from_=1, to=365, textvariable=days_var, width=5).pack(pady=5)

        def confirm_override():
            person = person_var.get()
            if not person:
                messagebox.showwarning("警告", "请选择覆盖人员")
                return
            try:
                days = int(days_var.get())
            except:
                days = 1
            until_date = (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")
            task_data['override_person'] = person
            task_data['override_until'] = until_date
            self.logger.info(f"【覆盖设置】任务 {task_data['name']} 设置覆盖: {person} (持续 {days} 天，至 {until_date})")
            self.save_data()
            self.update_all_floating_displays()
            self.update_main_window_display()
            messagebox.showinfo("成功", f"已设置 {person} 从今天起覆盖值日 {days} 天")
            override_win.destroy()

        btn_frame = ttk.Frame(override_win)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="确定", command=confirm_override).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=override_win.destroy).pack(side=tk.RIGHT, padx=10)

    def cancel_override(self):
        """取消当前任务的覆盖值日状态"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data.get('override_person'):
            messagebox.showinfo("提示", "当前没有覆盖设置")
            return
        if messagebox.askyesno("确认", f"确定要取消 {task_data['override_person']} 的覆盖吗？"):
            self.logger.info(f"【覆盖设置】任务 {task_data['name']} 取消覆盖: {task_data['override_person']}")
            task_data['override_person'] = None
            task_data['override_until'] = None
            self.save_data()
            self.update_all_floating_displays()
            self.update_main_window_display()
            messagebox.showinfo("成功", "覆盖已取消")

    # ========== 关于页面（保持原有代码）==========
    def show_about(self):
        """显示关于对话框"""
        parent = self.main_window if self.main_window and self.main_window.winfo_exists() else self.root
        about_win = tk.Toplevel(parent)
        about_win.title("关于")
        about_win.geometry("400x250")
        about_win.resizable(False, False)
        about_win.transient(parent)
        about_win.attributes('-topmost', True)
        about_win.grab_set()
        frame = ttk.Frame(about_win, padding="20")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="值日提醒", font=("微软雅黑", 14, "bold")).pack(pady=(0, 10))
        ttk.Label(frame, text="版本: 1.5", font=("微软雅黑", 10)).pack(anchor='w', pady=2)
        ttk.Label(frame, text="作者: teming(骗你的其实就是ai写的())", font=("微软雅黑", 10)).pack(anchor='w', pady=2)
        ttk.Label(frame, text="联系邮箱: temingmail@163.com", font=("微软雅黑", 10)).pack(anchor='w', pady=2)
        ttk.Label(frame, text="其实这个软件很烂，还有一堆bug没修", font=("微软雅黑", 10)).pack(anchor='w', pady=2)
        ttk.Label(frame, text="本人不怎么会编程，大家轻点喷qwq", font=("微软雅黑", 10)).pack(anchor='w', pady=2)
        ttk.Label(frame, text="感谢使用！", font=("微软雅黑", 10)).pack(pady=(10, 0))
        ttk.Button(frame, text="关闭", command=about_win.destroy).pack(pady=20)

    # ========== 窗口管理（保持原有代码）==========
    def hide_main_window(self, window):
        """隐藏设置窗口（不销毁）"""
        window.withdraw()
        self.logger.debug("【窗口管理】主设置窗口已隐藏")

    def on_task_selection_changed(self, event):
        """下拉框切换任务时更新界面"""
        self.update_main_window_display()

    def get_selected_task_key(self):
        """根据下拉框选中的任务名返回对应的键（如 'task1'）"""
        selected_task_name = self.selected_task_var.get()
        for key, data in self.task_groups.items():
            if data['name'] == selected_task_name:
                return key
        return next(iter(self.task_groups.keys()))

    def update_main_window_display(self):
        """根据当前选中的任务刷新主窗口所有显示信息"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if hasattr(self, 'current_duty_var'):
            self.current_duty_var.set(self.get_current_duty_person(task_key))
        if hasattr(self, 'tomorrow_duty_var'):
            self.tomorrow_duty_var.set(self.get_tomorrow_duty_person(task_key))
        if hasattr(self, 'hour_var'):
            self.hour_var.set(str(task_data['reminder_hour']))
        if hasattr(self, 'minute_var'):
            self.minute_var.set(str(task_data['reminder_minute']).zfill(2))
        if hasattr(self, 'listbox'):
            self.listbox.delete(0, tk.END)
            if not task_data['duty_list']:
                self.listbox.insert(tk.END, "无值日人员")
                return
            now = datetime.now()
            today = now.date()
            tomorrow = now.date() + timedelta(days=1)
            base_date = datetime(2024, 1, 1).date()
            days_today = (today - base_date).days
            days_tomorrow = (tomorrow - base_date).days
            idx_today = (task_data['starting_duty_index'] + days_today) % len(task_data['duty_list'])
            idx_tomorrow = (task_data['starting_duty_index'] + days_tomorrow) % len(task_data['duty_list'])
            for i, name in enumerate(task_data['duty_list']):
                if i == idx_today:
                    self.listbox.insert(tk.END, f"{i+1}. {name} ← 当前值日")
                elif i == idx_tomorrow:
                    self.listbox.insert(tk.END, f"{i+1}. {name} ← 明天值日")
                else:
                    self.listbox.insert(tk.END, f"{i+1}. {name}")
        if hasattr(self, 'info_label'):
            self.info_label.config(text=f"每天 {task_data['reminder_hour']:02d}:{task_data['reminder_minute']:02d} 提醒")
        if hasattr(self, 'top_btn'):
            status = "开" if task_data['always_on_top'] else "关"
            self.top_btn.config(text=f"置顶: {status}")
        if hasattr(self, 'floating_btn'):
            status = "开" if task_data['floating_enabled'] else "关"
            self.floating_btn.config(text=f"浮窗: {status}")
        # 更新锁定按钮状态
        if hasattr(self, 'lock_btn'):
            status = "开" if task_data.get('floating_locked', False) else "关"
            self.lock_btn.config(text=f"锁定浮窗: {status}")
        if hasattr(self, 'autostart_btn'):
            status = "开" if self.check_autostart() else "关"
            self.autostart_btn.config(text=f"开机自启: {status}")

    # ========== 任务操作（保持原有代码）==========
    def rename_task(self, window):
        """重命名当前选中的任务"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        rename_window = tk.Toplevel(window)
        rename_window.title("更改任务名称")
        rename_window.geometry("300x150")
        rename_window.transient(window)
        rename_window.grab_set()
        x = window.winfo_rootx() + (window.winfo_width() // 2) - 150
        y = window.winfo_rooty() + (window.winfo_height() // 2) - 75
        rename_window.geometry(f"300x150+{x}+{y}")
        ttk.Label(rename_window, text="请输入新任务名称:", font=("微软雅黑", 10)).pack(pady=10)
        name_entry = ttk.Entry(rename_window, width=25)
        name_entry.pack(pady=5)
        name_entry.insert(0, task_data['name'])
        name_entry.select_range(0, tk.END)
        name_entry.focus()

        def rename():
            new_name = name_entry.get().strip()
            if not new_name:
                messagebox.showwarning("警告", "请输入任务名称")
                return
            if new_name == task_data['name']:
                messagebox.showinfo("提示", "任务名称未更改")
                rename_window.destroy()
                return
            for key, data in self.task_groups.items():
                if key != task_key and data['name'] == new_name:
                    messagebox.showwarning("警告", f"任务名称 '{new_name}' 已存在")
                    return
            old_name = task_data['name']
            task_data['name'] = new_name
            self.logger.info(f"【任务管理】任务 '{old_name}' 已重命名为 '{new_name}'")
            if task_key in self.floating_widgets:
                self.floating_widgets[task_key].title(f"当前{new_name}")
            task_names = [data['name'] for data in self.task_groups.values()]
            self.task_combo['values'] = task_names
            self.task_combo.set(new_name)
            self.update_main_window_display()
            self.save_data()
            messagebox.showinfo("成功", f"任务名称已从 '{old_name}' 更改为 '{new_name}'")
            rename_window.destroy()

        button_frame = ttk.Frame(rename_window)
        button_frame.pack(pady=20)
        ttk.Button(button_frame, text="重命名", command=rename).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=rename_window.destroy).pack(side=tk.RIGHT, padx=10)
        name_entry.bind('<Return>', lambda e: rename())

    def add_new_task(self, window):
        """打开窗口创建新任务（包含姓名列表、提醒时间、语音模板）"""
        add_window = tk.Toplevel(window)
        add_window.title("添加新任务")
        add_window.geometry("400x350")
        add_window.transient(window)
        add_window.grab_set()
        x = window.winfo_rootx() + (window.winfo_width() // 2) - 200
        y = window.winfo_rooty() + (window.winfo_height() // 2) - 175
        add_window.geometry(f"400x350+{x}+{y}")
        frame = ttk.Frame(add_window, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="任务名称:", font=("微软雅黑", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        name_entry = ttk.Entry(frame, width=25)
        name_entry.grid(row=0, column=1, pady=5, padx=(10, 0))
        ttk.Label(frame, text="提醒时间:", font=("微软雅黑", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        time_frame = ttk.Frame(frame)
        time_frame.grid(row=1, column=1, pady=5, padx=(10, 0))
        hour_var = tk.StringVar(value="8")
        minute_var = tk.StringVar(value="00")
        ttk.Spinbox(time_frame, from_=0, to=23, width=5, textvariable=hour_var).pack(side=tk.LEFT)
        ttk.Label(time_frame, text=":").pack(side=tk.LEFT)
        ttk.Spinbox(time_frame, from_=0, to=59, width=5, textvariable=minute_var).pack(side=tk.LEFT)

        ttk.Label(frame, text="初始值日人员 (每行一人):", font=("微软雅黑", 10)).grid(row=2, column=0, sticky=(tk.W, tk.N), pady=5)
        duty_text = tk.Text(frame, width=25, height=6, font=("微软雅黑", 10))
        duty_text.grid(row=2, column=1, pady=5, padx=(10, 0))
        duty_scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=duty_text.yview)
        duty_scrollbar.grid(row=2, column=2, sticky=(tk.N, tk.S), pady=5)
        duty_text.configure(yscrollcommand=duty_scrollbar.set)

        ttk.Label(frame, text="语音模板:", font=("微软雅黑", 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
        voice_template = ttk.Entry(frame, width=25)
        voice_template.grid(row=3, column=1, pady=5, padx=(10, 0))
        voice_template.insert(0, '现在是%H:%M，明天是%Y年%m月%d日，请%DUTY%同学记得完成明天的%TASK%任务！')

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)

        def add_task():
            name = name_entry.get().strip()
            if not name:
                messagebox.showwarning("警告", "请输入任务名称")
                return
            for task_data in self.task_groups.values():
                if task_data['name'] == name:
                    messagebox.showwarning("警告", f"任务名称 '{name}' 已存在")
                    return
            duty_text_content = duty_text.get("1.0", tk.END).strip()
            duty_list = [line.strip() for line in duty_text_content.split('\n') if line.strip()]
            custom_template = voice_template.get().strip()
            if not custom_template:
                custom_template = '现在是%H:%M，明天是%Y年%m月%d日，请%DUTY%同学记得完成明天的%TASK%任务！'
            try:
                hour = int(hour_var.get())
                minute = int(minute_var.get())
                if not (0 <= hour <= 23 and 0 <= minute <= 59):
                    raise ValueError
            except ValueError:
                messagebox.showerror("错误", "请输入有效的提醒时间 (小时:0-23, 分钟:0-59)")
                return
            new_task_key = f"task{len(self.task_groups) + 1}"
            while new_task_key in self.task_groups:
                new_task_key = f"task{int(new_task_key[4:]) + 1}"
            new_task_data = {
                'name': name,
                'duty_list': duty_list,
                'starting_duty_index': 0,
                'reminder_hour': hour,
                'reminder_minute': minute,
                'always_on_top': True,
                'floating_x': None,
                'floating_y': None,
                'voice_enabled': True,
                'window_scale_factor': 1.0,
                'font_size_factor': 1.0,
                'floating_enabled': True,
                'custom_voice_template': custom_template,
                'override_person': None,
                'override_until': None,
                'floating_locked': False,        # 新任务默认不锁定
                'duty_label_color': 'blue'
            }
            self.task_groups[new_task_key] = new_task_data
            self.logger.info(f"【任务管理】已添加新任务: {name} (人员数: {len(duty_list)})")
            if new_task_data['floating_enabled']:
                self.create_single_floating_widget(new_task_key, new_task_data)
            task_names = [data['name'] for data in self.task_groups.values()]
            self.task_combo['values'] = task_names
            self.task_combo.set(name)
            self.save_data()
            messagebox.showinfo("成功", f"已添加任务: {name}")
            add_window.destroy()
            self.update_main_window_display()

        ttk.Button(button_frame, text="添加任务", command=add_task).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=add_window.destroy).pack(side=tk.RIGHT, padx=10)

    def delete_current_task(self, window):
        """删除当前选中的任务（至少保留一个）"""
        if len(self.task_groups) <= 1:
            messagebox.showwarning("警告", "至少需要保留一个任务")
            return
        task_key = self.get_selected_task_key()
        task_name = self.task_groups[task_key]['name']
        if not messagebox.askyesno("确认删除", f"确定要删除任务 '{task_name}' 吗？"):
            return
        self.logger.info(f"【任务管理】已删除任务: {task_name}")
        if task_key in self.floating_widgets:
            self.floating_widgets[task_key].destroy()
            del self.floating_widgets[task_key]
        del self.task_groups[task_key]
        task_names = [data['name'] for data in self.task_groups.values()]
        self.task_combo['values'] = task_names
        self.task_combo.set(task_names[0])
        self.save_data()
        messagebox.showinfo("成功", f"已删除任务: {task_name}")
        self.update_main_window_display()

    def import_from_excel(self, window):
        """从 Excel 文件导入值日人员列表（第一列姓名）"""
        try:
            file_path = filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            if not file_path:
                return
            self.logger.info(f"【Excel导入】开始导入文件: {file_path}")
            df = pd.read_excel(file_path)
            if df.empty:
                messagebox.showerror("错误", "Excel文件为空")
                return
            names_column = df.iloc[:, 0].tolist()
            filtered_names = []
            for name in names_column:
                if pd.notna(name):
                    name_str = str(name).strip()
                    if name_str:
                        try:
                            float(name_str)          # 跳过纯数字
                            continue
                        except ValueError:
                            filtered_names.append(name_str)
            if not filtered_names:
                messagebox.showerror("错误", "Excel文件中没有找到有效的姓名数据")
                return
            task_key = self.get_selected_task_key()
            task_data = self.task_groups[task_key]
            task_data['duty_list'] = filtered_names
            task_data['starting_duty_index'] = 0
            self.logger.info(f"【Excel导入】成功导入 {len(filtered_names)} 名值日人员")
            self.update_all_floating_displays()
            self.update_main_window_display()
            self.save_data()
            messagebox.showinfo("成功", f"成功导入 {len(filtered_names)} 名值日人员")
        except ImportError:
            messagebox.showerror("错误", "缺少必要的库，请安装pandas和openpyxl:\n\npip install pandas openpyxl")
        except Exception as e:
            self.logger.error(f"【Excel导入】导入失败: {e}")
            messagebox.showerror("错误", f"导入Excel文件失败: {str(e)}")

    def change_current_duty(self, window):
        """手动更改当前显示的值日人员（调整起始索引）"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            messagebox.showwarning("警告", "值日列表为空，请先添加值日人员")
            return
        change_window = tk.Toplevel(window)
        change_window.title("更改当前值日人员")
        change_window.geometry("300x200")
        change_window.transient(window)
        change_window.grab_set()
        x = window.winfo_rootx() + (window.winfo_width() // 2) - 150
        y = window.winfo_rooty() + (window.winfo_height() // 2) - 100
        change_window.geometry(f"300x200+{x}+{y}")
        ttk.Label(change_window, text="请选择新的值日人员:", font=("微软雅黑", 10)).pack(pady=10)
        duty_var = tk.StringVar()
        duty_combo = ttk.Combobox(change_window, textvariable=duty_var, state="readonly", width=25)
        duty_combo['values'] = task_data['duty_list']
        duty_combo.pack(pady=10)
        duty_combo.set(self.get_current_duty_person(task_key))

        def confirm_change():
            selected_name = duty_var.get()
            if not selected_name:
                messagebox.showwarning("警告", "请选择值日人员")
                return
            try:
                new_index = task_data['duty_list'].index(selected_name)
                current_day_index = self.get_current_day_index()
                task_data['starting_duty_index'] = (new_index - current_day_index) % len(task_data['duty_list'])
                self.logger.info(f"【值日管理】任务 {task_data['name']} 的当前值日人员已更改为: {selected_name}")
                self.update_all_floating_displays()
                self.update_main_window_display()
                self.save_data()
                messagebox.showinfo("成功", f"值日人员已更改为: {selected_name}")
                change_window.destroy()
            except ValueError:
                messagebox.showerror("错误", "选择的值日人员不在列表中")

        btn_frame = ttk.Frame(change_window)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确认更改", command=confirm_change).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=change_window.destroy).pack(side=tk.RIGHT, padx=10)

    def add_duty_person(self, window):
        """添加一个新值日人员到当前任务"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        add_window = tk.Toplevel(window)
        add_window.title("添加值日人员")
        add_window.geometry("300x150")
        add_window.transient(window)
        add_window.grab_set()
        x = window.winfo_rootx() + (window.winfo_width() // 2) - 150
        y = window.winfo_rooty() + (window.winfo_height() // 2) - 75
        add_window.geometry(f"300x150+{x}+{y}")
        ttk.Label(add_window, text="请输入值日人员姓名:", font=("微软雅黑", 10)).pack(pady=10)
        name_entry = ttk.Entry(add_window, width=25)
        name_entry.pack(pady=5)

        def add_person():
            name = name_entry.get().strip()
            if not name:
                messagebox.showwarning("警告", "请输入值日人员姓名")
                return
            if name in task_data['duty_list']:
                messagebox.showwarning("警告", "该人员已在值日列表中")
                return
            task_data['duty_list'].append(name)
            self.logger.info(f"【值日管理】任务 {task_data['name']} 已添加值日人员: {name}")
            self.update_all_floating_displays()
            self.update_main_window_display()
            self.save_data()
            messagebox.showinfo("成功", f"已添加值日人员: {name}")
            name_entry.delete(0, tk.END)

        btn_frame = ttk.Frame(add_window)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="添加", command=add_person).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=add_window.destroy).pack(side=tk.RIGHT, padx=10)

    def remove_duty_person(self, window):
        """从当前任务中移除一个值日人员"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            messagebox.showwarning("警告", "值日列表为空")
            return
        remove_window = tk.Toplevel(window)
        remove_window.title("移除值日人员")
        remove_window.geometry("300x200")
        remove_window.transient(window)
        remove_window.grab_set()
        x = window.winfo_rootx() + (window.winfo_width() // 2) - 150
        y = window.winfo_rooty() + (window.winfo_height() // 2) - 100
        remove_window.geometry(f"300x200+{x}+{y}")
        ttk.Label(remove_window, text="请选择要移除的值日人员:", font=("微软雅黑", 10)).pack(pady=10)
        duty_var = tk.StringVar()
        duty_combo = ttk.Combobox(remove_window, textvariable=duty_var, state="readonly", width=25)
        duty_combo['values'] = task_data['duty_list']
        duty_combo.pack(pady=10)
        duty_combo.set(self.get_current_duty_person(task_key))

        def remove_person():
            selected_name = duty_var.get()
            if not selected_name:
                messagebox.showwarning("警告", "请选择要移除的值日人员")
                return
            current_duty = self.get_current_duty_person(task_key)
            if selected_name == current_duty:
                if not messagebox.askyesno("确认", f"确定要移除当前值日人员 '{selected_name}' 吗？"):
                    return
            task_data['duty_list'].remove(selected_name)
            self.logger.info(f"【值日管理】任务 {task_data['name']} 已移除值日人员: {selected_name}")
            if not task_data['duty_list']:
                task_data['starting_duty_index'] = 0
            else:
                current_day_index = self.get_current_day_index()
                current_duty_after = self.get_current_duty_person(task_key)
                if current_duty_after != current_duty:
                    if current_duty in task_data['duty_list']:
                        idx = task_data['duty_list'].index(current_duty)
                        task_data['starting_duty_index'] = (idx - current_day_index) % len(task_data['duty_list'])
            self.update_all_floating_displays()
            self.update_main_window_display()
            self.save_data()
            messagebox.showinfo("成功", f"已移除值日人员: {selected_name}")
            if not task_data['duty_list']:
                remove_window.destroy()
            else:
                duty_combo['values'] = task_data['duty_list']
                duty_combo.set(self.get_current_duty_person(task_key))

        btn_frame = ttk.Frame(remove_window)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="移除", command=remove_person).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=remove_window.destroy).pack(side=tk.RIGHT, padx=10)

    def toggle_always_on_top(self, window):
        """切换当前任务浮窗和主窗口的置顶状态"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        task_data['always_on_top'] = not task_data['always_on_top']
        for w in self.floating_widgets.values():
            w.attributes('-topmost', task_data['always_on_top'])
        for w in self.root.winfo_children():
            if isinstance(w, tk.Toplevel) and w.title() == "值日提醒":
                w.attributes('-topmost', task_data['always_on_top'])
        self.logger.info(f"【窗口管理】任务 {task_data['name']} 的置顶状态已切换为: {'开' if task_data['always_on_top'] else '关'}")
        self.save_data()
        status = "开" if task_data['always_on_top'] else "关"
        self.top_btn.config(text=f"置顶: {status}")

    def toggle_floating(self, window):
        """切换当前任务的浮窗显示/隐藏"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        task_data['floating_enabled'] = not task_data['floating_enabled']
        if task_data['floating_enabled']:
            if task_key not in self.floating_widgets:
                self.create_single_floating_widget(task_key, task_data)
                self.logger.info(f"【窗口管理】任务 {task_data['name']} 的浮窗已显示")
        else:
            if task_key in self.floating_widgets:
                self.floating_widgets[task_key].destroy()
                del self.floating_widgets[task_key]
                self.logger.info(f"【窗口管理】任务 {task_data['name']} 的浮窗已隐藏")
        self.save_data()
        status = "开" if task_data['floating_enabled'] else "关"
        self.floating_btn.config(text=f"浮窗: {status}")

    def open_resize_window(self, window):
        """打开浮窗大小调整窗口（滑块）"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        original = task_data['window_scale_factor']
        resize_win = tk.Toplevel(window)
        resize_win.title("调整浮窗大小")
        resize_win.geometry("400x200")
        resize_win.transient(window)
        resize_win.grab_set()
        x = window.winfo_rootx() + (window.winfo_width() // 2) - 200
        y = window.winfo_rooty() + (window.winfo_height() // 2) - 100
        resize_win.geometry(f"400x200+{x}+{y}")
        ttk.Label(resize_win, text="拖动滑块调整浮窗大小:", font=("微软雅黑", 12)).pack(pady=10)
        current_scale_var = tk.StringVar()
        current_scale_var.set(f"当前缩放: {task_data['window_scale_factor']:.1f}x")
        ttk.Label(resize_win, textvariable=current_scale_var, font=("微软雅黑", 10)).pack(pady=5)
        scale_var = tk.DoubleVar(value=task_data['window_scale_factor'])
        scale_slider = ttk.Scale(resize_win, from_=0.5, to=2.0, orient='horizontal', variable=scale_var, length=300)
        scale_slider.pack(pady=10)

        def update_scale(*args):
            v = scale_var.get()
            current_scale_var.set(f"当前缩放: {v:.1f}x")
            task_data['window_scale_factor'] = v
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
        scale_var.trace_add('write', update_scale)

        btn_frame = ttk.Frame(resize_win)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确认", command=lambda: [self.save_data() and messagebox.showinfo("成功", "浮窗大小调整已保存"), resize_win.destroy()] if self.save_data() else messagebox.showerror("错误", "保存失败")).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=lambda: (
            setattr(task_data, 'window_scale_factor', original),
            self.update_floating_size_and_font(task_key) if task_data['floating_enabled'] and task_key in self.floating_widgets else None,
            resize_win.destroy()
        )).pack(side=tk.RIGHT, padx=10)
        ttk.Button(btn_frame, text="重置", command=lambda: (scale_var.set(1.0), setattr(task_data, 'window_scale_factor', 1.0), resize_win.destroy())).pack(side=tk.LEFT, padx=10)

    def open_font_resize_window(self, window):
        """打开浮窗字体大小调整窗口（滑块）"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        original = task_data['font_size_factor']
        font_win = tk.Toplevel(window)
        font_win.title("调整字体大小")
        font_win.geometry("400x200")
        font_win.transient(window)
        font_win.grab_set()
        x = window.winfo_rootx() + (window.winfo_width() // 2) - 200
        y = window.winfo_rooty() + (window.winfo_height() // 2) - 100
        font_win.geometry(f"400x200+{x}+{y}")
        ttk.Label(font_win, text="拖动滑块调整字体大小:", font=("微软雅黑", 12)).pack(pady=10)
        current_var = tk.StringVar()
        current_var.set(f"当前字体缩放: {task_data['font_size_factor']:.1f}x")
        ttk.Label(font_win, textvariable=current_var, font=("微软雅黑", 10)).pack(pady=5)
        scale_var = tk.DoubleVar(value=task_data['font_size_factor'])
        scale_slider = ttk.Scale(font_win, from_=0.5, to=2.0, orient='horizontal', variable=scale_var, length=300)
        scale_slider.pack(pady=10)

        

        def update_scale(*args):
            v = scale_var.get()
            current_var.set(f"当前字体缩放: {v:.1f}x")
            task_data['font_size_factor'] = v
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
        scale_var.trace_add('write', update_scale)

        btn_frame = ttk.Frame(font_win)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="确认", command=lambda: [self.save_data() and messagebox.showinfo("成功", "字体大小调整已保存"), font_win.destroy()] if self.save_data() else messagebox.showerror("错误", "保存失败")).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="取消", command=lambda: (
            setattr(task_data, 'font_size_factor', original),
            self.update_floating_size_and_font(task_key) if task_data['floating_enabled'] and task_key in self.floating_widgets else None,
            font_win.destroy()
        )).pack(side=tk.RIGHT, padx=10)
        ttk.Button(btn_frame, text="重置", command=lambda: (scale_var.set(1.0), setattr(task_data, 'font_size_factor', 1.0), font_win.destroy())).pack(side=tk.LEFT, padx=10)

    def update_all_floating_displays(self):
        """更新所有浮窗的显示内容"""
        for tk_key in self.task_groups:
            if self.task_groups[tk_key]['floating_enabled'] and tk_key in self.floating_widgets:
                self.update_floating_display(tk_key)

    def shuffle_order(self, window):
        """随机打乱当前任务的顺序并保持当前值日者不变"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data['duty_list'] or len(task_data['duty_list']) <= 1:
            messagebox.showinfo("提示", "值日人员不足，无法打乱顺序")
            return
        current_duty = self.get_current_duty_person(task_key)
        shuffled = task_data['duty_list'][:]
        random.shuffle(shuffled)
        idx = next((i for i, n in enumerate(shuffled) if n == current_duty), 0)
        task_data['duty_list'] = shuffled
        task_data['starting_duty_index'] = idx
        self.logger.info(f"【值日管理】任务 {task_data['name']} 的值日顺序已随机打乱")
        self.update_all_floating_displays()
        self.update_main_window_display()
        self.save_data()

    def save_current_order(self, window):
        """保存当前配置（并显示临时提示）"""
        self.save_data()
        for w in self.root.winfo_children():
            if isinstance(w, tk.Toplevel) and w.title() == "值日提醒":
                temp = tk.Label(w, text="已保存", fg="green")
                temp.place(relx=0.5, rely=0.1, anchor="center")
                w.after(1500, temp.destroy)

    def save_reminder_time(self, window):
        """保存当前选中任务的提醒时间"""
        try:
            hour = int(self.hour_var.get())
            minute = int(self.minute_var.get())
            if 0 <= hour <= 23 and 0 <= minute <= 59:
                task_key = self.get_selected_task_key()
                task_data = self.task_groups[task_key]
                old_time = f"{task_data['reminder_hour']:02d}:{task_data['reminder_minute']:02d}"
                task_data['reminder_hour'] = hour
                task_data['reminder_minute'] = minute
                self.logger.info(f"【提醒设置】任务 {task_data['name']} 的提醒时间已从 {old_time} 改为 {hour:02d}:{minute:02d}")
                self.reschedule_daily_reminder()
                self.save_data()
                self.update_main_window_display()
                # 临时提示
                for w in self.root.winfo_children():
                    if isinstance(w, tk.Toplevel) and w.title() == "值日提醒":
                        temp = tk.Label(w, text="提醒时间已保存", fg="green")
                        temp.place(relx=0.5, rely=0.15, anchor="center")
                        w.after(1500, temp.destroy)
            else:
                messagebox.showerror("错误", "请输入有效的小时(0-23)和分钟(0-59)")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")

    def reschedule_daily_reminder(self):
        """根据第一个任务的提醒时间重新设定每日定时器（所有任务共用同一时间）"""
        try:
            if hasattr(self, 'scheduler') and self.scheduler.running:
                try:
                    self.scheduler.remove_job('daily_duty_reminder')
                except:
                    pass
                first_task = next(iter(self.task_groups.values()))
                self.scheduler.add_job(
                    self.daily_reminder,
                    'cron',
                    hour=first_task['reminder_hour'],
                    minute=first_task['reminder_minute'],
                    second=0,
                    id='daily_duty_reminder'
                )
        except Exception as e:
            self.logger.error(f"【定时任务】重新安排定时任务失败: {e}")

    def test_speech(self, icon=None, item=None):
        """测试语音合成（手动触发）"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            messagebox.showwarning("警告", "请先添加值日人员")
            return
        template = task_data['custom_voice_template']
        tomorrow_duty = self.get_tomorrow_duty_person(task_key)
        tomorrow_date = (datetime.now() + timedelta(days=1)).strftime('%Y年%m月%d日')
        current_time = datetime.now().strftime('%H:%M')
        message = template.replace('%DUTY%', tomorrow_duty)\
                         .replace('%TASK%', task_data['name'])\
                         .replace('%TIME%', current_time)\
                         .replace('%DATE%', tomorrow_date)
        threading.Thread(target=self.speak_message, args=(message,), daemon=True).start()

    def speak_message(self, message):
        """播放语音消息（需 TTS 已初始化）"""
        if not self.tts_initialized:
            print("TTS未初始化，无法播放语音")
            return
        with self.tts_lock:
            try:
                self.tts_engine.say(message)
                self.tts_engine.runAndWait()
            except Exception as e:
                print(f"语音播放出错: {e}")

    def daily_reminder(self):
        """每日定时触发：为所有启用了语音的任务播放提醒"""
        self.logger.info("=" * 80)
        self.logger.info("【每日提醒】开始执行每日定时提醒任务...")
        reminder_count = 0
        for task_key, task_data in self.task_groups.items():
            if not task_data['voice_enabled']:
                continue
            tomorrow_duty = self.get_tomorrow_duty_person(task_key)
            template = task_data['custom_voice_template']
            tomorrow_date = (datetime.now() + timedelta(days=1)).strftime('%Y年%m月%d日')
            current_time = datetime.now().strftime('%H:%M')
            message = template.replace('%DUTY%', tomorrow_duty)\
                             .replace('%TASK%', task_data['name'])\
                             .replace('%TIME%', current_time)\
                             .replace('%DATE%', tomorrow_date)
            self.logger.info(f"【每日提醒】任务 {task_data['name']}: {message}")
            threading.Thread(target=self.speak_message, args=(message,), daemon=True).start()
            reminder_count += 1
        self.logger.info(f"【每日提醒】共触发 {reminder_count} 个任务的语音提醒")
        self.logger.info("=" * 80)

    def start_scheduler(self):
        """启动后台定时调度器"""
        self.scheduler = BackgroundScheduler()
        first_task = next(iter(self.task_groups.values()))
        self.scheduler.add_job(
            self.daily_reminder,
            'cron',
            hour=first_task['reminder_hour'],
            minute=first_task['reminder_minute'],
            second=0,
            id='daily_duty_reminder'
        )
        self.scheduler.start()

    def quit_app(self, icon=None, item=None):
        """退出整个应用"""
        self.self.logger.info("=" * 80)logger.info("=" * 80)
        self.logger.info("【程序退出】用户请求退出程序...")
        self.cleanup()
        if hasattr(self, 'icon'):
            self.icon.stop()
        self.self.logger.info("【程序退出】程序已完全退出")logger.info("【程序退出】程序已完全退出")
        self.self.logger.info("=" * 80)logger.info("=" * 80)
        self.root.quit()

       def清理(自我):def cleanup(self):
        """清理资源：关闭定时器、保存浮窗位置等"""
        try:
            if hasattr(self, 'scheduler') and self.scheduler.running:
                self.scheduler.shutdown(wait=True)
                self.logger.info("【资源清理】定时调度器已关闭")
        except Exception as   作为 e:
            self.logger.error(f"【资源清理】调度器关闭出错: {e}")
        for task_key, floating_widget in   在 self.floating_widgets.items():
            task_data = self.Task_data = self.task_groups[task_key]task_groups[task_key]
            task_dataTask_data ['floating_x'] = floating_widget.winfo_x（）['floating_x'] = floating_widget.winfo_x()
            task_dataTask_data ['floating_y'] = floating_widget.winfo_y（）['floating_y'] = floating_widget.winfo_y()
            如果‘font_size_factor’不在task_data中：if 'font_size_factor' not in   在 task_data:
                task_dataTask_data ['font_size_factor'] = 1.0['font_size_factor'] = 1.0
        self.save_data()
        self.logger.info("【资源清理】配置已保存，浮窗位置已记录")

def main   主要():
       试一试:try:
        app =    app = DutyReminderApp（）DutyReminderApp()
        app.root.mainloop()
    except KeyboardInterrupt:
        global_logger.info("【程序中断】检测到键盘中断 (Ctrl+C)")
        if 'app' in   在 locals():
            app.cleanup()
    except Exception as   作为 e:
        global_logger.error(f"【程序异常】程序运行出错: {e}", exc_info=True)
        raise

if __name__ == "__main__":
    main   主要()
