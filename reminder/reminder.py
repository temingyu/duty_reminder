# 此软件使用AI编写
# ai写的可能会有亿点点问题（ 
# 作者本人并不会写代码，全由AI生成，当然，也欢迎大家一起参与到此软件的维护当中，此软件将会在GitHub上共享源码
# 如果你发现了问题，可以发邮件或者提交issue给我，虽然不一定会修 嘻嘻（学业繁忙）
# 先这样吧，也没啥好说上的了（
# 哦对，我的邮箱:temingmail@163.com
# 就这样        
import ctypes
import tkinter as tk    
import pyttsx3.drivers
import pyttsx3.drivers.sapi5
import pystray._win32
from tkinter import ttk, messagebox, filedialog
import random
import pyttsx3
import json
import os
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler
import threading
import sys
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item
import atexit
import pandas as pd  # 需要pip install pandas openpyxl

class DutyReminderApp:
    
    def __init__(self):

        app_id = 'mycompany.myapp.subproduct.version'  # 可自定义唯一标识
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        
        # 初始化语音引擎标志
        self.tts_initialized = False
        self.tts_lock = threading.Lock()  # 语音播放锁
        
        # 加载或初始化数据
        self.load_data()
        
        # 初始化TTS引擎
        self.tts_engine = None
        #self.setup_tts()
        
        # 创建主窗口（隐藏）
        self.root = tk.Tk()
        self.root.withdraw()  # 初始隐藏主窗口
        ico = self.get_icon_path()
        try:
            self.root.iconbitmap(ico)
        except Exception as e:
            print(f"警告：无法设置窗口图标，原因: {e}")

        # 主窗口引用（修复关于页面的关键）
        self.main_window = None
        
        # 创建系统托盘图标
        self.create_system_tray()
        
        # 创建浮动小部件列表
        self.floating_widgets = {}
        self.create_floating_widgets()
        
        # 启动定时任务
        self.start_scheduler()
        
        # 注册退出处理
        atexit.register(self.cleanup)
    
    def load_data(self):
        """加载配置数据"""
        # 获取可执行文件所在目录
        if getattr(sys, 'frozen', False):
            # 如果是打包后的可执行文件
            app_dir = os.path.dirname(sys.executable)
        else:
            # 如果是开发环境
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.config_file = os.path.join(app_dir, "duty_config.json")
        
        # 默认模板
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
            'override_person': None,
            'override_until': None
        }

        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.task_groups = data.get('task_groups', {})
            except (json.JSONDecodeError, KeyError):
                self.task_groups = {}
        else:
            self.task_groups = {}

        # 确保至少有三个默认任务
        if not self.task_groups:
            self.task_groups = {
                'task1': {
                    **default_task,
                    'name': '刷勺',
                    'duty_list': ['张三', '李四', '王五', '赵六'],
                },
                'task2': {
                    **default_task,
                    'name': '打扫',
                    'duty_list': ['小明', '小红', '小刚', '小美'],
                },
                'task3': {
                    **default_task,
                    'name': '黑板',
                    'duty_list': ['阿强', '阿华', '阿丽', '阿军'],
                }
            }
            self.save_data()
        else:
            # 为每个任务组补充缺失字段
            for task_key, task_data in self.task_groups.items():
                # 确保名字存在
                if 'name' not in task_data:
                    task_data['name'] = f'任务{task_key}'
                # 确保列表存在
                if 'duty_list' not in task_data:
                    task_data['duty_list'] = []
                # 补充默认值
                for key, val in default_task.items():
                    if key not in task_data:
                        task_data[key] = val
                # 额外处理索引合法性
                if task_data['duty_list'] and (
                    task_data['starting_duty_index'] >= len(task_data['duty_list']) or 
                    task_data['starting_duty_index'] < 0
                ):
                    task_data['starting_duty_index'] = 0
            self.save_data()
    
    def save_data(self):
        """保存配置数据"""
        import copy
        data = {
            'task_groups': copy.deepcopy(self.task_groups)
        }
        try:
            import tempfile
            temp_fd, temp_path = tempfile.mkstemp(suffix='.json', dir=os.path.dirname(self.config_file))
            try:
                with os.fdopen(temp_fd, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                import shutil
                shutil.move(temp_path, self.config_file)
            except Exception as temp_error:
                try:
                    os.unlink(temp_path)
                except:
                    pass
                raise temp_error
        except Exception as e:
            print(f"保存配置失败: {e}")
            return False
        return True
    
    def get_icon_path(self, filename='reminder_icon.ico'):
        if getattr(sys, 'frozen', False):
            # 单文件 exe 的资源临时目录
            base_dir = sys._MEIPASS
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_dir, filename)
    
    #def setup_tts(self):
        """配置TTS引擎"""
        try:
            self.tts_engine = pyttsx3.init()
            voices = self.tts_engine.getProperty('voices')
            if voices:
                for voice in voices:
                    if 'Chinese' in voice.name or 'zh-CN' in voice.id or 'Microsoft' in voice.name:
                        self.tts_engine.setProperty('voice', voice.id)
                        break
            self.tts_engine.setProperty('rate', 180)
            self.tts_engine.setProperty('volume', 0.9)
            self.tts_initialized = True
        except Exception as e:
            print(f"TTS初始化失败: {e}")
            self.tts_initialized = False
    
    def get_current_day_index(self):
        """计算从基准日期开始的天数差"""
        base_date = datetime(2024, 1, 1).date()
        current_date = datetime.now().date()
        day_difference = (current_date - base_date).days
        return day_difference
    
    def is_override_active(self, task_key):
        """检查当前任务是否处于覆盖期"""
        task_data = self.task_groups[task_key]
        if task_data.get('override_person') and task_data.get('override_until'):
            try:
                until = datetime.strptime(task_data['override_until'], "%Y-%m-%d").date()
                return datetime.now().date() <= until
            except:
                pass
        return False

    def get_current_duty_person(self, task_key):
        """获取当前应该显示的值日人员"""
        if self.is_override_active(task_key):
            return self.task_groups[task_key]['override_person']
        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            return "无值日人员"
        now = datetime.now()
        reminder_time = now.replace(hour=task_data['reminder_hour'], minute=task_data['reminder_minute'], second=0, microsecond=0)
        current_day_index = self.get_current_day_index()
        actual_index = (task_data['starting_duty_index'] + current_day_index) % len(task_data['duty_list'])
        if now >= reminder_time:
            next_index = (actual_index + 1) % len(task_data['duty_list'])
            return task_data['duty_list'][next_index] if task_data['duty_list'] and next_index < len(task_data['duty_list']) else "无值日人员"
        else:
            return task_data['duty_list'][actual_index] if task_data['duty_list'] and actual_index < len(task_data['duty_list']) else "无值日人员"
    
    def get_tomorrow_duty_person(self, task_key):
        """获取明天的值日人员"""
        if self.is_override_active(task_key):
            return self.task_groups[task_key]['override_person']
        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            return "无值日人员"
        now = datetime.now()
        tomorrow = now + timedelta(days=1)
        tomorrow_date = tomorrow.date()
        base_date = datetime(2024, 1, 1).date()
        tomorrow_day_index = (tomorrow_date - base_date).days
        actual_index = (task_data['starting_duty_index'] + tomorrow_day_index) % len(task_data['duty_list'])
        return task_data['duty_list'][actual_index] if actual_index < len(task_data['duty_list']) else "无值日人员"
    
    def get_current_or_tomorrow_label(self, task_key):
        """根据时间返回合适的标签"""
        task_data = self.task_groups[task_key]
        now = datetime.now()
        reminder_time = now.replace(hour=task_data['reminder_hour'], minute=task_data['reminder_minute'], second=0, microsecond=0)
        if now >= reminder_time:
            return f"明日{task_data['name']}"
        else:
            return f"今日{task_data['name']}"
    
    def create_floating_widgets(self):
        """创建浮动小部件"""
        for task_key, task_data in self.task_groups.items():
            if task_data['floating_enabled']:
                self.create_single_floating_widget(task_key, task_data)
    
    def create_single_floating_widget(self, task_key, task_data):
        """创建单个浮动小部件"""
        floating_widget = tk.Toplevel(self.root)
        floating_widget.title(f"当前{task_data['name']}")
        floating_widget.geometry("220x100")
        floating_widget.overrideredirect(True)
        floating_widget.attributes('-topmost', task_data['always_on_top'])
        try:
            floating_widget.attributes('-alpha', 0.9)
        except tk.TclError:
            pass
        
        # 窗口拖拽功能
        floating_widget.bind('<Button-1>', lambda e, tk=task_key: self.start_drag(e, tk))
        floating_widget.bind('<B1-Motion>', lambda e, tk=task_key: self.drag_window(e, tk))
        # 新增：鼠标释放时自动保存位置
        floating_widget.bind('<ButtonRelease-1>', lambda e, tk=task_key: self.save_position_on_release(e, tk))
        
        floating_widget.task_key = task_key
        self.create_widget_ui(floating_widget, task_key)
        self.floating_widgets[task_key] = floating_widget
        self.update_floating_size_and_font(task_key)
        self.update_floating_display(task_key)
        floating_widget.after(60000, lambda tk=task_key: self.periodic_update(tk))
        self.set_initial_position(task_key)
    
    def set_initial_position(self, task_key):
        """设置初始位置（屏幕右上角）"""
        floating_widget = self.floating_widgets[task_key]
        task_data = self.task_groups[task_key]
        floating_widget.update_idletasks()
        screen_width = floating_widget.winfo_screenwidth()
        screen_height = floating_widget.winfo_screenheight()
        base_width = 220
        base_height = 100
        window_width = int(base_width * task_data['window_scale_factor'])
        window_height = int(base_height * task_data['window_scale_factor'])
        if task_data['floating_x'] is not None and task_data['floating_y'] is not None:
            if 0 <= task_data['floating_x'] <= screen_width - window_width and 0 <= task_data['floating_y'] <= screen_height - window_height:
                floating_widget.geometry(f"{window_width}x{window_height}+{task_data['floating_x']}+{task_data['floating_y']}")
                return
        offset_x = (ord(task_key[-1]) - ord('1')) * 230
        x = screen_width - window_width - 20 - offset_x
        y = 20
        floating_widget.geometry(f"{window_width}x{window_height}+{x}+{y}")

    def create_widget_ui(self, floating_widget, task_key):
        """创建浮动小部件界面"""
        main_frame = tk.Frame(floating_widget, bg='#f0f0f0', bd=2, relief='solid')
        main_frame.pack(fill='both', expand=True, padx=1, pady=1)
        label_text = tk.StringVar()
        label_text.set(self.get_current_or_tomorrow_label(task_key))
        title_label = tk.Label(main_frame, textvariable=label_text, bg='#f0f0f0', 
                              font=("微软雅黑", 10, "bold"), fg='gray')
        title_label.pack(pady=(5, 0))
        duty_label = tk.Label(main_frame, text="", bg='#f0f0f0', 
                             font=("微软雅黑", 12, "bold"), fg='red')
        duty_label.pack(pady=(0, 5))
        main_frame.bind('<Double-Button-1>', lambda e: self.show_main_window())
        floating_widget.label_text = label_text
        floating_widget.duty_label = duty_label
        floating_widget.title_label = title_label

    def start_drag(self, event, task_key):
        """开始拖拽"""
        floating_widget = self.floating_widgets[task_key]
        floating_widget.x = event.x
        floating_widget.y = event.y
    
    def drag_window(self, event, task_key):
        """拖拽窗口"""
        floating_widget = self.floating_widgets[task_key]
        task_data = self.task_groups[task_key]
        x = floating_widget.winfo_x() + event.x - floating_widget.x
        y = floating_widget.winfo_y() + event.y - floating_widget.y
        floating_widget.geometry(f"+{x}+{y}")
        task_data['floating_x'] = x
        task_data['floating_y'] = y

    def save_position_on_release(self, event, task_key):
        """鼠标释放时保存浮窗位置"""
        self.save_data()

    def update_floating_display(self, task_key):
        """更新浮动小部件显示"""
        if task_key not in self.floating_widgets:
            print(f"Floating widget for task {task_key} does not exist yet.")
            return
        floating_widget = self.floating_widgets[task_key]
        current_duty = self.get_current_duty_person(task_key)
        current_label = self.get_current_or_tomorrow_label(task_key)
        floating_widget.duty_label.config(text=current_duty)
        floating_widget.label_text.set(current_label)

    def periodic_update(self, task_key):
        """定期更新"""
        self.update_floating_display(task_key)
        self.floating_widgets[task_key].after(60000, lambda tk=task_key: self.periodic_update(tk))
    
    def update_floating_size_and_font(self, task_key):
        """根据缩放因子更新浮窗大小和字体"""
        if task_key not in self.floating_widgets:
            return
        floating_widget = self.floating_widgets[task_key]
        task_data = self.task_groups[task_key]
        base_width = 220
        base_height = 100
        base_title_font_size = 10
        base_duty_font_size = 12
        new_width = int(base_width * task_data['window_scale_factor'])
        new_height = int(base_height * task_data['window_scale_factor'])
        new_title_font_size = max(8, int(base_title_font_size * task_data['window_scale_factor'] * task_data['font_size_factor']))
        new_duty_font_size = max(10, int(base_duty_font_size * task_data['window_scale_factor'] * task_data['font_size_factor']))
        floating_widget.geometry(f"{new_width}x{new_height}")
        if hasattr(floating_widget, 'title_label'):
            floating_widget.title_label.config(font=("微软雅黑", new_title_font_size, "bold"))
        if hasattr(floating_widget, 'duty_label'):
            floating_widget.duty_label.config(font=("微软雅黑", new_duty_font_size, "bold"))

    def create_system_tray(self):
        """创建系统托盘图标（使用自定义图标）"""
        ico_path = self.get_icon_path('reminder_icon.ico')
        try:
            image = Image.open(ico_path)
        except:
            image = Image.new('RGB', (64, 64), color='lightblue')
            draw = ImageDraw.Draw(image)
            draw.rectangle([10, 10, 54, 54], outline='black', width=2)
            draw.text((20, 20), '值', fill='black', font_size=30)

        menu = (item('主页面', self.show_main_window),
                item('退出', self.quit_app))
        self.icon = pystray.Icon("值日提醒", image, "值日提醒", menu=menu)
        self.icon.run_detached()
    
    def show_main_window(self, icon=None, item=None):
        """显示主窗口"""
        if self.main_window is not None and self.main_window.winfo_exists():
            self.main_window.deiconify()
            self.main_window.lift()
            self.main_window.focus_force()
            self.update_main_window_display()
        else:
            self.main_window = self.create_main_window()
            self.main_window.deiconify()
            self.main_window.lift()
            self.main_window.focus_force()
    
    def create_main_window(self):
        """创建主窗口"""
        window = tk.Toplevel(self.root)
        window.title("值日提醒")
        window.geometry("700x780")
        window.resizable(True, True)
        ico = self.get_icon_path()
        try:
            window.iconbitmap(ico)
        except Exception as e:
            print(f"主窗口图标失败: {e}")

        first_task = next(iter(self.task_groups.values()))
        window.attributes('-topmost', first_task['always_on_top'])
        main_frame = ttk.Frame(window, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        title_label = ttk.Label(main_frame, text="值日提醒", font=("微软雅黑", 18, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 15))
        
        task_select_frame = ttk.LabelFrame(main_frame, text="任务管理", padding="10")
        task_select_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.selected_task_var = tk.StringVar()
        task_names = [data['name'] for data in self.task_groups.values()]
        self.task_combo = ttk.Combobox(task_select_frame, textvariable=self.selected_task_var, values=task_names, state="readonly", width=20)
        self.task_combo.grid(row=0, column=0, padx=(0, 10))
        self.task_combo.set(task_names[0])
        self.task_combo.bind('<<ComboboxSelected>>', self.on_task_selection_changed)
        
        refresh_btn = ttk.Button(task_select_frame, text="刷新显示", command=self.update_main_window_display)
        refresh_btn.grid(row=0, column=1)
        rename_btn = ttk.Button(task_select_frame, text="更改任务名称", command=lambda w=window: self.rename_task(w))
        rename_btn.grid(row=0, column=2, padx=(10, 0))
        add_task_btn = ttk.Button(task_select_frame, text="添加任务", command=lambda w=window: self.add_new_task(w))
        add_task_btn.grid(row=0, column=3, padx=(10, 0))
        delete_task_btn = ttk.Button(task_select_frame, text="删除当前任务", command=lambda w=window: self.delete_current_task(w))
        delete_task_btn.grid(row=0, column=4, padx=(10, 0))
        
        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        current_frame = ttk.LabelFrame(left_frame, text="值日信息", padding="10")
        current_frame.pack(fill='both', expand=True, pady=(0, 10))
        self.current_duty_var = tk.StringVar()
        ttk.Label(current_frame, text="当前值日:", font=("微软雅黑", 12, "bold")).pack(anchor='w', pady=(0, 5))
        ttk.Label(current_frame, textvariable=self.current_duty_var, font=("微软雅黑", 16, "bold"), foreground="red").pack(pady=5)
        self.tomorrow_duty_var = tk.StringVar()
        ttk.Label(current_frame, text="明天值日:", font=("微软雅黑", 12, "bold")).pack(anchor='w', pady=(10, 5))
        ttk.Label(current_frame, textvariable=self.tomorrow_duty_var, font=("微软雅黑", 16, "bold"), foreground="blue").pack(pady=5)
        
        time_frame = ttk.LabelFrame(right_frame, text="提醒时间设置", padding="10")
        time_frame.pack(fill='x', pady=(0, 10))
        time_row = ttk.Frame(time_frame)
        time_row.pack(fill='x', pady=5)
        ttk.Label(time_row, text="提醒时间:").pack(side=tk.LEFT, padx=(0, 5))
        self.hour_var = tk.StringVar(value="8")
        ttk.Spinbox(time_row, from_=0, to=23, width=5, textvariable=self.hour_var).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(time_row, text=":").pack(side=tk.LEFT)
        self.minute_var = tk.StringVar(value="00")
        ttk.Spinbox(time_row, from_=0, to=59, width=5, textvariable=self.minute_var).pack(side=tk.LEFT, padx=(5, 10))
        ttk.Button(time_row, text="保存提醒时间", command=lambda w=window: self.save_reminder_time(w)).pack(side=tk.LEFT)
        self.info_label = ttk.Label(time_frame, text="每天 08:00 自动提醒值日", font=("微软雅黑", 9))
        self.info_label.pack(pady=(10, 0))
        
        list_frame = ttk.LabelFrame(left_frame, text="值日顺序", padding="10")
        list_frame.pack(fill='both', expand=True)
        self.listbox = tk.Listbox(list_frame, height=10, font=("微软雅黑", 10))
        self.listbox.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill='y')
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        control_frame = ttk.LabelFrame(right_frame, text="功能控制", padding="10")
        control_frame.pack(fill='both', expand=True)
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill='x', pady=5)
        ttk.Button(button_frame, text="随机打乱顺序", command=lambda w=window: self.shuffle_order(w)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="保存当前顺序", command=lambda w=window: self.save_current_order(w)).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(control_frame, text="导入Excel", command=lambda w=window: self.import_from_excel(w)).pack(fill='x', pady=5)
        ttk.Button(control_frame, text="更改当前值日人员", command=lambda w=window: self.change_current_duty(w)).pack(fill='x', pady=5)
        ttk.Button(control_frame, text="添加值日人员", command=lambda w=window: self.add_duty_person(w)).pack(fill='x', pady=5)
        ttk.Button(control_frame, text="移除值日人员", command=lambda w=window: self.remove_duty_person(w)).pack(fill='x', pady=5)
        ttk.Button(control_frame, text="设置覆盖值日", command=lambda w=window: self.set_override_duty(w)).pack(fill='x', pady=5)
        ttk.Button(control_frame, text="取消覆盖", command=self.cancel_override).pack(fill='x', pady=5)
        
        switch_frame = ttk.LabelFrame(control_frame, text="显示控制", padding="10")
        switch_frame.pack(fill='x', pady=5)
        top_status = "开" if first_task['always_on_top'] else "关"
        self.top_btn = ttk.Button(switch_frame, text=f"置顶显示: {top_status}", command=lambda w=window: self.toggle_always_on_top(w))
        self.top_btn.pack(fill='x', pady=2)
        floating_status = "开" if first_task['floating_enabled'] else "关"
        self.floating_btn = ttk.Button(switch_frame, text=f"浮窗显示: {floating_status}", command=lambda w=window: self.toggle_floating(w))
        self.floating_btn.pack(fill='x', pady=2)
        self.resize_btn = ttk.Button(switch_frame, text="调整浮窗大小", command=lambda w=window: self.open_resize_window(w))
        self.resize_btn.pack(fill='x', pady=2)
        self.font_resize_btn = ttk.Button(switch_frame, text="调整字体大小", command=lambda w=window: self.open_font_resize_window(w))
        self.font_resize_btn.pack(fill='x', pady=2)
        ttk.Label(control_frame, text="Excel导入说明：文件中第一列应包含姓名，每行一个姓名", font=("微软雅黑", 8), foreground="gray").pack(pady=(10, 0))
        
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        left_frame.rowconfigure(1, weight=1)
        right_frame.rowconfigure(1, weight=1)
        
        # ---------- 右上角“关于”按钮 ----------
        about_btn = ttk.Button(window, text="关于", command=self.show_about)
        about_btn.place(relx=1.0, rely=0.0, anchor='ne', x=-10, y=10)
        
        self.update_main_window_display()
        window.protocol("WM_DELETE_WINDOW", lambda: self.hide_main_window(window))
        return window
    
    def set_override_duty(self, window):
        """设置覆盖值日人员"""
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
        days_spin = ttk.Spinbox(override_win, from_=1, to=365, textvariable=days_var, width=5)
        days_spin.pack(pady=5)

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
        """取消覆盖"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data.get('override_person'):
            messagebox.showinfo("提示", "当前没有覆盖设置")
            return
        if messagebox.askyesno("确认", f"确定要取消 {task_data['override_person']} 的覆盖吗？"):
            task_data['override_person'] = None
            task_data['override_until'] = None
            self.save_data()
            self.update_all_floating_displays()
            self.update_main_window_display()
            messagebox.showinfo("成功", "覆盖已取消")

    def show_about(self):
        """显示关于页面（修复版本）"""
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
    
    def hide_main_window(self, window):
        """隐藏主窗口而不是关闭"""
        window.withdraw()
    
    def on_task_selection_changed(self, event):
        """任务选择改变时更新显示"""
        self.update_main_window_display()
    
    def get_selected_task_key(self):
        """根据当前选择的任务名称获取任务键"""
        selected_task_name = self.selected_task_var.get()
        for key, data in self.task_groups.items():
            if data['name'] == selected_task_name:
                return key
        return next(iter(self.task_groups.keys()))
    
    def update_main_window_display(self):
        """更新主窗口显示"""
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
            days_since_start_today = (today - base_date).days
            actual_current_index = (task_data['starting_duty_index'] + days_since_start_today) % len(task_data['duty_list'])
            days_since_start_tomorrow = (tomorrow - base_date).days
            actual_next_index = (task_data['starting_duty_index'] + days_since_start_tomorrow) % len(task_data['duty_list'])
            for i, name in enumerate(task_data['duty_list']):
                if i == actual_current_index:
                    self.listbox.insert(tk.END, f"{i+1}. {name} ← 当前值日")
                elif i == actual_next_index:
                    self.listbox.insert(tk.END, f"{i+1}. {name} ← 明天值日")
                else:
                    self.listbox.insert(tk.END, f"{i+1}. {name}")
        
        if hasattr(self, 'info_label'):
            self.info_label.config(text=f"每天 {task_data['reminder_hour']:02d}:{task_data['reminder_minute']:02d} 自动提醒值日")
        
        if hasattr(self, 'top_btn'):
            status = "开" if task_data['always_on_top'] else "关"
            self.top_btn.config(text=f"置顶显示: {status}")
        if hasattr(self, 'floating_btn'):
            status = "开" if task_data['floating_enabled'] else "关"
            self.floating_btn.config(text=f"浮窗显示: {status}")
    
    def open_custom_voice_window(self, window):
        """打开自定义语音提醒内容窗口"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        custom_window = tk.Toplevel(window)
        custom_window.title("自定义语音提醒内容")
        custom_window.geometry("600x400")
        custom_window.transient(window)
        custom_window.grab_set()
        window_x = window.winfo_rootx()
        window_y = window.winfo_rooty()
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        x = window_x + (window_width // 2) - 300
        y = window_y + (window_height // 2) - 200
        custom_window.geometry(f"600x400+{x}+{y}")
        ttk.Label(custom_window, text="自定义语音提醒内容:", font=("微软雅黑", 12)).pack(pady=10)
        info_text = tk.Text(custom_window, height=4, wrap=tk.WORD, font=("微软雅黑", 9))
        info_text.pack(pady=5, padx=20, fill=tk.X)
        info_text.insert(tk.END, "变量说明：\n%DUTY% - 值日人员姓名\n%TASK% - 任务名称\n%TIME% - 当前时间\n%DATE% - 明天日期")
        info_text.config(state=tk.DISABLED)
        text_frame = ttk.Frame(custom_window)
        text_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        voice_text = tk.Text(text_frame, height=8, font=("微软雅黑", 10))
        voice_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        voice_scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=voice_text.yview)
        voice_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        voice_text.configure(yscrollcommand=voice_scrollbar.set)
        voice_text.insert(tk.END, task_data['custom_voice_template'])
        button_frame = ttk.Frame(custom_window)
        button_frame.pack(pady=20)
        def save_custom_voice():
            template = voice_text.get("1.0", tk.END).strip()
            if not template:
                messagebox.showwarning("警告", "请输入自定义语音内容")
                return
            task_data['custom_voice_template'] = template
            self.save_data()
            messagebox.showinfo("成功", "自定义语音内容已保存")
            custom_window.destroy()
        ttk.Button(button_frame, text="保存", command=save_custom_voice).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=custom_window.destroy).pack(side=tk.RIGHT, padx=10)
        def preview_voice():
            template = voice_text.get("1.0", tk.END).strip()
            if not template:
                messagebox.showwarning("警告", "请输入自定义语音内容")
                return
            tomorrow_duty = self.get_tomorrow_duty_person(task_key)
            tomorrow_date = (datetime.now() + timedelta(days=1)).strftime('%Y年%m月%d日')
            current_time = datetime.now().strftime('%H:%M')
            preview_text = template.replace('%DUTY%', tomorrow_duty)
            preview_text = preview_text.replace('%TASK%', task_data['name'])
            preview_text = preview_text.replace('%TIME%', current_time)
            preview_text = preview_text.replace('%DATE%', tomorrow_date)
            messagebox.showinfo("预览", f"语音内容预览：\n{preview_text}")
        ttk.Button(button_frame, text="预览", command=preview_voice).pack(side=tk.LEFT, padx=10)

    def rename_task(self, window):
        """重命名任务"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        rename_window = tk.Toplevel(window)
        rename_window.title("更改任务名称")
        rename_window.geometry("300x150")
        rename_window.transient(window)
        rename_window.grab_set()
        window_x = window.winfo_rootx()
        window_y = window.winfo_rooty()
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        x = window_x + (window_width // 2) - 150
        y = window_y + (window_height // 2) - 75
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
        """添加新任务"""
        add_window = tk.Toplevel(window)
        add_window.title("添加新任务")
        add_window.geometry("400x350")
        add_window.transient(window)
        add_window.grab_set()
        window_x = window.winfo_rootx()
        window_y = window.winfo_rooty()
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        x = window_x + (window_width // 2) - 200
        y = window_y + (window_height // 2) - 175
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
                'override_until': None
            }
            self.task_groups[new_task_key] = new_task_data
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
        """删除当前任务"""
        if len(self.task_groups) <= 1:
            messagebox.showwarning("警告", "至少需要保留一个任务")
            return
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        task_name = task_data['name']
        result = messagebox.askyesno("确认删除", f"确定要删除任务 '{task_name}' 吗？\n此操作不可撤销！")
        if not result:
            return
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
        """从Excel导入值日人员列表"""
        try:
            file_path = filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            if not file_path:
                return
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
                            float(name_str)
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
            self.update_all_floating_displays()
            self.update_main_window_display()
            self.save_data()
            messagebox.showinfo("成功", f"成功导入 {len(filtered_names)} 名值日人员")
        except ImportError:
            messagebox.showerror("错误", "缺少必要的库，请安装pandas和openpyxl:\n\npip install pandas openpyxl")
        except Exception as e:
            messagebox.showerror("错误", f"导入Excel文件失败: {str(e)}")
    
    def change_current_duty(self, window):
        """手动更改当前值日人员"""
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
        window_x = window.winfo_rootx()
        window_y = window.winfo_rooty()
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        x = window_x + (window_width // 2) - 150
        y = window_y + (window_height // 2) - 100
        change_window.geometry(f"300x200+{x}+{y}")
        ttk.Label(change_window, text="请选择新的值日人员:", font=("微软雅黑", 10)).pack(pady=10)
        duty_var = tk.StringVar()
        duty_combo = ttk.Combobox(change_window, textvariable=duty_var, state="readonly", width=25)
        duty_combo['values'] = task_data['duty_list']
        duty_combo.pack(pady=10)
        current_duty = self.get_current_duty_person(task_key)
        duty_combo.set(current_duty)
        button_frame = ttk.Frame(change_window)
        button_frame.pack(pady=20)
        def confirm_change():
            selected_name = duty_var.get()
            if not selected_name:
                messagebox.showwarning("警告", "请选择值日人员")
                return
            try:
                new_index = task_data['duty_list'].index(selected_name)
                current_day_index = self.get_current_day_index()
                calculated_starting_index = (new_index - current_day_index) % len(task_data['duty_list'])
                task_data['starting_duty_index'] = calculated_starting_index
                self.update_all_floating_displays()
                self.update_main_window_display()
                self.save_data()
                messagebox.showinfo("成功", f"值日人员已更改为: {selected_name}")
                change_window.destroy()
            except ValueError:
                messagebox.showerror("错误", "选择的值日人员不在列表中")
        ttk.Button(button_frame, text="确认更改", command=confirm_change).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=change_window.destroy).pack(side=tk.RIGHT, padx=10)
    
    def add_duty_person(self, window):
        """添加值日人员"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        add_window = tk.Toplevel(window)
        add_window.title("添加值日人员")
        add_window.geometry("300x150")
        add_window.transient(window)
        add_window.grab_set()
        window_x = window.winfo_rootx()
        window_y = window.winfo_rooty()
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        x = window_x + (window_width // 2) - 150
        y = window_y + (window_height // 2) - 75
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
            self.update_all_floating_displays()
            self.update_main_window_display()
            self.save_data()
            messagebox.showinfo("成功", f"已添加值日人员: {name}")
            name_entry.delete(0, tk.END)
        button_frame = ttk.Frame(add_window)
        button_frame.pack(pady=20)
        ttk.Button(button_frame, text="添加", command=add_person).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=add_window.destroy).pack(side=tk.RIGHT, padx=10)
    
    def remove_duty_person(self, window):
        """移除值日人员"""
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
        window_x = window.winfo_rootx()
        window_y = window.winfo_rooty()
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        x = window_x + (window_width // 2) - 150
        y = window_y + (window_height // 2) - 100
        remove_window.geometry(f"300x200+{x}+{y}")
        ttk.Label(remove_window, text="请选择要移除的值日人员:", font=("微软雅黑", 10)).pack(pady=10)
        duty_var = tk.StringVar()
        duty_combo = ttk.Combobox(remove_window, textvariable=duty_var, state="readonly", width=25)
        duty_combo['values'] = task_data['duty_list']
        duty_combo.pack(pady=10)
        current_duty = self.get_current_duty_person(task_key)
        duty_combo.set(current_duty)
        def remove_person():
            selected_name = duty_var.get()
            if not selected_name:
                messagebox.showwarning("警告", "请选择要移除的值日人员")
                return
            current_duty = self.get_current_duty_person(task_key)
            if selected_name == current_duty:
                result = messagebox.askyesno("确认", f"确定要移除当前值日人员 '{selected_name}' 吗？\n这将影响当前的值日安排。")
                if not result:
                    return
            task_data['duty_list'].remove(selected_name)
            if not task_data['duty_list']:
                task_data['starting_duty_index'] = 0
            else:
                current_day_index = self.get_current_day_index()
                current_duty_after_remove = self.get_current_duty_person(task_key)
                if current_duty_after_remove != current_duty:
                    if selected_name in task_data['duty_list']:
                        new_index = task_data['duty_list'].index(selected_name)
                    else:
                        new_index = 0
                    if current_duty in task_data['duty_list']:
                        target_index = task_data['duty_list'].index(current_duty)
                        task_data['starting_duty_index'] = (target_index - current_day_index) % len(task_data['duty_list'])
            self.update_all_floating_displays()
            self.update_main_window_display()
            self.save_data()
            messagebox.showinfo("成功", f"已移除值日人员: {selected_name}")
            if not task_data['duty_list']:
                remove_window.destroy()
            else:
                duty_combo['values'] = task_data['duty_list']
                if task_data['duty_list']:
                    current_duty_after = self.get_current_duty_person(task_key)
                    duty_combo.set(current_duty_after)
        button_frame = ttk.Frame(remove_window)
        button_frame.pack(pady=20)
        ttk.Button(button_frame, text="移除", command=remove_person).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="取消", command=remove_window.destroy).pack(side=tk.RIGHT, padx=10)
    
    def toggle_always_on_top(self, window):
        """切换置顶状态"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        task_data['always_on_top'] = not task_data['always_on_top']
        for widget in self.floating_widgets.values():
            widget.attributes('-topmost', task_data['always_on_top'])
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Toplevel) and widget.title() == "值日提醒":
                widget.attributes('-topmost', task_data['always_on_top'])
        self.save_data()
        status = "开" if task_data['always_on_top'] else "关"
        self.top_btn.config(text=f"置顶显示: {status}")

    def toggle_floating(self, window):
        """切换浮窗显示开关状态"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        task_data['floating_enabled'] = not task_data['floating_enabled']
        if task_data['floating_enabled']:
            if task_key not in self.floating_widgets:
                self.create_single_floating_widget(task_key, task_data)
        else:
            if task_key in self.floating_widgets:
                self.floating_widgets[task_key].destroy()
                del self.floating_widgets[task_key]
        self.save_data()
        status_text = "开" if task_data['floating_enabled'] else "关"
        self.floating_btn.config(text=f"浮窗显示: {status_text}")

    def open_resize_window(self, window):
        """打开调整浮窗大小的窗口"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        original_window_scale_factor = task_data['window_scale_factor']
        resize_window = tk.Toplevel(window)
        resize_window.title("调整浮窗大小")
        resize_window.geometry("400x200")
        resize_window.transient(window)
        resize_window.grab_set()
        window_x = window.winfo_rootx()
        window_y = window.winfo_rooty()
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        x = window_x + (window_width // 2) - 200
        y = window_y + (window_height // 2) - 100
        resize_window.geometry(f"400x200+{x}+{y}")
        ttk.Label(resize_window, text="拖动滑块调整浮窗大小:", font=("微软雅黑", 12)).pack(pady=10)
        current_scale_var = tk.StringVar()
        current_scale_var.set(f"当前缩放: {task_data['window_scale_factor']:.1f}x")
        ttk.Label(resize_window, textvariable=current_scale_var, font=("微软雅黑", 10)).pack(pady=5)
        scale_var = tk.DoubleVar(value=task_data['window_scale_factor'])
        scale_slider = ttk.Scale(resize_window, from_=0.5, to=2.0, orient='horizontal', variable=scale_var, length=300)
        scale_slider.pack(pady=10)
        def update_scale(*args):
            value = scale_var.get()
            current_scale_var.set(f"当前缩放: {value:.1f}x")
            task_data['window_scale_factor'] = value
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
        scale_var.trace_add('write', update_scale)
        button_frame = ttk.Frame(resize_window)
        button_frame.pack(pady=20)
        def confirm_resize():
            task_data['window_scale_factor'] = scale_var.get()
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
            if self.save_data():
                messagebox.showinfo("成功", "浮窗大小调整已保存")
                resize_window.destroy()
            else:
                messagebox.showerror("错误", "浮窗大小保存失败，请检查文件权限")
        ttk.Button(button_frame, text="确认", command=confirm_resize).pack(side=tk.LEFT, padx=10)
        def cancel_resize():
            task_data['window_scale_factor'] = original_window_scale_factor
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
            resize_window.destroy()
        ttk.Button(button_frame, text="取消", command=cancel_resize).pack(side=tk.RIGHT, padx=10)
        def reset_size():
            scale_var.set(1.0)
            task_data['window_scale_factor'] = 1.0
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
            if self.save_data():
                messagebox.showinfo("成功", "已重置为默认浮窗大小")
            else:
                messagebox.showerror("错误", "保存失败，请检查文件权限")
        ttk.Button(button_frame, text="重置", command=reset_size).pack(side=tk.LEFT, padx=10)
    
    def open_font_resize_window(self, window):
        """打开调整字体大小的窗口"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        original_font_size_factor = task_data['font_size_factor']
        font_resize_window = tk.Toplevel(window)
        font_resize_window.title("调整字体大小")
        font_resize_window.geometry("400x200")
        font_resize_window.transient(window)
        font_resize_window.grab_set()
        window_x = window.winfo_rootx()
        window_y = window.winfo_rooty()
        window_width = window.winfo_width()
        window_height = window.winfo_height()
        x = window_x + (window_width // 2) - 200
        y = window_y + (window_height // 2) - 100
        font_resize_window.geometry(f"400x200+{x}+{y}")
        ttk.Label(font_resize_window, text="拖动滑块调整字体大小:", font=("微软雅黑", 12)).pack(pady=10)
        current_scale_var = tk.StringVar()
        current_scale_var.set(f"当前字体缩放: {task_data['font_size_factor']:.1f}x")
        ttk.Label(font_resize_window, textvariable=current_scale_var, font=("微软雅黑", 10)).pack(pady=5)
        scale_var = tk.DoubleVar(value=task_data['font_size_factor'])
        scale_slider = ttk.Scale(font_resize_window, from_=0.5, to=2.0, orient='horizontal', variable=scale_var, length=300)
        scale_slider.pack(pady=10)
        def update_scale(*args):
            value = scale_var.get()
            current_scale_var.set(f"当前字体缩放: {value:.1f}x")
            task_data['font_size_factor'] = value
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
        scale_var.trace_add('write', update_scale)
        button_frame = ttk.Frame(font_resize_window)
        button_frame.pack(pady=20)
        def confirm_resize():
            task_data['font_size_factor'] = scale_var.get()
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
            if self.save_data():
                messagebox.showinfo("成功", "字体大小调整已保存")
                font_resize_window.destroy()
            else:
                messagebox.showerror("错误", "字体大小保存失败，请检查文件权限")
        ttk.Button(button_frame, text="确认", command=confirm_resize).pack(side=tk.LEFT, padx=10)
        def cancel_resize():
            task_data['font_size_factor'] = original_font_size_factor
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
            font_resize_window.destroy()
        ttk.Button(button_frame, text="取消", command=cancel_resize).pack(side=tk.RIGHT, padx=10)
        def reset_size():
            scale_var.set(1.0)
            task_data['font_size_factor'] = 1.0
            if task_data['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_size_and_font(task_key)
            if self.save_data():
                messagebox.showinfo("成功", "已重置为默认字体大小")
            else:
                messagebox.showerror("错误", "保存失败，请检查文件权限")
        ttk.Button(button_frame, text="重置", command=reset_size).pack(side=tk.LEFT, padx=10)

    def update_all_floating_displays(self):
        """更新所有浮动窗口的显示"""
        for task_key in self.task_groups.keys():
            if self.task_groups[task_key]['floating_enabled'] and task_key in self.floating_widgets:
                self.update_floating_display(task_key)

    def shuffle_order(self, window):
        """随机打乱值日顺序"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data['duty_list'] or len(task_data['duty_list']) <= 1:
            messagebox.showinfo("提示", "值日人员不足，无法打乱顺序")
            return
        current_duty = self.get_current_duty_person(task_key)
        shuffled_list = task_data['duty_list'][:]
        random.shuffle(shuffled_list)
        new_index = 0
        for i, name in enumerate(shuffled_list):
            if name == current_duty:
                new_index = i
                break
        task_data['duty_list'] = shuffled_list
        task_data['starting_duty_index'] = new_index
        self.update_all_floating_displays()
        self.update_main_window_display()
        self.save_data()
    
    def save_current_order(self, window):
        """保存当前顺序"""
        self.save_data()
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Toplevel) and widget.title() == "值日提醒":
                temp_label = tk.Label(widget, text="已保存", fg="green")
                temp_label.place(relx=0.5, rely=0.1, anchor="center")
                widget.after(1500, lambda lbl=temp_label: lbl.destroy() if lbl.winfo_exists() else None)
    
    def save_reminder_time(self, window):
        """保存提醒时间"""
        try:
            hour = int(self.hour_var.get())
            minute = int(self.minute_var.get())
            if 0 <= hour <= 23 and 0 <= minute <= 59:
                task_key = self.get_selected_task_key()
                task_data = self.task_groups[task_key]
                task_data['reminder_hour'] = hour
                task_data['reminder_minute'] = minute
                self.reschedule_daily_reminder()
                self.save_data()
                self.update_main_window_display()
                for widget in self.root.winfo_children():
                    if isinstance(widget, tk.Toplevel) and widget.title() == "值日提醒":
                        temp_label = tk.Label(widget, text="提醒时间已保存", fg="green")
                        temp_label.place(relx=0.5, rely=0.15, anchor="center")
                        widget.after(1500, lambda lbl=temp_label: lbl.destroy() if lbl.winfo_exists() else None)
            else:
                messagebox.showerror("错误", "请输入有效的小时(0-23)和分钟(0-59)")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
    
    def reschedule_daily_reminder(self):
        """重新安排每日提醒任务"""
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
            print(f"重新安排定时任务失败: {e}")
    
    def test_speech(self, icon=None, item=None):
        """测试语音提醒"""
        task_key = self.get_selected_task_key()
        task_data = self.task_groups[task_key]
        if not task_data['duty_list']:
            messagebox.showwarning("警告", "请先添加值日人员")
            return
        template = task_data['custom_voice_template']
        tomorrow_duty = self.get_tomorrow_duty_person(task_key)
        tomorrow_date = (datetime.now() + timedelta(days=1)).strftime('%Y年%m月%d日')
        current_time = datetime.now().strftime('%H:%M')
        message = template.replace('%DUTY%', tomorrow_duty)
        message = message.replace('%TASK%', task_data['name'])
        message = message.replace('%TIME%', current_time)
        message = message.replace('%DATE%', tomorrow_date)
        speech_thread = threading.Thread(target=self.speak_message, args=(message,))
        speech_thread.daemon = True
        speech_thread.start()
    
    def speak_message(self, message):
        """播放语音消息"""
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
        """每日提醒函数"""
        for task_key, task_data in self.task_groups.items():
            if not task_data['voice_enabled']:
                print(f"任务 {task_data['name']} 的语音提醒已关闭，跳过播放")
                continue
            tomorrow_duty = self.get_tomorrow_duty_person(task_key)
            template = task_data['custom_voice_template']
            tomorrow_date = (datetime.now() + timedelta(days=1)).strftime('%Y年%m月%d日')
            current_time = datetime.now().strftime('%H:%M')
            message = template.replace('%DUTY%', tomorrow_duty)
            message = message.replace('%TASK%', task_data['name'])
            message = message.replace('%TIME%', current_time)
            message = message.replace('%DATE%', tomorrow_date)
            speech_thread = threading.Thread(target=self.speak_message, args=(message,))
            speech_thread.daemon = True
            speech_thread.start()
    
    def start_scheduler(self):
        """启动定时任务"""
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
        """退出应用"""
        self.cleanup()
        if hasattr(self, 'icon'):
            self.icon.stop()
        self.root.quit()
    
    def cleanup(self):
        """清理资源"""
        try:
            if hasattr(self, 'scheduler') and self.scheduler.running:
                self.scheduler.shutdown(wait=True)
        except Exception as e:
            print(f"调度器关闭出错: {e}")
        for task_key, floating_widget in self.floating_widgets.items():
            task_data = self.task_groups[task_key]
            task_data['floating_x'] = floating_widget.winfo_x()
            task_data['floating_y'] = floating_widget.winfo_y()
            if 'font_size_factor' not in task_data:
                task_data['font_size_factor'] = 1.0
        self.save_data()

def main():
    app = DutyReminderApp()
    try:
        app.root.mainloop()
    except KeyboardInterrupt:
        app.cleanup()

if __name__ == "__main__":
    main()