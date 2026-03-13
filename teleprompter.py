"""
提词器主程序 - KTV风格
支持自动/手动播放模式，逐字高亮，分段录音，预览和导出
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Canvas
import sounddevice as sd
import numpy as np
import wave
import os
import threading
import time
import json
import shutil
from parser import parse_transcript_file, format_time


class ToolTip:
    """工具提示类 - 鼠标悬停显示说明"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)
    
    def show_tip(self, event=None):
        """显示提示"""
        if self.tip_window or not self.text:
            return
        
        x, y, _, _ = self.widget.bbox("insert") if hasattr(self.widget, 'bbox') else (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background="#ffffe0",
            foreground="#000000",
            relief=tk.SOLID,
            borderwidth=1,
            font=("Microsoft YaHei", 9)
        )
        label.pack(ipadx=8, ipady=4)
    
    def hide_tip(self, event=None):
        """隐藏提示"""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class TeleprompterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("录音提词器")
        self.root.geometry("1200x900")
        self.root.configure(bg='#000000')
        
        # 数据
        self.segments = []
        self.current_index = 0
        self.is_auto_mode = False
        self.is_playing = False
        self.is_recording = False
        self.is_previewing = False
        self.current_transcript_path = None  # 当前字幕文件路径
        self.current_project_path = None  # 当前工程存档目录
        
        # 录音相关
        self.recordings_dir = "recordings"
        self.recording_data = []
        self.recording_states = {}  # {index: True/False} 记录是否已录制
        self.skip_states = {}  # {index: True/False} 语气词跳过，合成时填静音
        self.merge_groups = []  # [[0],[1,2],[3],...] 合并组，同组内连续播放无间隔
        self.sample_rate = 44100
        self.preview_stream = None
        
        # AI演示相关
        self.is_ai_playing = False
        self.ai_demo_dir = "ai_demos"
        self.tts_available = False
        
        # 音量监控
        self.current_volume = 0.0
        self.volume_update_id = None
        
        # 播放速度控制
        self.playback_speed = 1.0
        self.speed_options = [0.5, 0.75, 1.0, 1.25, 1.5]
        self.current_speed_index = 2  # 默认1.0x
        
        # 音频缓存（性能优化）
        self.audio_cache = {}
        self.cache_size_limit = 10
        
        # 评分缓存（避免每次刷新都重新计算）
        self._score_cache = {}
        
        # 列表项控件引用（用于增量高亮更新，避免全量重建）
        self._list_item_widgets = {}  # index -> {"frame": frame, "widgets": [...]}
        
        # 线程安全
        self.ai_lock = threading.Lock()
        
        # 创建目录
        if not os.path.exists(self.recordings_dir):
            os.makedirs(self.recordings_dir)
        if not os.path.exists(self.ai_demo_dir):
            os.makedirs(self.ai_demo_dir)
        
        # KTV动画相关
        self.animation_id = None
        self.play_timer_id = None
        self.word_progress = 0.0  # 当前句子内的进度 0.0-1.0
        self.segment_start_time = 0  # 当前句子开始播放的时间戳
        
        # 录制防抖（避免双击/快捷键+点击导致快速切换）
        self._last_toggle_recording_time = 0.0
        
        self.setup_ui()
        self.bind_shortcuts()
        self.setup_drag_drop()
        self.init_tts_engine()
        
        # 绑定退出保护
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 初始化显示
        self.root.after(100, self.refresh_initial_display)
        
    def setup_drag_drop(self):
        """设置拖拽功能"""
        # 绑定拖拽事件到主窗口和主要组件
        def make_droppable(widget):
            widget.drop_target_register('DND_Files')
            widget.dnd_bind('<<Drop>>', self.handle_drop)
        
        # 尝试使用tkinterdnd2（如果可用）
        try:
            from tkinterdnd2 import DND_FILES
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.handle_drop)
            self.lyrics_canvas.drop_target_register(DND_FILES)
            self.lyrics_canvas.dnd_bind('<<Drop>>', self.handle_drop)
        except (ImportError, AttributeError):
            # 回退到简单的文本绑定（有限支持）
            pass
    
    def handle_drop(self, event):
        """处理拖拽文件"""
        try:
            # 处理拖拽的文件路径
            if hasattr(event, 'data'):
                file_path = event.data
            elif hasattr(event, 'widget'):
                return
            else:
                return
            
            # 清理路径（移除花括号和多余空格）
            file_path = file_path.strip('{}').strip()
            
            # 如果是多个文件，只取第一个
            if '\n' in file_path:
                file_path = file_path.split('\n')[0].strip()
            
            if not file_path:
                return
            
            # 检查文件是否存在
            if not os.path.exists(file_path):
                messagebox.showwarning("提示", "文件不存在")
                return
            
            # 检查文件扩展名
            if not file_path.lower().endswith('.txt'):
                messagebox.showwarning("提示", "请拖拽 .txt 格式的转录文件")
                return
            
            # 加载文件
            self.load_transcript_file(file_path)
            
        except Exception as e:
            print(f"拖拽处理错误: {e}")
            messagebox.showerror("错误", f"拖拽加载失败: {str(e)}")
        
    def setup_ui(self):
        """构建KTV风格GUI界面"""
        # 顶部控制栏
        top_frame = tk.Frame(self.root, bg='#1a1a1a', height=70)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        # 模式切换按钮
        self.mode_btn = tk.Button(
            top_frame, 
            text="📖 手动模式", 
            font=('Microsoft YaHei', 11, 'bold'),
            bg='#333333',
            fg='white',
            command=self.toggle_mode,
            width=12,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.mode_btn.pack(side=tk.LEFT, padx=15, pady=15)
        ToolTip(self.mode_btn, "切换播放模式\n手动模式: 手动控制切换句子\n自动模式: 按时间轴自动播放")
        
        # 文件加载按钮
        load_btn = tk.Button(
            top_frame,
            text="📂 加载文件",
            font=('Microsoft YaHei', 11),
            bg='#0078d4',
            fg='white',
            command=self.load_file,
            width=12,
            relief=tk.FLAT,
            cursor='hand2'
        )
        load_btn.pack(side=tk.LEFT, padx=5, pady=15)
        ToolTip(load_btn, "打开文件选择对话框\n加载Whisper转录文件")
        
        # 统计按钮
        stats_btn = tk.Button(
            top_frame,
            text="📊 统计",
            font=('Microsoft YaHei', 11),
            bg='#2d8659',
            fg='white',
            command=self.show_statistics,
            width=10,
            relief=tk.FLAT,
            cursor='hand2'
        )
        stats_btn.pack(side=tk.LEFT, padx=5, pady=15)
        ToolTip(stats_btn, "查看录制统计信息 (快捷键: F5)\n显示完成度、平均分等")
        
        # 批量操作按钮
        batch_btn = tk.Button(
            top_frame,
            text="🔄 低分重录",
            font=('Microsoft YaHei', 11),
            bg='#cc6600',
            fg='white',
            command=self.mark_low_scores_for_rerecord,
            width=12,
            relief=tk.FLAT,
            cursor='hand2'
        )
        batch_btn.pack(side=tk.LEFT, padx=5, pady=15)
        ToolTip(batch_btn, "标记低分句子\n批量跳转需要重录的句子")
        
        # 一键静音裁剪按钮
        auto_trim_btn = tk.Button(
            top_frame,
            text="✂️ 静音裁剪",
            font=('Microsoft YaHei', 11),
            bg='#9955aa',
            fg='white',
            command=self.batch_auto_trim_silence,
            width=12,
            relief=tk.FLAT,
            cursor='hand2'
        )
        auto_trim_btn.pack(side=tk.LEFT, padx=5, pady=15)
        ToolTip(auto_trim_btn, "一键裁剪所有已录音频首尾静音\n根据波形自动识别前后平坦区域")
        
        # 状态标签
        self.status_label = tk.Label(
            top_frame,
            text="请加载转录文件",
            font=('Microsoft YaHei', 10),
            bg='#1a1a1a',
            fg='#aaaaaa'
        )
        self.status_label.pack(side=tk.LEFT, padx=25)
        
        # 进度信息
        self.progress_label = tk.Label(
            top_frame,
            text="0/0",
            font=('Arial', 11, 'bold'),
            bg='#1a1a1a',
            fg='#00ff00'
        )
        self.progress_label.pack(side=tk.RIGHT, padx=15)
        
        # 主内容区域（左侧列表 + 右侧歌词）
        main_container = tk.Frame(self.root, bg='#000000')
        main_container.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 左侧语句列表面板
        list_frame = tk.Frame(main_container, bg='#1a1a1a', width=350)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(10, 5))
        list_frame.pack_propagate(False)
        
        # 列表标题
        list_title = tk.Label(
            list_frame,
            text="📝 语句列表",
            font=('Microsoft YaHei', 12, 'bold'),
            bg='#1a1a1a',
            fg='#ffffff'
        )
        list_title.pack(pady=10)
        
        # 创建滚动列表
        list_scroll_frame = tk.Frame(list_frame, bg='#1a1a1a')
        list_scroll_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.list_canvas = Canvas(
            list_scroll_frame,
            bg='#1a1a1a',
            highlightthickness=0
        )
        
        # 自定义样式的滚动条
        scrollbar = tk.Scrollbar(
            list_scroll_frame,
            orient=tk.VERTICAL,
            command=self.list_canvas.yview,
            width=16,
            bg='#2a2a2a',
            troughcolor='#1a1a1a',
            activebackground='#4a4a4a',
            highlightthickness=0,
            borderwidth=1,
            relief=tk.FLAT
        )
        
        self.list_scrollable_frame = tk.Frame(self.list_canvas, bg='#1a1a1a')
        
        self.list_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))
        )
        
        self.list_canvas.create_window((0, 0), window=self.list_scrollable_frame, anchor="nw")
        self.list_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.list_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 鼠标滚轮绑定
        self.list_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # KTV歌词显示区域（使用Canvas实现多行滚动和逐字高亮）
        lyrics_frame = tk.Frame(main_container, bg='#000000')
        lyrics_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 10))
        
        self.lyrics_canvas = Canvas(
            lyrics_frame,
            bg='#000000',
            highlightthickness=0,
            height=500
        )
        self.lyrics_canvas.pack(fill=tk.BOTH, expand=True)
        
        # 进度条
        progress_bar_frame = tk.Frame(self.root, bg='#000000')
        progress_bar_frame.pack(fill=tk.X, padx=50, pady=10)
        
        self.progress_bar = Canvas(
            progress_bar_frame,
            bg='#1a1a1a',
            height=8,
            highlightthickness=0
        )
        self.progress_bar.pack(fill=tk.X)
        
        # 时间显示
        time_frame = tk.Frame(self.root, bg='#000000')
        time_frame.pack(fill=tk.X, padx=50)
        
        self.time_label = tk.Label(
            time_frame,
            text="00:00.00 / 00:00.00",
            font=('Arial', 11),
            bg='#000000',
            fg='#666666'
        )
        self.time_label.pack()
        
        
        # 底部控制按钮区
        control_frame = tk.Frame(self.root, bg='#1a1a1a', height=100)
        control_frame.pack(fill=tk.X, side=tk.BOTTOM)
        control_frame.pack_propagate(False)
        
        # 按钮容器（居中）
        buttons_container = tk.Frame(control_frame, bg='#1a1a1a')
        buttons_container.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # 导航按钮
        self.prev_btn = tk.Button(
            buttons_container,
            text="⏮",
            font=('Arial', 18),
            bg='#333333',
            fg='white',
            command=self.prev_segment,
            width=4,
            height=1,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.prev_btn.grid(row=0, column=0, padx=5)
        ToolTip(self.prev_btn, "上一句 (快捷键: ←)\n跳转到前一个句子")
        
        # 播放按钮
        self.play_btn = tk.Button(
            buttons_container,
            text="▶",
            font=('Arial', 24, 'bold'),
            bg='#00aa00',
            fg='white',
            command=self.toggle_play,
            width=5,
            height=1,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.play_btn.grid(row=0, column=1, padx=10)
        ToolTip(self.play_btn, "播放/暂停 (快捷键: 空格)\n开始逐字高亮播放\n帮助你熟悉节奏")
        
        # 下一句按钮
        self.next_btn = tk.Button(
            buttons_container,
            text="⏭",
            font=('Arial', 18),
            bg='#333333',
            fg='white',
            command=self.next_segment,
            width=4,
            height=1,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.next_btn.grid(row=0, column=2, padx=5)
        ToolTip(self.next_btn, "下一句 (快捷键: →)\n跳转到后一个句子")
        
        # 录音按钮
        self.record_btn = tk.Button(
            buttons_container,
            text="⏺",
            font=('Arial', 20),
            bg='#cc0000',
            fg='white',
            command=self.toggle_recording,
            width=5,
            height=1,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.record_btn.grid(row=0, column=3, padx=20)
        ToolTip(self.record_btn, "开始/停止录制 (快捷键: R)\n录制时会显示逐字高亮提示\n跟着提示朗读文案")
        
        # AI演示按钮
        self.ai_demo_btn = tk.Button(
            buttons_container,
            text="🤖",
            font=('Arial', 16),
            bg='#9933cc',
            fg='white',
            command=self.play_ai_demo,
            width=4,
            height=1,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.ai_demo_btn.grid(row=0, column=4, padx=5)
        ToolTip(self.ai_demo_btn, "AI朗读演示 (快捷键: D)\n听听AI怎么读这句话\n可以模仿AI的节奏录制")
        
        # 预览按钮
        self.preview_btn = tk.Button(
            buttons_container,
            text="🔊",
            font=('Arial', 16),
            bg='#0078d4',
            fg='white',
            command=self.preview_recording,
            width=4,
            height=1,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.preview_btn.grid(row=0, column=5, padx=5)
        ToolTip(self.preview_btn, "预览录音 (快捷键: P)\n播放当前句子的录音\n检查录制质量")
        
        # 导出按钮
        self.export_btn = tk.Button(
            buttons_container,
            text="💾",
            font=('Arial', 16),
            bg='#ff8800',
            fg='white',
            command=self.export_audio,
            width=4,
            height=1,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.export_btn.grid(row=0, column=6, padx=5)
        ToolTip(self.export_btn, "导出完整音频 (快捷键: Enter)\n合并所有录音片段\n生成完整的WAV文件")
        
        # 试听合并版
        self.preview_merge_btn = tk.Button(
            buttons_container,
            text="📀",
            font=('Arial', 16),
            bg='#555555',
            fg='white',
            command=self.preview_merged_audio,
            width=4,
            height=1,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.preview_merge_btn.grid(row=0, column=7, padx=5)
        ToolTip(self.preview_merge_btn, "试听合并版\n预览导出后的完整音频效果")
        
        # 音量监控条（录制时显示）
        volume_frame = tk.Frame(control_frame, bg='#1a1a1a')
        volume_frame.place(relx=0.1, rely=0.8, anchor=tk.W)
        
        volume_label = tk.Label(
            volume_frame,
            text="音量:",
            font=('Microsoft YaHei', 9),
            bg='#1a1a1a',
            fg='#888888'
        )
        volume_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.volume_meter = Canvas(
            volume_frame,
            width=200,
            height=20,
            bg='#333333',
            highlightthickness=0
        )
        self.volume_meter.pack(side=tk.LEFT)
        self.volume_bar = self.volume_meter.create_rectangle(
            0, 0, 0, 20,
            fill='#00ff00',
            outline=''
        )
        
        # 播放速度控制按钮
        speed_frame = tk.Frame(control_frame, bg='#1a1a1a')
        speed_frame.place(relx=0.9, rely=0.8, anchor=tk.E)
        
        speed_label = tk.Label(
            speed_frame,
            text="速度:",
            font=('Microsoft YaHei', 9),
            bg='#1a1a1a',
            fg='#888888'
        )
        speed_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.speed_btn = tk.Button(
            speed_frame,
            text="1.0x",
            font=('Arial', 10, 'bold'),
            bg='#444444',
            fg='white',
            command=self.cycle_speed,
            width=6,
            relief=tk.FLAT,
            cursor='hand2'
        )
        self.speed_btn.pack(side=tk.LEFT)
        ToolTip(self.speed_btn, "切换播放速度\n点击循环: 0.5x → 0.75x → 1.0x → 1.25x → 1.5x")
        
    def _on_mousewheel(self, event):
        """鼠标滚轮滚动列表"""
        self.list_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def bind_shortcuts(self):
        """绑定快捷键"""
        self.root.bind('<space>', lambda e: self.toggle_play())
        self.root.bind('<Left>', lambda e: self.prev_segment())
        self.root.bind('<Right>', lambda e: self.next_segment())
        self.root.bind('r', lambda e: self._on_record_shortcut())
        self.root.bind('R', lambda e: self._on_record_shortcut())
        self.root.bind('d', lambda e: self.play_ai_demo())
        self.root.bind('D', lambda e: self.play_ai_demo())
        self.root.bind('p', lambda e: self.preview_recording())
        self.root.bind('P', lambda e: self.preview_recording())
        self.root.bind('<Return>', lambda e: self.export_audio())
        self.root.bind('s', lambda e: self.cycle_speed())
        self.root.bind('S', lambda e: self.cycle_speed())
        self.root.bind('<Home>', lambda e: self.jump_to_segment(0) if self.segments else None)
        self.root.bind('<End>', lambda e: self.jump_to_segment(len(self.segments)-1) if self.segments else None)
        self.root.bind('<Delete>', lambda e: self.delete_recording())
        self.root.bind('<F1>', lambda e: self.show_help_dialog())
        self.root.bind('<F5>', lambda e: self.show_statistics())
        self.root.bind('<Control-s>', lambda e: self.save_project())
        self.root.bind('<Control-o>', lambda e: self.load_project())
        
    def load_file(self):
        """加载转录文件"""
        file_path = filedialog.askopenfilename(
            title="选择转录文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialdir=os.getcwd()
        )
        
        if file_path:
            self.load_transcript_file(file_path)
    
    def load_transcript_file(self, file_path):
        """加载转录文件的通用方法"""
        try:
            self.current_transcript_path = file_path
            self.segments = parse_transcript_file(file_path)
            self.current_index = 0
            self.recording_states = {}
            self.skip_states = {}
            self.merge_groups = [[i] for i in range(len(self.segments))]
            self.word_progress = 0.0
            self.invalidate_score_cache()  # 新文件加载，清空评分缓存
            
            # 根据 recordings 目录中已存在的文件恢复录制状态
            self.restore_recording_states()
            
            self.status_label.config(
                text=f"已加载 {len(self.segments)} 个句子",
                fg='#aaaaaa'
            )
            self.update_display()
            self.refresh_list()
            
            messagebox.showinfo("成功", f"已加载 {len(self.segments)} 个句子")
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败: {str(e)}")
    
    def restore_recording_states(self):
        """根据 recordings 目录中的已有文件恢复录制完成标记"""
        if not self.segments:
            return
        
        for seg in self.segments:
            idx = seg.get('index')
            file_path = os.path.join(self.recordings_dir, f"segment_{idx:03d}.wav")
            if os.path.exists(file_path):
                self.recording_states[idx] = True
    
    def toggle_mode(self):
        """切换自动/手动模式"""
        self.is_auto_mode = not self.is_auto_mode
        if self.is_auto_mode:
            self.mode_btn.config(text="⏱ 自动模式", bg='#cc6600')
        else:
            self.mode_btn.config(text="📖 手动模式", bg='#333333')
    
    def update_display(self):
        """更新KTV风格显示内容"""
        if not self.segments:
            return
        
        # 更新进度
        recorded_mark = "✓" if self.recording_states.get(self.current_index) else ""
        self.progress_label.config(
            text=f"{self.current_index + 1}/{len(self.segments)} {recorded_mark}"
        )
        
        # 更新歌词显示
        self.render_ktv_lyrics()
        
        # 更新时间
        seg = self.segments[self.current_index]
        self.time_label.config(
            text=f"{format_time(seg['start_time'])} / {format_time(seg['end_time'])}"
        )
        
        # 更新进度条
        self.update_progress_bar()
    
    def render_ktv_lyrics(self):
        """渲染KTV风格歌词（多行显示，当前句逐字高亮）"""
        self.lyrics_canvas.delete("all")
        
        # 更新Canvas尺寸
        self.lyrics_canvas.update_idletasks()
        canvas_width = self.lyrics_canvas.winfo_width()
        canvas_height = self.lyrics_canvas.winfo_height()
        
        # 防止宽度为1（未初始化）
        if canvas_width < 100:
            canvas_width = 850
        if canvas_height < 100:
            canvas_height = 500
        
        if not self.segments:
            # 显示欢迎提示
            self.lyrics_canvas.create_text(
                canvas_width / 2, canvas_height / 2 - 60,
                text="🎤",
                font=('Arial', 80),
                fill='#333333',
                anchor=tk.CENTER
            )
            
            self.lyrics_canvas.create_text(
                canvas_width / 2, canvas_height / 2 + 40,
                text="KTV 提词器",
                font=('Microsoft YaHei', 32, 'bold'),
                fill='#666666',
                anchor=tk.CENTER
            )
            
            self.lyrics_canvas.create_text(
                canvas_width / 2, canvas_height / 2 + 100,
                text="拖拽转录文件到窗口，或点击上方加载按钮开始",
                font=('Microsoft YaHei', 14),
                fill='#444444',
                anchor=tk.CENTER
            )
            return
        
        # 获取Canvas尺寸（有内容时也使用相同逻辑）
        center_y = canvas_height / 2
        
        # 显示范围：当前句前后各2句
        display_range = 2
        start_idx = max(0, self.current_index - display_range)
        end_idx = min(len(self.segments), self.current_index + display_range + 1)
        
        y_offset = center_y - (self.current_index - start_idx) * 90
        
        for i in range(start_idx, end_idx):
            seg = self.segments[i]
            y_pos = y_offset + (i - start_idx) * 90
            
            is_current = (i == self.current_index)
            is_recorded = self.recording_states.get(i, False)
            
            if is_current and (self.is_playing or self.is_recording):
                # 当前播放或录制的句子 - 逐字高亮
                self.render_word_by_word(seg['text'], canvas_width / 2, y_pos, self.word_progress, is_recorded)
            else:
                # 其他句子 - 普通显示
                if is_current:
                    # 当前但未播放 - 亮白色
                    color = '#ffffff'
                    font_size = 40
                elif i < self.current_index:
                    # 已过的句子 - 暗灰色
                    color = '#444444'
                    font_size = 28
                else:
                    # 未到的句子 - 中灰色
                    color = '#777777'
                    font_size = 28
                
                # 添加录制标记
                display_text = f"✓ {seg['text']}" if is_recorded else seg['text']
                
                self.lyrics_canvas.create_text(
                    canvas_width / 2, y_pos,
                    text=display_text,
                    font=('Arial', font_size, 'bold' if is_current else 'normal'),
                    fill=color,
                    anchor=tk.CENTER,
                    width=canvas_width - 100
                )
    
    def render_word_by_word(self, text, x, y, progress, is_recorded):
        """逐字渲染，实现KTV效果"""
        words = text.split()
        total_chars = sum(len(word) + 1 for word in words) - 1  # 总字符数（包括空格）
        
        # 计算当前应该高亮到哪个字符
        current_char_idx = int(progress * total_chars)
        
        # 添加录制标记
        if is_recorded:
            self.lyrics_canvas.create_text(
                x - 500, y - 50,
                text="✓",
                font=('Arial', 20),
                fill='#00ff00',
                anchor=tk.W
            )
        
        # 录制中的闪烁指示
        if self.is_recording:
            blink = int(time.time() * 3) % 2  # 每秒闪烁3次
            if blink:
                self.lyrics_canvas.create_text(
                    x + 500, y - 50,
                    text="🔴 REC",
                    font=('Arial', 16, 'bold'),
                    fill='#ff0000',
                    anchor=tk.E
                )
        
        # 计算整行文本的宽度（用于居中）
        char_width = 18  # 近似字符宽度
        total_width = total_chars * char_width
        start_x = x - total_width / 2
        
        current_x = start_x
        char_count = 0
        
        for word_idx, word in enumerate(words):
            for char_idx, char in enumerate(word):
                # 确定颜色 - 优化配色方案
                if char_count < current_char_idx:
                    # 已读过 - 浅金色
                    color = '#ccaa00'
                    font_size = 38
                elif char_count == current_char_idx:
                    # 当前字 - 亮黄色，明显放大
                    color = '#ffff00'
                    font_size = 52
                else:
                    # 未读 - 亮白色
                    color = '#ffffff'
                    font_size = 38
                
                self.lyrics_canvas.create_text(
                    current_x, y,
                    text=char,
                    font=('Arial', font_size, 'bold'),
                    fill=color,
                    anchor=tk.W
                )
                
                current_x += char_width
                char_count += 1
            
            # 空格
            if word_idx < len(words) - 1:
                current_x += char_width
                char_count += 1
    
    def update_progress_bar(self):
        """更新进度条"""
        self.progress_bar.delete("all")
        
        if not self.segments:
            return
        
        # 更新Canvas尺寸
        self.progress_bar.update_idletasks()
        bar_width = self.progress_bar.winfo_width()
        bar_height = 8
        
        # 防止宽度未初始化
        if bar_width < 100:
            bar_width = 1100
        
        # 背景
        self.progress_bar.create_rectangle(
            0, 0, bar_width, bar_height,
            fill='#1a1a1a',
            outline=''
        )
        
        # 当前句子内的进度
        progress_width = bar_width * self.word_progress
        self.progress_bar.create_rectangle(
            0, 0, progress_width, bar_height,
            fill='#00ff00',
            outline=''
        )
        
    def prev_segment(self):
        """上一句"""
        if not self.segments:
            return
        
        if self.current_index > 0:
            self.current_index -= 1
            self.update_display()
            self.refresh_list()
    
    def next_segment(self):
        """下一句"""
        if not self.segments:
            return
        
        if self.current_index < len(self.segments) - 1:
            self.current_index += 1
            self.update_display()
            self.refresh_list()
    
    def toggle_play(self):
        """播放/暂停"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        
        self.is_playing = not self.is_playing
        
        if self.is_playing:
            self.play_btn.config(text="⏸", bg='#ff8800')
            self.start_playback()
        else:
            self.play_btn.config(text="▶", bg='#00aa00')
            self.stop_playback()
    
    def start_playback(self):
        """开始播放"""
        self.segment_start_time = time.time()
        self.word_progress = 0.0
        self.play_current_segment()
    
    def stop_playback(self):
        """停止播放"""
        if self.play_timer_id:
            self.root.after_cancel(self.play_timer_id)
            self.play_timer_id = None
        self.word_progress = 0.0
        self.update_display()
    
    def play_current_segment(self):
        """播放当前句子（动画更新）"""
        if not self.is_playing:
            return
        
        seg = self.segments[self.current_index]
        elapsed = time.time() - self.segment_start_time
        
        # 更新进度（应用播放速度）
        adjusted_duration = seg['duration'] / self.playback_speed
        if adjusted_duration > 0:
            self.word_progress = min(1.0, elapsed / adjusted_duration)
        else:
            self.word_progress = 1.0
        
        # 更新显示
        self.render_ktv_lyrics()
        self.update_progress_bar()
        
        # 更新时间显示
        current_time = seg['start_time'] + elapsed
        self.time_label.config(
            text=f"{format_time(current_time)} / {format_time(seg['end_time'])}"
        )
        
        # 检查是否播放完当前句
        if self.word_progress >= 1.0:
            if self.is_auto_mode:
                # 自动模式：等待间隔后切换到下一句
                gap_ms = int(seg['gap_after'] * 1000)
                self.play_timer_id = self.root.after(gap_ms, self.auto_next_segment)
            else:
                # 手动模式：停止在当前句
                self.word_progress = 1.0
                self.update_display()
        else:
            # 继续动画（约60fps）
            self.play_timer_id = self.root.after(16, self.play_current_segment)
    
    def auto_next_segment(self):
        """自动模式下切换到下一句"""
        if not self.is_playing:
            return
        
        if self.current_index < len(self.segments) - 1:
            self.current_index += 1
            self.segment_start_time = time.time()
            self.word_progress = 0.0
            self.update_display()
            self.play_current_segment()
        else:
            # 播放完毕
            self.is_playing = False
            self.play_btn.config(text="▶", bg='#00aa00')
            self.word_progress = 0.0
            self.update_display()
            messagebox.showinfo("完成", "播放完毕")
    
    def _on_record_shortcut(self):
        """录制快捷键，return break 防止按键重复触发"""
        self.toggle_recording()
        return "break"
    
    def toggle_recording(self):
        """开始/停止录制"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        
        # 防抖：400ms 内忽略重复触发（双击、快捷键+点击同时触发）
        now = time.time()
        if now - self._last_toggle_recording_time < 0.4:
            return
        self._last_toggle_recording_time = now
        
        try:
            self.record_btn.config(state=tk.DISABLED)
            if self.is_recording:
                self.stop_recording()
            else:
                self.start_recording()
        except Exception as e:
            self.is_recording = False
            self.record_btn.config(text="⏺", bg='#cc0000', state=tk.NORMAL)
            messagebox.showerror("录制错误", str(e))
        finally:
            self.record_btn.config(state=tk.NORMAL)
    
    def start_recording(self):
        """开始录制"""
        # 若已有未关闭的 stream，先尝试关闭（避免设备被占用）
        if hasattr(self, 'stream') and self.stream is not None:
            try:
                self.stream.stop()
                self.stream.close()
            except Exception:
                pass
            self.stream = None
        
        self.is_recording = True
        self.recording_data = []
        self.record_btn.config(text="⏹", bg='#ff0000')
        
        try:
            def callback(indata, frames, t, status):
                if status:
                    print(f"录音状态: {status}")
                self.recording_data.append(indata.copy())
            
            # blocksize 小一些、latency='low' 减少启动延迟，避免开头单词漏录
            self.stream = sd.InputStream(
                samplerate=self.sample_rate,
                channels=1,
                callback=callback,
                blocksize=512,
                latency='low'
            )
            self.stream.start()
        except Exception as e:
            self.is_recording = False
            self.record_btn.config(text="⏺", bg='#cc0000')
            raise RuntimeError(f"无法启动录音设备: {e}\n请检查麦克风是否被占用。")
        
        # 启动音量监控
        self.update_volume_meter()
        
        # 启动逐字高亮动画（帮助录制）
        self.segment_start_time = time.time()
        self.word_progress = 0.0
        self.start_recording_animation()
    
    def stop_recording(self):
        """停止录制并保存"""
        if not self.is_recording:
            return
        
        # 停止音量监控
        if self.volume_update_id:
            self.root.after_cancel(self.volume_update_id)
            self.volume_update_id = None
        self.current_volume = 0.0
        if hasattr(self, 'volume_meter') and hasattr(self, 'volume_bar'):
            self.volume_meter.coords(self.volume_bar, 0, 0, 0, 20)
        
        self.is_recording = False
        self.record_btn.config(text="⏺", bg='#cc0000')
        
        # 停止逐字高亮动画
        if self.play_timer_id:
            self.root.after_cancel(self.play_timer_id)
            self.play_timer_id = None
        self.word_progress = 0.0
        self.update_display()
        
        # 停止录音（容错：即使 stop/close 失败也继续保存）
        if hasattr(self, 'stream') and self.stream is not None:
            try:
                self.stream.stop()
                self.stream.close()
            except Exception as e:
                print(f"停止录音流时出错: {e}")
            self.stream = None
        
        # 保存录音
        if self.recording_data:
            recording = np.concatenate(self.recording_data, axis=0)
            file_path = os.path.join(
                self.recordings_dir,
                f"segment_{self.current_index:03d}.wav"
            )
            
            with wave.open(file_path, 'wb') as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(self.sample_rate)
                wf.writeframes((recording * 32767).astype(np.int16).tobytes())
            
            self.recording_states[self.current_index] = True
            self.invalidate_score_cache(self.current_index)  # 新录制需重新计算评分
            if self.current_index in self.audio_cache:
                del self.audio_cache[self.current_index]  # 重录覆盖后需清除缓存，否则预览会播旧音频
            self.update_display()
            
            # 录制完成后用完整刷新，避免单项更新导致列表项消失
            self.root.after(0, self.refresh_list)
            
            # 状态栏提示（2秒后自动消失）
            self.status_label.config(text=f"✓ 已保存第 {self.current_index + 1} 句", fg="green")
            self.root.after(2000, lambda: self.status_label.config(text="就绪", fg="black"))
    
    def start_recording_animation(self):
        """录制时的逐字高亮动画"""
        if not self.is_recording:
            return
        
        seg = self.segments[self.current_index]
        elapsed = time.time() - self.segment_start_time
        
        # 更新进度
        if seg['duration'] > 0:
            self.word_progress = min(1.0, elapsed / seg['duration'])
        else:
            self.word_progress = 1.0
        
        # 更新显示（显示逐字高亮）
        self.render_ktv_lyrics()
        self.update_progress_bar()
        
        # 更新时间显示
        current_time = seg['start_time'] + elapsed
        self.time_label.config(
            text=f"{format_time(current_time)} / {format_time(seg['end_time'])}"
        )
        
        # 检查是否录制完当前句
        if self.word_progress < 1.0:
            # 继续动画（约60fps）
            self.play_timer_id = self.root.after(16, self.start_recording_animation)
        else:
            # 录制时长到了，保持在100%
            self.word_progress = 1.0
            self.render_ktv_lyrics()
            self.update_progress_bar()
    
    def refresh_list(self):
        """刷新左侧语句列表"""
        self._do_refresh_list()
    
    def _create_list_item(self, i, pack_before=None, score_override=None):
        """创建单个列表项，返回 widget info 字典。score_override 用于异步传入评分避免阻塞"""
        seg = self.segments[i]
        is_current = (i == self.current_index)
        is_recorded = self.recording_states.get(i, False)
        score = score_override if score_override is not None else (self.calculate_score(i) if is_recorded else None)
        
        if is_current:
            bg_color = '#2d5a8a'
        elif is_recorded:
            bg_color = '#8a2d2d' if (score is not None and score < 60) else '#2a2a2a'
        else:
            bg_color = '#1a1a1a'
        
        item_frame = tk.Frame(
            self.list_scrollable_frame,
            bg=bg_color,
            relief=tk.RAISED,
            borderwidth=1
        )
        pack_kw = dict(fill=tk.X, padx=5, pady=2)
        if pack_before:
            pack_kw['before'] = pack_before
        item_frame.pack(**pack_kw)
        item_frame.bind("<Button-1>", lambda e, idx=i: self.jump_to_segment(idx))
        
        top_row = tk.Frame(item_frame, bg=bg_color)
        top_row.pack(fill=tk.X, padx=5, pady=3)
        
        num_label = tk.Label(top_row, text=f"#{i+1}", font=('Arial', 10, 'bold'), bg=bg_color, fg='#888888')
        num_label.pack(side=tk.LEFT)
        num_label.bind("<Button-1>", lambda e, idx=i: self.jump_to_segment(idx))
        
        status_icon = "✓" if is_recorded else "○"
        status_color = '#00ff00' if is_recorded else '#666666'
        status_label = tk.Label(top_row, text=status_icon, font=('Arial', 12, 'bold'), bg=bg_color, fg=status_color)
        status_label.pack(side=tk.LEFT, padx=5)
        status_label.bind("<Button-1>", lambda e, idx=i: self.jump_to_segment(idx))
        
        # 显示实际录音时长（已录制时）或预期时长
        dur = self.get_recording_duration(i) if is_recorded else None
        if dur is not None:
            dur_text = f"实际 {dur:.1f}s"
            if seg['duration'] > 0:
                dur_text += f" (预期{seg['duration']:.1f}s)"
        else:
            dur_text = f"预期 {seg['duration']:.1f}s"
        dur_label = tk.Label(top_row, text=dur_text, font=('Arial', 9), bg=bg_color, fg='#888888')
        dur_label.pack(side=tk.LEFT, padx=(0, 8))
        dur_label.bind("<Button-1>", lambda e, idx=i: self.jump_to_segment(idx))
        
        score_label, waveform_btn, trim_btn, skip_btn, merge_btn = None, None, None, None, None
        # 跳过按钮（语气词，合成时填静音）
        is_skip = self.skip_states.get(i, False)
        skip_btn = tk.Button(top_row, text="⏭" if is_skip else "跳", font=('Arial', 9), bg=bg_color,
            fg='#ff9900' if is_skip else '#666666', command=lambda idx=i: self.toggle_skip(idx),
            relief=tk.FLAT, cursor='hand2', padx=2)
        skip_btn.pack(side=tk.RIGHT)
        ToolTip(skip_btn, "跳过(语气词)：合成时用静音填充，不播放录音")
        # 合并按钮（与下一句合并）
        if i < len(self.segments) - 1:
            merge_btn = tk.Button(top_row, text="🔗", font=('Arial', 9), bg=bg_color, fg='#00aaff',
                command=lambda idx=i: self.merge_with_next(idx), relief=tk.FLAT, cursor='hand2', padx=2)
            merge_btn.pack(side=tk.RIGHT)
            ToolTip(merge_btn, "与下一句合并：合成时连续播放无间隔")
        if score is not None:
            score_color = '#00ff00' if score >= 80 else ('#ffaa00' if score >= 60 else '#ff3333')
            score_text = f"⭐ {score}" if score >= 80 else (f"★ {score}" if score >= 60 else f"⚠ {score}")
            score_label = tk.Label(top_row, text=score_text, font=('Arial', 10, 'bold'), bg=bg_color, fg=score_color)
            score_label.pack(side=tk.RIGHT)
            score_label.bind("<Button-1>", lambda e, idx=i: self.jump_to_segment(idx))
            trim_btn = tk.Button(top_row, text="✂️", font=('Arial', 10), bg=bg_color, fg='#ffaa00',
                command=lambda idx=i: self.show_trim_dialog(idx), relief=tk.FLAT, cursor='hand2', padx=2)
            trim_btn.pack(side=tk.RIGHT)
            ToolTip(trim_btn, "裁剪首尾空白 / 若超时则加速")
            waveform_btn = tk.Button(top_row, text="📊", font=('Arial', 10), bg=bg_color, fg='#00aaff',
                command=lambda idx=i: self.show_waveform(idx), relief=tk.FLAT, cursor='hand2', padx=2)
            waveform_btn.pack(side=tk.RIGHT, padx=(0, 5))
            ToolTip(waveform_btn, "查看音频波形")
        
        text_display = seg['text'][:40] + "..." if len(seg['text']) > 40 else seg['text']
        text_label = tk.Label(item_frame, text=text_display, font=('Arial', 9), bg=bg_color, fg='#cccccc',
            wraplength=320, justify=tk.LEFT, anchor=tk.W)
        text_label.pack(fill=tk.X, padx=5, pady=(0, 5))
        text_label.bind("<Button-1>", lambda e, idx=i: self.jump_to_segment(idx))
        
        info = {
            "frame": item_frame, "top_row": top_row, "num_label": num_label,
            "status_label": status_label, "dur_label": dur_label, "text_label": text_label,
            "score_label": score_label, "trim_btn": trim_btn, "waveform_btn": waveform_btn,
            "skip_btn": skip_btn, "merge_btn": merge_btn,
        }
        self._list_item_widgets[i] = info
        return info
    
    def _update_list_item_safe(self, index):
        """安全地更新单个列表项（供 after 回调用）"""
        if 0 <= index < len(self.segments) and self._list_item_widgets:
            self._update_list_item(index)
    
    def _update_list_item_async(self, index):
        """后台计算评分后更新单项（录制结束用，避免阻塞 UI）"""
        if not (0 <= index < len(self.segments) and self._list_item_widgets):
            return
        def compute_then_update():
            score = self.calculate_score(index)
            self.root.after(0, lambda: self._update_list_item_with_score(index, score))
        threading.Thread(target=compute_then_update, daemon=True).start()
    
    def _update_list_item_with_score(self, index, score):
        """用已知评分更新单项（主线程调用）"""
        if index not in self._list_item_widgets:
            return
        if not (0 <= index < len(self.segments)):
            return
        old_info = self._list_item_widgets.pop(index)
        old_frame = old_info["frame"]
        if not old_frame.winfo_exists():
            self._create_list_item(index, pack_before=None, score_override=score)
            return
        children = list(self.list_scrollable_frame.winfo_children())
        try:
            pos = children.index(old_frame)
        except ValueError:
            pos = -1
        next_sibling = children[pos + 1] if 0 <= pos < len(children) - 1 else None
        old_frame.destroy()
        if next_sibling and not next_sibling.winfo_exists():
            next_sibling = None
        try:
            self._create_list_item(index, pack_before=next_sibling, score_override=score)
        except Exception:
            self._do_refresh_list()
    
    def _update_list_item(self, index):
        """仅更新单个列表项（用于录制结束等场景，避免全量重建）"""
        if index not in self._list_item_widgets or not (0 <= index < len(self.segments)):
            return
        old_info = self._list_item_widgets.pop(index)
        old_frame = old_info["frame"]
        if not old_frame.winfo_exists():
            self._create_list_item(index, pack_before=None)
            return
        children = list(self.list_scrollable_frame.winfo_children())
        try:
            pos = children.index(old_frame)
        except ValueError:
            pos = -1
        next_sibling = children[pos + 1] if 0 <= pos < len(children) - 1 else None
        old_frame.destroy()
        if next_sibling and not next_sibling.winfo_exists():
            next_sibling = None
        try:
            self._create_list_item(index, pack_before=next_sibling)
        except Exception:
            self._do_refresh_list()
    
    def _do_refresh_list(self):
        """实际执行列表刷新"""
        # 清空现有列表
        self._list_item_widgets.clear()
        for widget in self.list_scrollable_frame.winfo_children():
            widget.destroy()
        
        if not self.segments:
            # 显示提示信息
            tip_label = tk.Label(
                self.list_scrollable_frame,
                text="📂\n\n拖拽文件到此处\n或\n点击上方加载按钮",
                font=('Microsoft YaHei', 12),
                bg='#1a1a1a',
                fg='#666666',
                justify=tk.CENTER
            )
            tip_label.pack(expand=True, pady=50)
            return
        
        # 创建每个句子的列表项
        for i, seg in enumerate(self.segments):
            self._create_list_item(i)
        
        self._list_current_index = self.current_index
    
    def update_list_highlight_only(self, prev_index, curr_index):
        """仅更新高亮状态（不重建列表），用于跳转时避免卡顿"""
        def apply_bg(widgets, bg_color):
            for w in widgets:
                if w and w.winfo_exists():
                    try:
                        w.config(bg=bg_color)
                    except tk.TclError:
                        pass
        
        for idx in (prev_index, curr_index):
            if idx is None or idx not in self._list_item_widgets:
                continue
            info = self._list_item_widgets[idx]
            is_current = (idx == self.current_index)
            is_recorded = self.recording_states.get(idx, False)
            score = self.calculate_score(idx) if is_recorded else None
            
            if is_current:
                bg_color = '#2d5a8a'
            elif is_recorded:
                bg_color = '#8a2d2d' if (score is not None and score < 60) else '#2a2a2a'
            else:
                bg_color = '#1a1a1a'
            
            apply_bg([
                info["frame"], info["top_row"], info["num_label"],
                info["status_label"], info.get("dur_label"), info["text_label"],
                info.get("score_label"), info.get("trim_btn"), info.get("waveform_btn")
            ], bg_color)
    
    def invalidate_score_cache(self, index=None):
        """使评分缓存失效（index 为 None 时清空全部）"""
        if index is None:
            self._score_cache.clear()
        else:
            self._score_cache.pop(index, None)
    
    def get_recording_duration(self, index):
        """获取指定段的实际录音时长（秒），无录音返回 None"""
        file_path = os.path.join(self.recordings_dir, f"segment_{index:03d}.wav")
        if not os.path.exists(file_path):
            return None
        try:
            with wave.open(file_path, 'rb') as wf:
                return wf.getnframes() / float(wf.getframerate())
        except Exception:
            return None
    
    def calculate_score(self, index):
        """计算录音质量评分（0-100），使用缓存避免重复计算"""
        if index in self._score_cache:
            return self._score_cache[index]
        
        file_path = os.path.join(self.recordings_dir, f"segment_{index:03d}.wav")
        
        if not os.path.exists(file_path):
            return None
        
        try:
            with wave.open(file_path, 'rb') as wf:
                frames = wf.getnframes()
                rate = wf.getframerate()
                duration = frames / float(rate)
                
                # 读取音频数据
                audio_data = np.frombuffer(wf.readframes(frames), dtype=np.int16)
                audio_float = audio_data.astype(np.float32) / 32767.0
            
            expected_duration = self.segments[index]['duration']
            
            # 评分标准
            score = 100
            
            # 1. 时长匹配度 (50分)
            duration_diff = abs(duration - expected_duration)
            if duration_diff > expected_duration * 0.3:
                score -= 30  # 时长差异超过30%
            elif duration_diff > expected_duration * 0.15:
                score -= 15  # 时长差异超过15%
            elif duration_diff > expected_duration * 0.05:
                score -= 5   # 时长差异超过5%
            
            # 2. 音量水平 (30分)
            rms = np.sqrt(np.mean(audio_float ** 2))
            if rms < 0.01:
                score -= 30  # 音量太小
            elif rms < 0.05:
                score -= 15  # 音量偏小
            elif rms > 0.8:
                score -= 15  # 音量过大（可能削波）
            
            # 3. 静音检测 (20分)
            silence_threshold = 0.01
            silence_frames = np.sum(np.abs(audio_float) < silence_threshold)
            silence_ratio = silence_frames / len(audio_float)
            
            if silence_ratio > 0.5:
                score -= 20  # 超过50%是静音
            elif silence_ratio > 0.3:
                score -= 10  # 超过30%是静音
            
            score = max(0, min(100, score))
            self._score_cache[index] = score
            return score
            
        except Exception as e:
            print(f"评分计算错误: {e}")
            return 50  # 默认分数
    
    def jump_to_segment(self, index):
        """跳转到指定句子"""
        if 0 <= index < len(self.segments):
            prev_index = self.current_index
            self.current_index = index
            self.word_progress = 0.0
            self.update_display()
            # 列表已构建时仅更新高亮，避免全量重建卡顿
            if self._list_item_widgets:
                self.update_list_highlight_only(prev_index, index)
            else:
                self.refresh_list()
    
    def init_tts_engine(self):
        """初始化TTS引擎"""
        try:
            import pyttsx3
            # 每次都创建新实例，避免状态问题
            # self.tts_engine会在每次使用时重新创建
            self.tts_available = True
        except Exception as e:
            print(f"TTS初始化警告: {e}")
            self.tts_available = False
    
    def play_ai_demo(self):
        """AI朗读演示（线程安全）"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        
        # 线程安全检查
        if not self.ai_lock.acquire(blocking=False):
            messagebox.showwarning("提示", "AI正在朗读中，请稍候...")
            return
        
        if self.is_ai_playing:
            # 停止播放
            sd.stop()
            self.is_ai_playing = False
            self.ai_demo_btn.config(text="🤖", bg='#9933cc')
            self.ai_lock.release()
            return
        
        if not self.tts_available:
            messagebox.showerror("错误", "TTS引擎不可用\n请安装: pip install pyttsx3")
            return
        
        seg = self.segments[self.current_index]
        text = seg['text']
        
        # 标记为播放中
        self.is_ai_playing = True
        self.ai_demo_btn.config(text="⏹", bg='#ff8800')
        
        # 生成AI音频文件
        ai_file = os.path.join(self.ai_demo_dir, f"ai_demo_{self.current_index:03d}.wav")
        
        def generate_and_play():
            engine = None
            try:
                import pyttsx3
                import time
                
                # 每次创建新的引擎实例
                engine = pyttsx3.init()
                engine.setProperty('rate', 150)
                engine.setProperty('volume', 0.9)
                
                # 尝试设置英文语音
                voices = engine.getProperty('voices')
                for voice in voices:
                    if 'english' in voice.name.lower() or 'en' in voice.id.lower():
                        engine.setProperty('voice', voice.id)
                        break
                
                # 生成音频文件
                engine.save_to_file(text, ai_file)
                engine.runAndWait()
                
                # 等待文件生成
                time.sleep(0.3)
                
                if os.path.exists(ai_file):
                    # 播放生成的音频
                    with wave.open(ai_file, 'rb') as wf:
                        audio_data = np.frombuffer(wf.readframes(wf.getnframes()), dtype=np.int16)
                        audio_float = audio_data.astype(np.float32) / 32767.0
                        sample_rate = wf.getframerate()
                    
                    # 播放
                    sd.play(audio_float, sample_rate)
                    sd.wait()
                
                # 播放完成后重置按钮
                self.root.after(0, lambda: self.ai_demo_btn.config(text="🤖", bg='#9933cc'))
                self.is_ai_playing = False
                self.ai_lock.release()  # 释放锁
                
            except Exception as e:
                print(f"AI演示错误: {e}")
                self.root.after(0, lambda: self.ai_demo_btn.config(text="🤖", bg='#9933cc'))
                self.is_ai_playing = False
                self.ai_lock.release()  # 确保异常时也释放锁
                self.root.after(0, lambda: messagebox.showerror("错误", f"AI朗读失败: {str(e)}"))
            finally:
                # 清理引擎
                if engine:
                    try:
                        del engine
                    except:
                        pass
        
        # 在后台线程中生成和播放
        thread = threading.Thread(target=generate_and_play, daemon=True)
        thread.start()
    
    def refresh_initial_display(self):
        """刷新初始显示"""
        # 强制更新窗口，确保Canvas获得正确的宽高
        self.root.update_idletasks()
        
        # 刷新显示
        if not self.segments:
            self.refresh_list()
            self.render_ktv_lyrics()
            self.update_progress_bar()
        
        # 绑定窗口大小改变事件
        self.root.bind('<Configure>', self.on_window_resize)
    
    def on_window_resize(self, event):
        """窗口大小改变时重新渲染"""
        # 只在窗口大小改变时才重新渲染，避免频繁刷新
        if event.widget == self.root:
            if hasattr(self, '_resize_after_id'):
                self.root.after_cancel(self._resize_after_id)
            self._resize_after_id = self.root.after(200, self.render_ktv_lyrics)
    
    def preview_recording(self):
        """预览当前句子的录音"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        
        file_path = os.path.join(
            self.recordings_dir,
            f"segment_{self.current_index:03d}.wav"
        )
        
        if not os.path.exists(file_path):
            messagebox.showwarning("提示", "当前句子还未录制")
            return
        
        if self.is_previewing:
            # 停止预览
            sd.stop()
            self.is_previewing = False
            self.preview_btn.config(text="🔊", bg='#0078d4')
        else:
            # 开始预览（使用缓存）
            try:
                audio_data = self.get_cached_audio(self.current_index)
                if audio_data is None:
                    messagebox.showerror("错误", "无法加载音频文件")
                    return
                
                self.is_previewing = True
                self.preview_btn.config(text="⏹", bg='#ff8800')
                
                def finished_callback():
                    self.is_previewing = False
                    self.preview_btn.config(text="🔊", bg='#0078d4')
                
                sd.play(audio_data, self.sample_rate, blocking=False)
                
                # 设置定时器检查播放完成
                duration_ms = int(len(audio_data) / self.sample_rate * 1000)
                self.root.after(duration_ms, finished_callback)
                
            except Exception as e:
                messagebox.showerror("错误", f"预览失败: {str(e)}")
    
    def preview_merged_audio(self):
        """试听合并版：预览导出后的完整音频"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        if not self.recording_states:
            messagebox.showwarning("提示", "还没有录制任何音频")
            return
        if self.is_previewing:
            sd.stop()
            self.is_previewing = False
            self._cancel_preview_timeline()
            self.preview_merge_btn.config(text="📀", bg='#555555')
            self.preview_btn.config(text="🔊", bg='#0078d4')
            return
        try:
            frames_bytes, rate, timeline = self._build_merged_audio_frames()
            if not frames_bytes:
                messagebox.showwarning("提示", "没有可播放的音频")
                return
            audio_data = np.frombuffer(frames_bytes, dtype=np.int16).astype(np.float32) / 32767.0
            self.is_previewing = True
            self._preview_after_ids = []
            self.preview_merge_btn.config(text="⏹", bg='#ff8800')
            # 开始时立即切换到第一句
            if timeline:
                self.jump_to_segment(timeline[0][0])
            def finished():
                self.is_previewing = False
                self._cancel_preview_timeline()
                self.preview_merge_btn.config(text="📀", bg='#555555')
            # 按时间点切换文字高亮
            for idx, start_sec in timeline:
                ms = int(start_sec * 1000)
                aid = self.root.after(ms, lambda i=idx: self._preview_jump_to_segment(i))
                self._preview_after_ids.append(aid)
            sd.play(audio_data, rate, blocking=False)
            duration_ms = int(len(audio_data) / rate * 1000)
            self.root.after(duration_ms, finished)
        except Exception as e:
            self.is_previewing = False
            self.preview_merge_btn.config(text="📀", bg='#555555')
            messagebox.showerror("错误", f"试听失败: {str(e)}")
    
    def _cancel_preview_timeline(self):
        """取消试听合并时预定的文字切换"""
        if hasattr(self, '_preview_after_ids'):
            for aid in self._preview_after_ids:
                try:
                    self.root.after_cancel(aid)
                except Exception:
                    pass
            self._preview_after_ids = []
    
    def _preview_jump_to_segment(self, index):
        """试听合并时切换到指定句并刷新显示"""
        if not self.is_previewing or not (0 <= index < len(self.segments)):
            return
        self.jump_to_segment(index)
    
    def save_project(self):
        """存档当前字幕+音频到工程目录"""
        if not self.segments or not self.current_transcript_path:
            messagebox.showwarning("提示", "请先加载字幕文件并完成部分录音后再存档")
            return
        project_root = filedialog.askdirectory(
            title="选择或创建一个工程存档目录（空目录最佳）",
            mustexist=False
        )
        if not project_root:
            return
        try:
            os.makedirs(project_root, exist_ok=True)
            # 保存字幕副本
            transcript_name = os.path.basename(self.current_transcript_path)
            project_transcript = os.path.join(project_root, transcript_name)
            shutil.copy2(self.current_transcript_path, project_transcript)
            # 复制音频
            audio_dir = os.path.join(project_root, "audio")
            os.makedirs(audio_dir, exist_ok=True)
            for seg in self.segments:
                idx = seg.get("index")
                src = os.path.join(self.recordings_dir, f"segment_{idx:03d}.wav")
                dst = os.path.join(audio_dir, f"segment_{idx:03d}.wav")
                if os.path.exists(src):
                    shutil.copy2(src, dst)
            # 保存元数据
            meta = {
                "transcript_file": transcript_name,
                "segments_count": len(self.segments),
                "recording_states": self.recording_states,
                "skip_states": self.skip_states,
                "merge_groups": self.merge_groups,
            }
            meta_path = os.path.join(project_root, "project_meta.json")
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False, indent=2)
            self.current_project_path = project_root
            messagebox.showinfo("成功", f"工程已存档到：\n{project_root}")
        except Exception as e:
            messagebox.showerror("错误", f"存档失败: {e}")
    
    def load_project(self):
        """从工程存档目录恢复字幕和音频"""
        project_root = filedialog.askdirectory(
            title="选择工程存档目录（含 project_meta.json）",
            mustexist=True
        )
        if not project_root:
            return
        meta_path = os.path.join(project_root, "project_meta.json")
        if not os.path.exists(meta_path):
            messagebox.showerror("错误", "所选目录中没有 project_meta.json，无法识别为工程存档")
            return
        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
            transcript_name = meta.get("transcript_file")
            if not transcript_name:
                raise ValueError("元数据缺少 transcript_file 字段")
            project_transcript = os.path.join(project_root, transcript_name)
            if not os.path.exists(project_transcript):
                raise FileNotFoundError(f"找不到字幕文件：{project_transcript}")
            # 覆盖 recordings 目录为该工程的音频
            audio_dir = os.path.join(project_root, "audio")
            if not os.path.exists(audio_dir):
                os.makedirs(audio_dir, exist_ok=True)
            if os.path.exists(self.recordings_dir):
                try:
                    for name in os.listdir(self.recordings_dir):
                        if name.endswith(".wav"):
                            os.remove(os.path.join(self.recordings_dir, name))
                except Exception:
                    pass
            else:
                os.makedirs(self.recordings_dir, exist_ok=True)
            for name in os.listdir(audio_dir):
                if name.endswith(".wav"):
                    shutil.copy2(os.path.join(audio_dir, name),
                                 os.path.join(self.recordings_dir, name))
            # 加载字幕并套用状态
            self.load_transcript_file(project_transcript)
            self.recording_states = meta.get("recording_states", {})
            self.skip_states = meta.get("skip_states", {})
            self.merge_groups = meta.get("merge_groups", [[i for i in range(len(self.segments))]])
            self.invalidate_score_cache()
            self.refresh_list()
            self.update_display()
            self.current_project_path = project_root
            messagebox.showinfo("成功", f"已从工程恢复：\n{project_root}")
        except Exception as e:
            messagebox.showerror("错误", f"读取工程失败: {e}")

    def export_audio(self):
        """导出合并的完整音频"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        
        # 检查是否有录音
        if not self.recording_states:
            messagebox.showwarning("提示", "还没有录制任何音频")
            return
        
        # 选择保存位置
        output_path = filedialog.asksaveasfilename(
            title="保存音频文件",
            defaultextension=".wav",
            filetypes=[("WAV文件", "*.wav")],
            initialfile="output.wav"
        )
        
        if not output_path:
            return
        
        try:
            self.merge_audio_segments(output_path)
            messagebox.showinfo("成功", f"音频已导出: {output_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")
    
    def _build_merged_audio_frames(self):
        """构建合并后的音频帧序列，供导出和试听使用。返回 (frames_bytes, sample_rate, timeline)
        timeline: [(segment_index, start_sec_in_output), ...] 各句在合并音频中的开始时间"""
        if not self.segments:
            return b'', self.sample_rate, []
        groups = self.merge_groups if self.merge_groups else [[i] for i in range(len(self.segments))]
        frames_list = []
        timeline = []
        output_time = 0.0
        last_end_time = 0
        for group in sorted(groups, key=lambda g: min(g)):
            seg_first = self.segments[group[0]]
            gap = seg_first['start_time'] - last_end_time
            if gap > 0:
                silence_frames = int(gap * self.sample_rate)
                frames_list.append(np.zeros(silence_frames, dtype=np.int16).tobytes())
                output_time += gap
            for i in group:
                seg = self.segments[i]
                timeline.append((i, output_time))
                file_path = os.path.join(self.recordings_dir, f"segment_{i:03d}.wav")
                if self.skip_states.get(i):
                    duration_frames = int(seg['duration'] * self.sample_rate)
                    frames_list.append(np.zeros(duration_frames, dtype=np.int16).tobytes())
                    output_time += seg['duration']
                elif os.path.exists(file_path):
                    with wave.open(file_path, 'rb') as wf:
                        data = wf.readframes(wf.getnframes())
                        frames_list.append(data)
                        output_time += len(data) / 2 / self.sample_rate
                else:
                    duration_frames = int(seg['duration'] * self.sample_rate)
                    frames_list.append(np.zeros(duration_frames, dtype=np.int16).tobytes())
                    output_time += seg['duration']
                last_end_time = seg['end_time']
        return b''.join(frames_list), self.sample_rate, timeline
    
    def merge_audio_segments(self, output_path):
        """合并所有录音片段，插入静音间隔。支持跳过(语气词填静音)与合并组(组内无间隔)"""
        frames_bytes, rate, _ = self._build_merged_audio_frames()
        if not frames_bytes:
            return
        with wave.open(output_path, 'wb') as output_wav:
            output_wav.setnchannels(1)
            output_wav.setsampwidth(2)
            output_wav.setframerate(rate)
            output_wav.writeframes(frames_bytes)

    
    def delete_recording(self):
        """删除当前句子的录音"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        
        if not self.recording_states.get(self.current_index):
            messagebox.showinfo("提示", "当前句子还未录制")
            return
        
        result = messagebox.askyesno(
            "确认删除",
            f"确定要删除第 {self.current_index + 1} 句的录音吗？\n\n这个操作不可撤销！"
        )
        
        if not result:
            return
        
        try:
            file_path = os.path.join(
                self.recordings_dir,
                f"segment_{self.current_index:03d}.wav"
            )
            
            if os.path.exists(file_path):
                os.remove(file_path)
            
            # 清除缓存
            if self.current_index in self.audio_cache:
                del self.audio_cache[self.current_index]
            
            # 更新状态
            self.recording_states[self.current_index] = False
            self.invalidate_score_cache(self.current_index)
            self.update_display()
            self.root.after(0, self.refresh_list)
            
            self.status_label.config(text=f"✓ 已删除第 {self.current_index + 1} 句", fg="orange")
            self.root.after(2000, lambda: self.status_label.config(text="就绪", fg="black"))
            
        except Exception as e:
            messagebox.showerror("错误", f"删除失败: {str(e)}")
    
    def toggle_skip(self, index):
        """切换跳过状态（语气词，合成时填静音）"""
        if not self.segments or not (0 <= index < len(self.segments)):
            return
        self.skip_states[index] = not self.skip_states.get(index, False)
        self._update_list_item_safe(index)
        skip_count = sum(1 for v in self.skip_states.values() if v)
        self.status_label.config(
            text=f"#{index+1} 已标记为{'跳过' if self.skip_states[index] else '不跳过'} (共{skip_count}个跳过)" if skip_count else f"#{index+1} 已取消跳过",
            fg='orange'
        )
        self.root.after(2000, lambda: self.status_label.config(text="就绪", fg="black"))
    
    def merge_with_next(self, index):
        """将当前句与下一句合并（合成时连续播放无间隔）"""
        if not self.segments or index >= len(self.segments) - 1:
            return
        gi = gj = -1
        for k, g in enumerate(self.merge_groups):
            if index in g:
                gi = k
            if index + 1 in g:
                gj = k
        if gi < 0 or gj < 0:
            return
        if gi == gj:
            messagebox.showinfo("提示", "这两句已经合并")
            return
        new_group = sorted(set(self.merge_groups[gi] + self.merge_groups[gj]))
        self.merge_groups = [g for k, g in enumerate(self.merge_groups) if k not in (gi, gj)]
        self.merge_groups.append(new_group)
        self.merge_groups.sort(key=lambda g: min(g))
        self.refresh_list()
        self.status_label.config(text=f"#{index+1} 与 #{index+2} 已合并", fg='#00aa00')
        self.root.after(2000, lambda: self.status_label.config(text="就绪", fg="black"))

    def show_waveform(self, index):
        """显示音频波形"""
        file_path = os.path.join(self.recordings_dir, f"segment_{index:03d}.wav")
        
        if not os.path.exists(file_path):
            messagebox.showwarning("提示", "该句子还未录制")
            return
        
        try:
            # 读取音频数据
            with wave.open(file_path, 'rb') as wf:
                frames = wf.getnframes()
                rate = wf.getframerate()
                audio_data = np.frombuffer(wf.readframes(frames), dtype=np.int16)
                audio_float = audio_data.astype(np.float32) / 32767.0
            
            # 创建波形窗口
            waveform_window = tk.Toplevel(self.root)
            waveform_window.title(f"波形 - 第 {index + 1} 句")
            waveform_window.geometry("800x400")
            waveform_window.configure(bg='#1a1a1a')
            
            # 信息标签
            seg = self.segments[index]
            info_frame = tk.Frame(waveform_window, bg='#1a1a1a')
            info_frame.pack(fill=tk.X, padx=10, pady=5)
            
            info_text = f"时长: {len(audio_float)/rate:.2f}秒 | 采样率: {rate}Hz | 峰值: {np.max(np.abs(audio_float)):.3f}"
            info_label = tk.Label(
                info_frame,
                text=info_text,
                font=('Arial', 10),
                bg='#1a1a1a',
                fg='#aaaaaa'
            )
            info_label.pack()
            
            text_label = tk.Label(
                info_frame,
                text=f'"{seg["text"]}"',
                font=('Microsoft YaHei', 11),
                bg='#1a1a1a',
                fg='#ffffff',
                wraplength=750
            )
            text_label.pack(pady=5)
            
            # 波形Canvas
            canvas = Canvas(
                waveform_window,
                bg='#000000',
                highlightthickness=1,
                highlightbackground='#444444'
            )
            canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 绘制波形
            def draw_waveform():
                width = canvas.winfo_width()
                height = canvas.winfo_height()
                
                if width < 10 or height < 10:
                    waveform_window.after(50, draw_waveform)
                    return
                
                # 下采样以适应显示宽度
                samples_per_pixel = max(1, len(audio_float) // width)
                downsampled = []
                
                for i in range(0, len(audio_float), samples_per_pixel):
                    chunk = audio_float[i:i+samples_per_pixel]
                    if len(chunk) > 0:
                        downsampled.append(np.max(np.abs(chunk)))
                
                # 绘制波形
                center_y = height / 2
                x_scale = width / len(downsampled)
                
                # 绘制中线
                canvas.create_line(0, center_y, width, center_y, fill='#333333', width=1)
                
                # 绘制波形线
                points = []
                for i, amplitude in enumerate(downsampled):
                    x = i * x_scale
                    y_offset = amplitude * (height / 2 - 20)
                    points.append((x, center_y - y_offset))
                
                # 上半部分
                for i in range(len(points) - 1):
                    x1, y1 = points[i]
                    x2, y2 = points[i + 1]
                    canvas.create_line(x1, y1, x2, y2, fill='#00ff88', width=1)
                
                # 下半部分（镜像）
                for i in range(len(points) - 1):
                    x1, y1 = points[i]
                    x2, y2 = points[i + 1]
                    y1_mirror = center_y + (center_y - y1)
                    y2_mirror = center_y + (center_y - y2)
                    canvas.create_line(x1, y1_mirror, x2, y2_mirror, fill='#00ff88', width=1)
                
                # 填充区域
                for i, amplitude in enumerate(downsampled):
                    x = i * x_scale
                    y_offset = amplitude * (height / 2 - 20)
                    canvas.create_line(
                        x, center_y - y_offset,
                        x, center_y + y_offset,
                        fill='#00aa66',
                        width=max(1, int(x_scale))
                    )
            
            waveform_window.after(100, draw_waveform)
            
            # 关闭按钮
            btn_frame = tk.Frame(waveform_window, bg='#1a1a1a')
            btn_frame.pack(fill=tk.X, padx=10, pady=10)
            
            close_btn = tk.Button(
                btn_frame,
                text="关闭",
                font=('Microsoft YaHei', 10),
                bg='#444444',
                fg='white',
                command=waveform_window.destroy,
                width=10,
                relief=tk.FLAT,
                cursor='hand2'
            )
            close_btn.pack()
            
        except Exception as e:
            messagebox.showerror("错误", f"无法显示波形: {str(e)}")
    
    def show_trim_dialog(self, index):
        """裁剪首尾空白，可选项：若仍超时则加速到预期时长"""
        file_path = os.path.join(self.recordings_dir, f"segment_{index:03d}.wav")
        if not os.path.exists(file_path):
            messagebox.showwarning("提示", "该句子还未录制")
            return
        
        try:
            with wave.open(file_path, 'rb') as wf:
                nframes = wf.getnframes()
                rate = wf.getframerate()
                audio_int16 = np.frombuffer(wf.readframes(nframes), dtype=np.int16)
            
            audio_float = audio_int16.astype(np.float32) / 32767.0
            total_dur = nframes / float(rate)
            seg = self.segments[index]
            expected_dur = seg['duration']
            
            dlg = tk.Toplevel(self.root)
            dlg.title(f"裁剪 - 第 {index + 1} 句")
            dlg.geometry("700x460")
            dlg.configure(bg='#1a1a1a')
            dlg.transient(self.root)
            dlg.grab_set()
            
            main_f = tk.Frame(dlg, bg='#1a1a1a', padx=15, pady=12)
            main_f.pack(fill=tk.BOTH, expand=True)
            
            tk.Label(main_f, text=f"总长 {total_dur:.2f}s | 预期 {expected_dur:.2f}s | 点击波形左半设开头、右半设结尾，或拖动绿线",
                font=('Microsoft YaHei', 10), bg='#1a1a1a', fg='#aaaaaa').pack(anchor=tk.W)
            
            canvas_h = 130
            canv = Canvas(main_f, bg='#0d0d0d', highlightthickness=1, highlightbackground='#444444', height=canvas_h)
            canv.pack(fill=tk.X, pady=(8, 0))
            # 时间轴
            axis_f = tk.Frame(main_f, bg='#1a1a1a')
            axis_f.pack(fill=tk.X)
            tk.Label(axis_f, text="0s", font=('Arial', 8), bg='#1a1a1a', fg='#555555').pack(side=tk.LEFT)
            tk.Label(axis_f, text=f"   {total_dur:.1f}s", font=('Arial', 8), bg='#1a1a1a', fg='#555555').pack(side=tk.RIGHT)
            
            var_trim_start = tk.DoubleVar(value=0)
            var_trim_end = tk.DoubleVar(value=0)
            drag_line = [None]
            
            def draw_all():
                t0 = var_trim_start.get()
                t1 = var_trim_end.get()
                w = canv.winfo_width()
                if w < 20:
                    dlg.after(50, draw_all)
                    return
                h = canvas_h
                center_y = h / 2
                canv.delete("all")
                spb = max(1, len(audio_float) // w)
                amp = [np.max(np.abs(audio_float[i:i+spb])) for i in range(0, len(audio_float), spb)]
                if not amp:
                    amp = [0]
                x0 = (t0 / total_dur) * w if total_dur > 0 else 0
                x1 = w - (t1 / total_dur) * w if total_dur > 0 else w
                canv.create_rectangle(0, 0, x0, h, fill='#2a1a1a', outline='')
                canv.create_rectangle(x1, 0, w, h, fill='#2a1a1a', outline='')
                xs = np.linspace(0, w - 1, len(amp))
                scale = (h / 2 - 10)
                for i in range(len(amp) - 1):
                    xa, xb = xs[i], xs[i + 1]
                    xm = (xa + xb) / 2
                    col = '#00aa66' if x0 <= xm <= x1 else '#554444'
                    ya = center_y - amp[i] * scale
                    yb = center_y - amp[i + 1] * scale
                    canv.create_line(xa, ya, xb, yb, fill=col, width=1)
                    canv.create_line(xa, center_y + (center_y - ya), xb, center_y + (center_y - yb), fill=col, width=1)
                for tag, x in [("line_start", x0), ("line_end", x1)]:
                    canv.create_line(x, 0, x, h, fill='#00ff88', width=2, tags=tag)
                    canv.create_rectangle(x - 8, 0, x + 8, h, fill='', outline='', tags=tag)
                canv.tag_raise("line_start", "all")
                canv.tag_raise("line_end", "all")
            
            def x_to_time(x):
                w = canv.winfo_width()
                return max(0, min(total_dur, (x / w) * total_dur)) if w > 0 else 0
            
            def on_click(e):
                x = max(0, min(canv.winfo_width(), e.x))
                mx = canv.winfo_width() / 2
                if x < mx:
                    t0 = x_to_time(x)
                    if t0 + var_trim_end.get() < total_dur - 0.02:
                        var_trim_start.set(t0)
                else:
                    t1 = total_dur - x_to_time(x)
                    if var_trim_start.get() + t1 < total_dur - 0.02 and t1 >= 0:
                        var_trim_end.set(t1)
                update_ui()
            
            def on_drag(e):
                if drag_line[0] is None:
                    return
                x = max(0, min(canv.winfo_width(), e.x))
                if drag_line[0] == "start":
                    t0 = x_to_time(x)
                    if t0 + var_trim_end.get() < total_dur - 0.02:
                        var_trim_start.set(t0)
                else:
                    t1 = total_dur - x_to_time(x)
                    if var_trim_start.get() + t1 < total_dur - 0.02 and t1 >= 0:
                        var_trim_end.set(t1)
                update_ui()
            
            def on_release(e):
                drag_line[0] = None
            
            def nearest_line(x):
                w = canv.winfo_width()
                x0 = (var_trim_start.get() / total_dur) * w if total_dur > 0 else 0
                x1 = w - (var_trim_end.get() / total_dur) * w if total_dur > 0 else w
                if abs(x - x0) < abs(x - x1):
                    return "start"
                return "end"
            
            def on_press(e):
                x = e.x
                drag_line[0] = nearest_line(x)
            
            canv.bind("<Button-1>", lambda e: (on_press(e), on_click(e)))
            canv.bind("<B1-Motion>", on_drag)
            canv.bind("<ButtonRelease-1>", on_release)
            
            def update_ui():
                t0 = var_trim_start.get()
                t1 = var_trim_end.get()
                after_dur = max(0.01, total_dur - t0 - t1)
                lbl_summary.config(text=f"去掉 开头 {t0:.2f}s + 结尾 {t1:.2f}s  →  保留 {after_dur:.2f}s", fg='#00ff88')
                need_speed = after_dur > expected_dur and expected_dur > 0
                chk_speed.config(state=tk.NORMAL if need_speed else tk.DISABLED)
                if not need_speed:
                    var_speed.set(False)
                draw_all()
            
            draw_all()
            canv.bind("<Configure>", lambda e: draw_all())
            
            lbl_summary = tk.Label(main_f, text=f"去掉 开头 0.00s + 结尾 0.00s  →  保留 {total_dur:.2f}s",
                font=('Microsoft YaHei', 11, 'bold'), bg='#1a1a1a', fg='#00ff88')
            lbl_summary.pack(anchor=tk.W, pady=6)
            
            ctrl_f = tk.Frame(main_f, bg='#1a1a1a')
            ctrl_f.pack(fill=tk.X, pady=4)
            tk.Label(ctrl_f, text="开头去掉:", font=('Microsoft YaHei', 9), bg='#1a1a1a', fg='#cccccc', width=8).pack(side=tk.LEFT)
            ent_start = tk.Entry(ctrl_f, textvariable=var_trim_start, width=6, font=('Arial', 10))
            ent_start.pack(side=tk.LEFT, padx=2)
            ent_start.bind("<Return>", lambda e: (_parse_entry(ent_start, var_trim_start, 0, total_dur), update_ui()))
            ent_start.bind("<FocusOut>", lambda e: (_parse_entry(ent_start, var_trim_start, 0, total_dur), update_ui()))
            tk.Label(ctrl_f, text="s    结尾去掉:", font=('Microsoft YaHei', 9), bg='#1a1a1a', fg='#cccccc').pack(side=tk.LEFT, padx=(10, 0))
            ent_end = tk.Entry(ctrl_f, textvariable=var_trim_end, width=6, font=('Arial', 10))
            ent_end.pack(side=tk.LEFT, padx=2)
            ent_end.bind("<Return>", lambda e: (_parse_entry(ent_end, var_trim_end, 0, total_dur), update_ui()))
            ent_end.bind("<FocusOut>", lambda e: (_parse_entry(ent_end, var_trim_end, 0, total_dur), update_ui()))
            tk.Label(ctrl_f, text="s  （也可直接输入数值）", font=('Microsoft YaHei', 9), bg='#1a1a1a', fg='#666666').pack(side=tk.LEFT)
            
            def _parse_entry(ent, var, lo, hi):
                try:
                    v = float(ent.get())
                    var.set(max(lo, min(hi, v)))
                except (ValueError, TypeError):
                    pass
            
            var_speed = tk.BooleanVar(value=True)
            need_speed_init = total_dur > expected_dur and expected_dur > 0
            chk_speed = tk.Checkbutton(main_f, text="若仍超时则自动加速到预期时长", variable=var_speed,
                font=('Microsoft YaHei', 10), bg='#1a1a1a', fg='#cccccc', selectcolor='#333333',
                activebackground='#1a1a1a', activeforeground='#cccccc',
                state=tk.NORMAL if need_speed_init else tk.DISABLED)
            chk_speed.pack(anchor=tk.W, pady=8)
            
            def compute_trimmed_audio():
                """按当前裁剪/加速设置计算输出音频（float32），失败返回 None"""
                try:
                    t0 = max(0, var_trim_start.get())
                    t1 = max(0, var_trim_end.get())
                    trim_start = int(t0 * rate)
                    trim_end = int(t1 * rate)
                    if trim_start + trim_end >= len(audio_int16):
                        return None
                    trimmed = audio_int16[trim_start:len(audio_int16) - trim_end].astype(np.float32) / 32767.0
                    dur_after = len(trimmed) / rate
                    do_speedup = var_speed.get() and expected_dur > 0 and dur_after > expected_dur
                    if do_speedup:
                        target_samples = int(expected_dur * rate)
                        if target_samples < 1:
                            target_samples = 1
                        x_old = np.linspace(0, 1, len(trimmed))
                        x_new = np.linspace(0, 1, target_samples)
                        out_float = np.interp(x_new, x_old, trimmed)
                    else:
                        out_float = trimmed
                    return np.clip(out_float, -1, 1).astype(np.float32)
                except Exception:
                    return None
            
            def do_preview():
                audio = compute_trimmed_audio()
                if audio is None:
                    messagebox.showwarning("提示", "裁剪参数无效或裁剪过多")
                    return
                try:
                    sd.play(audio, rate, blocking=False)
                except Exception as ex:
                    messagebox.showerror("错误", f"试听失败: {str(ex)}")
            
            def do_apply():
                out_float = compute_trimmed_audio()
                if out_float is None:
                    messagebox.showwarning("提示", "裁剪过多，至少保留 0.01 秒")
                    return
                out_int16 = (out_float * 32767).astype(np.int16)
                with wave.open(file_path, 'wb') as wf:
                    wf.setnchannels(1)
                    wf.setsampwidth(2)
                    wf.setframerate(rate)
                    wf.writeframes(out_int16.tobytes())
                self.invalidate_score_cache(index)
                if index in self.audio_cache:
                    del self.audio_cache[index]
                dlg.destroy()
                self.status_label.config(text=f"✓ 已裁剪第 {index + 1} 句", fg='#00aa00')
                self.root.after(2000, lambda: self.status_label.config(text="就绪", fg='black'))
                # 完整刷新列表，避免单项更新逻辑导致的列表项消失
                self.root.after(50, self.refresh_list)
            
            btn_f = tk.Frame(main_f, bg='#1a1a1a')
            btn_f.pack(fill=tk.X, pady=15)
            tk.Button(btn_f, text="▶ 试听", font=('Microsoft YaHei', 10), bg='#0078d4', fg='white',
                command=do_preview, width=10, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=(0, 5))
            tk.Button(btn_f, text="应用", font=('Microsoft YaHei', 10), bg='#00aa00', fg='white',
                command=do_apply, width=10, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=5)
            tk.Button(btn_f, text="取消", font=('Microsoft YaHei', 10), bg='#555555', fg='white',
                command=dlg.destroy, width=10, relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT)
            
        except Exception as e:
            messagebox.showerror("错误", f"裁剪失败: {str(e)}")

    def _auto_trim_silence_for_segment(self, index, threshold=0.01, min_silence_dur=0.12, margin_dur=0.03):
        """对单句音频自动裁剪首尾静音，返回是否实际发生了裁剪"""
        file_path = os.path.join(self.recordings_dir, f"segment_{index:03d}.wav")
        if not os.path.exists(file_path):
            return False
        
        try:
            with wave.open(file_path, 'rb') as wf:
                nframes = wf.getnframes()
                rate = wf.getframerate()
                audio_int16 = np.frombuffer(wf.readframes(nframes), dtype=np.int16)
            
            if nframes <= 0:
                return False
            
            audio_float = audio_int16.astype(np.float32) / 32767.0
            abs_audio = np.abs(audio_float)
            min_silence_frames = int(min_silence_dur * rate)
            margin_frames = int(margin_dur * rate)
            
            # 找前面连续静音区末尾
            start_idx = 0
            while start_idx < nframes and abs_audio[start_idx] < threshold:
                start_idx += 1
            # 找后面连续静音区起点
            end_idx = nframes - 1
            while end_idx > start_idx and abs_audio[end_idx] < threshold:
                end_idx -= 1
            
            # 加一点余量，避免剪得太紧
            start_idx = max(0, start_idx - margin_frames)
            end_idx = min(nframes - 1, end_idx + margin_frames)
            
            # 若有效内容太少或没有静音可剪，则跳过
            if end_idx <= start_idx or start_idx < min_silence_frames and nframes - 1 - end_idx < min_silence_frames:
                return False
            
            trimmed = audio_float[start_idx:end_idx + 1]
            if len(trimmed) <= 0 or len(trimmed) == len(audio_float):
                return False
            
            out_int16 = (np.clip(trimmed, -1, 1) * 32767).astype(np.int16)
            with wave.open(file_path, 'wb') as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(rate)
                wf.writeframes(out_int16.tobytes())
            
            # 更新缓存和评分
            self.invalidate_score_cache(index)
            if index in self.audio_cache:
                del self.audio_cache[index]
            return True
        except Exception as e:
            print(f"自动裁剪静音失败 index={index}: {e}")
            return False
    
    def batch_auto_trim_silence(self):
        """一键裁剪所有已录音频首尾静音"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        
        if not self.recording_states:
            messagebox.showinfo("提示", "还没有任何已录制的音频，无需裁剪")
            return
        
        total_recorded = sum(1 for i in range(len(self.segments)) if self.recording_states.get(i))
        if total_recorded == 0:
            messagebox.showinfo("提示", "还没有任何已录制的音频，无需裁剪")
            return
        
        if not messagebox.askyesno(
            "确认一键裁剪",
            f"将对所有已录制的 {total_recorded} 句音频自动裁剪首尾静音。\n\n"
            "仅根据波形判断前后平坦区域，可能会略微剪到讲话的吸气/尾音。\n"
            "该操作会直接覆盖原有音频文件，且不可撤销。\n\n是否继续？"
        ):
            return
        
        changed = 0
        for i in range(len(self.segments)):
            if self.recording_states.get(i):
                if self._auto_trim_silence_for_segment(i):
                    changed += 1
        
        # 全局刷新一次列表和显示
        self.invalidate_score_cache()
        self.refresh_list()
        self.update_display()
        
        messagebox.showinfo("完成", f"已自动裁剪 {changed} 段音频的首尾静音。\n(总已录制: {total_recorded})")
    
    def show_statistics(self):
        """显示录制统计信息"""
        if not self.segments:
            messagebox.showinfo("统计信息", "还没有加载文件")
            return
        
        total = len(self.segments)
        recorded = sum(1 for v in self.recording_states.values() if v)
        unrecorded = total - recorded
        
        # 计算平均分
        scores = [self.calculate_score(i) for i in range(total) if self.recording_states.get(i)]
        avg_score = np.mean(scores) if scores else 0
        
        # 分类统计
        excellent = sum(1 for s in scores if s >= 80)
        good = sum(1 for s in scores if 60 <= s < 80)
        poor = sum(1 for s in scores if s < 60)
        
        # 计算总时长
        total_duration = sum(seg['duration'] for seg in self.segments)
        recorded_duration = sum(seg['duration'] for i, seg in enumerate(self.segments) if self.recording_states.get(i))
        
        msg = f"""
📊 录制统计信息
{'='*40}

📝 进度统计:
   总句子数: {total}
   已录制: {recorded}
   未录制: {unrecorded}
   完成度: {recorded/total*100:.1f}%

⏱️ 时长统计:
   总时长: {total_duration:.1f}秒 ({total_duration/60:.1f}分钟)
   已录时长: {recorded_duration:.1f}秒 ({recorded_duration/60:.1f}分钟)

⭐ 质量统计:
   平均评分: {avg_score:.1f}
   优秀(≥80分): {excellent} 句
   良好(60-79分): {good} 句
   需重录(<60分): {poor} 句

{'='*40}
提示: 按Delete键可删除当前录音
      按F1键查看所有快捷键
        """
        
        messagebox.showinfo("统计信息", msg)
    
    def mark_low_scores_for_rerecord(self, threshold=60):
        """标记低分句子并跳转"""
        if not self.segments:
            messagebox.showwarning("提示", "请先加载转录文件")
            return
        
        if not self.recording_states:
            messagebox.showinfo("提示", "还没有录制任何音频")
            return
        
        # 找出所有低分句子
        low_scores = []
        for i in range(len(self.segments)):
            if self.recording_states.get(i):
                score = self.calculate_score(i)
                if score is not None and score < threshold:
                    low_scores.append((i, score))
        
        if not low_scores:
            messagebox.showinfo("太棒了！", f"没有低于{threshold}分的句子\n所有录音质量都很好！")
            return
        
        # 按分数排序（最低的在前）
        low_scores.sort(key=lambda x: x[1])
        
        # 构建提示消息
        msg = f"发现 {len(low_scores)} 个低分句子:\n\n"
        for idx, (i, score) in enumerate(low_scores[:5]):
            msg += f"   {idx+1}. 第 {i+1} 句 - 评分: {score:.0f}\n"
        
        if len(low_scores) > 5:
            msg += f"   ...还有 {len(low_scores)-5} 个\n"
        
        msg += f"\n是否跳转到第一个低分句子（第 {low_scores[0][0]+1} 句）？"
        
        if messagebox.askyesno("批量重录", msg):
            self.jump_to_segment(low_scores[0][0])
            self.status_label.config(text=f"准备重录第 {low_scores[0][0]+1} 句", fg="orange")
    
    def show_help_dialog(self):
        """显示帮助对话框"""
        help_text = """
⌨️ 快捷键列表
{'='*50}

播放控制:
   空格键      - 播放/暂停当前句子
   ← →        - 上一句/下一句
   Home       - 跳到第一句
   End        - 跳到最后一句
   S          - 切换播放速度 (0.5x ~ 1.5x)

录制控制:
   R          - 开始/停止录制
   Delete     - 删除当前录音
   P          - 预览当前录音

AI辅助:
   D          - AI演示朗读

导出:
   Enter      - 导出完整音频

工具:
   F1         - 显示帮助（本窗口）
   F5         - 显示统计信息

{'='*50}

💡 使用技巧:

1. 录制前可先按D听AI演示，把握节奏
2. 录制时注意音量条颜色：
   🔴 红色 = 太小声  🟢 绿色 = 正常  🟠 橙色 = 太大声
3. 点击左侧列表📊按钮可查看音频波形
4. 评分低于60分建议重录
5. 使用"低分重录"按钮快速找到需要改进的句子

{'='*50}
        """
        
        # 创建帮助窗口
        help_window = tk.Toplevel(self.root)
        help_window.title("帮助 - 快捷键与使用技巧")
        help_window.geometry("600x700")
        help_window.configure(bg='#1a1a1a')
        
        # 文本框
        text_frame = tk.Frame(help_window, bg='#1a1a1a')
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        text_widget = tk.Text(
            text_frame,
            font=('Consolas', 10),
            bg='#0a0a0a',
            fg='#e0e0e0',
            wrap=tk.WORD,
            padx=15,
            pady=15
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert('1.0', help_text)
        text_widget.config(state=tk.DISABLED)
        
        # 关闭按钮
        close_btn = tk.Button(
            help_window,
            text="关闭",
            font=('Microsoft YaHei', 11),
            bg='#444444',
            fg='white',
            command=help_window.destroy,
            width=15,
            relief=tk.FLAT,
            cursor='hand2'
        )
        close_btn.pack(pady=(0, 20))
    
    def get_cached_audio(self, index):
        """获取缓存的音频，如果没有则加载"""
        if index not in self.audio_cache:
            file_path = os.path.join(self.recordings_dir, f"segment_{index:03d}.wav")
            if os.path.exists(file_path):
                try:
                    with wave.open(file_path, 'rb') as wf:
                        audio_data = np.frombuffer(wf.readframes(wf.getnframes()), dtype=np.int16)
                        self.audio_cache[index] = audio_data.astype(np.float32) / 32767.0
                    
                    # 限制缓存大小
                    if len(self.audio_cache) > self.cache_size_limit:
                        oldest = min(self.audio_cache.keys())
                        del self.audio_cache[oldest]
                except Exception as e:
                    print(f"音频缓存加载错误: {e}")
                    return None
        
        return self.audio_cache.get(index)
    
    def update_volume_meter(self):
        """更新音量监控条"""
        if not self.is_recording:
            return
        
        try:
            # 计算当前音量（RMS）
            if self.recording_data:
                recent_data = self.recording_data[-1] if self.recording_data else np.array([[0]])
                rms = np.sqrt(np.mean(recent_data**2))
                self.current_volume = min(1.0, rms * 10)  # 放大并限制在0-1
            
            # 更新音量条
            width = int(self.current_volume * 200)
            self.volume_meter.coords(self.volume_bar, 0, 0, width, 20)
            
            # 根据音量调整颜色
            if self.current_volume < 0.1:
                color = '#ff0000'  # 太小声-红色
            elif self.current_volume < 0.7:
                color = '#00ff00'  # 正常-绿色
            else:
                color = '#ff8800'  # 太大声-橙色
            
            self.volume_meter.itemconfig(self.volume_bar, fill=color)
            
        except Exception as e:
            print(f"音量监控错误: {e}")
        
        # 30fps更新
        self.volume_update_id = self.root.after(33, self.update_volume_meter)
    
    def cycle_speed(self):
        """循环切换播放速度"""
        self.current_speed_index = (self.current_speed_index + 1) % len(self.speed_options)
        self.playback_speed = self.speed_options[self.current_speed_index]
        self.speed_btn.config(text=f"{self.playback_speed}x")
        
        # 如果正在播放，重新计算进度
        if self.is_playing:
            self.segment_start_time = time.time() - (self.word_progress * self.segments[self.current_index]['duration'] / self.playback_speed)
    
    def on_closing(self):
        """退出保护-检查是否有未导出的录音"""
        if self.recording_states:
            # 有录音
            recorded_count = sum(1 for v in self.recording_states.values() if v)
            if recorded_count > 0:
                result = messagebox.askyesnocancel(
                    "退出确认",
                    f"你已录制了 {recorded_count} 个片段\n\n确定要退出吗？\n（未导出的录音将保留在recordings文件夹）"
                )
                if result is None:  # 取消
                    return
                elif not result:  # 否
                    return
        
        # 清理资源
        if self.is_recording:
            self.stop_recording()
        if self.is_playing:
            self.toggle_play()
        
        self.root.destroy()


def main():
    # 尝试使用TkinterDnD（支持拖拽），如果不可用则使用普通Tk
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
        print("提示: 安装 tkinterdnd2 可启用拖拽功能 (pip install tkinterdnd2)")
    
    app = TeleprompterApp(root)
    
    # 尝试自动加载当前目录的文件
    default_file = "RA2_English_Colloquial.txt"
    if os.path.exists(default_file):
        app.segments = parse_transcript_file(default_file)
        app.current_index = 0
        app.status_label.config(text=f"已加载 {len(app.segments)} 个句子")
        app.update_display()
        app.refresh_list()
    
    root.mainloop()


if __name__ == "__main__":
    main()
