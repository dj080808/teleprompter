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
        
        # 录音相关
        self.recordings_dir = "recordings"
        self.recording_data = []
        self.recording_states = {}  # {index: True/False} 记录是否已录制
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
            self.segments = parse_transcript_file(file_path)
            self.current_index = 0
            self.recording_states = {}
            self.word_progress = 0.0
            self.invalidate_score_cache()  # 新文件加载，清空评分缓存
            
            self.status_label.config(
                text=f"已加载 {len(self.segments)} 个句子",
                fg='#aaaaaa'
            )
            self.update_display()
            self.refresh_list()
            
            messagebox.showinfo("成功", f"已加载 {len(self.segments)} 个句子")
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败: {str(e)}")
    
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
            
            self.stream = sd.InputStream(
                samplerate=self.sample_rate,
                channels=1,
                callback=callback
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
        
        score_label, waveform_btn = None, None
        if score is not None:
            score_color = '#00ff00' if score >= 80 else ('#ffaa00' if score >= 60 else '#ff3333')
            score_text = f"⭐ {score}" if score >= 80 else (f"★ {score}" if score >= 60 else f"⚠ {score}")
            score_label = tk.Label(top_row, text=score_text, font=('Arial', 10, 'bold'), bg=bg_color, fg=score_color)
            score_label.pack(side=tk.RIGHT)
            score_label.bind("<Button-1>", lambda e, idx=i: self.jump_to_segment(idx))
            waveform_btn = tk.Button(top_row, text="📊", font=('Arial', 10), bg=bg_color, fg='#00aaff',
                command=lambda idx=i: self.show_waveform(idx), relief=tk.FLAT, cursor='hand2', padx=5)
            waveform_btn.pack(side=tk.RIGHT, padx=(0, 5))
            ToolTip(waveform_btn, "查看音频波形")
        
        text_display = seg['text'][:40] + "..." if len(seg['text']) > 40 else seg['text']
        text_label = tk.Label(item_frame, text=text_display, font=('Arial', 9), bg=bg_color, fg='#cccccc',
            wraplength=320, justify=tk.LEFT, anchor=tk.W)
        text_label.pack(fill=tk.X, padx=5, pady=(0, 5))
        text_label.bind("<Button-1>", lambda e, idx=i: self.jump_to_segment(idx))
        
        info = {
            "frame": item_frame, "top_row": top_row, "num_label": num_label,
            "status_label": status_label, "text_label": text_label,
            "score_label": score_label, "waveform_btn": waveform_btn,
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
                info["status_label"], info["text_label"],
                info.get("score_label"), info.get("waveform_btn")
            ], bg_color)
    
    def invalidate_score_cache(self, index=None):
        """使评分缓存失效（index 为 None 时清空全部）"""
        if index is None:
            self._score_cache.clear()
        else:
            self._score_cache.pop(index, None)
    
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
    
    def merge_audio_segments(self, output_path):
        """合并所有录音片段，插入静音间隔"""
        with wave.open(output_path, 'wb') as output_wav:
            output_wav.setnchannels(1)
            output_wav.setsampwidth(2)
            output_wav.setframerate(self.sample_rate)
            
            last_end_time = 0
            
            for i, seg in enumerate(self.segments):
                # 插入静音（如果需要）
                gap = seg['start_time'] - last_end_time
                if gap > 0:
                    silence_frames = int(gap * self.sample_rate)
                    silence = np.zeros(silence_frames, dtype=np.int16)
                    output_wav.writeframes(silence.tobytes())
                
                # 添加录音片段
                file_path = os.path.join(self.recordings_dir, f"segment_{i:03d}.wav")
                
                if os.path.exists(file_path):
                    with wave.open(file_path, 'rb') as segment_wav:
                        frames = segment_wav.readframes(segment_wav.getnframes())
                        output_wav.writeframes(frames)
                    last_end_time = seg['end_time']
                else:
                    # 如果没有录音，插入静音
                    duration_frames = int(seg['duration'] * self.sample_rate)
                    silence = np.zeros(duration_frames, dtype=np.int16)
                    output_wav.writeframes(silence.tobytes())
                    last_end_time = seg['end_time']

    
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
