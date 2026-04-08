# TibiaFrames v1.2.4 - Tibia Screenshot Viewer and Organizer
# Save this file as: tibiaframes_v1.2.4.pyw

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import os
import shutil
import re
from datetime import datetime
import subprocess
import json
import time
import io
from collections import defaultdict, Counter

try:
    import win32clipboard
    _HAS_WIN32CLIPBOARD = True
except ImportError:
    _HAS_WIN32CLIPBOARD = False

class TibiaFrames:
    VERSION = "v1.2.4"
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"TibiaFrames - ({self.VERSION})")
        self.root.geometry("1600x900")
        
        # Hide console window
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleTitleW("TibiaFrames")
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd != 0:
                ctypes.windll.user32.ShowWindow(hwnd, 0)
        except:
            pass
        
        # Initialize variables
        self.screenshots_dir = ""
        self.screenshot_data = {}
        self.current_image = None
        self.current_image_path = None
        self.zoom_factor = 1.0
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.is_dragging = False
        
        # Cache
        self.image_cache = {}
        self.max_cache_size = 50
        self.resize_timer = None
        
        # Statistics panel
        self.stats_visible = False
        self.stats_frame = None
        
        # Compiled regex pattern
        self._filename_pattern = re.compile(
            r"(\d{4}-\d{2}-\d{2})_(\d{9})_([^_]+)_(.+)\.(png|jpg|jpeg)", 
            re.IGNORECASE
        )
        
        # File scanning cache
        self.file_list_cache = None
        self.last_scan_time = 0
        self.scan_interval = 2.0
        
        # UI variables
        self.direct_snip = tk.BooleanVar(value=True)
        self.selected_date = tk.StringVar()
        
        # Settings file path
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.settings_file = os.path.join(script_dir, "tibiaframes_settings.json")
        
        # Categories mapping
        self.categories = {
            "Achievement": "Achievement", 
            "BestiaryEntryCompleted": "Bestiary Entry Completed",
            "BestiaryEntryUnlocked": "Bestiary Entry Unlocked", 
            "BossDefeated": "Boss Defeated",
            "DeathPvE": "Death PvE", 
            "DeathPvP": "Death PvP", 
            "GiftofLifeTriggered": "Gift of Life Triggered",
            "HighestDamageDealt": "Highest Damage Dealt", 
            "HighestHealingDone": "Highest Healing Done",
            "Hotkey": "Hotkey", 
            "LevelUp": "Level Up", 
            "LowHealth": "Low Health",
            "PlayerAttacking": "Player Attacking", 
            "PlayerKill": "Player Kill + Assist",
            "PlayerKillAssist": "Player Kill + Assist", 
            "SkillUp": "Skill Up",
            "TreasureFound": "Treasure Found", 
            "ValuableLoot": "Valuable Loot"
        }
        
        # Setup UI and load settings
        self.setup_ui()
        self.load_settings()
        self.apply_theme()
        self.load_default_directory()
    
    def setup_ui(self):
        """Setup the main UI components"""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Control frame
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Directory frame (first row)
        dir_frame = ttk.Frame(control_frame)
        dir_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(dir_frame, text="Directory:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.dir_var = tk.StringVar()
        self.dir_entry = ttk.Entry(dir_frame, textvariable=self.dir_var, width=50)
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        ttk.Button(dir_frame, text="Browse", command=self.browse_directory).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(dir_frame, text="Refresh", command=self.refresh_screenshots).pack(side=tk.LEFT)
        
        # Info frame (right side of directory row)
        info_frame = ttk.Frame(dir_frame)
        info_frame.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.counter_label = ttk.Label(info_frame, text="No screenshots loaded", 
                                      font=('Arial', 10, 'bold'))
        self.counter_label.pack(anchor=tk.E)
        
        self.size_label = ttk.Label(info_frame, text="", font=('Arial', 9), 
                                   foreground='gray')
        self.size_label.pack(anchor=tk.E)
        
        # Stats button
        self.stats_button = ttk.Button(info_frame, text="Stats", command=self.toggle_stats_panel)
        self.stats_button.pack(anchor=tk.E, pady=(2, 0))
        
        # Action buttons frame (second row)
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Sort into Folders", command=self.sort_into_folders).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Snipping Tool", command=self.open_snipping_tool).pack(side=tk.LEFT, padx=(0, 5))
        
        self.direct_snip_check = ttk.Checkbutton(button_frame, text="Direct capture", 
                                                 variable=self.direct_snip, 
                                                 command=self.save_snip_preference)
        self.direct_snip_check.pack(side=tk.LEFT, padx=(10, 0))
        
        # Content frame
        self.content_frame = ttk.Frame(main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left panel
        left_frame = ttk.Frame(self.content_frame, width=350)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, expand=False, padx=(0, 5))
        left_frame.pack_propagate(False)
        
        # Screenshots tree
        ttk.Label(left_frame, text="Screenshots", font=('Arial', 12, 'bold')).pack(pady=(0, 5))
        
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.tree = ttk.Treeview(tree_frame, selectmode='browse')
        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Tree event bindings
        self.tree.bind('<ButtonRelease-1>', self.on_tree_select)
        self.tree.bind('<Button-3>', self.on_tree_right_click)
        self.tree.bind('<KeyRelease-Up>', self.on_tree_key_nav)
        self.tree.bind('<KeyRelease-Down>', self.on_tree_key_nav)
        
        # Categories by date section
        ttk.Label(left_frame, text="Categories by Date", font=('Arial', 12, 'bold')).pack(pady=(10, 5))
        
        date_frame = ttk.Frame(left_frame)
        date_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(date_frame, text="Date:").pack(side=tk.LEFT)
        self.date_combo = ttk.Combobox(date_frame, textvariable=self.selected_date, 
                                      width=20, state="readonly")
        self.date_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.date_combo.bind('<<ComboboxSelected>>', self.on_date_selected)
        
        # Category tree
        category_frame = ttk.Frame(left_frame)
        category_frame.pack(fill=tk.BOTH, expand=True)
        
        self.category_tree = ttk.Treeview(category_frame, selectmode='browse')
        category_scroll = ttk.Scrollbar(category_frame, orient=tk.VERTICAL, 
                                       command=self.category_tree.yview)
        self.category_tree.configure(yscrollcommand=category_scroll.set)
        
        self.category_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        category_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Category tree event bindings
        self.category_tree.bind('<ButtonRelease-1>', self.on_category_select)
        self.category_tree.bind('<Button-3>', self.on_category_right_click)
        self.category_tree.bind('<KeyRelease-Up>', self.on_category_key_nav)
        self.category_tree.bind('<KeyRelease-Down>', self.on_category_key_nav)
        
        # Right panel container (image viewer + stats)
        self.right_container = ttk.Frame(self.content_frame)
        self.right_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Image viewer frame
        self.image_frame = ttk.Frame(self.right_container)
        self.image_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.setup_image_viewer()
        
        # Context menus
        self.tree_context_menu = tk.Menu(self.root, tearoff=0)
        self.tree_context_menu.add_command(label="Copy Screenshot", command=self.copy_screenshot)
        self.tree_context_menu.add_command(label="Open Containing Folder", command=self.open_containing_folder)
        
        self.canvas_context_menu = tk.Menu(self.root, tearoff=0)
        self.canvas_context_menu.add_command(label="Copy Screenshot", command=self.copy_screenshot)
        self.canvas_context_menu.add_command(label="Open Containing Folder", command=self.open_containing_folder)
        
        self.root.focus_set()
    
    def setup_image_viewer(self):
        """Setup the image viewer components"""
        self.info_label = ttk.Label(self.image_frame, text="Select a screenshot to view", 
                                   font=('Arial', 12, 'bold'), anchor=tk.CENTER)
        self.info_label.pack(pady=(0, 5))
        
        # Image canvas frame
        canvas_frame = ttk.Frame(self.image_frame, relief=tk.SUNKEN, borderwidth=2)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg='white')
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
        
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Zoom controls
        zoom_frame = ttk.Frame(self.image_frame)
        zoom_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(zoom_frame, text="Zoom In (+)", command=self.zoom_in).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(zoom_frame, text="Zoom Out (-)", command=self.zoom_out).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(zoom_frame, text="Reset Zoom", command=self.reset_zoom).pack(side=tk.LEFT, padx=(0, 5))
        
        self.zoom_label = ttk.Label(zoom_frame, text="100%")
        self.zoom_label.pack(side=tk.LEFT, padx=(10, 0))
        
        # Canvas event bindings
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Button-4>", self.on_mouse_wheel)
        self.canvas.bind("<Button-5>", self.on_mouse_wheel)
        self.canvas.bind("<Button-1>", self.on_drag_start)
        self.canvas.bind("<B1-Motion>", self.on_drag_motion)
        self.canvas.bind("<ButtonRelease-1>", self.on_drag_end)
        self.canvas.bind("<Button-3>", self.on_canvas_right_click)
        
        # Window resize binding
        self.root.bind('<Configure>', self.on_window_resize)
        
        # Keyboard shortcuts
        self.root.bind('<Control-plus>', lambda e: self.zoom_in())
        self.root.bind('<Control-minus>', lambda e: self.zoom_out())
        self.root.bind('<Control-0>', lambda e: self.reset_zoom())
    
    def setup_stats_panel(self):
        """Setup the statistics panel"""
        if self.stats_frame:
            return
            
        self.stats_frame = ttk.Frame(self.right_container, width=300)
        self.stats_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        self.stats_frame.pack_propagate(False)
        
        # Stats header
        header_frame = ttk.Frame(self.stats_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(header_frame, text="Statistics", font=('Arial', 14, 'bold')).pack(side=tk.LEFT)
        ttk.Button(header_frame, text="✕", width=3, command=self.hide_stats_panel).pack(side=tk.RIGHT)
        
        # Stats notebook
        self.stats_notebook = ttk.Notebook(self.stats_frame)
        self.stats_notebook.pack(fill=tk.BOTH, expand=True)
        
        # Overview tab
        overview_frame = ttk.Frame(self.stats_notebook)
        self.stats_notebook.add(overview_frame, text="Summary")
        
        # Characters tab
        characters_frame = ttk.Frame(self.stats_notebook)
        self.stats_notebook.add(characters_frame, text="Chars")
        
        # Categories tab
        categories_frame = ttk.Frame(self.stats_notebook)
        self.stats_notebook.add(categories_frame, text="Types")
        
        # Character Details tab
        char_details_frame = ttk.Frame(self.stats_notebook)
        self.stats_notebook.add(char_details_frame, text="Details")
        
        # Activity tab
        activity_frame = ttk.Frame(self.stats_notebook)
        self.stats_notebook.add(activity_frame, text="Time")
        
        # Store tab frames for updates
        self.overview_frame = overview_frame
        self.characters_frame = characters_frame
        self.categories_frame = categories_frame
        self.char_details_frame = char_details_frame
        self.activity_frame = activity_frame
    
    def toggle_stats_panel(self):
        """Toggle the statistics panel visibility"""
        if self.stats_visible:
            self.hide_stats_panel()
        else:
            self.show_stats_panel()
    
    def show_stats_panel(self):
        """Show the statistics panel"""
        if not self.screenshot_data:
            messagebox.showwarning("Warning", "Please load screenshots first.")
            return
            
        self.setup_stats_panel()
        self.update_statistics()
        self.stats_visible = True
        self.stats_button.config(text="Hide Stats")
    
    def hide_stats_panel(self):
        """Hide the statistics panel"""
        if self.stats_frame:
            self.stats_frame.destroy()
            self.stats_frame = None
        self.stats_visible = False
        self.stats_button.config(text="Stats")
    
    def update_statistics(self):
        """Update all statistics in the panel"""
        if not self.stats_frame:
            return
            
        self.update_overview_stats()
        self.update_character_stats()
        self.update_category_stats()
        self.update_character_details_stats()
        self.update_activity_stats()
    
    def update_overview_stats(self):
        """Update overview statistics"""
        # Clear existing widgets
        for widget in self.overview_frame.winfo_children():
            widget.destroy()
        
        # Use the SAME counting method as the main window
        # Count all files that successfully parse, not data structure entries
        screenshot_files = self.get_screenshot_files()
        total_screenshots = 0
        characters = set()
        categories = set()
        dates = set()
        total_size = 0
        
        for file_path in screenshot_files:
            try:
                file_size = os.path.getsize(file_path)
                total_size += file_size
            except:
                pass
            
            parsed = self.parse_screenshot_filename(os.path.basename(file_path))
            if parsed:
                total_screenshots += 1
                characters.add(parsed['character'])
                categories.add(parsed['category'])
                dates.add(parsed['date'])
        
        # Display overview stats
        stats_text = tk.Text(self.overview_frame, wrap=tk.WORD, height=15, width=35)
        stats_scrollbar = ttk.Scrollbar(self.overview_frame, orient=tk.VERTICAL, command=stats_text.yview)
        stats_text.configure(yscrollcommand=stats_scrollbar.set)
        
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        stats_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        overview_data = f"""📊 OVERVIEW STATISTICS

📷 Total Screenshots: {total_screenshots:,}
👤 Characters: {len(characters)}
📂 Categories: {len(categories)}
📅 Days Captured: {len(dates)}
💾 Total Size: {self.format_file_size(total_size)}

🎯 Average per Character: {total_screenshots // max(len(characters), 1):,}
📈 Average per Day: {total_screenshots // max(len(dates), 1):,}

📁 Most Common Categories:"""
        
        # Get category counts
        category_counts = defaultdict(int)
        for character in self.screenshot_data:
            for category in self.screenshot_data[character]:
                for date in self.screenshot_data[character][category]:
                    times_data = self.screenshot_data[character][category][date]['times']
                    category_counts[category] += len(times_data)
        
        # Top 5 categories
        top_categories = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)[:5]
        for i, (category, count) in enumerate(top_categories, 1):
            overview_data += f"\n{i}. {category}: {count:,}"
        
        if dates:
            # Date range
            date_objects = []
            for character in self.screenshot_data:
                for category in self.screenshot_data[character]:
                    for date in self.screenshot_data[character][category]:
                        sort_date = self.screenshot_data[character][category][date]['sort_date']
                        if isinstance(sort_date, datetime):
                            date_objects.append(sort_date)
            
            if date_objects:
                first_date = min(date_objects).strftime("%d %B %Y")
                last_date = max(date_objects).strftime("%d %B %Y")
                total_days = (max(date_objects) - min(date_objects)).days + 1
                
                overview_data += f"\n\n📅 Date Range:\nFirst: {first_date}\nLast: {last_date}\nSpan: {total_days} days"
        
        stats_text.insert(tk.END, overview_data)
        stats_text.config(state=tk.DISABLED)
    
    def update_character_stats(self):
        """Update character statistics"""
        # Clear existing widgets
        for widget in self.characters_frame.winfo_children():
            widget.destroy()
        
        # Calculate character stats
        character_stats = {}
        for character in self.screenshot_data:
            total_shots = 0
            categories = set()
            for category in self.screenshot_data[character]:
                categories.add(category)
                for date in self.screenshot_data[character][category]:
                    times_data = self.screenshot_data[character][category][date]['times']
                    total_shots += len(times_data)
            
            character_stats[character] = {
                'total': total_shots,
                'categories': len(categories)
            }
        
        # Display character stats
        stats_text = tk.Text(self.characters_frame, wrap=tk.WORD, height=15, width=35)
        stats_scrollbar = ttk.Scrollbar(self.characters_frame, orient=tk.VERTICAL, command=stats_text.yview)
        stats_text.configure(yscrollcommand=stats_scrollbar.set)
        
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        stats_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        character_data = "👤 CHARACTER STATISTICS\n\n"
        
        # Sort characters by total screenshots
        sorted_chars = sorted(character_stats.items(), key=lambda x: x[1]['total'], reverse=True)
        
        for i, (character, stats) in enumerate(sorted_chars, 1):
            character_data += f"{i}. {character}\n"
            character_data += f"   📷 Screenshots: {stats['total']:,}\n"
            character_data += f"   📂 Categories: {stats['categories']}\n\n"
        
        stats_text.insert(tk.END, character_data)
        stats_text.config(state=tk.DISABLED)
    
    def update_category_stats(self):
        """Update category statistics"""
        # Clear existing widgets
        for widget in self.categories_frame.winfo_children():
            widget.destroy()
        
        # Calculate category stats
        category_counts = defaultdict(int)
        category_characters = defaultdict(set)
        
        for character in self.screenshot_data:
            for category in self.screenshot_data[character]:
                category_characters[category].add(character)
                for date in self.screenshot_data[character][category]:
                    times_data = self.screenshot_data[character][category][date]['times']
                    category_counts[category] += len(times_data)
        
        # Display category stats
        stats_text = tk.Text(self.categories_frame, wrap=tk.WORD, height=15, width=35)
        stats_scrollbar = ttk.Scrollbar(self.categories_frame, orient=tk.VERTICAL, command=stats_text.yview)
        stats_text.configure(yscrollcommand=stats_scrollbar.set)
        
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        stats_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        category_data = "📂 CATEGORY STATISTICS\n\n"
        
        # Sort categories by count
        sorted_categories = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)
        
        for i, (category, count) in enumerate(sorted_categories, 1):
            character_count = len(category_characters[category])
            percentage = (count / sum(category_counts.values())) * 100
            
            category_data += f"{i}. {category}\n"
            category_data += f"   📷 Count: {count:,} ({percentage:.1f}%)\n"
            category_data += f"   👤 Characters: {character_count}\n\n"
        
        stats_text.insert(tk.END, category_data)
        stats_text.config(state=tk.DISABLED)
    
    def update_character_details_stats(self):
        """Update character-specific detailed statistics"""
        # Clear existing widgets
        for widget in self.char_details_frame.winfo_children():
            widget.destroy()
        
        # Character selection frame
        selection_frame = ttk.Frame(self.char_details_frame)
        selection_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(selection_frame, text="Character:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT)
        
        # Character dropdown
        characters = sorted(self.screenshot_data.keys()) if self.screenshot_data else []
        self.char_details_var = tk.StringVar()
        self.char_details_combo = ttk.Combobox(selection_frame, textvariable=self.char_details_var,
                                              values=characters, state="readonly", width=20)
        self.char_details_combo.pack(side=tk.LEFT, padx=(5, 0))
        self.char_details_combo.bind('<<ComboboxSelected>>', self.on_character_details_selected)
        
        # Details display frame
        self.char_details_display = ttk.Frame(self.char_details_frame)
        self.char_details_display.pack(fill=tk.BOTH, expand=True)
        
        # Set default selection
        if characters:
            self.char_details_var.set(characters[0])
            self.update_character_details_display(characters[0])
    
    def on_character_details_selected(self, event):
        """Handle character selection in details tab"""
        selected_char = self.char_details_var.get()
        if selected_char:
            self.update_character_details_display(selected_char)
    
    def update_character_details_display(self, character):
        """Update the character details display for selected character"""
        # Clear existing display
        for widget in self.char_details_display.winfo_children():
            widget.destroy()
        
        if character not in self.screenshot_data:
            return
        
        # Create text widget for character details
        details_text = tk.Text(self.char_details_display, wrap=tk.WORD, height=12, width=35)
        details_scrollbar = ttk.Scrollbar(self.char_details_display, orient=tk.VERTICAL, command=details_text.yview)
        details_text.configure(yscrollcommand=details_scrollbar.set)
        
        details_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Calculate character-specific stats
        char_data = self.screenshot_data[character]
        total_screenshots = 0
        category_counts = {}
        date_counts = {}
        hourly_activity = defaultdict(int)
        first_date = None
        last_date = None
        
        for category in char_data:
            category_count = 0
            for date in char_data[category]:
                date_data = char_data[category][date]
                times_data = date_data['times']
                screenshot_count = len(times_data)
                
                category_count += screenshot_count
                total_screenshots += screenshot_count
                
                # Track dates
                sort_date = date_data['sort_date']
                if isinstance(sort_date, datetime):
                    if first_date is None or sort_date < first_date:
                        first_date = sort_date
                    if last_date is None or sort_date > last_date:
                        last_date = sort_date
                
                # Count dates
                if date not in date_counts:
                    date_counts[date] = 0
                date_counts[date] += screenshot_count
                
                # Hourly activity
                for time in times_data:
                    try:
                        hour = int(time[:2])
                        hourly_activity[hour] += 1
                    except:
                        pass
            
            category_counts[category] = category_count
        
        # Build details text
        details_data = f"👤 {character.upper()}\n\n"
        details_data += f"📷 Total Screenshots: {total_screenshots:,}\n"
        details_data += f"📂 Categories Used: {len(category_counts)}\n"
        details_data += f"📅 Active Days: {len(date_counts)}\n"
        
        if first_date and last_date:
            span_days = (last_date - first_date).days + 1
            details_data += f"📊 Activity Span: {span_days} days\n"
            details_data += f"📈 Avg per Day: {total_screenshots / span_days:.1f}\n"
        
        # Category breakdown
        details_data += f"\n📂 CATEGORY BREAKDOWN:\n"
        sorted_categories = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)
        
        for category, count in sorted_categories:
            percentage = (count / total_screenshots) * 100
            details_data += f"\n• {category}\n"
            details_data += f"  📷 {count:,} shots ({percentage:.1f}%)\n"
        
        # Most active days
        if date_counts:
            details_data += f"\n📅 MOST ACTIVE DAYS:\n"
            top_dates = sorted(date_counts.items(), key=lambda x: x[1], reverse=True)[:5]
            for i, (date, count) in enumerate(top_dates, 1):
                details_data += f"{i}. {date}: {count} screenshots\n"
        
        # Peak activity hours
        if hourly_activity:
            peak_hour = max(hourly_activity.items(), key=lambda x: x[1])
            details_data += f"\n⏰ Peak Hour: {peak_hour[0]:02d}:00 ({peak_hour[1]} screenshots)\n"
            
            # Show hourly distribution (compact version)
            details_data += f"\n⏰ HOURLY ACTIVITY:\n"
            sorted_hours = sorted(hourly_activity.items())
            for hour, count in sorted_hours:
                if count > 0:
                    bar_length = min(10, count // max(1, max(hourly_activity.values()) // 10))
                    bar = "█" * bar_length
                    details_data += f"{hour:02d}h │{bar:<10} {count}\n"
        
        details_text.insert(tk.END, details_data)
        details_text.config(state=tk.DISABLED)
    
    def update_activity_stats(self):
        """Update activity statistics with long-term usage support"""
        # Clear existing widgets
        for widget in self.activity_frame.winfo_children():
            widget.destroy()
        
        # Calculate activity stats
        hourly_activity = defaultdict(int)
        daily_activity = defaultdict(int)
        monthly_activity = defaultdict(int)
        yearly_activity = defaultdict(int)
        all_dates = []
        
        for character in self.screenshot_data:
            for category in self.screenshot_data[character]:
                for date in self.screenshot_data[character][category]:
                    sort_date = self.screenshot_data[character][category][date]['sort_date']
                    times_data = self.screenshot_data[character][category][date]['times']
                    
                    for time in times_data:
                        # Extract hour from time
                        try:
                            hour = int(time[:2])
                            hourly_activity[hour] += 1
                        except:
                            pass
                    
                    # Daily and temporal activity
                    if isinstance(sort_date, datetime):
                        all_dates.append(sort_date)
                        day_name = sort_date.strftime("%A")
                        month_year = sort_date.strftime("%b %Y")
                        year = sort_date.year
                        
                        daily_activity[day_name] += len(times_data)
                        monthly_activity[month_year] += len(times_data)
                        yearly_activity[year] += len(times_data)
        
        # Display activity stats
        stats_text = tk.Text(self.activity_frame, wrap=tk.WORD, height=15, width=35)
        stats_scrollbar = ttk.Scrollbar(self.activity_frame, orient=tk.VERTICAL, command=stats_text.yview)
        stats_text.configure(yscrollcommand=stats_scrollbar.set)
        
        stats_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        stats_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        activity_data = "⏰ ACTIVITY PATTERNS\n\n"
        
        # Peak hours
        if hourly_activity:
            peak_hour = max(hourly_activity.items(), key=lambda x: x[1])
            activity_data += f"🌟 Peak Hour: {peak_hour[0]:02d}:00 ({peak_hour[1]} screenshots)\n\n"
            
            activity_data += "📊 Hourly Distribution:\n"
            sorted_hours = sorted(hourly_activity.items())
            for hour, count in sorted_hours:
                if count > 0:
                    bar = "█" * min(20, count // max(1, max(hourly_activity.values()) // 20))
                    activity_data += f"{hour:02d}:00 │{bar} {count}\n"
        
        # Day of week activity
        if daily_activity:
            activity_data += "\n📅 Day of Week:\n"
            days_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            for day in days_order:
                if day in daily_activity:
                    count = daily_activity[day]
                    bar = "█" * min(15, count // max(1, max(daily_activity.values()) // 15))
                    activity_data += f"{day[:3]} │{bar} {count}\n"
        
        # Smart temporal display based on data span
        if all_dates and yearly_activity:
            earliest_date = min(all_dates)
            latest_date = max(all_dates)
            total_years = latest_date.year - earliest_date.year + 1
            
            if total_years > 3:
                # Show yearly overview for long-term users
                activity_data += "\n📊 Yearly Activity:\n"
                sorted_years = sorted(yearly_activity.items())
                max_yearly = max(yearly_activity.values()) if yearly_activity else 1
                
                for year, count in sorted_years:
                    bar = "█" * min(15, count // max(1, max_yearly // 15))
                    activity_data += f"{year} │{bar} {count:,}\n"
                
                # Recent months (last 12 months)
                activity_data += "\n📆 Recent Activity (Last 12 Months):\n"
                recent_months = []
                current_date = latest_date
                for i in range(12):
                    month_key = current_date.strftime("%b %Y")
                    if month_key in monthly_activity:
                        recent_months.append((month_key, monthly_activity[month_key], current_date))
                    
                    # Go to previous month
                    if current_date.month == 1:
                        current_date = current_date.replace(year=current_date.year - 1, month=12)
                    else:
                        current_date = current_date.replace(month=current_date.month - 1)
                
                recent_months.reverse()  # Show chronologically
                if recent_months:
                    max_recent = max(count for _, count, _ in recent_months)
                    for month, count, _ in recent_months:
                        bar = "█" * min(12, count // max(1, max_recent // 12))
                        activity_data += f"{month} │{bar} {count}\n"
                        
            elif total_years > 1:
                # Show last 18 months for medium-term users
                activity_data += "\n📆 Monthly Activity (Last 18 Months):\n"
                recent_months = []
                current_date = latest_date
                for i in range(18):
                    month_key = current_date.strftime("%b %Y")
                    if month_key in monthly_activity:
                        recent_months.append((month_key, monthly_activity[month_key]))
                    
                    # Go to previous month
                    if current_date.month == 1:
                        current_date = current_date.replace(year=current_date.year - 1, month=12)
                    else:
                        current_date = current_date.replace(month=current_date.month - 1)
                
                recent_months.reverse()  # Show chronologically
                if recent_months:
                    max_monthly = max(count for _, count in recent_months)
                    for month, count in recent_months:
                        bar = "█" * min(15, count // max(1, max_monthly // 15))
                        activity_data += f"{month} │{bar} {count}\n"
            else:
                # Show all months for short-term users
                activity_data += "\n📆 Monthly Activity:\n"
                sorted_months = sorted(monthly_activity.items(), key=lambda x: datetime.strptime(x[0], "%b %Y"))
                for month, count in sorted_months:
                    bar = "█" * min(15, count // max(1, max(monthly_activity.values()) // 15))
                    activity_data += f"{month} │{bar} {count}\n"
            
            # Data span summary
            span_days = (latest_date - earliest_date).days + 1
            activity_data += f"\n📈 SUMMARY:\n"
            activity_data += f"First Screenshot: {earliest_date.strftime('%d %b %Y')}\n"
            activity_data += f"Latest Screenshot: {latest_date.strftime('%d %b %Y')}\n"
            activity_data += f"Total Span: {span_days:,} days ({total_years} years)\n"
            
            total_screenshots = sum(yearly_activity.values())
            activity_data += f"Screenshots/Day: {total_screenshots / span_days:.1f}\n"
        
        stats_text.insert(tk.END, activity_data)
        stats_text.config(state=tk.DISABLED)
    
    def apply_theme(self):
        """Apply dark theme to the application"""
        colors = {
            'bg': '#2b2b2b', 'fg': '#ffffff', 'canvas_bg': '#1e1e1e',
            'entry_bg': '#404040', 'entry_fg': '#ffffff', 'button_bg': '#404040',
            'button_fg': '#ffffff', 'tree_bg': '#2b2b2b', 'tree_fg': '#ffffff',
            'tree_select': '#0078d4', 'frame_bg': '#2b2b2b', 'label_fg': '#ffffff',
            'gray_fg': '#cccccc'
        }
        
        self.root.configure(bg=colors['bg'])
        self.canvas.configure(bg=colors['canvas_bg'])
        
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure all ttk widgets
        style.configure('Treeview', background=colors['tree_bg'], foreground=colors['tree_fg'], 
                       fieldbackground=colors['tree_bg'], borderwidth=0)
        style.configure('Treeview.Heading', background=colors['button_bg'], 
                       foreground=colors['button_fg'])
        style.map('Treeview', background=[('selected', colors['tree_select'])])
        
        style.configure('TFrame', background=colors['frame_bg'])
        style.configure('TLabel', background=colors['frame_bg'], foreground=colors['label_fg'])
        style.configure('TButton', background=colors['button_bg'], foreground=colors['button_fg'])
        style.configure('TEntry', fieldbackground=colors['entry_bg'], foreground=colors['entry_fg'])
        style.configure('TCheckbutton', background=colors['frame_bg'], foreground=colors['label_fg'])
        style.configure('TCombobox', fieldbackground=colors['entry_bg'], 
                       foreground=colors['entry_fg'])
        style.configure('TNotebook', background=colors['frame_bg'])
        style.configure('TNotebook.Tab', background=colors['button_bg'], foreground=colors['button_fg'])
        
        # Combobox styling
        style.map('TCombobox', 
                 fieldbackground=[('readonly', colors['entry_bg'])],
                 foreground=[('readonly', colors['entry_fg'])],
                 selectbackground=[('readonly', colors['tree_select'])],
                 selectforeground=[('readonly', '#ffffff')])
        
        style.map('TButton', background=[('active', '#505050')])
        style.map('TCheckbutton', background=[('active', colors['frame_bg'])])
        style.map('TNotebook.Tab', background=[('selected', colors['tree_select'])])
        
        # Update specific labels
        if hasattr(self, 'size_label'):
            self.size_label.configure(foreground=colors['gray_fg'])
    
    def load_settings(self):
        """Load settings from file"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                    
                    saved_dir = settings.get('screenshots_directory', '')
                    if saved_dir and os.path.exists(saved_dir):
                        self.dir_var.set(saved_dir)
                    
                    self.direct_snip.set(settings.get('direct_snip', True))
                    
                    window_geometry = settings.get('window_geometry', '')
                    if window_geometry and 'x' in window_geometry:
                        try:
                            self.root.after(10, lambda g=window_geometry: self.root.geometry(g))
                        except:
                            pass
        except:
            pass
    
    def save_settings(self):
        """Save settings to file"""
        try:
            settings = {
                'screenshots_directory': self.dir_var.get(),
                'direct_snip': self.direct_snip.get(),
                'window_geometry': self.root.geometry()
            }
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
        except:
            pass
    
    def save_snip_preference(self):
        """Save snipping tool preference"""
        self.save_settings()
    
    def load_default_directory(self):
        """Load default Tibia screenshot directory"""
        if self.dir_var.get():
            self.load_screenshots()
            return
            
        username = os.environ.get('USERNAME', '')
        if username:
            default_path = f"C:\\Users\\{username}\\AppData\\Local\\Tibia\\packages\\Tibia"
            if os.path.exists(default_path):
                self.dir_var.set(default_path)
                self.load_screenshots()
                self.save_settings()
    
    def browse_directory(self):
        """Open directory browser dialog"""
        directory = filedialog.askdirectory(title="Select Tibia Screenshots Directory")
        if directory:
            self.dir_var.set(directory)
            self.load_screenshots()
            self.save_settings()
    
    def parse_screenshot_filename(self, filename):
        """Parse screenshot filename and return metadata"""
        match = self._filename_pattern.match(filename)
        
        if match:
            date_str, time_str, character, category, ext = match.groups()
            
            try:
                year, month, day = map(int, date_str.split('-'))
                date_obj = datetime(year, month, day)
                formatted_date = date_obj.strftime("%d %B %Y")
                sort_date = date_obj
            except:
                formatted_date = date_str
                sort_date = date_str
            
            try:
                hours = int(time_str[:2])
                minutes = int(time_str[2:4])
                seconds = int(time_str[4:6])
                formatted_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                sort_time = formatted_time
            except:
                formatted_time = time_str
                sort_time = time_str
            
            display_category = self.categories.get(category, category)
            
            return {
                'date': formatted_date, 
                'time': formatted_time, 
                'character': character,
                'category': display_category, 
                'raw_category': category,
                'full_datetime': f"{formatted_date} {formatted_time}",
                'sort_date': sort_date, 
                'sort_time': sort_time
            }
        
        return None
    
    def get_screenshot_files(self, force_rescan=False):
        """Get list of screenshot files with caching"""
        current_time = time.time()
        
        if (not force_rescan and self.file_list_cache is not None and 
            current_time - self.last_scan_time < self.scan_interval):
            return self.file_list_cache
        
        directory = self.dir_var.get()
        if not directory or not os.path.exists(directory):
            self.file_list_cache = []
            return []
        
        screenshot_files = []
        extensions = ('.png', '.jpg', '.jpeg')
        
        for root, dirs, files in os.walk(directory):
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            for file in files:
                if file.lower().endswith(extensions):
                    screenshot_files.append(os.path.join(root, file))
        
        self.file_list_cache = screenshot_files
        self.last_scan_time = current_time
        return screenshot_files
    
    def format_file_size(self, size_bytes):
        """Format file size in human-readable format"""
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB"]
        import math
        i = int(math.floor(math.log(size_bytes, 1024)))
        p = math.pow(1024, i)
        s = round(size_bytes / p, 1)
        
        if s == int(s):
            s = int(s)
            
        return f"{s} {size_names[i]}"
    
    def load_screenshots(self):
        """Load screenshots from directory"""
        directory = self.dir_var.get()
        if not directory or not os.path.exists(directory):
            return
        
        self.screenshots_dir = directory
        self.screenshot_data = {}
        self.image_cache.clear()
        
        # Hide stats panel if visible during reload
        if self.stats_visible:
            self.hide_stats_panel()
        
        screenshot_files = self.get_screenshot_files(force_rescan=True)
        
        if not screenshot_files:
            self.counter_label.config(text="No screenshots found")
            self.update_tree()
            return
        
        total_count = 0
        processed_count = 0
        total_size_bytes = 0
        batch_size = 25
        
        def process_batch(start_idx):
            nonlocal total_count, processed_count, total_size_bytes
            
            end_idx = min(start_idx + batch_size, len(screenshot_files))
            
            for i in range(start_idx, end_idx):
                file_path = screenshot_files[i]
                processed_count += 1
                
                try:
                    file_size = os.path.getsize(file_path)
                    total_size_bytes += file_size
                except OSError:
                    pass
                
                parsed = self.parse_screenshot_filename(os.path.basename(file_path))
                if parsed:
                    character = parsed['character']
                    category = parsed['category']
                    date = parsed['date']
                    time = parsed['time']
                    sort_date = parsed['sort_date']
                    sort_time = parsed['sort_time']
                    
                    if character not in self.screenshot_data:
                        self.screenshot_data[character] = {}
                    if category not in self.screenshot_data[character]:
                        self.screenshot_data[character][category] = {}
                    if date not in self.screenshot_data[character][category]:
                        self.screenshot_data[character][category][date] = {
                            'sort_date': sort_date, 'times': {}
                        }
                    
                    self.screenshot_data[character][category][date]['times'][time] = {
                        'file_path': file_path, 'sort_time': sort_time
                    }
                    total_count += 1
            
            current_size = self.format_file_size(total_size_bytes)
            progress = f"Loading... {processed_count}/{len(screenshot_files)}"
            self.counter_label.config(text=progress)
            self.size_label.config(text=f"Size: {current_size}")
            self.root.update_idletasks()
            
            if end_idx < len(screenshot_files):
                self.root.after(1, lambda: process_batch(end_idx))
            else:
                self.update_tree()
                self.update_date_combo()
                final_size = self.format_file_size(total_size_bytes)
                self.counter_label.config(text=f"Total screenshots: {total_count}")
                self.size_label.config(text=f"Total size: {final_size}")
        
        if screenshot_files:
            process_batch(0)
        else:
            self.counter_label.config(text="No screenshots found")
            self.size_label.config(text="")
    
    def update_tree(self):
        """Update the main screenshot tree view"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for character in sorted(self.screenshot_data.keys()):
            char_item = self.tree.insert('', 'end', text=character, open=False)
            
            for category in sorted(self.screenshot_data[character].keys()):
                cat_item = self.tree.insert(char_item, 'end', text=category, open=False)
                
                date_items = []
                for date in self.screenshot_data[character][category]:
                    date_data = self.screenshot_data[character][category][date]
                    date_items.append((date, date_data['sort_date']))
                
                date_items.sort(key=lambda x: x[1] if isinstance(x[1], datetime) else datetime.min, 
                              reverse=True)
                
                for date, sort_date in date_items:
                    date_item = self.tree.insert(cat_item, 'end', text=date, open=False)
                    
                    time_items = []
                    times_data = self.screenshot_data[character][category][date]['times']
                    for time in times_data:
                        time_data = times_data[time]
                        time_items.append((time, time_data['sort_time'], time_data['file_path']))
                    
                    time_items.sort(key=lambda x: x[1], reverse=True)
                    
                    for time, sort_time, file_path in time_items:
                        self.tree.insert(date_item, 'end', text=time, values=(file_path,))
    
    def update_date_combo(self):
        """Update the date combobox with available dates"""
        dates = set()
        date_objects = []
        
        for character in self.screenshot_data:
            for category in self.screenshot_data[character]:
                for date in self.screenshot_data[character][category]:
                    if date not in dates:
                        dates.add(date)
                        sort_date = self.screenshot_data[character][category][date]['sort_date']
                        if isinstance(sort_date, datetime):
                            date_objects.append((date, sort_date))
                        else:
                            date_objects.append((date, datetime.min))
        
        # Sort by actual date objects (newest first)
        date_objects.sort(key=lambda x: x[1], reverse=True)
        sorted_dates = [date_text for date_text, _ in date_objects]
        
        self.date_combo['values'] = sorted_dates
        
        if sorted_dates:
            self.selected_date.set(sorted_dates[0])
            self.update_category_tree()
    
    def update_category_tree(self):
        """Update the category tree view based on selected date"""
        for item in self.category_tree.get_children():
            self.category_tree.delete(item)
        
        selected_date = self.selected_date.get()
        if not selected_date or not self.screenshot_data:
            return
        
        category_screenshots = {}
        
        for character in self.screenshot_data:
            for category in self.screenshot_data[character]:
                for date in self.screenshot_data[character][category]:
                    if date == selected_date:
                        if category not in category_screenshots:
                            category_screenshots[category] = []
                        
                        times_data = self.screenshot_data[character][category][date]['times']
                        for time in times_data:
                            time_data = times_data[time]
                            category_screenshots[category].append({
                                'time': time, 
                                'character': character,
                                'file_path': time_data['file_path'], 
                                'sort_time': time_data['sort_time']
                            })
        
        for category in sorted(category_screenshots.keys()):
            screenshots = category_screenshots[category]
            cat_item = self.category_tree.insert('', 'end', 
                                               text=f"{category} ({len(screenshots)})", 
                                               open=True)
            screenshots.sort(key=lambda x: x['sort_time'], reverse=True)
            
            for screenshot in screenshots:
                display_text = f"{screenshot['time']} - {screenshot['character']}"
                self.category_tree.insert(cat_item, 'end', text=display_text, 
                                        values=(screenshot['file_path'],))
    
    def on_date_selected(self, event):
        """Handle date selection"""
        self.update_category_tree()
    
    def on_tree_select(self, event):
        """Handle tree selection"""
        self.load_selected_image()
    
    def on_tree_key_nav(self, event):
        """Handle arrow key navigation in tree"""
        self.load_selected_image()
    
    def on_tree_right_click(self, event):
        """Handle right click on tree"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            values = self.tree.item(item, 'values')
            has_file = bool(values and values[0])
            self.tree_context_menu.entryconfig("Copy Screenshot", state=tk.NORMAL if has_file else tk.DISABLED)
            self.tree_context_menu.post(event.x_root, event.y_root)
    
    def on_category_select(self, event):
        """Handle category tree selection"""
        selection = self.category_tree.selection()
        if not selection:
            return
        item = selection[0]
        values = self.category_tree.item(item, 'values')
        if values and values[0]:
            file_path = values[0]
            self.display_image(file_path)
    
    def on_category_right_click(self, event):
        """Handle right click on category tree"""
        item = self.category_tree.identify_row(event.y)
        if item:
            self.category_tree.selection_set(item)
            values = self.category_tree.item(item, 'values')
            has_file = bool(values and values[0])
            self.tree_context_menu.entryconfig("Copy Screenshot", state=tk.NORMAL if has_file else tk.DISABLED)
            self.tree_context_menu.post(event.x_root, event.y_root)
    
    def on_category_key_nav(self, event):
        """Handle arrow key navigation in category tree"""
        self.on_category_select(event)
    
    def on_canvas_right_click(self, event):
        """Handle right click on canvas"""
        if self.current_image_path:
            self.canvas_context_menu.post(event.x_root, event.y_root)
    
    def load_selected_image(self):
        """Load the selected image from tree"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        
        if values and values[0]:
            file_path = values[0]
            self.display_image(file_path)
    
    def display_image(self, file_path):
        """Display an image file"""
        try:
            # Verify file exists
            if not os.path.exists(file_path):
                messagebox.showerror("Error", "Screenshot file no longer exists.")
                return
                
            self.current_image_path = file_path
            
            filename = os.path.basename(file_path)
            parsed = self.parse_screenshot_filename(filename)
            
            if parsed:
                info_text = f"{parsed['character']} - {parsed['category']} - {parsed['full_datetime']}"
                self.info_label.config(text=info_text)
            else:
                self.info_label.config(text=filename)
            
            self.load_and_display_image()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load image:\n{str(e)}")
    
    def load_and_display_image(self):
        """Load and display the current image with caching"""
        if not self.current_image_path:
            return
        
        try:
            self.canvas.update_idletasks()
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            if canvas_width <= 1 or canvas_height <= 1:
                canvas_width, canvas_height = 800, 600
            
            cache_key = f"{self.current_image_path}_{self.zoom_factor}_{canvas_width}x{canvas_height}"
            
            if cache_key in self.image_cache:
                self.current_image = self.image_cache[cache_key]
                self.display_cached_image()
                return
            
            with Image.open(self.current_image_path) as img:
                original_width, original_height = img.size
                
                scale_w = canvas_width / original_width
                scale_h = canvas_height / original_height
                auto_fit_scale = min(scale_w, scale_h, 1.0)
                final_scale = auto_fit_scale * self.zoom_factor
                
                new_width = int(original_width * final_scale)
                new_height = int(original_height * final_scale)
                
                if final_scale < 0.25:
                    resample = Image.Resampling.NEAREST
                elif final_scale < 0.75:
                    resample = Image.Resampling.BILINEAR
                else:
                    resample = Image.Resampling.LANCZOS
                
                img_resized = img.resize((new_width, new_height), resample)
                self.current_image = ImageTk.PhotoImage(img_resized)
                
                self.image_cache[cache_key] = self.current_image
                
                if len(self.image_cache) > self.max_cache_size:
                    keys_to_remove = list(self.image_cache.keys())[:-self.max_cache_size + 3]
                    for key in keys_to_remove:
                        del self.image_cache[key]
                
                self.display_cached_image()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to display image:\n{str(e)}")
    
    def display_cached_image(self):
        """Display the cached image on canvas"""
        if not self.current_image:
            return
            
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width = self.current_image.width()
        img_height = self.current_image.height()
        
        self.canvas.delete("all")
        
        if img_width < canvas_width and img_height < canvas_height:
            x_offset = (canvas_width - img_width) // 2
            y_offset = (canvas_height - img_height) // 2
            self.canvas.create_image(x_offset, y_offset, anchor=tk.NW, image=self.current_image)
            self.canvas.configure(scrollregion=(0, 0, canvas_width, canvas_height))
        else:
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.current_image)
            self.canvas.configure(scrollregion=(0, 0, img_width, img_height))
        
        # Update zoom label
        try:
            with Image.open(self.current_image_path) as img:
                original_width, original_height = img.size
        except:
            return
        
        canvas_width = max(canvas_width, 800)
        canvas_height = max(canvas_height, 600)
        scale_w = canvas_width / original_width
        scale_h = canvas_height / original_height
        auto_fit_scale = min(scale_w, scale_h, 1.0)
        effective_zoom = int(auto_fit_scale * self.zoom_factor * 100)
        
        if self.zoom_factor == 1.0:
            self.zoom_label.config(text=f"{effective_zoom}% (Auto-fit)")
        else:
            self.zoom_label.config(text=f"{effective_zoom}%")
    
    def zoom_in(self):
        """Increase zoom level by 20%"""
        self.zoom_factor = min(self.zoom_factor * 1.2, 5.0)
        self.load_and_display_image()
    
    def zoom_out(self):
        """Decrease zoom level by 20%"""
        self.zoom_factor = max(self.zoom_factor / 1.2, 0.1)
        self.load_and_display_image()
    
    def reset_zoom(self):
        """Reset zoom to auto-fit"""
        self.zoom_factor = 1.0
        self.load_and_display_image()
    
    def on_mouse_wheel(self, event):
        """Handle mouse wheel for zooming"""
        if event.delta > 0 or event.num == 4:
            self.zoom_in()
        else:
            self.zoom_out()
    
    def on_drag_start(self, event):
        """Start dragging"""
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        self.is_dragging = False
        self.canvas.config(cursor="hand2")
    
    def on_drag_motion(self, event):
        """Handle drag motion for panning"""
        if self.current_image and self.zoom_factor > 1.0:
            self.is_dragging = True
            
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y
            
            x_scroll = self.canvas.canvasx(0)
            y_scroll = self.canvas.canvasy(0)
            
            new_x = x_scroll - dx
            new_y = y_scroll - dy
            
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            scroll_region = self.canvas.cget("scrollregion").split()
            
            if len(scroll_region) == 4:
                scroll_width = float(scroll_region[2])
                scroll_height = float(scroll_region[3])
                
                max_x = max(0, scroll_width - canvas_width)
                max_y = max(0, scroll_height - canvas_height)
                
                new_x = max(0, min(new_x, max_x))
                new_y = max(0, min(new_y, max_y))
                
                if scroll_width > canvas_width:
                    self.canvas.xview_moveto(new_x / scroll_width)
                if scroll_height > canvas_height:
                    self.canvas.yview_moveto(new_y / scroll_height)
            
            self.drag_start_x = event.x
            self.drag_start_y = event.y
    
    def on_drag_end(self, event):
        """End dragging"""
        self.canvas.config(cursor="")
        self.is_dragging = False
    
    def on_window_resize(self, event):
        """Handle window resize"""
        if event.widget == self.root and self.current_image_path:
            if self.resize_timer:
                self.root.after_cancel(self.resize_timer)
            self.resize_timer = self.root.after(200, self.refresh_image_display)
    
    def refresh_image_display(self):
        """Refresh image display after resize"""
        self.resize_timer = None
        if self.current_image_path:
            cache_keys_to_remove = [k for k in self.image_cache.keys() if k.startswith(self.current_image_path)]
            for key in cache_keys_to_remove:
                del self.image_cache[key]
            self.load_and_display_image()
    
    def copy_screenshot(self):
        """Copy current screenshot to clipboard"""
        if not self.current_image_path:
            messagebox.showwarning("Warning", "No screenshot selected.")
            return

        if not os.path.exists(self.current_image_path):
            messagebox.showerror("Error", "Screenshot file no longer exists.")
            return

        try:
            if _HAS_WIN32CLIPBOARD:
                img = Image.open(self.current_image_path).convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="BMP")
                bmp_data = buf.getvalue()[14:]  # strip 14-byte BMP file header; CF_DIB expects BITMAPINFOHEADER onward
                buf.close()
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, bmp_data)
                win32clipboard.CloseClipboard()
            else:
                powershell_cmd = f'''
                Add-Type -AssemblyName System.Windows.Forms
                $image = [System.Drawing.Image]::FromFile("{self.current_image_path}")
                [System.Windows.Forms.Clipboard]::SetImage($image)
                $image.Dispose()
                '''
                subprocess.run(["powershell", "-Command", powershell_cmd],
                             capture_output=True, text=True, check=True)
            messagebox.showinfo("Success", "Screenshot copied to clipboard!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy screenshot:\n{str(e)}")
    
    def open_containing_folder(self):
        """Open containing folder"""
        # First try to get path from currently displayed image
        file_path = self.current_image_path
        
        # If no displayed image, get from tree selection
        if not file_path:
            selection = self.tree.selection()
            if selection:
                values = self.tree.item(selection[0], 'values')
                if values and values[0]:
                    file_path = values[0]
        
        # Try category tree as fallback
        if not file_path:
            selection = self.category_tree.selection()
            if selection:
                values = self.category_tree.item(selection[0], 'values')
                if values and values[0]:
                    file_path = values[0]
        
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("Warning", "No valid screenshot selected.")
            return
        
        folder_path = os.path.dirname(file_path)
        os.startfile(folder_path)
    
    def sort_into_folders(self):
        """Sort screenshots into organized folders"""
        if not self.screenshots_dir or not self.screenshot_data:
            messagebox.showwarning("Warning", "Please load screenshots first.")
            return
        
        result = messagebox.askyesno("Sort Screenshots", 
                                   "This will create folders and move screenshots.\n"
                                   "Do you want to proceed?")
        if not result:
            return
        
        try:
            moved_count = 0
            for character in self.screenshot_data:
                for category in self.screenshot_data[character]:
                    for date in self.screenshot_data[character][category]:
                        date_data = self.screenshot_data[character][category][date]
                        times_data = date_data['times']
                        
                        for time in times_data:
                            time_data = times_data[time]
                            source_path = time_data['file_path']
                            
                            target_dir = os.path.join(self.screenshots_dir, character, category)
                            os.makedirs(target_dir, exist_ok=True)
                            
                            filename = os.path.basename(source_path)
                            target_path = os.path.join(target_dir, filename)
                            
                            if source_path != target_path and os.path.exists(source_path):
                                shutil.move(source_path, target_path)
                                moved_count += 1
            
            messagebox.showinfo("Success", f"Moved {moved_count} screenshots into organized folders.")
            self.refresh_screenshots()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to sort screenshots:\n{str(e)}")
    
    def open_snipping_tool(self):
        """Open the snipping tool"""
        try:
            if self.direct_snip.get():
                try:
                    subprocess.Popen(["powershell", "-command", "start ms-screenclip:"])
                    return
                except:
                    try:
                        import pyautogui
                        pyautogui.hotkey('win', 'shift', 's')
                        return
                    except ImportError:
                        pass
            
            try:
                subprocess.Popen(["snippingtool"])
            except:
                try:
                    subprocess.Popen([r"C:\Windows\System32\SnippingTool.exe"])
                except:
                    messagebox.showinfo("Info", 
                                      "Could not open snipping tool automatically.\n"
                                      "Try pressing Windows + Shift + S for direct capture.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open snipping tool:\n{str(e)}")
    
    def refresh_screenshots(self):
        """Refresh the screenshot list"""
        self.get_screenshot_files(force_rescan=True)
        self.load_screenshots()

def main():
    """Main entry point"""
    root = tk.Tk()
    app = TibiaFrames(root)
    
    def on_closing():
        """Handle window closing"""
        app.save_settings()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    style = ttk.Style()
    if 'clam' in style.theme_names():
        style.theme_use('clam')
    
    root.mainloop()

if __name__ == "__main__":
    main()