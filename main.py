import customtkinter as ctk
from tkinter import messagebox
import win32gui
import win32con
import win32process
import psutil
import json
import os
import re
import threading

# Set appearance mode and color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class WindowResizerTool:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Smart Window Manager Pro")
        self.root.geometry("900x800")
        self.root.minsize(800, 600)
        
        # Get screen dimensions
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        
        # File to store layouts
        self.layouts_file = "window_layouts.json"
        self.layouts = self.load_layouts()
        
        # Window data
        self.windows = []
        self.selected_windows = []
        self.window_groups = {}  # Group windows by application
        
        # Collapsible groups state
        self.collapsed_groups = {}
        
        self.create_widgets()
        self.refresh_windows()
    
    def get_window_info(self, hwnd):
        """Get comprehensive window information for smart matching"""
        try:
            # Basic window info
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            
            # Process info
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(pid)
                process_name = process.name()
                exe_path = process.exe()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                process_name = "Unknown"
                exe_path = ""
            
            # Window position and size
            rect = win32gui.GetWindowRect(hwnd)
            
            return {
                'hwnd': hwnd,
                'title': title,
                'class_name': class_name,
                'process_name': process_name,
                'exe_path': exe_path,
                'pid': pid,
                'rect': rect,
                'width': rect[2] - rect[0],
                'height': rect[3] - rect[1]
            }
        except Exception:
            return None
    
    def create_smart_identifier(self, window_info):
        """Create a smart identifier for a window with better multi-instance support"""
        # Extract core application name from process
        process_base = window_info['process_name'].lower().replace('.exe', '')
        
        # Create title patterns for matching
        title_words = re.findall(r'\b\w+\b', window_info['title'].lower())
        
        # Enhanced app identifiers with better matching
        app_identifiers = {
            'brave': ['brave', 'brave-browser'],
            'chrome': ['chrome', 'google chrome'],
            'firefox': ['firefox', 'mozilla'],
            'code': ['visual studio code', 'code', 'vscode'],
            'notepad': ['notepad'],
            'notepad++': ['notepad++', 'npp'],
            'explorer': ['file explorer', 'windows explorer'],
            'cmd': ['command prompt', 'cmd'],
            'powershell': ['powershell'],
            'terminal': ['terminal', 'windows terminal'],
            'discord': ['discord'],
            'spotify': ['spotify'],
            'steam': ['steam'],
            'obs': ['obs studio', 'obs'],
            'slack': ['slack'],
            'teams': ['microsoft teams', 'teams'],
            'excel': ['excel', 'microsoft excel'],
            'word': ['word', 'microsoft word'],
            'outlook': ['outlook', 'microsoft outlook'],
            'pycharm': ['pycharm'],
            'intellij': ['intellij', 'idea'],
            'sublime': ['sublime text', 'sublime'],
            'atom': ['atom'],
            'git': ['git', 'github desktop']
        }
        
        # Find app type with process name priority
        app_type = process_base
        for app_key, keywords in app_identifiers.items():
            if any(keyword in process_base for keyword in keywords):
                app_type = app_key
                break
            elif any(keyword in window_info['title'].lower() for keyword in keywords):
                app_type = app_key
                break
        
        # Extract meaningful parts of title for better matching
        # Remove common prefixes/suffixes that change frequently
        clean_title = window_info['title']
        # Remove common browser suffixes
        clean_title = re.sub(r' - (Google Chrome|Mozilla Firefox|Brave|Microsoft Edge)$', '', clean_title)
        # Remove VSCode workspace indicators
        clean_title = re.sub(r' - Visual Studio Code$', '', clean_title)
        # Extract file/folder names for editors
        if app_type in ['code', 'sublime', 'atom', 'notepad++']:
            # Try to extract the main file/folder being edited
            title_parts = clean_title.split(' - ')
            if len(title_parts) > 1:
                clean_title = title_parts[0]  # Usually the file/project name
        
        # For browsers, try to extract domain or main content
        if app_type in ['brave', 'chrome', 'firefox']:
            # Look for domain patterns or meaningful content identifiers
            url_match = re.search(r'https?://([^/\s]+)', clean_title)
            if url_match:
                clean_title = url_match.group(1)
            else:
                # Take first few words of title
                words = clean_title.split()
                clean_title = ' '.join(words[:3]) if len(words) > 3 else clean_title
        
        return {
            'app_type': app_type,
            'process_name': window_info['process_name'],
            'class_name': window_info['class_name'],
            'title_keywords': re.findall(r'\b\w+\b', clean_title.lower())[:5],
            'clean_title': clean_title,
            'title_length': len(window_info['title']),
            'original_title': window_info['title'],
            'process_pid': window_info['pid'],
            'exe_path': window_info.get('exe_path', ''),
            # Add position info to help distinguish windows
            'position_x': window_info['rect'][0],
            'position_y': window_info['rect'][1]
        }
    
    def match_window_smart(self, identifier, current_windows):
        """Enhanced smart matching algorithm for better multi-instance support"""
        matches = []
        
        for window_info in current_windows:
            score = 0
            current_identifier = self.create_smart_identifier(window_info)
            
            # Process name match (highest priority)
            if identifier['process_name'] == window_info['process_name']:
                score += 60
            
            # App type match (high priority)
            if identifier['app_type'] == current_identifier['app_type']:
                score += 50
            
            # Class name match (high priority)
            if identifier['class_name'] == window_info['class_name']:
                score += 40
            
            # Executable path match (high priority for distinguishing instances)
            if identifier.get('exe_path') and identifier['exe_path'] == current_identifier.get('exe_path'):
                score += 35
            
            # Clean title matching (medium-high priority)
            if identifier.get('clean_title') and current_identifier.get('clean_title'):
                if identifier['clean_title'].lower() == current_identifier['clean_title'].lower():
                    score += 45
                elif identifier['clean_title'].lower() in current_identifier['clean_title'].lower():
                    score += 25
            
            # Title keyword matching (medium priority)
            if identifier.get('title_keywords') and current_identifier.get('title_keywords'):
                matching_keywords = set(identifier['title_keywords']) & set(current_identifier['title_keywords'])
                score += len(matching_keywords) * 8
            
            # Position similarity (low-medium priority - windows tend to stay in similar areas)
            if identifier.get('position_x') and identifier.get('position_y'):
                x_diff = abs(identifier['position_x'] - current_identifier.get('position_x', 0))
                y_diff = abs(identifier['position_y'] - current_identifier.get('position_y', 0))
                if x_diff < 100 and y_diff < 100:  # Within 100 pixels
                    score += 15
                elif x_diff < 300 and y_diff < 300:  # Within 300 pixels
                    score += 8
            
            # Title length similarity (low priority)
            if identifier.get('title_length'):
                length_diff = abs(identifier['title_length'] - len(window_info['title']))
                if length_diff < 10:
                    score += 5
                elif length_diff < 50:
                    score += 2
            
            # Exact title match (bonus for perfect matches)
            if identifier['original_title'] == window_info['title']:
                score += 100
            
            # Store potential match with score
            if score > 0:
                matches.append((window_info, score, current_identifier))
        
        # Sort by score and return all matches above threshold
        matches.sort(key=lambda x: x[1], reverse=True)
        
        # For debugging and multi-instance handling, return the best match
        if matches:
            return matches[0][0], matches[0][1]
        
        return None, 0
    
    def create_widgets(self):
        # Configure grid weights for responsive design
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Main container with padding
        main_frame = ctk.CTkFrame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)
        main_frame.grid_rowconfigure(2, weight=1)  # Make window list expandable
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Title with improved styling
        title_label = ctk.CTkLabel(main_frame, text="🪟 Smart Window Manager Pro", 
                                 font=ctk.CTkFont(size=28, weight="bold"))
        title_label.grid(row=0, column=0, pady=(10, 20), sticky="ew")
        
        # Layouts section (collapsible)
        self.create_layouts_section(main_frame)
        
        # Create notebook for tabbed interface (only 2 tabs now)
        self.notebook = ctk.CTkTabview(main_frame)
        self.notebook.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        
        # Windows tab
        self.windows_tab = self.notebook.add("🗂️ Windows")
        self.create_windows_tab()
        
        # Quick Actions tab (now includes position controls)
        self.quick_tab = self.notebook.add("⚡ Quick Actions")
        self.create_quick_actions_tab()
    
    def create_layouts_section(self, parent):
        """Create collapsible layouts section at top"""
        self.layouts_frame = ctk.CTkFrame(parent)
        self.layouts_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        self.layouts_frame.grid_columnconfigure(0, weight=1)
        
        # Header with collapse/expand button
        header_frame = ctk.CTkFrame(self.layouts_frame)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        header_frame.grid_columnconfigure(1, weight=1)
        
        self.layouts_collapsed = False
        self.layouts_toggle_btn = ctk.CTkButton(header_frame, text="▼", width=30, height=30,
                                              command=self.toggle_layouts_section)
        self.layouts_toggle_btn.grid(row=0, column=0, padx=5)
        
        layouts_title = ctk.CTkLabel(header_frame, text="💾 Smart Layout Manager", 
                                   font=ctk.CTkFont(size=18, weight="bold"))
        layouts_title.grid(row=0, column=1, sticky="w", padx=10)
        
        # Collapsible content frame
        self.layouts_content = ctk.CTkFrame(self.layouts_frame)
        self.layouts_content.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.layouts_content.grid_columnconfigure(0, weight=1)
        
        # Layout controls
        controls_frame = ctk.CTkFrame(self.layouts_content)
        controls_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        # Layout name input
        name_frame = ctk.CTkFrame(controls_frame)
        name_frame.pack(pady=8)
        
        ctk.CTkLabel(name_frame, text="Layout Name:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=8)
        self.layout_name_entry = ctk.CTkEntry(name_frame, width=200, placeholder_text="Enter layout name...")
        self.layout_name_entry.pack(side="left", padx=8)
        
        # Action buttons
        btn_frame = ctk.CTkFrame(controls_frame)
        btn_frame.pack(pady=8)
        
        ctk.CTkButton(btn_frame, text="💾 Save Layout", command=self.save_layout, 
                     width=140, height=35).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="📂 Load Layout", command=self.show_load_layout_dialog, 
                     width=140, height=35).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="🗑️ Delete Layout", command=self.show_delete_layout_dialog, 
                     width=140, height=35).pack(side="left", padx=5)
        
        # Matching threshold
        threshold_frame = ctk.CTkFrame(controls_frame)
        threshold_frame.pack(pady=10)
        
        ctk.CTkLabel(threshold_frame, text="Smart Matching Sensitivity:", 
                    font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)
        self.match_threshold = ctk.CTkSlider(threshold_frame, from_=10, to=80, number_of_steps=14)
        self.match_threshold.set(40)
        self.match_threshold.pack(side="left", padx=10)
        
        self.threshold_label = ctk.CTkLabel(threshold_frame, text="40%")
        self.threshold_label.pack(side="left", padx=5)
        self.match_threshold.configure(command=self.update_threshold_label)
        
        # Layouts list
        list_frame = ctk.CTkFrame(self.layouts_content)
        list_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        self.layouts_listbox = ctk.CTkScrollableFrame(list_frame, height=150)
        self.layouts_listbox.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        self.refresh_layouts_display()
    
    def toggle_layouts_section(self):
        """Toggle the layouts section visibility"""
        self.layouts_collapsed = not self.layouts_collapsed
        
        if self.layouts_collapsed:
            self.layouts_content.grid_remove()
            self.layouts_toggle_btn.configure(text="▶")
        else:
            self.layouts_content.grid()
            self.layouts_toggle_btn.configure(text="▼")
    
    def create_windows_tab(self):
        """Create the windows management tab with improved UI"""
        # Configure grid
        self.windows_tab.grid_rowconfigure(1, weight=1)
        self.windows_tab.grid_columnconfigure(0, weight=1)
        
        # Search and filter section
        search_frame = ctk.CTkFrame(self.windows_tab)
        search_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        search_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(search_frame, text="🔍", font=ctk.CTkFont(size=16)).grid(row=0, column=0, padx=10, pady=10)
        
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="Search windows by title or app name...")
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=10)
        self.search_entry.bind("<KeyRelease>", self.on_search_change)
        
        refresh_btn = ctk.CTkButton(search_frame, text="🔄 Refresh", 
                                  command=self.refresh_windows, width=100)
        refresh_btn.grid(row=0, column=2, padx=10, pady=10)
        
        # Window list section with improved scrolling
        list_frame = ctk.CTkFrame(self.windows_tab)
        list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        list_frame.grid_rowconfigure(1, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        
        # Header with select all option
        header_frame = ctk.CTkFrame(list_frame)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        header_frame.grid_columnconfigure(1, weight=1)
        
        self.select_all_var = ctk.BooleanVar()
        select_all_cb = ctk.CTkCheckBox(header_frame, text="Select All", 
                                       variable=self.select_all_var,
                                       command=self.toggle_select_all)
        select_all_cb.grid(row=0, column=0, padx=10, pady=5)
        
        self.selection_label = ctk.CTkLabel(header_frame, text="0 windows selected", 
                                          font=ctk.CTkFont(size=12))
        self.selection_label.grid(row=0, column=1, padx=10, pady=5, sticky="e")
        
        # Scrollable window list with better height
        self.window_listbox = ctk.CTkScrollableFrame(list_frame, height=400)
        self.window_listbox.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
    
    def create_quick_actions_tab(self):
        """Create quick actions tab with position controls included"""
        # Configure grid
        self.quick_tab.grid_rowconfigure(0, weight=1)
        self.quick_tab.grid_columnconfigure(0, weight=1)
        
        main_quick_frame = ctk.CTkScrollableFrame(self.quick_tab)
        main_quick_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # Position & Size Controls section
        pos_frame = ctk.CTkFrame(main_quick_frame)
        pos_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(pos_frame, text="📐 Position & Size Controls", 
                    font=ctk.CTkFont(size=18, weight="bold")).pack(pady=15)
        
        # Input fields in improved grid layout
        input_frame = ctk.CTkFrame(pos_frame)
        input_frame.pack(pady=10)
        
        # X, Y inputs
        xy_frame = ctk.CTkFrame(input_frame)
        xy_frame.pack(pady=8)
        
        ctk.CTkLabel(xy_frame, text="X:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=8)
        self.x_entry = ctk.CTkEntry(xy_frame, width=100, placeholder_text="X position")
        self.x_entry.pack(side="left", padx=5)
        
        ctk.CTkLabel(xy_frame, text="Y:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=8)
        self.y_entry = ctk.CTkEntry(xy_frame, width=100, placeholder_text="Y position")
        self.y_entry.pack(side="left", padx=5)
        
        # Width, Height inputs
        wh_frame = ctk.CTkFrame(input_frame)
        wh_frame.pack(pady=8)
        
        ctk.CTkLabel(wh_frame, text="Width:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=8)
        self.width_entry = ctk.CTkEntry(wh_frame, width=100, placeholder_text="Width")
        self.width_entry.pack(side="left", padx=5)
        
        ctk.CTkLabel(wh_frame, text="Height:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=8)
        self.height_entry = ctk.CTkEntry(wh_frame, width=100, placeholder_text="Height")
        self.height_entry.pack(side="left", padx=5)
        
        # Apply button with better styling
        apply_btn = ctk.CTkButton(pos_frame, text="✨ Apply Position", 
                                command=self.apply_position, height=45, 
                                font=ctk.CTkFont(size=14, weight="bold"))
        apply_btn.pack(pady=15)
        
        # Quick position presets section
        ctk.CTkLabel(main_quick_frame, text="⚡ Quick Position Presets", 
                    font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(30, 20))
        
        # Quick position buttons in organized grid
        quick_btn_frame = ctk.CTkFrame(main_quick_frame)
        quick_btn_frame.pack(pady=20, padx=20)
        
        btn_style = {"width": 140, "height": 45, "font": ctk.CTkFont(size=12, weight="bold")}
        
        # Row 1 - Screen halves
        row1 = ctk.CTkFrame(quick_btn_frame)
        row1.pack(pady=8)
        
        ctk.CTkButton(row1, text="⬅️ Left Half", command=lambda: self.quick_position("left_half"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row1, text="➡️ Right Half", command=lambda: self.quick_position("right_half"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row1, text="⬆️ Top Half", command=lambda: self.quick_position("top_half"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row1, text="⬇️ Bottom Half", command=lambda: self.quick_position("bottom_half"), **btn_style).pack(side="left", padx=8)
        
        # Row 2 - Quarters
        row2 = ctk.CTkFrame(quick_btn_frame)
        row2.pack(pady=8)
        
        ctk.CTkButton(row2, text="↖️ Top Left", command=lambda: self.quick_position("top_left"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row2, text="↗️ Top Right", command=lambda: self.quick_position("top_right"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row2, text="↙️ Bottom Left", command=lambda: self.quick_position("bottom_left"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row2, text="↘️ Bottom Right", command=lambda: self.quick_position("bottom_right"), **btn_style).pack(side="left", padx=8)
        
        # Row 3 - Special actions
        row3 = ctk.CTkFrame(quick_btn_frame)
        row3.pack(pady=8)
        
        ctk.CTkButton(row3, text="🔲 Maximize", command=lambda: self.quick_position("maximize"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row3, text="🎯 Center", command=lambda: self.quick_position("center"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row3, text="🔄 Minimize", command=lambda: self.quick_position("minimize"), **btn_style).pack(side="left", padx=8)
        ctk.CTkButton(row3, text="📺 Restore", command=lambda: self.quick_position("restore"), **btn_style).pack(side="left", padx=8)
    
    def get_windows(self):
        """Get all visible windows with comprehensive info"""
        windows = []
        
        def enum_windows_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                window_info = self.get_window_info(hwnd)
                if window_info and window_info['title'] and window_info['title'] != "Program Manager":
                    windows.append(window_info)
            return True
        
        win32gui.EnumWindows(enum_windows_callback, None)
        return windows
    
    def toggle_group_collapse(self, app_type):
        """Toggle collapse state for an app group"""
        self.collapsed_groups[app_type] = not self.collapsed_groups.get(app_type, False)
        self.refresh_windows(self.search_entry.get().lower() if hasattr(self, 'search_entry') else "")
    
    def refresh_windows(self, search_filter=""):
        """Refresh the window list with detailed information and grouping"""
        # Clear existing checkboxes
        for widget in self.window_listbox.winfo_children():
            widget.destroy()
        
        self.windows = self.get_windows()
        self.window_checkboxes = {}
        
        # Filter windows based on search
        filtered_windows = []
        for window_info in self.windows:
            if search_filter:
                if (search_filter in window_info['title'].lower() or 
                    search_filter in window_info['process_name'].lower() or
                    search_filter in self.create_smart_identifier(window_info)['app_type'].lower()):
                    filtered_windows.append(window_info)
            else:
                filtered_windows.append(window_info)
        
        # Group windows by application
        grouped_windows = self.group_windows_by_app(filtered_windows)
        
        # Display grouped windows
        for app_type, app_windows in sorted(grouped_windows.items()):
            if len(app_windows) == 0:
                continue
                
            # App group header with collapse button
            group_header = ctk.CTkFrame(self.window_listbox)
            group_header.pack(fill="x", padx=5, pady=(10, 5))
            
            # Collapse/expand button
            is_collapsed = self.collapsed_groups.get(app_type, False)
            collapse_btn = ctk.CTkButton(group_header, text="▶" if is_collapsed else "▼", 
                                       width=30, height=30,
                                       command=lambda at=app_type: self.toggle_group_collapse(at))
            collapse_btn.pack(side="left", padx=10, pady=8)
            
            # App icon and name
            app_display_name = self.get_app_display_name(app_type)
            app_count = len(app_windows)
            
            header_label = ctk.CTkLabel(group_header, 
                                      text=f"{app_display_name} ({app_count})",
                                      font=ctk.CTkFont(size=14, weight="bold"))
            header_label.pack(side="left", padx=5, pady=8)
            
            # Group select button
            group_select_btn = ctk.CTkButton(group_header, text="Select All", width=80, height=25,
                                           command=lambda windows=app_windows: self.select_app_group(windows))
            group_select_btn.pack(side="right", padx=15, pady=8)
            
            # Individual windows in this group (only if not collapsed)
            if not is_collapsed:
                for i, window_info in enumerate(app_windows):
                    self.create_window_entry(window_info, i, len(app_windows))
        
        # Update selection label
        try:
            self.update_selection_label()
        except:
            pass  # Selection label might not exist yet
        
        # Update layouts display if we're on that tab
        try:
            self.refresh_layouts_display()
        except:
            pass  # Layouts tab might not be created yet
    
    def on_window_select(self, hwnd):
        """Handle window selection"""
        if hwnd in self.selected_windows:
            self.selected_windows.remove(hwnd)
        else:
            self.selected_windows.append(hwnd)
        
        # Update selection label
        try:
            self.update_selection_label()
        except:
            pass
    
    def get_selected_windows(self):
        """Get currently selected windows"""
        return self.selected_windows
    
    def apply_position(self):
        """Apply custom position to selected windows"""
        selected = self.get_selected_windows()
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one window")
            return
        
        try:
            x = int(self.x_entry.get()) if self.x_entry.get() else 100
            y = int(self.y_entry.get()) if self.y_entry.get() else 100
            width = int(self.width_entry.get()) if self.width_entry.get() else 800
            height = int(self.height_entry.get()) if self.height_entry.get() else 600
            
            for hwnd in selected:
                self.move_window(hwnd, x, y, width, height)
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numbers")
    
    def quick_position(self, position):
        """Apply quick position to selected windows"""
        selected = self.get_selected_windows()
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one window")
            return
        
        positions = {
            "left_half": (0, 0, self.screen_width // 2, self.screen_height),
            "right_half": (self.screen_width // 2, 0, self.screen_width // 2, self.screen_height),
            "top_half": (0, 0, self.screen_width, self.screen_height // 2),
            "bottom_half": (0, self.screen_height // 2, self.screen_width, self.screen_height // 2),
            "maximize": (0, 0, self.screen_width, self.screen_height),
            "center": (self.screen_width // 4, self.screen_height // 4, 
                      self.screen_width // 2, self.screen_height // 2),
            "top_left": (0, 0, self.screen_width // 2, self.screen_height // 2),
            "top_right": (self.screen_width // 2, 0, self.screen_width // 2, self.screen_height // 2),
            "bottom_left": (0, self.screen_height // 2, self.screen_width // 2, self.screen_height // 2),
            "bottom_right": (self.screen_width // 2, self.screen_height // 2, 
                           self.screen_width // 2, self.screen_height // 2)
        }
        
        for hwnd in selected:
            if position == "minimize":
                try:
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                except Exception as e:
                    print(f"Failed to minimize window: {e}")
            elif position == "restore":
                try:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                except Exception as e:
                    print(f"Failed to restore window: {e}")
            elif position in positions:
                x, y, width, height = positions[position]
                self.move_window(hwnd, x, y, width, height)
    
    def move_window(self, hwnd, x, y, width, height):
        """Move and resize a window"""
        try:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, width, height, 0)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to move window: {str(e)}")
    
    def save_layout(self):
        """Save current window positions as a smart layout"""
        layout_name = self.layout_name_entry.get().strip()
        if not layout_name:
            messagebox.showwarning("No Name", "Please enter a layout name")
            return
        
        selected = self.get_selected_windows()
        if not selected:
            messagebox.showwarning("No Selection", "Please select windows to save in layout")
            return
        
        # Check if layout already exists
        if layout_name in self.layouts:
            if not messagebox.askyesno("Layout Exists", 
                                     f"Layout '{layout_name}' already exists. Overwrite it?"):
                return
        
        layout_data = {}
        for hwnd in selected:
            # Find window info
            window_info = None
            for w in self.windows:
                if w['hwnd'] == hwnd:
                    window_info = w
                    break
            
            if window_info:
                identifier = self.create_smart_identifier(window_info)
                layout_data[f"window_{len(layout_data)}"] = {
                    "identifier": identifier,
                    "position": {
                        "x": window_info['rect'][0],
                        "y": window_info['rect'][1], 
                        "width": window_info['width'],
                        "height": window_info['height']
                    }
                }
        
        if layout_data:
            self.layouts[layout_name] = layout_data
            self.save_layouts()
            messagebox.showinfo("Success", f"Smart layout '{layout_name}' saved with {len(layout_data)} windows!")
            self.layout_name_entry.delete(0, "end")
            
            # Refresh layouts display
            try:
                self.refresh_layouts_display()
            except:
                pass
    
    def show_load_layout_dialog(self):
        """Show dialog to load a layout with match preview"""
        if not self.layouts:
            messagebox.showinfo("No Layouts", "No saved layouts found")
            return
        
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Load Smart Layout")
        dialog.geometry("600x400")
        
        ctk.CTkLabel(dialog, text="Select Layout to Load:", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=20)
        
        layout_frame = ctk.CTkScrollableFrame(dialog, height=200)
        layout_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        current_windows = self.get_windows()
        
        for layout_name, layout_data in self.layouts.items():
            # Create layout entry with match preview
            layout_entry = ctk.CTkFrame(layout_frame)
            layout_entry.pack(fill="x", pady=5)
            
            # Layout name button
            btn = ctk.CTkButton(layout_entry, text=layout_name, width=150,
                               command=lambda name=layout_name: self.load_layout(name, dialog))
            btn.pack(side="left", padx=5)
            
            # Match preview
            matches = 0
            total = len(layout_data)
            for window_key, window_data in layout_data.items():
                if 'identifier' in window_data:
                    match, score = self.match_window_smart(window_data['identifier'], current_windows)
                    if score >= self.match_threshold.get():
                        matches += 1
            
            match_text = f"Matches: {matches}/{total}"
            match_color = "green" if matches == total else "yellow" if matches > 0 else "red"
            
            match_label = ctk.CTkLabel(layout_entry, text=match_text, text_color=match_color)
            match_label.pack(side="left", padx=10)
        
        ctk.CTkButton(dialog, text="Cancel", command=dialog.destroy).pack(pady=20)
    
    def load_layout(self, layout_name, dialog):
        """Load a saved layout using smart matching"""
        if layout_name not in self.layouts:
            messagebox.showerror("Error", "Layout not found")
            return
        
        layout_data = self.layouts[layout_name]
        current_windows = self.get_windows()
        
        applied_count = 0
        failed_matches = []
        
        for window_key, window_data in layout_data.items():
            if 'identifier' in window_data:
                # Smart matching
                match, score = self.match_window_smart(window_data['identifier'], current_windows)
                
                if match and score >= self.match_threshold.get():
                    pos = window_data['position']
                    self.move_window(match['hwnd'], pos['x'], pos['y'], pos['width'], pos['height'])
                    applied_count += 1
                else:
                    failed_matches.append(window_data['identifier']['original_title'])
        
        dialog.destroy()
        
        # Show results
        result_msg = f"Layout '{layout_name}' loaded!\nApplied to {applied_count} windows."
        if failed_matches:
            result_msg += f"\n\nCouldn't match {len(failed_matches)} windows:\n" + "\n".join(failed_matches[:3])
            if len(failed_matches) > 3:
                result_msg += f"\n... and {len(failed_matches) - 3} more"
        
        messagebox.showinfo("Smart Layout Loaded", result_msg)
    
    def load_layout_direct(self, layout_name):
        """Load a layout directly without dialog"""
        if layout_name not in self.layouts:
            messagebox.showerror("Error", "Layout not found")
            return
        
        layout_data = self.layouts[layout_name]
        current_windows = self.get_windows()
        
        applied_count = 0
        failed_matches = []
        
        for window_key, window_data in layout_data.items():
            if 'identifier' in window_data:
                # Smart matching
                match, score = self.match_window_smart(window_data['identifier'], current_windows)
                
                if match and score >= self.match_threshold.get():
                    pos = window_data['position']
                    self.move_window(match['hwnd'], pos['x'], pos['y'], pos['width'], pos['height'])
                    applied_count += 1
                else:
                    failed_matches.append(window_data['identifier']['original_title'])
        
        # Show results
        result_msg = f"Layout '{layout_name}' loaded!\nApplied to {applied_count} windows."
        if failed_matches:
            result_msg += f"\n\nCouldn't match {len(failed_matches)} windows:\n" + "\n".join(failed_matches[:3])
            if len(failed_matches) > 3:
                result_msg += f"\n... and {len(failed_matches) - 3} more"
        
        messagebox.showinfo("Smart Layout Loaded", result_msg)
        
        # Refresh layouts display
        self.refresh_layouts_display()
    
    def show_delete_layout_dialog(self):
        """Show dialog to delete a layout"""
        if not self.layouts:
            messagebox.showinfo("No Layouts", "No saved layouts found")
            return
        
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Delete Layout")
        dialog.geometry("400x300")
        
        ctk.CTkLabel(dialog, text="Select Layout to Delete:", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=20)
        
        layout_frame = ctk.CTkScrollableFrame(dialog, height=150)
        layout_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        for layout_name in self.layouts.keys():
            btn = ctk.CTkButton(layout_frame, text=layout_name, 
                               command=lambda name=layout_name: self.delete_layout(name, dialog))
            btn.pack(fill="x", pady=5)
        
        ctk.CTkButton(dialog, text="Cancel", command=dialog.destroy).pack(pady=20)
    
    def delete_layout(self, layout_name, dialog):
        """Delete a saved layout"""
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete layout '{layout_name}'?"):
            del self.layouts[layout_name]
            self.save_layouts()
            dialog.destroy()
            messagebox.showinfo("Success", f"Layout '{layout_name}' deleted!")
    
    def delete_layout_direct(self, layout_name):
        """Delete a layout directly with confirmation"""
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete layout '{layout_name}'?"):
            del self.layouts[layout_name]
            self.save_layouts()
            messagebox.showinfo("Success", f"Layout '{layout_name}' deleted!")
            self.refresh_layouts_display()
    
    def load_layouts(self):
        """Load layouts from file"""
        try:
            if os.path.exists(self.layouts_file):
                with open(self.layouts_file, 'r') as f:
                    return json.load(f)
        except Exception:
            pass
        return {}
    
    def save_layouts(self):
        """Save layouts to file"""
        try:
            with open(self.layouts_file, 'w') as f:
                json.dump(self.layouts, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save layouts: {str(e)}")
    
    def run(self):
        """Start the application"""
        self.root.mainloop()
    
    def on_search_change(self, event=None):
        """Handle search input changes"""
        search_term = self.search_entry.get().lower()
        self.refresh_windows(search_filter=search_term)
    
    def toggle_select_all(self):
        """Toggle selection of all visible windows"""
        select_all = self.select_all_var.get()
        
        for hwnd, checkbox in self.window_checkboxes.items():
            if select_all:
                if hwnd not in self.selected_windows:
                    self.selected_windows.append(hwnd)
                checkbox.select()
            else:
                if hwnd in self.selected_windows:
                    self.selected_windows.remove(hwnd)
                checkbox.deselect()
        
        self.update_selection_label()
    
    def update_selection_label(self):
        """Update the selection count label"""
        count = len(self.selected_windows)
        self.selection_label.configure(text=f"{count} window{'s' if count != 1 else ''} selected")
    
    def update_threshold_label(self, value):
        """Update threshold percentage label"""
        self.threshold_label.configure(text=f"{int(value)}%")
    
    def refresh_layouts_display(self):
        """Refresh the layouts display"""
        # Clear existing layout widgets
        for widget in self.layouts_listbox.winfo_children():
            widget.destroy()
        
        if not self.layouts:
            no_layouts_label = ctk.CTkLabel(self.layouts_listbox, 
                                          text="No saved layouts found.\nCreate your first layout in the Windows tab!",
                                          font=ctk.CTkFont(size=14),
                                          text_color="gray")
            no_layouts_label.pack(pady=50)
            return
        
        current_windows = self.get_windows()
        
        for layout_name, layout_data in self.layouts.items():
            # Create layout card
            layout_card = ctk.CTkFrame(self.layouts_listbox)
            layout_card.pack(fill="x", padx=10, pady=8)
            
            # Layout header
            header_frame = ctk.CTkFrame(layout_card)
            header_frame.pack(fill="x", padx=15, pady=15)
            
            # Layout name and info
            name_frame = ctk.CTkFrame(header_frame)
            name_frame.pack(fill="x")
            
            name_label = ctk.CTkLabel(name_frame, text=f"📋 {layout_name}", 
                                    font=ctk.CTkFont(size=16, weight="bold"))
            name_label.pack(side="left", padx=10, pady=5)
            
            # Calculate matches
            matches = 0
            total = len(layout_data)
            for window_key, window_data in layout_data.items():
                if 'identifier' in window_data:
                    match, score = self.match_window_smart(window_data['identifier'], current_windows)
                    if score >= self.match_threshold.get():
                        matches += 1
            
            # Match status
            match_text = f"{matches}/{total} matches"
            match_color = "#00ff00" if matches == total else "#ffaa00" if matches > 0 else "#ff6666"
            
            match_label = ctk.CTkLabel(name_frame, text=match_text, 
                                     text_color=match_color, font=ctk.CTkFont(weight="bold"))
            match_label.pack(side="right", padx=10, pady=5)
            
            # Action buttons
            btn_frame = ctk.CTkFrame(header_frame)
            btn_frame.pack(fill="x", pady=(10, 0))
            
            load_btn = ctk.CTkButton(btn_frame, text="📂 Load", width=80, height=30,
                                   command=lambda name=layout_name: self.load_layout_direct(name))
            load_btn.pack(side="left", padx=5)
            
            delete_btn = ctk.CTkButton(btn_frame, text="🗑️ Delete", width=80, height=30,
                                     command=lambda name=layout_name: self.delete_layout_direct(name))
            delete_btn.pack(side="right", padx=5)
            
            # Layout details (expandable)
            details_text = f"Contains {total} window configurations"
            details_label = ctk.CTkLabel(header_frame, text=details_text, 
                                       font=ctk.CTkFont(size=11), text_color="gray")
            details_label.pack(pady=(5, 0))
    
    def get_app_display_name(self, app_type):
        """Get a friendly display name for the app type"""
        display_names = {
            'brave': '🌐 Brave Browser',
            'chrome': '🌐 Google Chrome', 
            'firefox': '🦊 Firefox',
            'code': '💻 VS Code',
            'notepad': '📝 Notepad',
            'notepad++': '📝 Notepad++',
            'explorer': '📁 File Explorer',
            'cmd': '⚫ Command Prompt',
            'powershell': '🔵 PowerShell',
            'terminal': '💻 Terminal',
            'discord': '💬 Discord',
            'spotify': '🎵 Spotify',
            'steam': '🎮 Steam',
            'obs': '🎥 OBS Studio',
            'slack': '💼 Slack',
            'teams': '📞 Microsoft Teams',
            'excel': '📊 Excel',
            'word': '📄 Word',
            'outlook': '📧 Outlook',
            'pycharm': '🐍 PyCharm',
            'sublime': '✨ Sublime Text',
            'git': '🐙 Git'
        }
        return display_names.get(app_type, f"📱 {app_type.title()}")
    
    def create_window_entry(self, window_info, index, total_in_group):
        """Create an individual window entry with improved styling"""
        # Create window entry frame
        window_entry = ctk.CTkFrame(self.window_listbox)
        window_entry.pack(fill="x", padx=15, pady=2)
        
        # Checkbox
        checkbox = ctk.CTkCheckBox(window_entry, text="", 
                                 command=lambda h=window_info['hwnd']: self.on_window_select(h))
        checkbox.pack(side="left", padx=10, pady=8)
        
        # Window info container
        info_frame = ctk.CTkFrame(window_entry)
        info_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        
        # Main title with better truncation
        title_text = window_info['title']
        if len(title_text) > 60:
            title_text = title_text[:57] + "..."
        
        title_label = ctk.CTkLabel(info_frame, text=title_text, 
                                 font=ctk.CTkFont(size=12, weight="bold"))
        title_label.pack(anchor="w", padx=10, pady=(5, 2))
        
        # Subtitle with enhanced info
        identifier = self.create_smart_identifier(window_info)
        clean_title = identifier.get('clean_title', '')
        if clean_title and clean_title != window_info['title']:
            if len(clean_title) > 40:
                clean_title = clean_title[:37] + "..."
            subtitle = f"Content: {clean_title} • PID: {window_info['pid']}"
        else:
            subtitle = f"{window_info['process_name']} • PID: {window_info['pid']}"
        
        subtitle_label = ctk.CTkLabel(info_frame, text=subtitle, 
                                    font=ctk.CTkFont(size=10), 
                                    text_color="gray")
        subtitle_label.pack(anchor="w", padx=10, pady=(0, 5))
        
        # Window size info
        size_text = f"{window_info['width']}×{window_info['height']} at ({window_info['rect'][0]}, {window_info['rect'][1]})"
        size_label = ctk.CTkLabel(window_entry, text=size_text, 
                                font=ctk.CTkFont(size=9), 
                                text_color="lightgray")
        size_label.pack(side="right", padx=10, pady=8)
        
        self.window_checkboxes[window_info['hwnd']] = checkbox
        
        # If this window is already selected, check the box
        if window_info['hwnd'] in self.selected_windows:
            checkbox.select()
    
    def select_app_group(self, app_windows):
        """Select all windows in an application group"""
        for window_info in app_windows:
            hwnd = window_info['hwnd']
            if hwnd not in self.selected_windows:
                self.selected_windows.append(hwnd)
                if hwnd in self.window_checkboxes:
                    self.window_checkboxes[hwnd].select()
        
        self.update_selection_label()

    def group_windows_by_app(self, windows):
        """Group windows by application for better organization"""
        groups = {}
        for window_info in windows:
            identifier = self.create_smart_identifier(window_info)
            app_type = identifier['app_type']
            
            if app_type not in groups:
                groups[app_type] = []
            groups[app_type].append(window_info)
        
        return groups

if __name__ == "__main__":
    app = WindowResizerTool()
    app.run()