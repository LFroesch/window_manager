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
        self.root.title("Window Resizer Tool")
        self.root.geometry("700x700")
        
        # Get screen dimensions
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        
        # File to store layouts
        self.layouts_file = "window_layouts.json"
        self.layouts = self.load_layouts()
        
        # Window data
        self.windows = []
        self.selected_windows = []
        
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
        """Create a smart identifier for a window"""
        # Extract core application name from process
        process_base = window_info['process_name'].lower().replace('.exe', '')
        
        # Create title patterns for matching
        title_words = re.findall(r'\b\w+\b', window_info['title'].lower())
        
        # Common app identifiers
        app_identifiers = {
            'chrome': ['chrome', 'google'],
            'firefox': ['firefox', 'mozilla'],
            'code': ['visual studio code', 'code'],
            'notepad': ['notepad'],
            'explorer': ['file explorer', 'windows explorer'],
            'cmd': ['command prompt', 'cmd'],
            'powershell': ['powershell'],
            'terminal': ['terminal'],
            'discord': ['discord'],
            'spotify': ['spotify'],
            'steam': ['steam'],
            'obs': ['obs studio', 'obs'],
            'slack': ['slack'],
            'teams': ['microsoft teams', 'teams']
        }
        
        # Find app type
        app_type = process_base
        for app_key, keywords in app_identifiers.items():
            if any(keyword in window_info['title'].lower() or keyword in process_base for keyword in keywords):
                app_type = app_key
                break
        
        return {
            'app_type': app_type,
            'process_name': window_info['process_name'],
            'class_name': window_info['class_name'],
            'title_keywords': title_words[:3],  # First 3 words for partial matching
            'title_length': len(window_info['title']),
            'original_title': window_info['title']
        }
    
    def match_window_smart(self, identifier, current_windows):
        """Smart matching algorithm to find the best window match"""
        best_match = None
        best_score = 0
        
        for window_info in current_windows:
            score = 0
            
            # Process name match (high priority)
            if identifier['process_name'] == window_info['process_name']:
                score += 50
            
            # App type match (high priority)
            current_identifier = self.create_smart_identifier(window_info)
            if identifier['app_type'] == current_identifier['app_type']:
                score += 40
            
            # Class name match (medium priority)
            if identifier['class_name'] == window_info['class_name']:
                score += 30
            
            # Title keyword matching (medium priority)
            current_keywords = re.findall(r'\b\w+\b', window_info['title'].lower())
            matching_keywords = set(identifier['title_keywords']) & set(current_keywords)
            score += len(matching_keywords) * 10
            
            # Title length similarity (low priority)
            if abs(identifier['title_length'] - len(window_info['title'])) < 10:
                score += 5
            
            # Exact title match (bonus)
            if identifier['original_title'] == window_info['title']:
                score += 100
            
            if score > best_score:
                best_score = score
                best_match = window_info
        
        return best_match, best_score
    
    def create_widgets(self):
        # Main container
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = ctk.CTkLabel(main_frame, text="Smart Window Resizer Tool", 
                                 font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=10)
        
        # Window selection section
        window_frame = ctk.CTkFrame(main_frame)
        window_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(window_frame, text="Select Windows:", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=20, pady=(20,5))
        
        # Window list with detailed info
        self.window_listbox = ctk.CTkScrollableFrame(window_frame, height=200)
        self.window_listbox.pack(fill="x", padx=20, pady=10)
        
        refresh_btn = ctk.CTkButton(window_frame, text="Refresh Windows", 
                                  command=self.refresh_windows)
        refresh_btn.pack(pady=10)
        
        # Position controls
        pos_frame = ctk.CTkFrame(main_frame)
        pos_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(pos_frame, text="Position & Size", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        # Input fields in a grid
        input_frame = ctk.CTkFrame(pos_frame)
        input_frame.pack(padx=20, pady=10)
        
        # X, Y inputs
        xy_frame = ctk.CTkFrame(input_frame)
        xy_frame.pack(pady=5)
        
        ctk.CTkLabel(xy_frame, text="X:").pack(side="left", padx=5)
        self.x_entry = ctk.CTkEntry(xy_frame, width=80, placeholder_text="100")
        self.x_entry.pack(side="left", padx=5)
        
        ctk.CTkLabel(xy_frame, text="Y:").pack(side="left", padx=5)
        self.y_entry = ctk.CTkEntry(xy_frame, width=80, placeholder_text="100")
        self.y_entry.pack(side="left", padx=5)
        
        # Width, Height inputs
        wh_frame = ctk.CTkFrame(input_frame)
        wh_frame.pack(pady=5)
        
        ctk.CTkLabel(wh_frame, text="Width:").pack(side="left", padx=5)
        self.width_entry = ctk.CTkEntry(wh_frame, width=80, placeholder_text="800")
        self.width_entry.pack(side="left", padx=5)
        
        ctk.CTkLabel(wh_frame, text="Height:").pack(side="left", padx=5)
        self.height_entry = ctk.CTkEntry(wh_frame, width=80, placeholder_text="600")
        self.height_entry.pack(side="left", padx=5)
        
        # Apply button
        apply_btn = ctk.CTkButton(pos_frame, text="Apply Position", 
                                command=self.apply_position, height=40)
        apply_btn.pack(pady=10)
        
        # Quick positions
        quick_frame = ctk.CTkFrame(main_frame)
        quick_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(quick_frame, text="Quick Positions", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        # Quick position buttons
        quick_btn_frame = ctk.CTkFrame(quick_frame)
        quick_btn_frame.pack(pady=10)
        
        btn_style = {"width": 120, "height": 35}
        
        # Top row
        top_row = ctk.CTkFrame(quick_btn_frame)
        top_row.pack(pady=5)
        
        ctk.CTkButton(top_row, text="Left Half", command=lambda: self.quick_position("left_half"), **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(top_row, text="Right Half", command=lambda: self.quick_position("right_half"), **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(top_row, text="Top Half", command=lambda: self.quick_position("top_half"), **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(top_row, text="Bottom Half", command=lambda: self.quick_position("bottom_half"), **btn_style).pack(side="left", padx=5)
        
        # Bottom row
        bottom_row = ctk.CTkFrame(quick_btn_frame)
        bottom_row.pack(pady=5)
        
        ctk.CTkButton(bottom_row, text="Maximize", command=lambda: self.quick_position("maximize"), **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(bottom_row, text="Center", command=lambda: self.quick_position("center"), **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(bottom_row, text="Top Left", command=lambda: self.quick_position("top_left"), **btn_style).pack(side="left", padx=5)
        ctk.CTkButton(bottom_row, text="Top Right", command=lambda: self.quick_position("top_right"), **btn_style).pack(side="left", padx=5)
        
        # Layouts section
        layout_frame = ctk.CTkFrame(main_frame)
        layout_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(layout_frame, text="Smart Layouts", 
                    font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)
        
        layout_btn_frame = ctk.CTkFrame(layout_frame)
        layout_btn_frame.pack(pady=10)
        
        # Layout name input
        layout_input_frame = ctk.CTkFrame(layout_btn_frame)
        layout_input_frame.pack(pady=5)
        
        ctk.CTkLabel(layout_input_frame, text="Layout Name:").pack(side="left", padx=5)
        self.layout_name_entry = ctk.CTkEntry(layout_input_frame, width=150, placeholder_text="My Layout")
        self.layout_name_entry.pack(side="left", padx=5)
        
        # Layout buttons
        layout_actions_frame = ctk.CTkFrame(layout_btn_frame)
        layout_actions_frame.pack(pady=5)
        
        ctk.CTkButton(layout_actions_frame, text="Save Layout", 
                     command=self.save_layout, width=120).pack(side="left", padx=5)
        ctk.CTkButton(layout_actions_frame, text="Load Layout", 
                     command=self.show_load_layout_dialog, width=120).pack(side="left", padx=5)
        ctk.CTkButton(layout_actions_frame, text="Delete Layout", 
                     command=self.show_delete_layout_dialog, width=120).pack(side="left", padx=5)
        
        # Matching options
        match_frame = ctk.CTkFrame(layout_frame)
        match_frame.pack(pady=10)
        
        ctk.CTkLabel(match_frame, text="Smart Matching:").pack(side="left", padx=5)
        self.match_threshold = ctk.CTkSlider(match_frame, from_=0, to=100, number_of_steps=20)
        self.match_threshold.set(30)  # Default minimum match score
        self.match_threshold.pack(side="left", padx=5)
        
        ctk.CTkLabel(match_frame, text="Min Score: 30").pack(side="left", padx=5)
    
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
    
    def refresh_windows(self):
        """Refresh the window list with detailed information"""
        # Clear existing checkboxes
        for widget in self.window_listbox.winfo_children():
            widget.destroy()
        
        self.windows = self.get_windows()
        self.window_checkboxes = {}
        
        for window_info in self.windows:
            # Create detailed display text
            display_text = f"{window_info['title']}"
            subtitle = f"({window_info['process_name']} • {window_info['class_name']})"
            
            # Create frame for each window entry
            window_entry = ctk.CTkFrame(self.window_listbox)
            window_entry.pack(fill="x", padx=5, pady=2)
            
            checkbox = ctk.CTkCheckBox(window_entry, text="", 
                                     command=lambda h=window_info['hwnd']: self.on_window_select(h))
            checkbox.pack(side="left", padx=5)
            
            # Main title
            title_label = ctk.CTkLabel(window_entry, text=display_text, 
                                     font=ctk.CTkFont(size=12, weight="bold"))
            title_label.pack(side="left", padx=5)
            
            # Subtitle with process info
            subtitle_label = ctk.CTkLabel(window_entry, text=subtitle, 
                                        font=ctk.CTkFont(size=10), 
                                        text_color="gray")
            subtitle_label.pack(side="left", padx=5)
            
            self.window_checkboxes[window_info['hwnd']] = checkbox
    
    def on_window_select(self, hwnd):
        """Handle window selection"""
        if hwnd in self.selected_windows:
            self.selected_windows.remove(hwnd)
        else:
            self.selected_windows.append(hwnd)
    
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
            "top_right": (self.screen_width // 2, 0, self.screen_width // 2, self.screen_height // 2)
        }
        
        if position in positions:
            x, y, width, height = positions[position]
            for hwnd in selected:
                self.move_window(hwnd, x, y, width, height)
    
    def move_window(self, hwnd, x, y, width, height):
        """Move and resize a window"""
        try:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, width, height, 0)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to move window: {str(e)}")
    
    def save_layout(self):
        """Save current window positions as a smart layout"""
        layout_name = self.layout_name_entry.get()
        if not layout_name:
            messagebox.showwarning("No Name", "Please enter a layout name")
            return
        
        selected = self.get_selected_windows()
        if not selected:
            messagebox.showwarning("No Selection", "Please select windows to save in layout")
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

if __name__ == "__main__":
    app = WindowResizerTool()
    app.run()