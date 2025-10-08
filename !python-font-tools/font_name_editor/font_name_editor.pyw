import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
from fontTools.ttLib import TTFont, TTCollection
import traceback
import os
import sys
import ctypes

# Windows OLE drag and drop support
if sys.platform == 'win32':
    try:
        import pywintypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # DPI awareness
    except:
        pass

# Try to import tkinterdnd2 for proper drag and drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    print("tkinterdnd2 not found. Install with: pip install tkinterdnd2")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# Create base class depending on DnD availability
if HAS_DND:
    class BaseWindow(TkinterDnD.Tk):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            # Apply CTk styling manually
            ctk.set_appearance_mode("dark")
            ctk.set_default_color_theme("blue")
else:
    BaseWindow = ctk.CTk

class CompactFontEditor(BaseWindow):
    def __init__(self):
        super().__init__()
        
        self.title("Font Name Editor")
        
        # Calculate center position before showing window
        window_width = 750
        window_height = 500
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # Set geometry with center position directly
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.minsize(700, 450)
        
        # Set application icon
        self.set_application_icon()
        
        # Set dark background for the main window
        if HAS_DND:
            # For TkinterDnD.Tk, use standard tkinter configure
            self.configure(bg="#1a1a1a")
        else:
            # For CTk, use fg_color
            self.configure(fg_color="#1a1a1a")
        
        # Apply dark title bar for Windows 10/11
        self.apply_dark_title_bar()
        
        # Data storage
        self.font_files = []
        self.current_font_index = None
        self.current_font_data = {}
        
        # Pin/Unpin state
        self.is_pinned = False
        
        self.setup_compact_ui()
    
    def set_application_icon(self):
        """Set the window and taskbar icon"""
        try:
            # Get the directory where the script is located
            if getattr(sys, 'frozen', False):
                # If running as compiled executable (PyInstaller)
                # PyInstaller extracts to a temp folder, use sys._MEIPASS
                application_path = sys._MEIPASS
            else:
                # If running as script
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(application_path, 'icon', 'font_name_editor.ico')
            
            if os.path.exists(icon_path):
                # Set window icon (this must be done before calling iconbitmap)
                self.iconbitmap(default=icon_path)
                
                # Windows-specific taskbar icon fix
                if sys.platform == 'win32':
                    try:
                        import ctypes
                        from ctypes import wintypes
                        
                        # Set AppUserModelID to separate from Python
                        myappid = 'fonteditor.nameviewer.application.v1'
                        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
                        
                        # Force update the window icon for taskbar
                        # This needs to happen after the window is created
                        self.after(100, lambda: self._update_taskbar_icon(icon_path))
                        
                    except Exception as e:
                        print(f"Taskbar icon error: {e}")
                
                print(f"icon loaded: {icon_path}")
            else:
                print(f"icon not found at: {icon_path}")
        except Exception as e:
            print(f"Could not set icon: {e}")
    
    def _update_taskbar_icon(self, icon_path):
        """Force update the taskbar icon on Windows"""
        try:
            import ctypes
            from ctypes import wintypes
            
            # Get window handle
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            
            # Load icon with proper size
            # LR_LOADFROMFILE = 0x00000010
            # IMAGE_icon = 1
            # LR_DEFAULTSIZE = 0x00000040
            hicon = ctypes.windll.user32.LoadImageW(
                None, 
                icon_path, 
                1,  # IMAGE_icon
                0,  # Use default size
                0,  # Use default size
                0x00000010 | 0x00000040  # LR_LOADFROMFILE | LR_DEFAULTSIZE
            )
            
            if hicon:
                # WM_SETicon = 0x0080
                # icon_SMALL = 0
                # icon_BIG = 1
                ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 0, hicon)  # Small icon
                ctypes.windll.user32.SendMessageW(hwnd, 0x0080, 1, hicon)  # Large icon
                
        except Exception as e:
            print(f"Failed to update taskbar icon: {e}")
    
    def apply_dark_title_bar(self):
        """Apply dark mode to the window title bar on Windows 10/11"""
        try:
            if sys.platform == 'win32':
                # Update the window to get the HWND
                self.update()
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                
                # Windows 10 version 1809 and later support dark title bars
                DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                
                # Try Windows 11 attribute first (19)
                rendering_policy = ctypes.c_int(2)  # Dark mode
                try:
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd,
                        19,  # DWMWA_USE_IMMERSIVE_DARK_MODE for Windows 11
                        ctypes.byref(rendering_policy),
                        ctypes.sizeof(rendering_policy)
                    )
                except:
                    pass
                
                # Try Windows 10 attribute (20)
                try:
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd,
                        DWMWA_USE_IMMERSIVE_DARK_MODE,
                        ctypes.byref(rendering_policy),
                        ctypes.sizeof(rendering_policy)
                    )
                except:
                    pass
                    
                # Redraw title bar
                self.update()
        except Exception as e:
            print(f"Could not apply dark title bar: {e}")
        
        self.setup_compact_ui()
        
    def setup_compact_ui(self):
        # Main grid layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # LEFT PANEL - Font List (Compact)
        left_frame = ctk.CTkFrame(self, width=200, fg_color="#242424")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        left_frame.grid_propagate(False)
        
        # RIGHT PANEL - Edit Area
        right_frame = ctk.CTkFrame(self, fg_color="#242424")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        right_frame.grid_columnconfigure(0, weight=1)
        
        self.setup_left_panel(left_frame)
        self.setup_right_panel(right_frame)
        
    def setup_left_panel(self, parent):
        # Title with counter
        title_frame = ctk.CTkFrame(parent, fg_color="transparent")
        title_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        title = ctk.CTkLabel(title_frame, text="Font Files", font=ctk.CTkFont(size=14, weight="bold"))
        title.pack(side="left")
        
        self.font_count_label = ctk.CTkLabel(title_frame, text="(0)", text_color="#888888")
        self.font_count_label.pack(side="left", padx=(5, 0))
        
        # Action buttons - all in one row with same size
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent", height=30)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        add_btn = ctk.CTkButton(btn_frame, text="+ Add Fonts", width=55, height=24, 
                               command=self.add_font, fg_color="#1f538d", hover_color="#14375e")
        add_btn.pack(side="left", padx=(0, 3))
        
        remove_btn = ctk.CTkButton(btn_frame, text="− Remove", width=55, height=24, 
                                  command=self.remove_font, fg_color="#8b0000", hover_color="#660000")
        remove_btn.pack(side="left", padx=(0, 3))
        
        clear_btn = ctk.CTkButton(btn_frame, text="✕ Clear", width=55, height=24, 
                                 command=self.clear_all_fonts, fg_color="#8b4513", hover_color="#654321")
        clear_btn.pack(side="left")
        
        # Drop area for drag and drop
        self.drop_frame = ctk.CTkFrame(parent, height=60, border_width=2, 
                                      border_color="#1f538d", fg_color="#1a1a1a")
        self.drop_frame.pack(fill="x", padx=10, pady=5)
        self.drop_frame.pack_propagate(False)
        
        dnd_text = "📁 Drag & Drop or Click to Add Fonts" if HAS_DND else "📁 Click to Add Fonts (DnD unavailable)"
        self.drop_label = ctk.CTkLabel(self.drop_frame, text=dnd_text, 
                                 text_color="#1f538d", font=ctk.CTkFont(size=11))
        self.drop_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # Make clickable
        self.drop_frame.bind("<Button-1>", lambda e: self.add_font())
        self.drop_label.bind("<Button-1>", lambda e: self.add_font())
        
        # Setup drag and drop
        self.setup_drag_drop()
        
        # Font list
        self.font_list_frame = ctk.CTkScrollableFrame(parent, height=280, fg_color="#1a1a1a")
        self.font_list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
    def setup_drag_drop(self):
        # Visual feedback on hover
        self.drop_frame.bind("<Enter>", self.on_hover_enter)
        self.drop_frame.bind("<Leave>", self.on_hover_leave)
        
        if HAS_DND:
            # Register the drop target
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
            self.drop_frame.dnd_bind('<<DragEnter>>', self.on_drag_enter)
            self.drop_frame.dnd_bind('<<DragLeave>>', self.on_drag_leave)
            print("tkinterdnd2 drag and drop enabled")
        else:
            print("Drag and drop not available - tkinterdnd2 not installed")
        
    def on_hover_enter(self, event):
        self.drop_frame.configure(fg_color="#2d2d2d", border_color="#3d7cc9")
        
    def on_hover_leave(self, event):
        self.drop_frame.configure(fg_color="#1a1a1a", border_color="#1f538d")
    
    def on_drag_enter(self, event):
        self.drop_frame.configure(fg_color="#2d2d2d", border_color="#3d7cc9")
        return event.action
        
    def on_drag_leave(self, event):
        self.drop_frame.configure(fg_color="#1a1a1a", border_color="#1f538d")
        return event.action
    
    def on_drop(self, event):
        self.drop_frame.configure(fg_color="#1a1a1a", border_color="#1f538d")
        
        # Get dropped files - tkinterdnd2 returns them as a string
        files = self.parse_drop_files(event.data)
        
        for file_path in files:
            if os.path.exists(file_path) and os.path.isfile(file_path):
                self.load_font(file_path)
        
        self.update_font_count()
        return event.action
    
    def parse_drop_files(self, data):
        """Parse the file paths from drag and drop data"""
        # tkinterdnd2 returns files in a special format
        # Could be: {file1} {file2} or file1 file2
        files = []
        current = []
        in_brace = False
        
        for char in data:
            if char == '{':
                in_brace = True
            elif char == '}':
                in_brace = False
                if current:
                    files.append(''.join(current).strip())
                    current = []
            elif char == ' ' and not in_brace:
                if current:
                    files.append(''.join(current).strip())
                    current = []
            else:
                current.append(char)
        
        if current:
            files.append(''.join(current).strip())
        
        return [f for f in files if f]
        
    def setup_right_panel(self, parent):
        # Header with current font info and Pin button
        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")
        header_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(header_frame, text="Editing:", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky="w")
        self.current_font_label = ctk.CTkLabel(header_frame, text="No font selected", text_color="#888888")
        self.current_font_label.grid(row=0, column=1, sticky="w", padx=(5, 0))
        
        # Pin/Unpin button in top-right corner
        self.pin_btn = ctk.CTkButton(header_frame, text="📌 Pin", width=70, height=24, 
                                    command=self.toggle_pin, fg_color="#2b2b2b", hover_color="#3d3d3d")
        self.pin_btn.grid(row=0, column=2, sticky="e")
        
        # Compact form - Vertical layout with full width fields
        form_frame = ctk.CTkFrame(parent, fg_color="#2b2b2b")
        form_frame.grid(row=1, column=0, padx=15, pady=5, sticky="ew")
        form_frame.grid_columnconfigure(1, weight=1)
        
        # Family Name
        ctk.CTkLabel(form_frame, text="Family Name:", text_color="#e0e0e0").grid(row=0, column=0, padx=(10, 5), pady=8, sticky="w")
        self.family_name_entry = ctk.CTkEntry(form_frame, placeholder_text="Arial", fg_color="#1a1a1a", border_color="#3d3d3d")
        self.family_name_entry.grid(row=0, column=1, padx=(0, 10), pady=8, sticky="ew")
        
        # Style
        ctk.CTkLabel(form_frame, text="Style:", text_color="#e0e0e0").grid(row=1, column=0, padx=(10, 5), pady=8, sticky="w")
        self.subfamily_name_entry = ctk.CTkEntry(form_frame, placeholder_text="Regular", fg_color="#1a1a1a", border_color="#3d3d3d")
        self.subfamily_name_entry.grid(row=1, column=1, padx=(0, 10), pady=8, sticky="ew")
        
        # Full Name
        ctk.CTkLabel(form_frame, text="Full Name:", text_color="#e0e0e0").grid(row=2, column=0, padx=(10, 5), pady=8, sticky="w")
        self.full_name_entry = ctk.CTkEntry(form_frame, placeholder_text="Arial Regular", fg_color="#1a1a1a", border_color="#3d3d3d")
        self.full_name_entry.grid(row=2, column=1, padx=(0, 10), pady=8, sticky="ew")
        
        # PostScript Name
        ctk.CTkLabel(form_frame, text="PostScript:", text_color="#e0e0e0").grid(row=3, column=0, padx=(10, 5), pady=8, sticky="w")
        self.postscript_name_entry = ctk.CTkEntry(form_frame, placeholder_text="ArialRegular", fg_color="#1a1a1a", border_color="#3d3d3d")
        self.postscript_name_entry.grid(row=3, column=1, padx=(0, 10), pady=8, sticky="ew")
        
        # Version
        ctk.CTkLabel(form_frame, text="Version:", text_color="#e0e0e0").grid(row=4, column=0, padx=(10, 5), pady=8, sticky="w")
        self.version_entry = ctk.CTkEntry(form_frame, placeholder_text="Version 1.0", fg_color="#1a1a1a", border_color="#3d3d3d")
        self.version_entry.grid(row=4, column=1, padx=(0, 10), pady=8, sticky="ew")
        
        # Action buttons
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.grid(row=2, column=0, padx=15, pady=15, sticky="e")
        
        self.revert_btn = ctk.CTkButton(btn_frame, text="↶ Revert", width=80, 
                                       command=self.revert_changes, state="disabled",
                                       fg_color="#2b2b2b", border_width=1, border_color="#3d3d3d",
                                       hover_color="#3d3d3d",
                                       text_color="#e0e0e0")
        self.revert_btn.pack(side="left", padx=5)
        
        self.apply_btn = ctk.CTkButton(btn_frame, text="💾 Save Changes", width=120, 
                                      command=self.apply_changes, state="disabled",
                                      fg_color="#0d7a0d", hover_color="#0a5e0a")
        self.apply_btn.pack(side="left", padx=5)
        
    def add_font(self):
        try:
            file_paths = filedialog.askopenfilenames(
                title="Select Font Files",
                filetypes=[
                    ("Font Files", "*.ttf *.otf *.ttc *.woff *.woff2"),
                    ("All Files", "*.*")
                ]
            )
            
            if not file_paths:
                return
                
            for file_path in file_paths:
                self.load_font(file_path)
            self.update_font_count()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file dialog:\n{str(e)}")
        
    def load_font(self, file_path):
        try:
            font_path = Path(file_path)
            
            if not font_path.exists():
                messagebox.showerror("Error", f"File not found: {font_path.name}")
                return
            
            if any(f['path'] == font_path for f in self.font_files):
                messagebox.showinfo("Already Loaded", f"{font_path.name} is already in the list.")
                return
            
            # Load font info
            if font_path.suffix.lower() == '.ttc':
                ttc = TTCollection(str(font_path))
                font_info = {
                    'path': font_path,
                    'type': 'collection',
                    'count': len(ttc.fonts),
                    'name': font_path.stem
                }
                ttc.close()
            else:
                font = TTFont(str(font_path))
                name_table = font['name']
                family_name = name_table.getDebugName(1) or font_path.stem
                font_info = {
                    'path': font_path,
                    'type': 'single',
                    'name': family_name
                }
                font.close()
            
            self.font_files.append(font_info)
            self.add_font_to_list(font_info, len(self.font_files) - 1)
            print(f"Successfully loaded: {font_path.name}")
            
        except Exception as e:
            error_details = traceback.format_exc()
            messagebox.showerror("Error", f"Failed to load {os.path.basename(file_path)}:\n{str(e)}")
            print(f"Error loading font:\n{error_details}")
            
    def add_font_to_list(self, font_info, index):
        item_frame = ctk.CTkFrame(self.font_list_frame, height=32, fg_color="transparent")
        item_frame.pack(fill="x", padx=2, pady=1)
        item_frame.grid_propagate(False)
        
        display_name = font_info['path'].name
        if font_info['type'] == 'collection':
            display_name += f" ({font_info['count']} fonts)"
            
        btn = ctk.CTkButton(item_frame, text=display_name, 
                           command=lambda idx=index: self.select_font(idx),
                           anchor="w", height=28, 
                           fg_color="#2b2b2b",
                           hover_color="#3d3d3d",
                           text_color="#e0e0e0")
        btn.pack(fill="x", padx=2, pady=2)
        
    def select_font(self, index):
        if index >= len(self.font_files):
            return
            
        self.current_font_index = index
        font_info = self.font_files[index]
        
        try:
            if font_info['type'] == 'collection':
                messagebox.showinfo("Collection Font", 
                    f"This font collection contains {font_info['count']} variants.\n"
                    "Please select individual font files for editing.")
                return
                
            font = TTFont(str(font_info['path']))
            name_table = font['name']
            
            self.current_font_data = {
                'family': name_table.getDebugName(1) or "",
                'subfamily': name_table.getDebugName(2) or "",
                'full_name': name_table.getDebugName(4) or "",
                'postscript': name_table.getDebugName(6) or "",
                'version': name_table.getDebugName(5) or ""
            }
            
            font.close()
            
            self.current_font_label.configure(text=font_info['path'].name)
            self.family_name_entry.delete(0, 'end')
            self.family_name_entry.insert(0, self.current_font_data['family'])
            
            self.subfamily_name_entry.delete(0, 'end')
            self.subfamily_name_entry.insert(0, self.current_font_data['subfamily'])
            
            self.full_name_entry.delete(0, 'end')
            self.full_name_entry.insert(0, self.current_font_data['full_name'])
            
            self.postscript_name_entry.delete(0, 'end')
            self.postscript_name_entry.insert(0, self.current_font_data['postscript'])
            
            self.version_entry.delete(0, 'end')
            self.version_entry.insert(0, self.current_font_data['version'])
            
            self.revert_btn.configure(state="normal")
            self.apply_btn.configure(state="normal")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load font data:\n{str(e)}")
            
    def revert_changes(self):
        if self.current_font_index is not None and self.current_font_data:
            self.family_name_entry.delete(0, 'end')
            self.family_name_entry.insert(0, self.current_font_data['family'])
            
            self.subfamily_name_entry.delete(0, 'end')
            self.subfamily_name_entry.insert(0, self.current_font_data['subfamily'])
            
            self.full_name_entry.delete(0, 'end')
            self.full_name_entry.insert(0, self.current_font_data['full_name'])
            
            self.postscript_name_entry.delete(0, 'end')
            self.postscript_name_entry.insert(0, self.current_font_data['postscript'])
            
            self.version_entry.delete(0, 'end')
            self.version_entry.insert(0, self.current_font_data['version'])
            
    def apply_changes(self):
        if self.current_font_index is None:
            return
            
        try:
            font_info = self.font_files[self.current_font_index]
            font_path = font_info['path']
            
            new_family = self.family_name_entry.get().strip()
            new_subfamily = self.subfamily_name_entry.get().strip()
            new_full_name = self.full_name_entry.get().strip()
            new_postscript = self.postscript_name_entry.get().strip()
            new_version = self.version_entry.get().strip()
            
            if not new_family:
                messagebox.showwarning("Warning", "Family name cannot be empty!")
                return
                
            font = TTFont(str(font_path))
            name_table = font['name']
            
            for record in name_table.names:
                if record.nameID == 1:
                    record.string = new_family
                elif record.nameID == 2:
                    record.string = new_subfamily
                elif record.nameID == 4:
                    record.string = new_full_name
                elif record.nameID == 6:
                    record.string = new_postscript.replace(" ", "")
                elif record.nameID == 5 and new_version:
                    record.string = new_version
                    
            font.save(str(font_path))
            font.close()
            
            messagebox.showinfo("Success", f"✓ Font saved successfully!\n{font_path.name}")
            
            self.current_font_data = {
                'family': new_family,
                'subfamily': new_subfamily,
                'full_name': new_full_name,
                'postscript': new_postscript,
                'version': new_version
            }
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to apply changes:\n{str(e)}")
            
    def remove_font(self):
        if self.current_font_index is not None:
            self.font_files.pop(self.current_font_index)
            
            for widget in self.font_list_frame.winfo_children():
                widget.destroy()
                
            for i, font_info in enumerate(self.font_files):
                self.add_font_to_list(font_info, i)
            
            self.current_font_index = None
            self.current_font_label.configure(text="No font selected")
            self.clear_form()
            self.update_font_count()
            
    def clear_form(self):
        for entry in [self.family_name_entry, self.subfamily_name_entry, 
                     self.full_name_entry, self.postscript_name_entry, self.version_entry]:
            entry.delete(0, 'end')
            
        self.revert_btn.configure(state="disabled")
        self.apply_btn.configure(state="disabled")
        
    def update_font_count(self):
        count = len(self.font_files)
        self.font_count_label.configure(text=f"({count})")
    
    def clear_all_fonts(self):
        """Clear all loaded fonts at once"""
        if not self.font_files:
            messagebox.showinfo("Info", "No fonts to clear.")
            return
        
        # Ask for confirmation
        if messagebox.askyesno("Clear All Fonts", 
                              f"Are you sure you want to clear all {len(self.font_files)} font(s)?"):
            self.font_files.clear()
            
            # Clear the list display
            for widget in self.font_list_frame.winfo_children():
                widget.destroy()
            
            # Reset current selection
            self.current_font_index = None
            self.current_font_label.configure(text="No font selected")
            self.clear_form()
            self.update_font_count()
    
    def toggle_pin(self):
        """Toggle window always on top"""
        self.is_pinned = not self.is_pinned
        self.attributes('-topmost', self.is_pinned)
        
        if self.is_pinned:
            self.pin_btn.configure(text="📍 Unpin", fg_color="#1f538d", hover_color="#14375e")
        else:
            self.pin_btn.configure(text="📌 Pin", fg_color="#2b2b2b", hover_color="#3d3d3d")

if __name__ == "__main__":
    # Hide console window on Windows
    if sys.platform == 'win32':
        try:
            import ctypes
            ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
        except:
            pass
    
    app = CompactFontEditor()
    app.mainloop()
