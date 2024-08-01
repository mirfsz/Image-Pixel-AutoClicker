import pyautogui
import time
import win32api, win32con
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import keyboard
import os
import subprocess
import threading
import sys
import cv2
import numpy as np
from PIL import Image, ImageTk, ImageGrab
import json

class AutomationScriptGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Automation Script Generator")
        self.root.geometry("1200x800")
        self.root.configure(bg='#1e1e1e')
        self.setup_styles()
        self.positions = []
        self.script_path = None
        self.script_process = None
        self.script_thread = None
        self.emergency_stop_flag = threading.Event()
        self.search_region = None
        self.setup_gui()
        self.setup_keyboard_hooks()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TButton', font=('Arial', 10), background='#3a3a3a', foreground='white')
        self.style.configure('TLabel', font=('Arial', 10), background='#1e1e1e', foreground='white')
        self.style.configure('Header.TLabel', font=('Arial', 16, 'bold'), background='#1e1e1e', foreground='#00aaff')
        self.style.configure('TFrame', background='#1e1e1e')
        self.style.configure('TLabelframe', background='#1e1e1e', foreground='white')
        self.style.configure('TLabelframe.Label', background='#1e1e1e', foreground='white')
        self.style.configure('Treeview', background='#2a2a2a', foreground='white', fieldbackground='#2a2a2a')
        self.style.map('Treeview', background=[('selected', '#00aaff')])

    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Automation Script Generator", style='Header.TLabel').pack(pady=10)

        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.setup_action_list(left_frame)
        self.setup_buttons(left_frame)
        self.setup_instructions(left_frame)
        self.setup_script_editor(right_frame)
        self.setup_log_view(right_frame)

        self.status_var = tk.StringVar()
        status_label = ttk.Label(left_frame, textvariable=self.status_var, font=('Arial', 10, 'italic'))
        status_label.pack(pady=5)

    def setup_action_list(self, parent):
        action_frame = ttk.LabelFrame(parent, text="Recorded Actions", padding="10")
        action_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.action_list = ttk.Treeview(action_frame, columns=("Position", "Color", "Action", "Key", "Image", "Region", "Timeout"), show="headings")
        self.action_list.heading("Position", text="Position")
        self.action_list.heading("Color", text="Color")
        self.action_list.heading("Action", text="Action")
        self.action_list.heading("Key", text="Key")
        self.action_list.heading("Image", text="Image")
        self.action_list.heading("Region", text="Region")
        self.action_list.heading("Timeout", text="Timeout")
        self.action_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(action_frame, orient=tk.VERTICAL, command=self.action_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.action_list.configure(yscrollcommand=scrollbar.set)

        # Enable drag and drop reordering
        self.action_list.bind('<Button-1>', self.on_tree_select)
        self.action_list.bind('<B1-Motion>', self.on_tree_drag)
        self.action_list.bind('<ButtonRelease-1>', self.on_tree_drop)

    def setup_buttons(self, parent):
        button_frame = ttk.Frame(parent)
        button_frame.pack(pady=10)

        buttons = [
            ("Generate Script", self.generate_script, "<Control-g>"),
            ("Run Script", self.run_script, "<Control-r>"),
            ("Pause Script", self.pause_script, "<Control-p>"),
            ("Clear All", self.clear_all, "<Control-d>"),
            ("Edit Selected", self.edit_selected, "<Control-e>"),
            ("Delete Selected", self.delete_selected, "<Delete>"),
            ("Open Script Directory", self.open_script_directory, "<Control-o>"),
            ("Save Project", self.save_project, "<Control-s>"),
            ("Load Project", self.load_project, "<Control-l>"),
            ("Quick Tutorial", self.show_tutorial, "<F1>")
        ]

        for text, command, shortcut in buttons:
            btn = ttk.Button(button_frame, text=text, command=command)
            btn.pack(side=tk.LEFT, padx=5)
            self.create_tooltip(btn, f"{text} ({shortcut})")
            self.root.bind(shortcut, command)

    def setup_instructions(self, parent):
        instruction_frame = ttk.LabelFrame(parent, text="Instructions", padding="10")
        instruction_frame.pack(fill=tk.X, pady=10)
        instructions = (
            "• Press F1 to record an action position (default: click)\n"
            "• Press F2 to capture a color for the last recorded action\n"
            "• For image actions, select image first, then press F4 four times to select search region\n"
            "• Edit actions to change type (click/press/find image) and add key presses\n"
            "• Set custom timeouts for color detection and image recognition\n"
            "• Generate your script when ready, then run, pause, or edit as needed\n"
            "• Press F9 to emergency stop the running script"
        )
        ttk.Label(instruction_frame, text=instructions, justify=tk.LEFT, wraplength=500).pack()

    def setup_script_editor(self, parent):
        script_frame = ttk.LabelFrame(parent, text="Script Editor", padding="10")
        script_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.script_editor = scrolledtext.ScrolledText(script_frame, wrap=tk.WORD, bg='#2a2a2a', fg='white', insertbackground='white')
        self.script_editor.pack(fill=tk.BOTH, expand=True)

    def setup_log_view(self, parent):
        log_frame = ttk.LabelFrame(parent, text="Script Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.log_view = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, bg='#2a2a2a', fg='white', insertbackground='white')
        self.log_view.pack(fill=tk.BOTH, expand=True)

    def setup_keyboard_hooks(self):
        keyboard.on_press_key("f1", self.record_position)
        keyboard.on_press_key("f2", self.capture_color)
        keyboard.on_press_key("f9", self.emergency_stop)

    def create_tooltip(self, widget, text):
        def enter(event):
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            label = tk.Label(self.tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()

        def leave(event):
            if self.tooltip:
                self.tooltip.destroy()

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def on_tree_select(self, event):
        self.tree_drag_source = self.action_list.identify_row(event.y)

    def on_tree_drag(self, event):
        if self.action_list.identify_row(event.y) != self.tree_drag_source:
            self.action_list.selection_set(self.tree_drag_source)

    def on_tree_drop(self, event):
        target = self.action_list.identify_row(event.y)
        if self.tree_drag_source and target:
            source_index = self.action_list.index(self.tree_drag_source)
            target_index = self.action_list.index(target)
            self.positions.insert(target_index, self.positions.pop(source_index))
            self.update_gui()

    def record_position(self, event):
        x, y = pyautogui.position()
        new_action = {'x_action': x, 'y_action': y, 'color_x': None, 'color_y': None, 'color': None, 'action': 'click',
                      'key': None, 'image_path': None, 'search_region': None, 'timeout': None}
        self.positions.append(new_action)
        self.update_gui()
        self.status_var.set(f"Action recorded at ({x}, {y})")

    def capture_color(self, event):
        if self.positions and self.positions[-1]['color'] is None:
            x, y = pyautogui.position()
            r, g, b = pyautogui.pixel(x, y)
            self.positions[-1]['color_x'] = x
            self.positions[-1]['color_y'] = y
            self.positions[-1]['color'] = (r, g, b)
            self.update_gui()
            self.status_var.set(f"Color captured: R={r}, G={g}, B={b}")

    def select_image_and_region(self):
        image_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        if not image_path:
            return None, None

        self.status_var.set("Image selected. Now press F4 four times to select the search region.")
        self.root.update()

        self.region_points = []
        self.temp_keyboard_hook = keyboard.on_press_key("f4", self.select_search_region_point)

        while len(self.region_points) < 4:
            self.root.update()
            time.sleep(0.1)

        keyboard.unhook(self.temp_keyboard_hook)

        if len(self.region_points) == 4:
            x1 = min(p[0] for p in self.region_points)
            y1 = min(p[1] for p in self.region_points)
            x2 = max(p[0] for p in self.region_points)
            y2 = max(p[1] for p in self.region_points)
            search_region = (x1, y1, x2 - x1, y2 - y1)
            self.draw_region_box(search_region)
            return image_path, search_region
        else:
            return None, None

    def select_search_region_point(self, event):
        x, y = pyautogui.position()
        self.region_points.append((x, y))
        self.status_var.set(f"Point {len(self.region_points)} selected at ({x}, {y})")
        self.root.update()

    def draw_region_box(self, region):
        x, y, w, h = region
        box_window = tk.Toplevel(self.root)
        box_window.overrideredirect(True)
        box_window.geometry(f"{w}x{h}+{x}+{y}")
        box_window.attributes('-alpha', 0.3)
        box_window.configure(bg='red')
        box_window.attributes('-topmost', True)

        def close_box():
            box_window.destroy()

        box_window.after(2000, close_box)  # Close after 2 seconds

    def update_gui(self):
        self.action_list.delete(*self.action_list.get_children())
        for idx, pos in enumerate(self.positions, start=1):
            color_info = f"({pos['color_x']}, {pos['color_y']}) - RGB: {pos['color']}" if pos[
                                                                                              'color'] is not None else "None"
            image_info = os.path.basename(pos['image_path']) if pos['image_path'] else "N/A"
            region_info = str(pos['search_region']) if pos['search_region'] else "N/A"
            timeout_info = pos['timeout'] if pos['timeout'] is not None else "Indefinite"
            self.action_list.insert("", "end", values=(
                f"({pos['x_action']}, {pos['y_action']})",
                color_info,
                pos['action'].capitalize(),
                pos['key'] if pos['action'] == 'press' else "N/A",
                image_info,
                region_info,
                timeout_info
            ))

    def edit_selected(self, event=None):
        selected_item = self.action_list.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select an action to edit.")
            return

        index = self.action_list.index(selected_item)
        self.edit_action(index)

    def delete_selected(self, event=None):
        selected_item = self.action_list.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select an action to delete.")
            return

        index = self.action_list.index(selected_item)
        del self.positions[index]
        self.update_gui()

    def edit_action(self, index):
        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"Edit Action {index + 1}")
        edit_window.geometry("400x550")
        edit_window.configure(bg='#1e1e1e')

        pos = self.positions[index]

        ttk.Label(edit_window, text="Action X:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        x_action_entry = ttk.Entry(edit_window)
        x_action_entry.grid(row=0, column=1, padx=5, pady=5)
        x_action_entry.insert(0, pos['x_action'])

        ttk.Label(edit_window, text="Action Y:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        y_action_entry = ttk.Entry(edit_window)
        y_action_entry.grid(row=1, column=1, padx=5, pady=5)
        y_action_entry.insert(0, pos['y_action'])

        ttk.Label(edit_window, text="Color X:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        color_x_entry = ttk.Entry(edit_window)
        color_x_entry.grid(row=2, column=1, padx=5, pady=5)
        color_x_entry.insert(0, pos['color_x'] if pos['color_x'] is not None else '')

        ttk.Label(edit_window, text="Color Y:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        color_y_entry = ttk.Entry(edit_window)
        color_y_entry.grid(row=3, column=1, padx=5, pady=5)
        color_y_entry.insert(0, pos['color_y'] if pos['color_y'] is not None else '')

        ttk.Label(edit_window, text="Color (R,G,B):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        color_entry = ttk.Entry(edit_window)
        color_entry.grid(row=4, column=1, padx=5, pady=5)
        color_entry.insert(0, str(pos['color']) if pos['color'] is not None else '')

        ttk.Label(edit_window, text="Action Type:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        action_var = tk.StringVar(value=pos['action'])
        ttk.Radiobutton(edit_window, text="Click", variable=action_var, value='click').grid(row=5, column=1, padx=5,
                                                                                            pady=5, sticky="w")
        ttk.Radiobutton(edit_window, text="Press", variable=action_var, value='press').grid(row=5, column=2, padx=5,
                                                                                            pady=5, sticky="w")
        ttk.Radiobutton(edit_window, text="Find Image", variable=action_var, value='find_image').grid(row=5, column=3,
                                                                                                      padx=5, pady=5,
                                                                                                      sticky="w")

        ttk.Label(edit_window, text="Key (if Press):").grid(row=6, column=0, padx=5, pady=5, sticky="w")
        key_entry = ttk.Entry(edit_window)
        key_entry.grid(row=6, column=1, padx=5, pady=5)
        key_entry.insert(0, pos['key'] if pos['key'] is not None else '')

        ttk.Label(edit_window, text="Timeout (seconds):").grid(row=7, column=0, padx=5, pady=5, sticky="w")
        timeout_entry = ttk.Entry(edit_window)
        timeout_entry.grid(row=7, column=1, padx=5, pady=5)
        timeout_entry.insert(0, pos['timeout'] if pos['timeout'] is not None else '')

        def select_image():
            image_path, search_region = self.select_image_and_region()
            if image_path and search_region:
                pos['image_path'] = image_path
                pos['search_region'] = search_region
                image_label.config(text=f"Image: {os.path.basename(image_path)}")
                region_label.config(text=f"Region: {search_region}")

        ttk.Button(edit_window, text="Select Image", command=select_image).grid(row=8, column=0, columnspan=2, pady=10)

        image_label = ttk.Label(edit_window,
                                text=f"Image: {os.path.basename(pos['image_path']) if pos['image_path'] else 'None'}")
        image_label.grid(row=9, column=0, columnspan=2, pady=5)

        region_label = ttk.Label(edit_window, text=f"Region: {pos['search_region']}")
        region_label.grid(row=10, column=0, columnspan=2, pady=5)

        def save_edits():
            try:
                pos['x_action'] = int(x_action_entry.get())
                pos['y_action'] = int(y_action_entry.get())
                pos['color_x'] = int(color_x_entry.get()) if color_x_entry.get() else None
                pos['color_y'] = int(color_y_entry.get()) if color_y_entry.get() else None
                pos['color'] = eval(color_entry.get()) if color_entry.get() else None
                pos['action'] = action_var.get()
                pos['key'] = key_entry.get() if action_var.get() == 'press' else None
                pos['timeout'] = float(timeout_entry.get()) if timeout_entry.get() else None
                edit_window.destroy()
                self.update_gui()
            except ValueError as e:
                messagebox.showerror("Invalid Input", f"Please check your inputs: {str(e)}")

        save_button = ttk.Button(edit_window, text="Save", command=save_edits)
        save_button.grid(row=11, column=0, columnspan=2, pady=10)

        def test_action():
            self.test_action(pos)

        test_button = ttk.Button(edit_window, text="Test Action", command=test_action)
        test_button.grid(row=12, column=0, columnspan=2, pady=10)

    def generate_script(self, event=None):
        script = self.create_script()
        script_dir = os.path.join(os.path.expanduser("~"), "AutomationScripts")
        os.makedirs(script_dir, exist_ok=True)
        self.script_path = os.path.join(script_dir, "automation_script.py")

        with open(self.script_path, "w") as f:
            f.write(script)

        self.script_editor.delete('1.0', tk.END)
        self.script_editor.insert(tk.END, script)
        self.status_var.set(f"Script generated and saved to: {self.script_path}")

    def create_script(self):
        script_lines = [
            "import pyautogui",
            "import time",
            "import win32api, win32con",
            "import sys",
            "import cv2",
            "import numpy as np",
            "",
            "def click(x, y):",
            "    win32api.SetCursorPos((x, y))",
            "    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)",
            "    time.sleep(0.1)",
            "    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)",
            "",
            "def wait_for_color(x, y, target_color, timeout=None):",
            "    start_time = time.time()",
            "    while timeout is None or time.time() - start_time < timeout:",
            "        if pyautogui.pixel(x, y) == target_color:",
            "            return True",
            "        time.sleep(0.1)",
            "    return False",
            "",
            "def find_image(image_path, search_region, timeout=None):",
            "    target = cv2.imread(image_path, 0)",
            "    start_time = time.time()",
            "    while timeout is None or time.time() - start_time < timeout:",
            "        screenshot = pyautogui.screenshot(region=search_region)",
            "        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)",
            "        result = cv2.matchTemplate(screenshot, target, cv2.TM_CCOEFF_NORMED)",
            "        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)",
            "        if max_val > 0.8:",
            "            return (max_loc[0] + search_region[0], max_loc[1] + search_region[1])",
            "        time.sleep(0.1)",
            "    return None",
            "",
            "positions = [",
        ]

        for pos in self.positions:
            script_lines.append(f"    {pos},")

        script_lines.append("]")
        script_lines.append("")
        script_lines.append("for pos in positions:")
        script_lines.append("    if pos['action'] == 'find_image':")
        script_lines.append(
            "        print(f\"Searching for image: {pos['image_path']} in region {pos['search_region']}\")")
        script_lines.append("        sys.stdout.flush()")
        script_lines.append("        location = find_image(pos['image_path'], pos['search_region'], pos['timeout'])")
        script_lines.append("        if location:")
        script_lines.append("            print(f\"Object found at: {location}\")")
        script_lines.append("            sys.stdout.flush()")
        script_lines.append("            print(f\"Executing action: click at {location}\")")
        script_lines.append("            sys.stdout.flush()")
        script_lines.append("            click(location[0], location[1])")
        script_lines.append("        else:")
        script_lines.append("            print(f\"Object not found within timeout period\")")
        script_lines.append("            sys.stdout.flush()")
        script_lines.append("    elif pos['color'] is not None:")
        script_lines.append(
            "        print(f\"Waiting for color {pos['color']} at ({pos['color_x']}, {pos['color_y']})\")")
        script_lines.append("        sys.stdout.flush()")
        script_lines.append("        if wait_for_color(pos['color_x'], pos['color_y'], pos['color'], pos['timeout']):")
        script_lines.append(
            "            print(f\"Color condition met for action at ({pos['x_action']}, {pos['y_action']})\")")
        script_lines.append("            sys.stdout.flush()")
        script_lines.append("            if pos['action'] == 'click':")
        script_lines.append(
            "                print(f\"Executing action: click at ({pos['x_action']}, {pos['y_action']})\")")
        script_lines.append("                sys.stdout.flush()")
        script_lines.append("                click(pos['x_action'], pos['y_action'])")
        script_lines.append("            elif pos['action'] == 'press':")
        script_lines.append("                print(f\"Executing action: press key {pos['key']}\")")
        script_lines.append("                sys.stdout.flush()")
        script_lines.append("                pyautogui.press(pos['key'])")
        script_lines.append("        else:")
        script_lines.append("            print(f\"Timeout waiting for color at ({pos['color_x']}, {pos['color_y']})\")")
        script_lines.append("            print(\"Skipping this action and continuing with the next\")")
        script_lines.append("            sys.stdout.flush()")
        script_lines.append("    else:")
        script_lines.append("        print(f\"Executing action at ({pos['x_action']}, {pos['y_action']})\")")
        script_lines.append("        sys.stdout.flush()")
        script_lines.append("        if pos['action'] == 'click':")
        script_lines.append("            click(pos['x_action'], pos['y_action'])")
        script_lines.append("        elif pos['action'] == 'press':")
        script_lines.append("            pyautogui.press(pos['key'])")
        script_lines.append("    time.sleep(0.5)  # Adjust delay if needed")
        script_lines.append("")
        script_lines.append("print(\"Script execution completed.\")")
        script_lines.append("sys.stdout.flush()")

        return "\n".join(script_lines)

    def run_script(self, event=None):
        if not self.script_path:
            messagebox.showerror("Error", "Please generate a script first.")
            return

        if self.script_thread and self.script_thread.is_alive():
            messagebox.showinfo("Info", "Script is already running.")
            return

        self.emergency_stop_flag.clear()
        self.clear_log()
        self.script_thread = threading.Thread(target=self.run_script_thread)
        self.script_thread.start()
        self.status_var.set("Script is running...")

    def run_script_thread(self):
        try:
            process = subprocess.Popen(["python", self.script_path],
                                       stdout=subprocess.PIPE,
                                       stderr=subprocess.PIPE,
                                       creationflags=subprocess.CREATE_NO_WINDOW)

            while True:
                if self.emergency_stop_flag.is_set():
                    process.terminate()
                    self.log_message("Script execution stopped by user.")
                    break

                output = process.stdout.readline()
                if output == b'' and process.poll() is not None:
                    break
                if output:
                    self.log_message(output.strip().decode())

            rc = process.poll()
            if rc == 0:
                self.log_message("Script execution completed successfully.")
            elif not self.emergency_stop_flag.is_set():
                self.log_message("Script execution failed with return code: " + str(rc))
        except Exception as e:
            self.log_message(f"An error occurred: {str(e)}")
        finally:
            self.status_var.set("Script execution finished.")

    def log_message(self, message):
        self.log_view.insert(tk.END, message + "\n")
        self.log_view.see(tk.END)
        self.log_view.update()

    def pause_script(self, event=None):
        if not self.script_thread or not self.script_thread.is_alive():
            messagebox.showinfo("Info", "No script is currently running.")
            return

        self.emergency_stop_flag.set()
        self.status_var.set("Script paused.")

    def emergency_stop(self, event):
        if self.script_thread and self.script_thread.is_alive():
            self.emergency_stop_flag.set()
            self.status_var.set("Emergency stop triggered. Stopping script...")

    def save_script(self):
        if not self.script_path:
            messagebox.showerror("Error", "No script has been generated yet.")
            return

        script_content = self.script_editor.get('1.0', tk.END)
        with open(self.script_path, 'w') as f:
            f.write(script_content)
        self.status_var.set(f"Script saved to: {self.script_path}")

    def open_script_directory(self, event=None):
        script_dir = os.path.join(os.path.expanduser("~"), "AutomationScripts")
        if not os.path.exists(script_dir):
            os.makedirs(script_dir)
        os.startfile(script_dir)

    def clear_all(self, event=None):
        if messagebox.askyesno("Clear All", "Are you sure you want to clear all actions?"):
            self.positions.clear()
            self.update_gui()
            self.status_var.set("All actions cleared.")

    def clear_log(self):
        self.log_view.delete('1.0', tk.END)

    def save_project(self, event=None):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'w') as f:
                json.dump(self.positions, f)
            self.status_var.set(f"Project saved to: {file_path}")

    def load_project(self, event=None):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, 'r') as f:
                self.positions = json.load(f)
            self.update_gui()
            self.status_var.set(f"Project loaded from: {file_path}")

    def show_tutorial(self, event=None):
        tutorial_text = """
                Quick Start Guide:
                1. Press F1 to record a mouse position for an action.
                2. Press F2 to capture a color at the current mouse position.
                3. Use the 'Edit Selected' button to modify action details.
                4. For image recognition, select an image and define a search region.
                5. Generate the script when you're ready.
                6. Run the script and monitor its progress in the log view.
                7. Use F9 for emergency stop during script execution.

                Tips:
                - You can drag and drop actions in the list to reorder them.
                - Use the 'Test Action' button in the edit window to verify individual actions.
                - Save your project regularly to avoid losing your work.
                """
        messagebox.showinfo("Tutorial", tutorial_text)

    def test_action(self, pos):
        def execute_test():
            if pos['action'] == 'click':
                pyautogui.click(pos['x_action'], pos['y_action'])
                self.log_message(f"Clicked at ({pos['x_action']}, {pos['y_action']})")
            elif pos['action'] == 'press':
                pyautogui.press(pos['key'])
                self.log_message(f"Pressed key: {pos['key']}")
            elif pos['action'] == 'find_image':
                location = self.find_image_test(pos['image_path'], pos['search_region'], pos['timeout'])
                if location:
                    self.log_message(f"Image found at: {location}")
                    pyautogui.click(location[0], location[1])
                else:
                    self.log_message("Image not found")

            if pos['color'] is not None:
                color_match = self.wait_for_color_test(pos['color_x'], pos['color_y'], pos['color'], pos['timeout'])
                if color_match:
                    self.log_message(f"Color condition met at ({pos['color_x']}, {pos['color_y']})")
                else:
                    self.log_message(f"Color condition not met within timeout")

        test_thread = threading.Thread(target=execute_test)
        test_thread.start()

    def find_image_test(self, image_path, search_region, timeout):
        start_time = time.time()
        while timeout is None or time.time() - start_time < timeout:
            try:
                location = pyautogui.locateOnScreen(image_path, region=search_region, confidence=0.8)
                if location:
                    return pyautogui.center(location)
            except pyautogui.ImageNotFoundException:
                pass
            time.sleep(0.1)
        return None

    def wait_for_color_test(self, x, y, target_color, timeout):
        start_time = time.time()
        while timeout is None or time.time() - start_time < timeout:
            if pyautogui.pixel(x, y) == target_color:
                return True
            time.sleep(0.1)
        return False

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = AutomationScriptGenerator()
    app.run()
