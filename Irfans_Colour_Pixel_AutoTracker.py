import pyautogui
import time
import win32api, win32con
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import keyboard
import os
import subprocess
import threading
import sys
import io

class AutomationScriptGenerator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Automation Script Generator")
        self.root.geometry("1200x800")
        self.root.configure(bg='#1e1e1e')

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

        self.positions = []
        self.script_path = None
        self.script_process = None
        self.script_thread = None
        self.emergency_stop_flag = threading.Event()
        self.setup_gui()
        self.setup_keyboard_hooks()

    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Automation Script Generator", style='Header.TLabel').pack(pady=10)

        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        action_frame = ttk.LabelFrame(left_frame, text="Recorded Actions", padding="10")
        action_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.action_list = ttk.Treeview(action_frame, columns=("Position", "Color", "Action", "Key"), show="headings")
        self.action_list.heading("Position", text="Position")
        self.action_list.heading("Color", text="Color")
        self.action_list.heading("Action", text="Action")
        self.action_list.heading("Key", text="Key")
        self.action_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(action_frame, orient=tk.VERTICAL, command=self.action_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.action_list.configure(yscrollcommand=scrollbar.set)

        button_frame = ttk.Frame(left_frame)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Generate Script", command=self.generate_script).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Run Script", command=self.run_script).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Pause Script", command=self.pause_script).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear All", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Edit Selected", command=self.edit_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete Selected", command=self.delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Open Script Directory", command=self.open_script_directory).pack(side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar()
        status_label = ttk.Label(left_frame, textvariable=self.status_var, font=('Arial', 10, 'italic'))
        status_label.pack(pady=5)

        instruction_frame = ttk.LabelFrame(left_frame, text="Instructions", padding="10")
        instruction_frame.pack(fill=tk.X, pady=10)
        instructions = (
            "• Press F1 to record an action position (default: click)\n"
            "• Press F2 to capture a color for the last recorded action\n"
            "• Edit actions to change type (click/press) and add key presses\n"
            "• Generate your script when ready, then run, pause, or edit as needed\n"
            "• Press F9 to emergency stop the running script"
        )
        ttk.Label(instruction_frame, text=instructions, justify=tk.LEFT, wraplength=500).pack()

        script_frame = ttk.LabelFrame(right_frame, text="Script Editor", padding="10")
        script_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.script_editor = scrolledtext.ScrolledText(script_frame, wrap=tk.WORD, bg='#2a2a2a', fg='white', insertbackground='white')
        self.script_editor.pack(fill=tk.BOTH, expand=True)

        log_frame = ttk.LabelFrame(right_frame, text="Script Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.log_view = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, bg='#2a2a2a', fg='white', insertbackground='white')
        self.log_view.pack(fill=tk.BOTH, expand=True)

        ttk.Button(right_frame, text="Save Script", command=self.save_script).pack(pady=5)
        ttk.Button(right_frame, text="Clear Log", command=self.clear_log).pack(pady=5)

    def setup_keyboard_hooks(self):
        keyboard.on_press_key("f1", self.record_position)
        keyboard.on_press_key("f2", self.capture_color)
        keyboard.on_press_key("f9", self.emergency_stop)

    def record_position(self, event):
        x, y = pyautogui.position()
        new_action = {'x_action': x, 'y_action': y, 'color_x': None, 'color_y': None, 'color': None, 'action': 'click', 'key': None}
        self.positions.append(new_action)
        self.update_gui()

    def capture_color(self, event):
        if self.positions and self.positions[-1]['color'] is None:
            x, y = pyautogui.position()
            r, g, b = pyautogui.pixel(x, y)
            self.positions[-1]['color_x'] = x
            self.positions[-1]['color_y'] = y
            self.positions[-1]['color'] = (r, g, b)
            self.update_gui()
            self.status_var.set(f"Color captured: R={r}, G={g}, B={b}")

    def update_gui(self):
        self.action_list.delete(*self.action_list.get_children())
        for idx, pos in enumerate(self.positions, start=1):
            color_info = f"({pos['color_x']}, {pos['color_y']}) - RGB: {pos['color']}" if pos['color'] is not None else "None"
            self.action_list.insert("", "end", values=(
                f"({pos['x_action']}, {pos['y_action']})",
                color_info,
                pos['action'].capitalize(),
                pos['key'] if pos['action'] == 'press' else "N/A"
            ))

    def edit_selected(self):
        selected_item = self.action_list.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select an action to edit.")
            return
        
        index = self.action_list.index(selected_item)
        self.edit_action(index)

    def delete_selected(self):
        selected_item = self.action_list.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select an action to delete.")
            return
        
        index = self.action_list.index(selected_item)
        del self.positions[index]
        self.update_gui()

    def edit_action(self, index):
        edit_window = tk.Toplevel(self.root)
        edit_window.title(f"Edit Action {index+1}")
        edit_window.geometry("400x300")
        edit_window.configure(bg='#1e1e1e')

        pos = self.positions[index]

        ttk.Label(edit_window, text="Action X:").grid(row=0, column=0, padx=5, pady=5)
        x_action_entry = ttk.Entry(edit_window)
        x_action_entry.grid(row=0, column=1, padx=5, pady=5)
        x_action_entry.insert(0, pos['x_action'])

        ttk.Label(edit_window, text="Action Y:").grid(row=1, column=0, padx=5, pady=5)
        y_action_entry = ttk.Entry(edit_window)
        y_action_entry.grid(row=1, column=1, padx=5, pady=5)
        y_action_entry.insert(0, pos['y_action'])

        ttk.Label(edit_window, text="Color X:").grid(row=2, column=0, padx=5, pady=5)
        color_x_entry = ttk.Entry(edit_window)
        color_x_entry.grid(row=2, column=1, padx=5, pady=5)
        color_x_entry.insert(0, pos['color_x'] if pos['color_x'] is not None else '')

        ttk.Label(edit_window, text="Color Y:").grid(row=3, column=0, padx=5, pady=5)
        color_y_entry = ttk.Entry(edit_window)
        color_y_entry.grid(row=3, column=1, padx=5, pady=5)
        color_y_entry.insert(0, pos['color_y'] if pos['color_y'] is not None else '')

        ttk.Label(edit_window, text="Color (R,G,B):").grid(row=4, column=0, padx=5, pady=5)
        color_entry = ttk.Entry(edit_window)
        color_entry.grid(row=4, column=1, padx=5, pady=5)
        color_entry.insert(0, str(pos['color']) if pos['color'] is not None else '')

        ttk.Label(edit_window, text="Action Type:").grid(row=5, column=0, padx=5, pady=5)
        action_var = tk.StringVar(value=pos['action'])
        ttk.Radiobutton(edit_window, text="Click", variable=action_var, value='click').grid(row=5, column=1, padx=5, pady=5)
        ttk.Radiobutton(edit_window, text="Press", variable=action_var, value='press').grid(row=5, column=2, padx=5, pady=5)

        ttk.Label(edit_window, text="Key (if Press):").grid(row=6, column=0, padx=5, pady=5)
        key_entry = ttk.Entry(edit_window)
        key_entry.grid(row=6, column=1, padx=5, pady=5)
        key_entry.insert(0, pos['key'] if pos['key'] is not None else '')

        def save_edits():
            pos['x_action'] = int(x_action_entry.get())
            pos['y_action'] = int(y_action_entry.get())
            pos['color_x'] = int(color_x_entry.get()) if color_x_entry.get() else None
            pos['color_y'] = int(color_y_entry.get()) if color_y_entry.get() else None
            pos['color'] = eval(color_entry.get()) if color_entry.get() else None
            pos['action'] = action_var.get()
            pos['key'] = key_entry.get() if action_var.get() == 'press' else None
            edit_window.destroy()
            self.update_gui()

        save_button = ttk.Button(edit_window, text="Save", command=save_edits)
        save_button.grid(row=7, column=0, columnspan=2, pady=10)

    def generate_script(self):
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
            "",
            "def click(x, y):",
            "    win32api.SetCursorPos((x, y))",
            "    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)",
            "    time.sleep(0.1)",
            "    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)",
            "",
            "def wait_for_color(x, y, target_color, timeout=30):",
            "    start_time = time.time()",
            "    while time.time() - start_time < timeout:",
            "        if pyautogui.pixel(x, y) == target_color:",
            "            return True",
            "        time.sleep(0.1)",
            "    return False",
            "",
            "positions = [",
        ]

        for pos in self.positions:
            script_lines.append(f"    {pos},")
        
        script_lines.append("]")
        script_lines.append("")
        script_lines.append("for pos in positions:")
        script_lines.append("    if pos['color'] is not None:")
        script_lines.append("        print(f\"Waiting for color {pos['color']} at ({pos['color_x']}, {pos['color_y']})\")")
        script_lines.append("        sys.stdout.flush()")
        script_lines.append("        if wait_for_color(pos['color_x'], pos['color_y'], pos['color']):")
        script_lines.append("            print(f\"Color condition met for action at ({pos['x_action']}, {pos['y_action']})\")")
        script_lines.append("            sys.stdout.flush()")
        script_lines.append("            if pos['action'] == 'click':")
        script_lines.append("                click(pos['x_action'], pos['y_action'])")
        script_lines.append("            elif pos['action'] == 'press':")
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

    def run_script(self):
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

    def pause_script(self):
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

    def open_script_directory(self):
        script_dir = os.path.join(os.path.expanduser("~"), "AutomationScripts")
        if not os.path.exists(script_dir):
            os.makedirs(script_dir)
        os.startfile(script_dir)

    def clear_all(self):
        self.positions.clear()
        self.update_gui()
        self.status_var.set("All actions cleared.")

    def clear_log(self):
        self.log_view.delete('1.0', tk.END)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = AutomationScriptGenerator()
    app.run()
