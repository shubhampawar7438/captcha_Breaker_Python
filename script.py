import time
import pyautogui
import sys
import random
import string
import os
import json
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog, scrolledtext
import threading
import webbrowser
from pynput import mouse

# Configure pyautogui settings
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

class AutomationRecorder:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Professional Automation Recorder")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Apply modern theme
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Variables
        self.actions = []
        self.recording = False
        self.browser_shortcut = ""
        self.target_url = ""
        self.selected_element_type = "email"
        self.waiting_for_click = False
        self.mouse_listener = None
        self.stop_replay_flag = False  # Flag to stop replay
        self.replay_thread = None
        
        # Create main interface
        self.create_main_interface()
        
        # Load saved settings
        self.load_settings()
        
    def create_main_interface(self):
        """Create the main application interface"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: Record Session
        self.record_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.record_frame, text="Record Session")
        self.create_record_tab()
        
        # Tab 2: Replay Session
        self.replay_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.replay_frame, text="Replay Session")
        self.create_replay_tab()
        
        # Tab 3: View Accounts
        self.accounts_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.accounts_frame, text="Created Accounts")
        self.create_accounts_tab()
        
        # Tab 4: Settings
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Settings")
        self.create_settings_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def create_record_tab(self):
        """Create the recording session tab"""
        # Browser selection
        browser_frame = ttk.LabelFrame(self.record_frame, text="Browser Selection", padding=10)
        browser_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(browser_frame, text="Browser Shortcut:").grid(row=0, column=0, sticky='w', padx=5)
        self.browser_path_var = tk.StringVar()
        self.browser_entry = ttk.Entry(browser_frame, textvariable=self.browser_path_var, width=50)
        self.browser_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(browser_frame, text="Browse...", command=self.browse_browser).grid(row=0, column=2, padx=5)
        
        # URL input
        url_frame = ttk.LabelFrame(self.record_frame, text="Target URL", padding=10)
        url_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(url_frame, text="Website URL:").grid(row=0, column=0, sticky='w', padx=5)
        self.url_var = tk.StringVar()
        url_entry = ttk.Entry(url_frame, textvariable=self.url_var, width=50)
        url_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Recording controls
        control_frame = ttk.LabelFrame(self.record_frame, text="Recording Controls", padding=10)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        self.record_btn = ttk.Button(control_frame, text="Start Recording", command=self.start_recording)
        self.record_btn.pack(side='left', padx=5)
        
        self.stop_btn = ttk.Button(control_frame, text="Stop Recording", command=self.stop_recording, state='disabled')
        self.stop_btn.pack(side='left', padx=5)
        
        ttk.Button(control_frame, text="Test Browser", command=self.test_browser).pack(side='left', padx=5)
        
        # Element type selection
        element_frame = ttk.LabelFrame(self.record_frame, text="Element Type Selection", padding=10)
        element_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(element_frame, text="Select element type, then click on the element in browser:").pack(anchor='w', padx=5)
        
        type_frame = ttk.Frame(element_frame)
        type_frame.pack(fill='x', padx=5, pady=5)
        
        self.element_type_var = tk.StringVar(value="email")
        
        types = [
            ("Email Field", "email"), 
            ("Password Field", "password"), 
            ("Submit Button", "button"), 
            ("CAPTCHA", "captcha"), 
            ("Other", "other")
        ]
        
        for i, (text, value) in enumerate(types):
            ttk.Radiobutton(type_frame, text=text, variable=self.element_type_var, 
                           value=value, command=self.on_element_type_selected).grid(row=i//3, column=i%3, sticky='w', padx=5, pady=2)
        
        # CAPTCHA wait time setting
        captcha_frame = ttk.Frame(element_frame)
        captcha_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(captcha_frame, text="CAPTCHA Wait Time (seconds):").pack(side='left', padx=5)
        self.captcha_wait_time_var = tk.IntVar(value=7)
        captcha_spinbox = ttk.Spinbox(captcha_frame, from_=3, to=30, textvariable=self.captcha_wait_time_var, width=10)
        captcha_spinbox.pack(side='left', padx=5)
        
        # Recording status
        status_frame = ttk.LabelFrame(self.record_frame, text="Recording Status", padding=10)
        status_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.recording_status = tk.StringVar(value="Not recording")
        status_label = ttk.Label(status_frame, textvariable=self.recording_status, font=('Arial', 12))
        status_label.pack(pady=5)
        
        self.click_status = tk.StringVar(value="Select an element type, then click on the element")
        click_label = ttk.Label(status_frame, textvariable=self.click_status, font=('Arial', 10))
        click_label.pack(pady=5)
        
        # Action list
        list_frame = ttk.LabelFrame(status_frame, text="Recorded Actions", padding=5)
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.action_listbox = tk.Listbox(list_frame, height=8)
        self.action_listbox.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.action_listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.action_listbox.config(yscrollcommand=scrollbar.set)
        
        # Save/Load buttons
        file_frame = ttk.Frame(self.record_frame)
        file_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(file_frame, text="Save Session", command=self.save_session).pack(side='left', padx=5)
        ttk.Button(file_frame, text="Load Session", command=self.load_session).pack(side='left', padx=5)
        ttk.Button(file_frame, text="Clear Actions", command=self.clear_actions).pack(side='left', padx=5)
        
    def create_replay_tab(self):
        """Create the replay session tab"""
        # Session selection
        session_frame = ttk.LabelFrame(self.replay_frame, text="Session Selection", padding=10)
        session_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(session_frame, text="Loaded Session:").grid(row=0, column=0, sticky='w', padx=5)
        self.session_label = ttk.Label(session_frame, text="No session loaded")
        self.session_label.grid(row=0, column=1, sticky='w', padx=5)
        
        ttk.Button(session_frame, text="Load Session", command=self.load_session_for_replay).grid(row=0, column=2, padx=5)
        
        # Account creation settings
        settings_frame = ttk.LabelFrame(self.replay_frame, text="Account Creation Settings", padding=10)
        settings_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(settings_frame, text="Number of Accounts:").grid(row=0, column=0, sticky='w', padx=5)
        self.account_count_var = tk.IntVar(value=1)
        account_spinbox = ttk.Spinbox(settings_frame, from_=1, to=100, textvariable=self.account_count_var, width=10)
        account_spinbox.grid(row=0, column=1, sticky='w', padx=5)
        
        # Email domain settings
        ttk.Label(settings_frame, text="Email Domain:").grid(row=1, column=0, sticky='w', padx=5)
        self.domain_var = tk.StringVar(value="gmail.com")
        domain_combo = ttk.Combobox(settings_frame, textvariable=self.domain_var, width=20)
        domain_combo['values'] = ('gmail.com', 'yahoo.com', 'outlook.com', 'example.com')
        domain_combo.grid(row=1, column=1, sticky='w', padx=5)
        
        # Repeat operation settings
        repeat_frame = ttk.LabelFrame(self.replay_frame, text="Repeat Operation Settings", padding=10)
        repeat_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(repeat_frame, text="Repeat Count:").grid(row=0, column=0, sticky='w', padx=5)
        self.repeat_count_var = tk.IntVar(value=1)
        repeat_spinbox = ttk.Spinbox(repeat_frame, from_=1, to=100, textvariable=self.repeat_count_var, width=10)
        repeat_spinbox.grid(row=0, column=1, sticky='w', padx=5)
        
        ttk.Label(repeat_frame, text="Wait Between Iterations (seconds):").grid(row=1, column=0, sticky='w', padx=5)
        self.iteration_wait_var = tk.IntVar(value=3)
        iteration_spinbox = ttk.Spinbox(repeat_frame, from_=1, to=30, textvariable=self.iteration_wait_var, width=10)
        iteration_spinbox.grid(row=1, column=1, sticky='w', padx=5)
        
        # Replay controls
        control_frame = ttk.LabelFrame(self.replay_frame, text="Replay Controls", padding=10)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        self.replay_btn = ttk.Button(control_frame, text="Start Replay", command=self.start_replay)
        self.replay_btn.pack(side='left', padx=5)
        
        self.stop_replay_btn = ttk.Button(control_frame, text="Stop Replay", command=self.stop_replay, state='disabled')
        self.stop_replay_btn.pack(side='left', padx=5)
        
        # Preview actions
        preview_frame = ttk.LabelFrame(self.replay_frame, text="Action Preview", padding=10)
        preview_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.preview_listbox = tk.Listbox(preview_frame, height=10)
        self.preview_listbox.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Add scrollbar
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient='vertical', command=self.preview_listbox.yview)
        preview_scrollbar.pack(side='right', fill='y')
        self.preview_listbox.config(yscrollcommand=preview_scrollbar.set)
        
    def create_accounts_tab(self):
        """Create the view accounts tab"""
        # Account list
        list_frame = ttk.LabelFrame(self.accounts_frame, text="Created Accounts", padding=10)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.accounts_text = scrolledtext.ScrolledText(list_frame, height=20, width=80)
        self.accounts_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Control buttons
        button_frame = ttk.Frame(self.accounts_frame)
        button_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(button_frame, text="Refresh", command=self.refresh_accounts).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Clear All", command=self.clear_accounts).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Export to CSV", command=self.export_accounts).pack(side='left', padx=5)
        
        # Load accounts on startup
        self.refresh_accounts()
        
    def create_settings_tab(self):
        """Create the settings tab"""
        # General settings
        general_frame = ttk.LabelFrame(self.settings_frame, text="General Settings", padding=10)
        general_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(general_frame, text="Default Wait Time (seconds):").grid(row=0, column=0, sticky='w', padx=5)
        self.wait_time_var = tk.DoubleVar(value=1.0)
        wait_spinbox = ttk.Spinbox(general_frame, from_=0.1, to=5.0, increment=0.1, 
                                   textvariable=self.wait_time_var, width=10)
        wait_spinbox.grid(row=0, column=1, sticky='w', padx=5)
        
        ttk.Label(general_frame, text="Manual CAPTCHA Wait Time (seconds):").grid(row=1, column=0, sticky='w', padx=5)
        self.captcha_wait_var = tk.IntVar(value=15)
        captcha_spinbox = ttk.Spinbox(general_frame, from_=5, to=60, textvariable=self.captcha_wait_var, width=10)
        captcha_spinbox.grid(row=1, column=1, sticky='w', padx=5)
        
        # Save settings button
        ttk.Button(general_frame, text="Save Settings", command=self.save_settings).grid(row=2, column=0, columnspan=2, pady=10)
        
    def browse_browser(self):
        """Browse for browser shortcut"""
        filename = filedialog.askopenfilename(
            title="Select Browser Shortcut",
            filetypes=[("Shortcut files", "*.lnk"), ("All files", "*.*")]
        )
        if filename:
            self.browser_path_var.set(filename)
            self.browser_shortcut = filename
            
    def test_browser(self):
        """Test browser opening with URL"""
        if not self.browser_path_var.get():
            messagebox.showerror("Error", "Please select a browser shortcut first")
            return
            
        if not self.url_var.get():
            messagebox.showerror("Error", "Please enter a target URL first")
            return
            
        try:
            subprocess.Popen(f'cmd /c start "" "{self.browser_path_var.get()}"', shell=True)
            time.sleep(3)
            pyautogui.hotkey('ctrl', 'l')
            time.sleep(0.5)
            pyautogui.write(self.url_var.get())
            pyautogui.press('enter')
            self.status_var.set("Browser opened successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open browser: {str(e)}")
            
    def on_element_type_selected(self):
        """Handle element type selection"""
        self.selected_element_type = self.element_type_var.get()
        if self.recording:
            self.waiting_for_click = True
            self.click_status.set(f"Selected {self.selected_element_type}. Now click on the element in the browser.")
            self.status_var.set(f"Waiting for click on {self.selected_element_type}")
            
    def on_click(self, x, y, button, pressed):
        """Handle mouse click events"""
        if pressed and self.recording and self.waiting_for_click:
            # Add action
            action = {
                'type': 'click',
                'coordinates': (x, y),
                'element_type': self.selected_element_type,
                'timestamp': time.time()
            }
            
            # Add special handling for CAPTCHA
            if self.selected_element_type == 'captcha':
                action['wait_before'] = self.captcha_wait_time_var.get()
                action['wait_after'] = 7  # Always wait 7 seconds after clicking CAPTCHA
            
            self.actions.append(action)
            
            # Update UI
            wait_text = f" (wait {action.get('wait_before', 0)}s before, {action.get('wait_after', 0)}s after)" if self.selected_element_type == 'captcha' else ""
            action_text = f"{self.selected_element_type} at ({x}, {y}){wait_text}"
            self.action_listbox.insert(tk.END, action_text)
            self.action_listbox.see(tk.END)
            
            # Reset waiting state
            self.waiting_for_click = False
            self.click_status.set(f"Recorded {self.selected_element_type}. Select next element type.")
            self.status_var.set(f"Recorded {self.selected_element_type} at ({x}, {y})")
            
    def start_recording(self):
        """Start recording user actions"""
        if not self.browser_path_var.get():
            messagebox.showerror("Error", "Please select a browser shortcut first")
            return
            
        if not self.url_var.get():
            messagebox.showerror("Error", "Please enter a target URL first")
            return
            
        # Clear previous actions
        self.actions = []
        self.action_listbox.delete(0, tk.END)
        
        # Open browser
        self.test_browser()
        
        # Update UI
        self.record_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.recording = True
        self.recording_status.set("Recording in progress...")
        self.click_status.set("Select an element type, then click on the element")
        self.status_var.set("Recording started. Select element type and click on elements.")
        
        # Start mouse listener
        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.mouse_listener.start()
        
    def stop_recording(self):
        """Stop recording user actions"""
        self.recording = False
        self.waiting_for_click = False
        
        # Stop mouse listener
        if self.mouse_listener:
            self.mouse_listener.stop()
            self.mouse_listener = None
            
        # Update UI
        self.record_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.recording_status.set("Recording stopped")
        self.click_status.set(f"Recording stopped. {len(self.actions)} actions recorded.")
        self.status_var.set(f"Recording stopped. {len(self.actions)} actions recorded.")
        
    def clear_actions(self):
        """Clear all recorded actions"""
        self.actions = []
        self.action_listbox.delete(0, tk.END)
        self.status_var.set("Actions cleared")
        
    def save_session(self):
        """Save recorded session to file"""
        if not self.actions:
            messagebox.showerror("Error", "No actions to save")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                session_data = {
                    'browser_shortcut': self.browser_path_var.get(),
                    'target_url': self.url_var.get(),
                    'actions': self.actions,
                    'created_at': datetime.now().isoformat()
                }
                
                with open(filename, 'w') as f:
                    json.dump(session_data, f, indent=2)
                    
                messagebox.showinfo("Success", f"Session saved to {filename}")
                self.status_var.set(f"Session saved: {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save session: {str(e)}")
                
    def load_session(self):
        """Load a saved session"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    session_data = json.load(f)
                    
                self.browser_path_var.set(session_data.get('browser_shortcut', ''))
                self.url_var.set(session_data.get('target_url', ''))
                self.actions = session_data.get('actions', [])
                
                # Update action list
                self.action_listbox.delete(0, tk.END)
                for action in self.actions:
                    element_type = action.get('element_type', 'unknown')
                    x, y = action.get('coordinates', (0, 0))
                    wait_before = action.get('wait_before', 0)
                    wait_after = action.get('wait_after', 0)
                    wait_text = ""
                    if element_type == 'captcha':
                        wait_text = f" (wait {wait_before}s before, {wait_after}s after)"
                    action_text = f"{element_type} at ({x}, {y}){wait_text}"
                    self.action_listbox.insert(tk.END, action_text)
                
                messagebox.showinfo("Success", f"Session loaded: {len(self.actions)} actions")
                self.status_var.set(f"Session loaded: {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load session: {str(e)}")
                
    def load_session_for_replay(self):
        """Load a session for replay"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    session_data = json.load(f)
                    
                self.browser_shortcut = session_data.get('browser_shortcut', '')
                self.target_url = session_data.get('target_url', '')
                self.actions = session_data.get('actions', [])
                
                # Update UI
                self.session_label.config(text=os.path.basename(filename))
                self.preview_listbox.delete(0, tk.END)
                
                for i, action in enumerate(self.actions):
                    element_type = action.get('element_type', 'unknown')
                    x, y = action.get('coordinates', (0, 0))
                    wait_before = action.get('wait_before', 0)
                    wait_after = action.get('wait_after', 0)
                    wait_text = ""
                    if element_type == 'captcha':
                        wait_text = f" (wait {wait_before}s before, {wait_after}s after)"
                    action_text = f"{i+1}. {element_type} at ({x}, {y}){wait_text}"
                    self.preview_listbox.insert(tk.END, action_text)
                    
                messagebox.showinfo("Success", f"Session loaded: {len(self.actions)} actions")
                self.status_var.set(f"Session loaded for replay: {os.path.basename(filename)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load session: {str(e)}")
                
    def start_replay(self):
        """Start replaying the recorded session"""
        if not self.actions:
            messagebox.showerror("Error", "No session loaded for replay")
            return
            
        if not self.browser_shortcut:
            messagebox.showerror("Error", "No browser shortcut in session")
            return
            
        # Reset stop flag
        self.stop_replay_flag = False
        
        # Update UI
        self.replay_btn.config(state='disabled')
        self.stop_replay_btn.config(state='normal')
        
        # Start replay in a separate thread
        self.replay_thread = threading.Thread(target=self.replay_session)
        self.replay_thread.daemon = True
        self.replay_thread.start()
        
    def stop_replay(self):
        """Stop the replay process"""
        self.stop_replay_flag = True
        self.status_var.set("Stopping replay...")
        
        # Update UI
        self.replay_btn.config(state='normal')
        self.stop_replay_btn.config(state='disabled')
        
    def replay_session(self):
        """Replay the recorded session with random data"""
        num_accounts = self.account_count_var.get()
        repeat_count = self.repeat_count_var.get()
        iteration_wait = self.iteration_wait_var.get()
        
        # Open browser once at the beginning
        try:
            subprocess.Popen(f'cmd /c start "" "{self.browser_shortcut}"', shell=True)
            time.sleep(3)
            
            # Navigate to URL
            pyautogui.hotkey('ctrl', 'l')
            time.sleep(0.5)
            pyautogui.write(self.target_url)
            pyautogui.press('enter')
            time.sleep(3)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to open browser: {str(e)}"))
            return
        
        for repeat in range(repeat_count):
            if self.stop_replay_flag:  # Check if stop was requested
                break
                
            self.root.after(0, lambda r=repeat: self.status_var.set(f"Repeat iteration {r+1}/{repeat_count}"))
            
            for i in range(num_accounts):
                if self.stop_replay_flag:  # Check if stop was requested
                    break
                    
                self.root.after(0, lambda r=repeat, a=i: self.status_var.set(f"Repeat {r+1}/{repeat_count}, Account {a+1}/{num_accounts}"))
                
                # Generate random credentials
                email = self.generate_random_email()
                password = self.generate_random_password()
                
                # Replay actions
                email_filled = False
                password_filled = False
                
                for action in self.actions:
                    if self.stop_replay_flag:  # Check if stop was requested
                        break
                        
                    x, y = action.get('coordinates', (0, 0))
                    element_type = action.get('element_type', 'other')
                    
                    if element_type == 'email':
                        pyautogui.click(x, y)
                        time.sleep(0.5)
                        pyautogui.hotkey('ctrl', 'a')
                        pyautogui.write(email)
                        email_filled = True
                    elif element_type == 'password':
                        pyautogui.click(x, y)
                        time.sleep(0.5)
                        pyautogui.hotkey('ctrl', 'a')
                        pyautogui.write(password)
                        password_filled = True
                    elif element_type == 'captcha':
                        # Wait before clicking on CAPTCHA to ensure it loads
                        wait_before = action.get('wait_before', 7)
                        self.root.after(0, lambda wt=wait_before: self.status_var.set(f"Waiting {wt} seconds for CAPTCHA to load..."))
                        time.sleep(wait_before)
                        
                        pyautogui.click(x, y)
                        
                        # Wait after clicking CAPTCHA
                        wait_after = action.get('wait_after', 7)
                        self.root.after(0, lambda wt=wait_after: self.status_var.set(f"Waiting {wt} seconds after clicking CAPTCHA..."))
                        time.sleep(wait_after)
                    elif element_type == 'button':
                        pyautogui.click(x, y)
                        time.sleep(self.wait_time_var.get())
                    else:
                        # Default action - just click
                        pyautogui.click(x, y)
                        time.sleep(self.wait_time_var.get())
                        
                # Save account if both email and password were filled
                if email_filled and password_filled:
                    self.save_account(email, password)
                    # Update UI to show the account was saved
                    self.root.after(0, self.refresh_accounts)
                        
                # Wait between accounts
                if i < num_accounts - 1 and not self.stop_replay_flag:
                    time.sleep(2)
            
            # Refresh tab and wait between iterations
            if repeat < repeat_count - 1 and not self.stop_replay_flag:
                self.root.after(0, lambda: self.status_var.set(f"Refreshing tab and waiting {iteration_wait} seconds..."))
                # Refresh the current tab instead of opening a new one
                pyautogui.hotkey('ctrl', 'r')  # Refresh the current tab
                time.sleep(iteration_wait)
                
        # Update UI
        if not self.stop_replay_flag:
            self.root.after(0, lambda: self.status_var.set("Replay completed"))
        else:
            self.root.after(0, lambda: self.status_var.set("Replay stopped by user"))
            
        self.root.after(0, lambda: self.replay_btn.config(state='normal'))
        self.root.after(0, lambda: self.stop_replay_btn.config(state='disabled'))
        
    def generate_random_email(self):
        """Generate a random email address"""
        username = ''.join(random.choices(string.ascii_lowercase + string.digits, k=10))
        domain = self.domain_var.get()
        return f"{username}@{domain}"
        
    def generate_random_password(self):
        """Generate a random password"""
        chars = string.ascii_letters + string.digits + "!@#$%^&*"
        return ''.join(random.choices(chars, k=12))
        
    def save_account(self, email, password):
        """Save account credentials to file"""
        filename = "created_accounts.txt"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        with open(filename, "a") as f:
            f.write(f"[{timestamp}] Email: {email}, Password: {password}\n")
            
    def refresh_accounts(self):
        """Refresh the accounts list"""
        self.accounts_text.delete(1.0, tk.END)
        
        if os.path.exists("created_accounts.txt"):
            with open("created_accounts.txt", "r") as f:
                content = f.read()
                self.accounts_text.insert(1.0, content)
        else:
            self.accounts_text.insert(1.0, "No accounts created yet.")
            
    def clear_accounts(self):
        """Clear all accounts"""
        if messagebox.askyesno("Confirm", "Are you sure you want to clear all accounts?"):
            if os.path.exists("created_accounts.txt"):
                os.remove("created_accounts.txt")
            self.refresh_accounts()
            self.status_var.set("Accounts cleared")
            
    def export_accounts(self):
        """Export accounts to CSV"""
        if not os.path.exists("created_accounts.txt"):
            messagebox.showerror("Error", "No accounts to export")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open("created_accounts.txt", "r") as infile:
                    with open(filename, "w") as outfile:
                        outfile.write("Timestamp,Email,Password\n")
                        for line in infile:
                            if line.strip():
                                # Parse the line and format as CSV
                                parts = line.strip().split("Email: ")
                                if len(parts) > 1:
                                    timestamp = parts[0].strip("[]")
                                    email_pass = parts[1].split(", Password: ")
                                    if len(email_pass) > 1:
                                        email = email_pass[0]
                                        password = email_pass[1]
                                        outfile.write(f"{timestamp},{email},{password}\n")
                                        
                messagebox.showinfo("Success", f"Accounts exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export accounts: {str(e)}")
            
    def save_settings(self):
        """Save application settings"""
        settings = {
            'browser_shortcut': self.browser_path_var.get(),
            'target_url': self.url_var.get(),
            'wait_time': self.wait_time_var.get(),
            'captcha_wait': self.captcha_wait_var.get(),
            'email_domain': self.domain_var.get(),
            'captcha_wait_time': self.captcha_wait_time_var.get(),
            'repeat_count': self.repeat_count_var.get(),
            'iteration_wait': self.iteration_wait_var.get()
        }
        
        try:
            with open("automation_settings.json", "w") as f:
                json.dump(settings, f, indent=2)
                
            messagebox.showinfo("Success", "Settings saved")
            self.status_var.set("Settings saved")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
        
    def load_settings(self):
        """Load application settings"""
        if os.path.exists("automation_settings.json"):
            try:
                with open("automation_settings.json", "r") as f:
                    settings = json.load(f)
                    
                self.browser_path_var.set(settings.get('browser_shortcut', ''))
                self.url_var.set(settings.get('target_url', ''))
                self.wait_time_var.set(settings.get('wait_time', 1.0))
                self.captcha_wait_var.set(settings.get('captcha_wait', 15))
                self.domain_var.set(settings.get('email_domain', 'gmail.com'))
                self.captcha_wait_time_var.set(settings.get('captcha_wait_time', 7))
                self.repeat_count_var.set(settings.get('repeat_count', 1))
                self.iteration_wait_var.set(settings.get('iteration_wait', 3))
            except:
                pass
                
    def run(self):
        """Run the application"""
        self.root.mainloop()

if __name__ == "__main__":
    # Install pynput if not available
    try:
        import pynput
    except ImportError:
        print("Installing required package: pynput")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pynput"])
        import pynput
        
    app = AutomationRecorder()
    app.run()