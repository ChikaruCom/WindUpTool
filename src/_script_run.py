import tkinter as tk
from tkinter import scrolledtext, messagebox
import sys
import os
import subprocess
import psutil
from datetime import datetime
import pyperclip

def get_python_version():
    """Get the current Python version."""
    return sys.version

def get_python_path():
    """Get the Python executable path."""
    return sys.executable

def get_installed_modules():
    """Get the list of installed modules."""
    result = subprocess.run([sys.executable, '-m', 'pip', 'freeze'], stdout=subprocess.PIPE, text=True)
    return result.stdout

def get_cpu_info():
    """Get the CPU information."""
    return f"CPU: {psutil.cpu_count(logical=True)} cores"

def get_memory_info():
    """Get the memory (RAM) information."""
    memory = psutil.virtual_memory()
    return f"Total Memory: {round(memory.total / (1024 ** 3), 2)} GB"

def get_system_info():
    """Get complete system information, including CPU, memory, and Python details."""
    python_version = get_python_version()
    python_path = get_python_path()
    installed_modules = get_installed_modules()
    cpu_info = get_cpu_info()
    memory_info = get_memory_info()
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    return (f"System Information as of {current_time}:\n\n"
            f"Python Version:\n{python_version}\n\n"
            f"Python Executable Path:\n{python_path}\n\n"
            f"CPU Information:\n{cpu_info}\n\n"
            f"Memory Information:\n{memory_info}\n\n"
            "Installed Modules:\n"
            f"{installed_modules}")

def display_info():
    """Display the gathered system information in the GUI."""
    system_info = get_system_info()
    text_area.delete(1.0, tk.END)  # Clear the text area
    text_area.insert(tk.END, system_info)

def copy_to_clipboard():
    """Copy the current contents of the text area to the clipboard."""
    text = text_area.get(1.0, tk.END)
    pyperclip.copy(text)
    messagebox.showinfo("Copy to Clipboard", "Information copied to clipboard successfully!")

# Create the main window
root = tk.Tk()
root.title("System and Python Environment Info")

# Create a text area to display the information
text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=30)
text_area.pack(padx=10, pady=10)

# Create a button to refresh and display the Python environment information
refresh_button = tk.Button(root, text="Refresh Info", command=display_info)
refresh_button.pack(pady=10)

# Create a button to copy the displayed information to the clipboard
copy_button = tk.Button(root, text="Copy to Clipboard", command=copy_to_clipboard)
copy_button.pack(pady=10)

# Initially display the info
display_info()

# Start the GUI event loop
root.mainloop()
