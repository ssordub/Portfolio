import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import subprocess
import tempfile
import os
from datetime import datetime
import json
from operator import itemgetter
import pathlib
from typing import Optional


class FileSystemNavigator:
    """Handles directory traversal and tree population to view the file system"""
    
    def __init__(self, treeview):
        self.tree = treeview
        self.hidden_files = False
        self.populated_nodes = set()  # Track which nodes have been populated
        self.current_path = None  # Track current directory path
        self.navigation_history = []  # Track directory history for back navigation
        
    def populate_root(self, path: str) -> None:
        """Populate the root directory and its immediate children"""
        # Clear existing items and populated nodes tracking
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.populated_nodes.clear()
        
        # Reset navigation state
        self.current_path = path
        self.navigation_history = [path]
            
        # Add root node
        root_node = self.tree.insert('', 'end', text=path, values=(path, ''),
                                   open=True)
        
        try:
            # Add back navigation item if not at root
            if self.current_path != path:
                self.tree.insert('', 'end', text='..',
                               values=(os.path.dirname(self.current_path), 'Directory'),
                               tags=('directory',))
            
            # Populate first level
            self._populate_directory(root_node, path)
            self.populated_nodes.add(root_node)
        except PermissionError:
            self.tree.insert(root_node, 'end', text='Access Denied',
                           values=('Access Denied', ''))
        except Exception as e:
            self.tree.insert(root_node, 'end', text=f'Error: {str(e)}',
                           values=(f'Error: {str(e)}', ''))
    
    def _populate_directory(self, parent: str, path: str) -> None:
        """Recursively populate directory structure under given parent node"""
        # Clear existing children
        for child in self.tree.get_children():
            self.tree.delete(child)
        
        try:
            has_entries = False
            
            # Add back navigation if not at root
            if path != self.navigation_history[0]:
                self.tree.insert('', 'end', text='..',
                               values=(os.path.dirname(path), 'Directory'),
                               tags=('directory',))
            
            self.tree.tag_configure('oddrow', background='lightgray')
            self.tree.tag_configure('evenrow', background='white')
            # List directory contents
            with os.scandir(path) as entries:
                sorted_entries = sorted(entries, 
                                     key=lambda e: (not e.is_dir(), e.name.lower()))
                count = 0
                for entry in sorted_entries:
                    if not self.hidden_files and entry.name.startswith('.'):
                        continue
                        
                    has_entries = True
                    is_directory = entry.is_dir()
                    full_path = entry.path
                    
                    # Create node
                    if count % 2 == 0:
                        self.tree.insert(
                            '', 'end',
                            text=entry.name,
                            values=(full_path, 'Directory' if is_directory else 'File'),
                            tags=('directory' if is_directory else 'file', 'evenrow')
                        )
                    else:
                        self.tree.insert(
                            '', 'end',
                            text=entry.name,
                            values=(full_path, 'Directory' if is_directory else 'File'),
                            tags=('directory' if is_directory else 'file', 'oddrow')
                        )
                    count += 1

            # Mark as populated
            self.populated_nodes.add(parent)
            self.current_path = path
            self.tree.tag_configure('colored', background='lightblue')
            # If directory is empty, add indicator
            if not has_entries:
                self.tree.insert('', 'end', text='(Empty)', tags=("colored",),
                               values=('', ''))
                
        except PermissionError:
            self.tree.insert('', 'end', text='Access Denied',
                           values=('Access Denied', ''))
        except Exception as e:
            self.tree.insert('', 'end', text=f'Error: {str(e)}',
                           values=(f'Error: {str(e)}', ''))

    def navigate_to_directory(self, path: str) -> None:
        """Navigate to a specific directory"""
        if os.path.isdir(path):
            # Add current path to history if it's a forward navigation
            if path != os.path.dirname(self.current_path):
                self.navigation_history.append(path)
            self._populate_directory('', path)

    def navigate_back(self) -> None:
        """Navigate to the previous directory"""
        if len(self.navigation_history) > 1:
            self.navigation_history.pop()  # Remove current directory
            previous_path = self.navigation_history[-1]  # Get previous directory
            self._populate_directory('', previous_path)

class TreeViewContextMenu:
    """A context menu manager that can be attached to existing TreeView widgets"""
    
    def __init__(self, treeview):
        """
        Initialize the context menu manager
        
        Parameters:
        treeview: The existing ttk.Treeview widget to attach the context menu to
        """
        self.tree = treeview
        self.menu = tk.Menu(self.tree, tearoff=0)
        self.actions = {}
        
        # Bind right-click event
        self.tree.bind('<Button-3>', self._show_popup)
        
    def add_command(self, label, callback, **kwargs):
        """
        Add a new command to the context menu
        
        Parameters:
        label: The text to show in the menu
        callback: The function to call when the item is selected
        **kwargs: Additional arguments to pass to menu.add_command()
        """
        self.actions[label] = callback
        self.menu.add_command(label=label, command=callback, **kwargs)
        
    def add_separator(self):
        """Add a separator line to the context menu"""
        self.menu.add_separator()
        
    def remove_command(self, label):
        """
        Remove a command from the context menu
        
        Parameters:
        label: The label of the command to remove
        """
        if label in self.actions:
            index = self._find_menu_index(label)
            if index is not None:
                self.menu.delete(index)
                del self.actions[label]
                
    def _find_menu_index(self, label):
        """Find the index of a menu item by its label"""
        last_index = self.menu.index('end')
        if last_index is None:
            return None
            
        for i in range(last_index + 1):
            try:
                if self.menu.entrycget(i, 'label') == label:
                    return i
            except tk.TclError:
                continue
        return None
                
    def _show_popup(self, event):
        """Display the context menu"""
        # Get the item under cursor
        item = self.tree.identify_row(event.y)
        
        if item:
            # Select the item before showing menu
            self.tree.selection_set(item)
            try:
                self.menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.menu.grab_release()
                
    def enable_command(self, label):
        """Enable a command in the context menu"""
        index = self._find_menu_index(label)
        if index is not None:
            self.menu.entryconfig(index, state='normal')
            
    def disable_command(self, label):
        """Disable a command in the context menu"""
        index = self._find_menu_index(label)
        if index is not None:
            self.menu.entryconfig(index, state='disabled')
            
    def clear_menu(self):
        """Remove all items from the context menu"""
        self.menu.delete(0, 'end')
        self.actions.clear()


class SystemUtilityGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("System Configuration Utility")
        
        # Set window properties
        self.root.geometry("1000x800")
        self.root.minsize(800, 600)
        
        # Initialize sorting state for hardware devices
        self.sort_column = "Name"
        self.sort_reverse = False
        
        # Initialize the default hardware scan script (hidden from UI)
        self.hardware_scan_script = '''function Get-HardwareDevices {
    Get-WmiObject -Class Win32_PnPEntity | 
        Select-Object Name, Manufacturer, DeviceID |
        Format-Table -AutoSize
}
Get-HardwareDevices'''
        
        # Set up styling
        self.setup_styles()
        
        # Create main container
        self.main_frame = ttk.Frame(self.root, style="Main.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create GUI sections
        self.create_header()
        self.create_notebook()
        self.create_status_bar()
        
    def setup_styles(self):
        """Configure the GUI styling"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colors
        bg_color = "#f0f0f0"
        accent_color = "#2c3e50"
        button_color = "lightgray"
        
        # Configure widget styles
        style.configure("Main.TFrame", background=bg_color)
        style.configure("Header.TLabel",
                       font=("Helvetica", 16, "bold"),
                       background=bg_color,
                       foreground=accent_color,
                       padding=10)
        style.configure("Section.TLabelframe",
                       background=bg_color,
                       padding=10)
        style.configure("Action.TButton",
                       font=("Helvetica", 10, "bold"),
                       background=button_color,
                       padding=10)
        style.configure("Treeview",
                       background="white",
                       foreground="black",
                       fieldbackground="white",
                       font=("Helvetica", 9))
        
    def create_header(self):
        """Create the header section"""
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title = ttk.Label(header_frame,
                         text="System Configuration Utility",
                         style="Header.TLabel")
        title.pack(side=tk.LEFT)
        
        
        
    def create_notebook(self):
        """Create tabbed interface"""
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create tabs
        self.create_hardware_tab()
        self.create_network_tab()
        self.create_system_tab()
        self.create_files_tab()
        # self.create_script_tab()
        
    def create_hardware_tab(self):
        """Create hardware devices tab with consistent layout management"""
        # Create main container frame
        hardware_frame = ttk.Frame(self.notebook)
        self.notebook.add(hardware_frame, text="Hardware Devices")
        
        # Create button frame at the top
        button_frame = ttk.Frame(hardware_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        context_menu = TreeViewContextMenu(hardware_frame)
        context_menu.add_command("Refresh", refresh)
        # Add scan and export buttons
        scan_button = ttk.Button(
            button_frame,
            text="Scan Hardware",
            style="Action.TButton",
            command=self.scan_hardware
        )
        scan_button.pack(side=tk.LEFT, padx=5)
        
        export_button = ttk.Button(
            button_frame,
            text="Export Hardware List",
            style="Action.TButton",
            command=self.export_hardware
        )
        export_button.pack(side=tk.LEFT, padx=5)
        
        # Store button references
        self.scan_button = scan_button
        self.export_button = export_button
        
        # Create frame for treeview and scrollbars
        tree_container = ttk.Frame(hardware_frame)
        tree_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create and configure treeview
        columns = ("Name", "Manufacturer", "DeviceID")
        self.tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings"
        )
        context_menu = TreeViewContextMenu(self.tree)
        context_menu.add_command("Refresh", self.scan_hardware)
        # Configure columns
        for col in columns:
            self.tree.heading(col, 
                            text=col,
                            command=lambda c=col: self.sort_treeview(c))
            self.tree.column(col, minwidth=100, width=200)
        
        # Create scrollbars
        vsb = ttk.Scrollbar(
            tree_container,
            orient="vertical",
            command=self.tree.yview
        )
        
        # Configure treeview scrolling
        self.tree.configure(
            yscrollcommand=vsb.set,
            
        )
        
        # Pack the treeview and scrollbars
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        
    def create_network_tab(self):
        """Create network configuration tab"""
        network_frame = ttk.Frame(self.notebook)
        self.notebook.add(network_frame, text="Network Setup")
        
        # DHCP Section
        dhcp_frame = ttk.LabelFrame(network_frame, text="DHCP Configuration", style="Section.TLabelframe")
        dhcp_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(dhcp_frame, 
                  text="Enable DHCP",
                  command=self.enable_dhcp,
                  style="Action.TButton").pack(pady=5)
        
        # Static IP Section
        static_frame = ttk.LabelFrame(network_frame, text="Static IP Configuration", style="Section.TLabelframe")
        static_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # IP Address
        ip_frame = ttk.Frame(static_frame)
        ip_frame.pack(fill=tk.X, pady=2)
        ttk.Label(ip_frame, text="IP Address:").pack(side=tk.LEFT)
        self.ip_entry = ttk.Entry(ip_frame)
        self.ip_entry.pack(side=tk.LEFT, padx=5)
        
        # Subnet Mask
        subnet_frame = ttk.Frame(static_frame)
        subnet_frame.pack(fill=tk.X, pady=2)
        ttk.Label(subnet_frame, text="Subnet Mask:").pack(side=tk.LEFT)
        self.subnet_entry = ttk.Entry(subnet_frame)
        self.subnet_entry.pack(side=tk.LEFT, padx=5)
        
        # Gateway
        gateway_frame = ttk.Frame(static_frame)
        gateway_frame.pack(fill=tk.X, pady=2)
        ttk.Label(gateway_frame, text="Gateway:").pack(side=tk.LEFT)
        self.gateway_entry = ttk.Entry(gateway_frame)
        self.gateway_entry.pack(side=tk.LEFT, padx=5)
        
        # DNS Servers
        dns_frame1 = ttk.Frame(static_frame)
        dns_frame1.pack(fill=tk.X, pady=2)
        ttk.Label(dns_frame1, text="DNS Server 1:").pack(side=tk.LEFT)
        self.dns_entry1 = ttk.Entry(dns_frame1)
        self.dns_entry1.pack(side=tk.LEFT, padx=5)

        dns_frame2 = ttk.Frame(static_frame)
        dns_frame2.pack(fill=tk.X, pady=2)
        ttk.Label(dns_frame2, text="DNS Server 2:").pack(side=tk.LEFT)
        self.dns_entry2 = ttk.Entry(dns_frame2)
        self.dns_entry2.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(static_frame,
                  text="Apply Static IP",
                  command=self.apply_static_ip,
                  style="Action.TButton").pack(pady=5)
        
    def create_system_tab(self):
        """Create system configuration tab"""
        system_frame = ttk.Frame(self.notebook)
        self.notebook.add(system_frame, text="System Setup")
        
        # Computer Name Section
        name_frame = ttk.LabelFrame(system_frame, text="Computer Name", style="Section.TLabelframe")
        name_frame.pack(fill=tk.X, padx=5, pady=5)
        
        name_input_frame = ttk.Frame(name_frame)
        name_input_frame.pack(fill=tk.X, pady=2)
        ttk.Label(name_input_frame, text="New Name:").pack(side=tk.LEFT)
        self.computer_name_entry = ttk.Entry(name_input_frame)
        self.computer_name_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(name_frame,
                  text="Rename Computer",
                  command=self.rename_computer,
                  style="Action.TButton").pack(pady=5)
        
        # Time Zone Section
        timezone_frame = ttk.LabelFrame(system_frame, text="Time Zone", style="Section.TLabelframe")
        timezone_frame.pack(fill=tk.X, padx=5, pady=5)
        
        timezone_input_frame = ttk.Frame(timezone_frame)
        timezone_input_frame.pack(fill=tk.X, pady=2)
        ttk.Label(timezone_input_frame, text="Time Zone:").pack(side=tk.LEFT)
        
        time_zones = self.get_available_timezones()
        self.timezone_combo = ttk.Combobox(timezone_input_frame, values=time_zones)
        self.timezone_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Button(timezone_frame,
                  text="Set Time Zone",
                  command=self.set_timezone,
                  style="Action.TButton").pack(pady=5)
        
        # Windows Activation Section
        activation_frame = ttk.LabelFrame(system_frame, text="Windows Activation", style="Section.TLabelframe")
        activation_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(activation_frame,
                  text="Check/Apply Activation",
                  command=self.check_activation,
                  style="Action.TButton").pack(pady=5)
        

    def create_files_tab(self):
        """Create the file transfer tab with drive selection and dual directory trees"""
        files_frame = ttk.Frame(self.notebook)
        self.notebook.add(files_frame, text="Transfer Files")
        
        # Drive A Frame with drive selector
        drive_a_frame = ttk.LabelFrame(files_frame, text="Source Drive", padding="5 5 5 5")
        drive_a_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Drive selector A
        drive_selector_a = ttk.Frame(drive_a_frame)
        drive_selector_a.pack(fill=tk.X, pady=2)
        ttk.Label(drive_selector_a, text="Select Drive:").pack(side=tk.LEFT, padx=5)
        self.drive_combo_a = ttk.Combobox(drive_selector_a, state="readonly", width=5)
        self.drive_combo_a.pack(side=tk.LEFT, padx=5)
        
        # Create container for tree and scrollbars A
        tree_container_a = ttk.Frame(drive_a_frame)
        tree_container_a.pack(fill=tk.BOTH, expand=True)
        
        # Button row (middle)
        button_frame = ttk.Frame(files_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=2)

        copy_button = ttk.Button(button_frame,
                        text="Copy File",
                        command=self.copy_file,
                        style="Action.TButton")
        copy_button.pack(side=tk.LEFT, pady=5)

        cut_button = ttk.Button(button_frame,
                            text="Cut/Paste File",
                            command=self.cut_paste_file,
                            style="Action.TButton")
        cut_button.pack(side=tk.RIGHT, pady=5)
        
        # Drive B Frame with drive selector
        drive_b_frame = ttk.LabelFrame(files_frame, text="Destination Drive", padding="5 5 5 5")
        drive_b_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Drive selector B
        drive_selector_b = ttk.Frame(drive_b_frame)
        drive_selector_b.pack(fill=tk.X, pady=2)
        ttk.Label(drive_selector_b, text="Select Drive:").pack(side=tk.LEFT, padx=5)
        self.drive_combo_b = ttk.Combobox(drive_selector_b, state="readonly", width=5)
        self.drive_combo_b.pack(side=tk.LEFT, padx=5)
        
        # Create container for tree and scrollbars B
        tree_container_b = ttk.Frame(drive_b_frame)
        tree_container_b.pack(fill=tk.BOTH, expand=True)
        
        # Configure TreeView for both drives
        columns = ("path", "type")
        
        # Create source drive tree
        self.tree_source = ttk.Treeview(tree_container_a, columns=columns, show="tree")
        self.tree_source.heading("#0", text="Name")
        self.tree_source.heading("path", text="Path")
        self.tree_source.heading("type", text="Type")
        context_menu = TreeViewContextMenu(self.tree_source)
        context_menu.add_command("Copy", refresh)
        context_menu.add_command("Rename", refresh)

        # Create destination drive tree
        self.tree_dest = ttk.Treeview(tree_container_b, columns=columns, show="tree")
        self.tree_dest.heading("#0", text="Name")
        self.tree_dest.heading("path", text="Path")
        self.tree_dest.heading("type", text="Type")
        context_menu = TreeViewContextMenu(self.tree_dest)
        context_menu.add_command("Rename", refresh)
        
        # Add scrollbars for both trees
        for tree, container in [(self.tree_source, tree_container_a), 
                            (self.tree_dest, tree_container_b)]:
            # Create scrollbars
            vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
            
            
            # Configure tree
            tree.configure(yscrollcommand=vsb.set)
            
            # Pack in order: tree, vertical scrollbar, horizontal scrollbar
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            
        
        # Create navigators for both drives
        self.nav_source = FileSystemNavigator(self.tree_source)
        self.nav_dest = FileSystemNavigator(self.tree_dest)
        
        # Bind double-click event for directory navigation
        self.tree_source.bind('<Double-1>', lambda e: self._on_double_click(e, self.tree_source))
        self.tree_dest.bind('<Double-1>', lambda e: self._on_double_click(e, self.tree_dest))
        
        # Bind drive selection change events
        self.drive_combo_a.bind('<<ComboboxSelected>>', lambda e: self._on_drive_changed(e, self.nav_source))
        self.drive_combo_b.bind('<<ComboboxSelected>>', lambda e: self._on_drive_changed(e, self.nav_dest))
        
        # Bind tab selection event to populate trees and drive lists
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)

    def _on_tab_changed(self, event):
        """Handle tab selection to populate file trees and drive lists"""
        current = self.notebook.select()
        if self.notebook.index(current) == 3:  # Transfer Files tab index
            # Get available drives
            available_drives = self._get_available_drives()
            
            # Update drive comboboxes
            self.drive_combo_a['values'] = available_drives
            self.drive_combo_b['values'] = available_drives
            
            # Set default selections if not already set
            if not self.drive_combo_a.get() and available_drives:
                self.drive_combo_a.set(available_drives[0])
                self.nav_source.populate_root(f"{available_drives[0]}")
                
            if not self.drive_combo_b.get() and len(available_drives) > 1:
                self.drive_combo_b.set(available_drives[1])
                self.nav_dest.populate_root(f"{available_drives[1]}")
            elif not self.drive_combo_b.get() and available_drives:
                self.drive_combo_b.set(available_drives[0])
                self.nav_dest.populate_root(f"{available_drives[0]}")
    
    def _get_available_drives(self):
        """Get list of available drives"""
        # import win32api
        
        drives = []
        try:
            # Get list of drives using win32api
            bitmask = win32api.GetLogicalDrives()
            for letter in range(65, 91):  # A-Z
                if bitmask & (1 << (letter - 65)):
                    drive = chr(letter) + ":"
                    try:
                        # Check if drive is ready
                        win32api.GetVolumeInformation(drive + "\\")
                        drives.append(drive)
                    except:
                        continue
        except:
            # Fallback to basic drive detection
            import string
            for letter in string.ascii_uppercase:
                drive = f"{letter}:"
                if os.path.exists(drive):
                    drives.append(drive)
        
        return drives

    def _on_drive_changed(self, event, navigator):
        """Handle drive selection change"""
        combobox = event.widget
        selected_drive = combobox.get()
        if selected_drive:
            navigator.populate_root(selected_drive)

    def _on_double_click(self, event, tree):
        """Handle double-click events for directory navigation"""
        region = tree.identify("region", event.x, event.y)
        if region != "nothing":
            item = tree.identify('item', event.x, event.y)
            if item:
                values = tree.item(item)['values']
                if values:
                    path = values[0]
                    item_type = values[1]
                    
                    # Get the appropriate navigator based on which tree was clicked
                    navigator = self.nav_source if tree == self.tree_source else self.nav_dest
                    
                    # Handle directory navigation
                    if item_type == 'Directory':
                        if tree.item(item)['text'] == '..':
                            navigator.navigate_back()
                        else:
                            navigator.navigate_to_directory(path)

    def _get_selected_path(self, tree):
        """Get the full path of the selected item in the tree"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a file to transfer")
            return None
            
        item = selection[0]
        values = tree.item(item)['values']
        if not values:
            return None
            
        return values[0]  # Return the full path

    def copy_file(self):
        """Copy selected file from source to destination"""
        # Get selected file path from source tree
        source_path = self._get_selected_path(self.tree_source)
        if not source_path:
            return
            
        # Get current directory path from destination tree
        dest_dir = self.nav_dest.current_path
        if not dest_dir:
            messagebox.showerror("Error", "Please select a destination directory")
            return
            
        try:
            # Confirm operation with user
            filename = os.path.basename(source_path)
            dest_path = os.path.join(dest_dir, filename)
            
            if os.path.exists(dest_path):
                if not messagebox.askyesno("File Exists", 
                    f"File '{filename}' already exists in destination.\nDo you want to replace it?"):
                    return
            
            # Construct and execute PowerShell copy command
            ps_command = f'''
            $source = '{source_path}'
            $dest = '{dest_path}'
            Copy-Item -Path $source -Destination $dest -Force
            '''
            
            result = self.run_powershell_command(ps_command)
            
            if result is not False:
                self.status_var.set(f"File copied: {filename}")
                # Refresh destination tree
                self.nav_dest.navigate_to_directory(dest_dir)
            else:
                messagebox.showerror("Error", "Failed to copy file")
                
        except Exception as e:
            messagebox.showerror("Copy Error", str(e))
            self.status_var.set("File copy failed")

    def cut_paste_file(self):
        """Move selected file from source to destination"""
        # Get selected file path from source tree
        source_path = self._get_selected_path(self.tree_source)
        if not source_path:
            return
            
        # Get current directory path from destination tree
        dest_dir = self.nav_dest.current_path
        if not dest_dir:
            messagebox.showerror("Error", "Please select a destination directory")
            return
            
        try:
            # Confirm operation with user
            filename = os.path.basename(source_path)
            dest_path = os.path.join(dest_dir, filename)
            
            if os.path.exists(dest_path):
                if not messagebox.askyesno("File Exists", 
                    f"File '{filename}' already exists in destination.\nDo you want to replace it?"):
                    return
            
            # Construct and execute PowerShell move command
            ps_command = f'''
            $source = '{source_path}'
            $dest = '{dest_path}'
            Move-Item -Path $source -Destination $dest -Force
            '''
            
            result = self.run_powershell_command(ps_command)
            
            if result is not False:
                self.status_var.set(f"File moved: {filename}")
                # Refresh both trees
                source_dir = os.path.dirname(source_path)
                self.nav_source.navigate_to_directory(source_dir)
                self.nav_dest.navigate_to_directory(dest_dir)
            else:
                messagebox.showerror("Error", "Failed to move file")
                
        except Exception as e:
            messagebox.showerror("Move Error", str(e))
            self.status_var.set("File move failed")
    

    def create_script_tab(self):
        """Create PowerShell script view/edit tab"""
        script_frame = ttk.Frame(self.notebook)
        self.notebook.add(script_frame, text="PowerShell Script")
        
        self.script_text = scrolledtext.ScrolledText(script_frame,
                                                   wrap=tk.WORD,
                                                   font=("Consolas", 10))
        self.script_text.pack(fill=tk.BOTH, expand=True)
        
        # Load default staging setup script content
        self.load_staging_script()
        
    def create_status_bar(self):
        """Create the status bar"""
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_label = ttk.Label(self.main_frame,
                                    textvariable=self.status_var,
                                    style="Status.TLabel")
        self.status_label.pack(fill=tk.X)
        
    # Hardware scanning functions
    def scan_hardware(self):
        """Scan hardware devices"""
        self.status_var.set("Scanning hardware devices...")
        self.scan_button.configure(state="disabled")
        self.root.update()
        
        try:
            # Clear existing results
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Create temporary file for the script
            with tempfile.NamedTemporaryFile(delete=False,
                                           suffix='.ps1',
                                           mode='w') as temp_file:
                temp_file.write(self.hardware_scan_script)
                temp_file_path = temp_file.name
            
            # Execute PowerShell script
            process = subprocess.Popen(
                ['powershell.exe', '-ExecutionPolicy', 'Bypass',
                 '-File', temp_file_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True
            )
            
            stdout, stderr = process.communicate()
            
            if stderr:
                messagebox.showerror("Error", stderr)
            else:
                self.process_hardware_output(stdout)
            
            # Cleanup
            os.unlink(temp_file_path)
            self.status_var.set(f"Hardware scan completed - {datetime.now().strftime('%H:%M:%S')}")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("Hardware scan failed")
            
        finally:
            self.scan_button.configure(state="normal")
            
    def process_hardware_output(self, output):
        """Process the hardware scan output and populate the treeview"""
        lines = output.split('\n')
        items = []

        for line in lines:
            line = line.strip()
            if line and not line.startswith("Name") and not line.startswith("-"):
                    parts = line.split()
                    if len(parts) >= 3:
                        name = " ".join(parts[:-2])
                        manufacturer = parts[-2]
                        device_id = parts[-1]
                        items.append((name, manufacturer, device_id))
        
        # Sort items based on current sort settings
        col_index = ["Name", "Manufacturer", "DeviceID"].index(self.sort_column)
        items.sort(key=itemgetter(col_index), reverse=self.sort_reverse)
        
        count = 0
        self.tree.tag_configure('oddrow', background='lightgray')
        self.tree.tag_configure('evenrow', background='white')
        # Populate treeview
        for item in items:
            if count % 2 == 0:
                self.tree.insert("", tk.END, values=item, tags="evenrow")
                count += 1
            else:
                self.tree.insert("", tk.END, values=item, tags="oddrow")
                count += 1
            
    def sort_treeview(self, column):
        """Sort treeview contents when a column header is clicked"""
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        
        # Get all items from treeview
        items = []
        for item_id in self.tree.get_children(""):
            values = self.tree.item(item_id)["values"]
            items.append(values)
        
        # Sort items
        col_index = ["Name", "Manufacturer", "DeviceID"].index(column)
        items.sort(key=itemgetter(col_index), reverse=self.sort_reverse)
        
        # Clear and repopulate treeview
        for item_id in self.tree.get_children(""):
            self.tree.delete(item_id)
        
        for item in items:
            self.tree.insert("", tk.END, values=item)
            
        # Update column headers to show sort direction
        self.update_sort_indicators()
        
    def update_sort_indicators(self):
        """Update column headers to show current sort column and direction"""
        for col in ["Name", "Manufacturer", "DeviceID"]:
            text = col
            if col == self.sort_column:
                text = f"{col} {'↓' if self.sort_reverse else '↑'}"
            self.tree.heading(col, text=text)
            
    def export_hardware(self):
        """Export the hardware scan results to a JSON file"""
        try:
            results = []
            for item in self.tree.get_children():
                values = self.tree.item(item)["values"]
                results.append({
                    "Name": values[0],
                    "Manufacturer": values[1],
                    "DeviceID": values[2]
                })
            
            filename = f"hardware_scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(filename, 'w') as f:
                json.dump(results, f, indent=4)
            
            self.status_var.set(f"Results exported to {filename}")
            messagebox.showinfo("Export Complete",
                              f"Results have been exported to {filename}")
            
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
            
    # Network configuration functions
    def enable_dhcp(self):
        """Enable DHCP on the network interface"""
        if messagebox.askyesno("Confirm", "Enable DHCP on the network interface?"):
            result = self.run_powershell_command(
                "Set-NetIPInterface -InterfaceAlias 'Ethernet' -Dhcp Enabled"
            )
            if result is not False:
                self.status_var.set("DHCP enabled successfully")
                
    def apply_static_ip(self):
        """Apply static IP configuration with support for two DNS servers"""
        # Validate required fields
        if not all([self.ip_entry.get(), self.subnet_entry.get(), self.gateway_entry.get()]):
            messagebox.showwarning("Input Error", "IP address, subnet mask, and gateway are required")
            return
            
        if messagebox.askyesno("Confirm", "Apply static IP configuration?"):
            try:
                # Step 1: Configure IP and gateway
                commands = [
                    f"New-NetIPAddress -IPAddress {self.ip_entry.get()} "
                    f"-PrefixLength {self.subnet_entry.get()} "
                    f"-DefaultGateway {self.gateway_entry.get()} "
                    f"-InterfaceAlias 'Ethernet'"
                ]
                
                # Step 2: Build DNS server list from provided entries
                dns_servers = []
                if self.dns_entry1.get().strip():
                    dns_servers.append(self.dns_entry1.get().strip())
                if self.dns_entry2.get().strip():
                    dns_servers.append(self.dns_entry2.get().strip())
                
                # Step 3: Configure DNS if any servers are provided
                if dns_servers:
                    dns_list = "'" + "','".join(dns_servers) + "'"
                    commands.append(
                        f"Set-DnsClientServerAddress -InterfaceAlias 'Ethernet' "
                        f"-ServerAddresses {dns_list}"
                    )
                
                # Step 4: Execute each command in sequence
                for cmd in commands:
                    result = self.run_powershell_command(cmd)
                    if result is False:  # Command failed
                        self.status_var.set("Static IP configuration failed")
                        return
                
                # Step 5: Update status on success
                self.status_var.set("Static IP configuration applied successfully")
                messagebox.showinfo("Success", 
                    "Network configuration has been updated with the following:\n"
                    f"IP: {self.ip_entry.get()}\n"
                    f"Subnet: {self.subnet_entry.get()}\n"
                    f"Gateway: {self.gateway_entry.get()}\n"
                    f"Primary DNS: {self.dns_entry1.get() or 'Not set'}\n"
                    f"Secondary DNS: {self.dns_entry2.get() or 'Not set'}")
                
            except Exception as e:
                messagebox.showerror("Configuration Error", 
                    f"Failed to apply network configuration:\n{str(e)}")
                self.status_var.set("Static IP configuration failed")
            
    # System configuration functions
    def get_available_timezones(self):
        """Get list of available time zones from PowerShell"""
        result = self.run_powershell_command(
            "Get-TimeZone -ListAvailable | Select-Object -ExpandProperty Id"
        )
        if result:
            return sorted(result.split('\n'))
        return []
        
    def rename_computer(self):
        """Rename the computer"""
        new_name = self.computer_name_entry.get()
        if not new_name:
            messagebox.showwarning("Input Error", "Computer name is required")
            return
            
        if messagebox.askyesno("Confirm", f"Rename computer to '{new_name}'?"):
            result = self.run_powershell_command(
                f"Rename-Computer -NewName '{new_name}' -Force"
            )
            if result is not False:
                self.status_var.set("Computer renamed - restart required")
                if messagebox.askyesno("Restart", "Restart computer now?"):
                    self.run_powershell_command("Restart-Computer -Force")
                    
    def set_timezone(self):
        """Set the system time zone"""
        timezone = self.timezone_combo.get()
        if not timezone:
            messagebox.showwarning("Input Error", "Please select a time zone")
            return
            
        if messagebox.askyesno("Confirm", f"Set time zone to '{timezone}'?"):
            result = self.run_powershell_command(
                f"Set-TimeZone -Id '{timezone}'"
            )
            if result is not False:
                self.status_var.set("Time zone updated successfully")
                
    def check_activation(self):
        """Check and handle Windows activation"""
        result = self.run_powershell_command(
            "Get-WmiObject -Query \"SELECT LicenseStatus FROM "
            "SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL\""
        )
        
        if result and "1" in result:
            messagebox.showinfo("Activation", "Windows is already activated")
        else:
            if messagebox.askyesno("Activation", 
                                 "Windows is not activated. Attempt activation now?"):
                result = self.run_powershell_command("slmgr.vbs /ato")
                if result is not False:
                    self.status_var.set("Windows activation attempted")
                    
    def run_powershell_command(self, command):
        """Execute a PowerShell command and return the output"""
        try:
            process = subprocess.Popen(
                ['powershell.exe', '-Command', command],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True
            )
            stdout, stderr = process.communicate()
            
            if stderr:
                messagebox.showerror("Error", stderr)
                return False
                
            return stdout.strip()
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return False
            
    def load_staging_script(self):
        """Load the staging setup script into the script editor"""
        try:
            with open('simpleStagingSetup.bat', 'r') as f:
                self.script_text.delete(1.0, tk.END)
                self.script_text.insert(tk.END, f.read())
        except Exception as e:
            self.script_text.insert(tk.END, "# Default staging setup script will be loaded here")

def refresh():
    return

def main():
    root = tk.Tk()
    app = SystemUtilityGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()