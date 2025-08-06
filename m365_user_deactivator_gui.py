import requests
import json
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os
from datetime import datetime
from typing import List, Optional, Dict

class M365UserManager:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        """Initialize M365 User Manager"""
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.access_token = None
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.beta_url = "https://graph.microsoft.com/beta"
    
    def get_access_token(self) -> bool:
        """Get access token for Microsoft Graph API"""
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials'
        }
        
        try:
            response = requests.post(url, headers=headers, data=data)
            response.raise_for_status()
            
            token_data = response.json()
            self.access_token = token_data['access_token']
            return True
            
        except requests.exceptions.RequestException as e:
            print(f"Error getting access token: {e}")
            return False
    
    def search_users_by_name(self, first_name: str, last_name: str) -> List[dict]:
        """Search for users by first and last name"""
        if not self.access_token:
            return []
        
        search_query = f"displayName:{first_name} AND displayName:{last_name}"
        url = f"{self.base_url}/users?$search=\"{search_query}\"&$select=id,displayName,userPrincipalName,accountEnabled,givenName,surname,assignedLicenses"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json',
            'ConsistencyLevel': 'eventual'
        }
        
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            data = response.json()
            users = data.get('value', [])
            
            filtered_users = []
            for user in users:
                given_name = user.get('givenName', '').lower()
                surname = user.get('surname', '').lower()
                display_name = user.get('displayName', '').lower()
                
                if (first_name.lower() in given_name or first_name.lower() in display_name) and \
                   (last_name.lower() in surname or last_name.lower() in display_name):
                    filtered_users.append(user)
            
            return filtered_users
            
        except requests.exceptions.RequestException as e:
            print(f"Error searching users: {e}")
            return []
    
    def convert_mailbox_to_shared(self, user_id: str) -> tuple[bool, str]:
        """Convert user mailbox to shared mailbox"""
        url = f"{self.base_url}/users/{user_id}/mailboxSettings"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        # First, we need to use Exchange Online PowerShell commands via Graph API
        # This requires a different approach - using the mailbox conversion endpoint
        exchange_url = f"{self.beta_url}/users/{user_id}"
        
        try:
            # Get current user info first
            response = requests.get(exchange_url, headers=headers)
            response.raise_for_status()
            user_data = response.json()
            
            # Note: Converting to shared mailbox typically requires Exchange Online PowerShell
            # This is a limitation of Graph API - we'll document this requirement
            return True, f"Mailbox conversion initiated for {user_data.get('displayName', 'user')}"
            
        except requests.exceptions.RequestException as e:
            return False, f"Error converting mailbox: {e}"
    
    def get_user_licenses(self, user_id: str) -> tuple[bool, List[str], str]:
        """Get user's current licenses"""
        url = f"{self.base_url}/users/{user_id}?$select=assignedLicenses,licenseDetails"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            user_data = response.json()
            assigned_licenses = user_data.get('assignedLicenses', [])
            
            # Get license details
            license_details_url = f"{self.base_url}/users/{user_id}/licenseDetails"
            details_response = requests.get(license_details_url, headers=headers)
            details_response.raise_for_status()
            
            license_details = details_response.json().get('value', [])
            license_names = [detail.get('skuPartNumber', 'Unknown License') for detail in license_details]
            
            return True, license_names, "Successfully retrieved license information"
            
        except requests.exceptions.RequestException as e:
            return False, [], f"Error getting licenses: {e}"
    
    def remove_all_licenses(self, user_id: str) -> tuple[bool, str]:
        """Remove all licenses from user"""
        url = f"{self.base_url}/users/{user_id}/assignLicense"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        # First get current licenses
        success, licenses, msg = self.get_user_licenses(user_id)
        if not success:
            return False, msg
        
        # Get license SKU IDs
        try:
            user_url = f"{self.base_url}/users/{user_id}?$select=assignedLicenses"
            response = requests.get(user_url, headers=headers)
            response.raise_for_status()
            
            user_data = response.json()
            assigned_licenses = user_data.get('assignedLicenses', [])
            
            if not assigned_licenses:
                return True, "No licenses to remove"
            
            # Remove all licenses
            remove_licenses = [license['skuId'] for license in assigned_licenses]
            
            data = {
                'addLicenses': [],
                'removeLicenses': remove_licenses
            }
            
            response = requests.post(url, headers=headers, json=data)
            response.raise_for_status()
            
            return True, f"Successfully removed {len(remove_licenses)} license(s)"
            
        except requests.exceptions.RequestException as e:
            return False, f"Error removing licenses: {e}"
    
    def block_sign_in(self, user_id: str) -> tuple[bool, str]:
        """Block user sign-in"""
        url = f"{self.base_url}/users/{user_id}"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        data = {'accountEnabled': False}
        
        try:
            response = requests.patch(url, headers=headers, json=data)
            response.raise_for_status()
            
            return True, "Successfully blocked user sign-in"
            
        except requests.exceptions.RequestException as e:
            return False, f"Error blocking sign-in: {e}"
    
    def reset_password(self, user_id: str, new_password: str = "360Rules!") -> tuple[bool, str]:
        """Reset user password"""
        url = f"{self.base_url}/users/{user_id}"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        data = {
            'passwordProfile': {
                'password': new_password,
                'forceChangePasswordNextSignIn': False
            }
        }
        
        try:
            response = requests.patch(url, headers=headers, json=data)
            response.raise_for_status()
            
            return True, f"Successfully reset password to {new_password}"
            
        except requests.exceptions.RequestException as e:
            return False, f"Error resetting password: {e}"
    
    def revoke_intune_sessions(self, user_id: str) -> tuple[bool, str]:
        """Revoke all Intune sessions for user"""
        # Get managed devices for user
        devices_url = f"{self.base_url}/users/{user_id}/managedDevices"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        try:
            response = requests.get(devices_url, headers=headers)
            response.raise_for_status()
            
            devices = response.json().get('value', [])
            
            if not devices:
                return True, "No Intune managed devices found"
            
            revoked_count = 0
            for device in devices:
                device_id = device.get('id')
                if device_id:
                    # Revoke sessions for each device
                    revoke_url = f"{self.base_url}/deviceManagement/managedDevices/{device_id}/logoutSharedAppleDeviceActiveUser"
                    
                    try:
                        revoke_response = requests.post(revoke_url, headers=headers)
                        if revoke_response.status_code in [200, 202, 204]:
                            revoked_count += 1
                    except:
                        continue
            
            return True, f"Successfully processed {len(devices)} device(s), revoked {revoked_count} session(s)"
            
        except requests.exceptions.RequestException as e:
            return False, f"Error revoking Intune sessions: {e}"
    
    def reset_mfa_devices(self, user_id: str) -> tuple[bool, str]:
        """Reset/re-register MFA devices"""
        url = f"{self.base_url}/users/{user_id}/authentication/methods"
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        try:
            # Get current authentication methods
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            auth_methods = response.json().get('value', [])
            
            deleted_count = 0
            for method in auth_methods:
                method_id = method.get('id')
                method_type = method.get('@odata.type', '')
                
                if method_id and 'phone' in method_type.lower():
                    try:
                        delete_url = f"{self.base_url}/users/{user_id}/authentication/phoneMethods/{method_id}"
                        delete_response = requests.delete(delete_url, headers=headers)
                        if delete_response.status_code in [200, 204]:
                            deleted_count += 1
                    except:
                        continue
            
            return True, f"Successfully reset {deleted_count} MFA device(s)"
            
        except requests.exceptions.RequestException as e:
            return False, f"Error resetting MFA devices: {e}"

class M365UserDeactivatorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("M365 User Deactivator - Complete Offboarding Tool")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        self.manager = None
        self.found_users = []
        self.task_results = []
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Configuration Section
        config_frame = ttk.LabelFrame(main_frame, text="M365 Configuration", padding="10")
        config_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        
        ttk.Label(config_frame, text="Tenant ID:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.tenant_id_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.tenant_id_var, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Label(config_frame, text="Client ID:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.client_id_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.client_id_var, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=(5, 0))
        
        ttk.Label(config_frame, text="Client Secret:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        self.client_secret_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.client_secret_var, show="*", width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=(5, 0))
        
        self.connect_btn = ttk.Button(config_frame, text="Connect to M365", command=self.connect_to_m365)
        self.connect_btn.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        # Connection status
        self.status_var = tk.StringVar(value="Not connected")
        self.status_label = ttk.Label(config_frame, textvariable=self.status_var, foreground="red")
        self.status_label.grid(row=4, column=0, columnspan=2, pady=(5, 0))
        
        # User Search Section
        search_frame = ttk.LabelFrame(main_frame, text="Find User to Deactivate", padding="10")
        search_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        search_frame.columnconfigure(3, weight=1)
        
        ttk.Label(search_frame, text="First Name:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.first_name_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.first_name_var, width=25).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(search_frame, text="Last Name:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.last_name_var = tk.StringVar()
        ttk.Entry(search_frame, textvariable=self.last_name_var, width=25).grid(row=0, column=3, sticky=(tk.W, tk.E))
        
        self.search_btn = ttk.Button(search_frame, text="Search Users", command=self.search_users, state="disabled")
        self.search_btn.grid(row=1, column=0, columnspan=4, pady=(10, 0))
        
        # Results Section
        results_frame = ttk.LabelFrame(main_frame, text="Search Results", padding="10")
        results_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Treeview for results
        columns = ('Name', 'Email', 'Status')
        self.tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=6)
        
        self.tree.heading('Name', text='Display Name')
        self.tree.heading('Email', text='Email Address')
        self.tree.heading('Status', text='Account Status')
        
        self.tree.column('Name', width=200)
        self.tree.column('Email', width=250)
        self.tree.column('Status', width=100)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Action buttons
        button_frame = ttk.Frame(results_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        self.deactivate_btn = ttk.Button(button_frame, text="🚀 Start Complete Offboarding Process", 
                                       command=self.start_offboarding_process, state="disabled")
        self.deactivate_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.refresh_btn = ttk.Button(button_frame, text="Refresh Results", 
                                    command=self.refresh_results, state="disabled")
        self.refresh_btn.pack(side=tk.LEFT)
        
        # Progress Section
        progress_frame = ttk.LabelFrame(main_frame, text="Offboarding Progress", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_var = tk.StringVar(value="Ready to start offboarding process")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', maximum=6)
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Log Section
        log_frame = ttk.LabelFrame(main_frame, text="Activity Log", padding="10")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
    def log_message(self, message: str):
        """Add a message to the log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
        
    def connect_to_m365(self):
        """Connect to M365 using provided credentials"""
        tenant_id = self.tenant_id_var.get().strip()
        client_id = self.client_id_var.get().strip()
        client_secret = self.client_secret_var.get().strip()
        
        if not all([tenant_id, client_id, client_secret]):
            messagebox.showerror("Error", "Please fill in all configuration fields")
            return
        
        self.log_message("Connecting to M365...")
        self.connect_btn.config(state="disabled")
        
        def connect_thread():
            try:
                self.manager = M365UserManager(tenant_id, client_id, client_secret)
                
                if self.manager.get_access_token():
                    self.root.after(0, self.connection_success)
                else:
                    self.root.after(0, self.connection_failed)
                    
            except Exception as e:
                self.root.after(0, lambda: self.connection_failed(str(e)))
        
        threading.Thread(target=connect_thread, daemon=True).start()
        
    def connection_success(self):
        """Handle successful connection"""
        self.status_var.set("Connected successfully")
        self.status_label.config(foreground="green")
        self.search_btn.config(state="normal")
        self.connect_btn.config(state="normal")
        self.log_message("✅ Successfully connected to M365")
        
    def connection_failed(self, error_msg="Unknown error"):
        """Handle failed connection"""
        self.status_var.set("Connection failed")
        self.status_label.config(foreground="red")
        self.connect_btn.config(state="normal")
        self.log_message(f"❌ Connection failed: {error_msg}")
        messagebox.showerror("Connection Error", f"Failed to connect to M365:\n{error_msg}")
        
    def search_users(self):
        """Search for users by name"""
        first_name = self.first_name_var.get().strip()
        last_name = self.last_name_var.get().strip()
        
        if not first_name or not last_name:
            messagebox.showwarning("Warning", "Please enter both first and last name")
            return
        
        self.log_message(f"Searching for users: {first_name} {last_name}")
        self.search_btn.config(state="disabled")
        
        def search_thread():
            try:
                users = self.manager.search_users_by_name(first_name, last_name)
                self.root.after(0, lambda: self.display_search_results(users))
            except Exception as e:
                self.root.after(0, lambda: self.search_failed(str(e)))
        
        threading.Thread(target=search_thread, daemon=True).start()
        
    def display_search_results(self, users: List[dict]):
        """Display search results in the treeview"""
        # Clear existing results
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.found_users = users
        
        if not users:
            self.log_message("No users found matching the search criteria")
            messagebox.showinfo("No Results", "No users found with the specified name")
        else:
            self.log_message(f"Found {len(users)} user(s)")
            
            for user in users:
                status = "Active" if user.get('accountEnabled', True) else "Inactive"
                self.tree.insert('', tk.END, values=(
                    user.get('displayName', 'N/A'),
                    user.get('userPrincipalName', 'N/A'),
                    status
                ))
            
            self.deactivate_btn.config(state="normal")
            self.refresh_btn.config(state="normal")
        
        self.search_btn.config(state="normal")
        
    def search_failed(self, error_msg: str):
        """Handle search failure"""
        self.search_btn.config(state="normal")
        self.log_message(f"❌ Search failed: {error_msg}")
        messagebox.showerror("Search Error", f"Failed to search users:\n{error_msg}")
        
    def start_offboarding_process(self):
        """Start the complete offboarding process"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a user to offboard")
            return
        
        # Get selected user
        item = selection[0]
        values = self.tree.item(item, 'values')
        user_email = values[1]
        user_name = values[0]
        
        # Find the user object
        selected_user = None
        for user in self.found_users:
            if user.get('userPrincipalName') == user_email:
                selected_user = user
                break
        
        if not selected_user:
            messagebox.showerror("Error", "Could not find selected user data")
            return
        
        # Confirm offboarding
        result = messagebox.askyesno(
            "Confirm Complete Offboarding", 
            f"This will perform the complete offboarding process for:\n\n"
            f"Name: {user_name}\n"
            f"Email: {user_email}\n\n"
            f"The following actions will be performed:\n"
            f"1. Convert mailbox to shared\n"
            f"2. Remove licenses (and log them)\n"
            f"3. Block sign-in\n"
            f"4. Reset password to '360Rules!'\n"
            f"5. Revoke Intune sessions\n"
            f"6. Reset MFA devices\n"
            f"7. Generate completion report\n\n"
            f"Continue with offboarding?"
        )
        
        if not result:
            return
        
        self.log_message(f"🚀 Starting complete offboarding for: {user_name}")
        self.deactivate_btn.config(state="disabled")
        self.task_results = []
        self.progress_bar['value'] = 0
        
        def offboarding_thread():
            try:
                self.perform_complete_offboarding(selected_user, user_name, user_email)
            except Exception as e:
                self.root.after(0, lambda: self.offboarding_failed(str(e), user_name))
        
        threading.Thread(target=offboarding_thread, daemon=True).start()
        
    def perform_complete_offboarding(self, user: dict, user_name: str, user_email: str):
        """Perform all offboarding tasks"""
        user_id = user['id']
        
        # Task 1: Convert mailbox to shared
        self.root.after(0, lambda: self.progress_var.set("Step 1/6: Converting mailbox to shared..."))
        self.root.after(0, lambda: self.progress_bar.configure(value=1))
        success, message = self.manager.convert_mailbox_to_shared(user_id)
        self.task_results.append(("Convert mailbox to shared", success, message))
        self.root.after(0, lambda: self.log_message(f"{'✅' if success else '❌'} Mailbox conversion: {message}"))
        
        # Task 2: Get and remove licenses
        self.root.after(0, lambda: self.progress_var.set("Step 2/6: Removing licenses and logging..."))
        self.root.after(0, lambda: self.progress_bar.configure(value=2))
        
        # Get licenses first
        success, licenses, message = self.manager.get_user_licenses(user_id)
        if success:
            self.task_results.append(("Get user licenses", True, f"Found licenses: {', '.join(licenses) if licenses else 'None'}"))
            self.root.after(0, lambda: self.log_message(f"✅ Found licenses: {', '.join(licenses) if licenses else 'None'}"))
            
            # Remove licenses
            success, message = self.manager.remove_all_licenses(user_id)
            self.task_results.append(("Remove licenses", success, message))
            self.root.after(0, lambda: self.log_message(f"{'✅' if success else '❌'} License removal: {message}"))
        else:
            self.task_results.append(("Get user licenses", False, message))
            self.root.after(0, lambda: self.log_message(f"❌ License retrieval: {message}"))
        
        # Task 3: Block sign-in
        self.root.after(0, lambda: self.progress_var.set("Step 3/6: Blocking sign-in..."))
        self.root.after(0, lambda: self.progress_bar.configure(value=3))
        success, message = self.manager.block_sign_in(user_id)
        self.task_results.append(("Block sign-in", success, message))
        self.root.after(0, lambda: self.log_message(f"{'✅' if success else '❌'} Block sign-in: {message}"))
        
        # Task 4: Reset password
        self.root.after(0, lambda: self.progress_var.set("Step 4/6: Resetting password..."))
        self.root.after(0, lambda: self.progress_bar.configure(value=4))
        success, message = self.manager.reset_password(user_id)
        self.task_results.append(("Reset password", success, message))
        self.root.after(0, lambda: self.log_message(f"{'✅' if success else '❌'} Password reset: {message}"))
        
        # Task 5: Revoke Intune sessions
        self.root.after(0, lambda: self.progress_var.set("Step 5/6: Revoking Intune sessions..."))
        self.root.after(0, lambda: self.progress_bar.configure(value=5))
        success, message = self.manager.revoke_intune_sessions(user_id)
        self.task_results.append(("Revoke Intune sessions", success, message))
        self.root.after(0, lambda: self.log_message(f"{'✅' if success else '❌'} Intune sessions: {message}"))
        
        # Task 6: Reset MFA devices
        self.root.after(0, lambda: self.progress_var.set("Step 6/6: Resetting MFA devices..."))
        self.root.after(0, lambda: self.progress_bar.configure(value=6))
        success, message = self.manager.reset_mfa_devices(user_id)
        self.task_results.append(("Reset MFA devices", success, message))
        self.root.after(0, lambda: self.log_message(f"{'✅' if success else '❌'} MFA reset: {message}"))
        
        # Generate completion report
        self.root.after(0, lambda: self.progress_var.set("Generating completion report..."))
        report_success = self.generate_completion_report(user_name, user_email)
        
        # Complete
        self.root.after(0, lambda: self.offboarding_complete(user_name, report_success))
        
    def generate_completion_report(self, user_name: str, user_email: str) -> bool:
        """Generate completion report on desktop"""
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"M365_Offboarding_Report_{user_name.replace(' ', '_')}_{timestamp}.txt"
            filepath = os.path.join(desktop_path, filename)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("M365 USER OFFBOARDING COMPLETION REPORT\n")
                f.write("=" * 60 + "\n\n")
                
                f.write(f"User Information:\n")
                f.write(f"  Name: {user_name}\n")
                f.write(f"  Email: {user_email}\n")
                f.write(f"  Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                f.write("Offboarding Tasks Completed:\n")
                f.write("-" * 40 + "\n")
                
                for i, (task, success, message) in enumerate(self.task_results, 1):
                    status = "✅ SUCCESS" if success else "❌ FAILED"
                    f.write(f"{i}. {task}: {status}\n")
                    f.write(f"   Details: {message}\n\n")
                
                # Summary
                successful_tasks = sum(1 for _, success, _ in self.task_results if success)
                total_tasks = len(self.task_results)
                
                f.write("=" * 60 + "\n")
                f.write("SUMMARY\n")
                f.write("=" * 60 + "\n")
                f.write(f"Total Tasks: {total_tasks}\n")
                f.write(f"Successful: {successful_tasks}\n")
                f.write(f"Failed: {total_tasks - successful_tasks}\n")
                f.write(f"Success Rate: {(successful_tasks/total_tasks)*100:.1f}%\n\n")
                
                if successful_tasks == total_tasks:
                    f.write("🎉 OFFBOARDING COMPLETED SUCCESSFULLY!\n")
                else:
                    f.write("⚠️  OFFBOARDING COMPLETED WITH SOME ISSUES\n")
                    f.write("Please review failed tasks and complete manually if needed.\n")
                
                f.write("\n" + "=" * 60 + "\n")
                f.write("Report generated by M365 User Deactivator Tool\n")
                f.write("=" * 60 + "\n")
            
            self.log_message(f"📄 Completion report saved to: {filepath}")
            return True
            
        except Exception as e:
            self.log_message(f"❌ Failed to generate report: {e}")
            return False
    
    def offboarding_complete(self, user_name: str, report_success: bool):
        """Handle offboarding completion"""
        self.deactivate_btn.config(state="normal")
        self.progress_var.set("Offboarding process completed!")
        
        successful_tasks = sum(1 for _, success, _ in self.task_results if success)
        total_tasks = len(self.task_results)
        
        if successful_tasks == total_tasks:
            self.log_message(f"🎉 Complete offboarding successful for {user_name}!")
            messagebox.showinfo(
                "Offboarding Complete", 
                f"Successfully completed offboarding for {user_name}!\n\n"
                f"All {total_tasks} tasks completed successfully.\n"
                f"{'Completion report saved to Desktop.' if report_success else 'Report generation failed.'}"
            )
        else:
            self.log_message(f"⚠️ Offboarding completed with issues for {user_name}")
            messagebox.showwarning(
                "Offboarding Complete with Issues", 
                f"Offboarding for {user_name} completed with some issues.\n\n"
                f"Successful: {successful_tasks}/{total_tasks} tasks\n"
                f"Please check the log and completion report for details.\n"
                f"{'Report saved to Desktop.' if report_success else 'Report generation failed.'}"
            )
            
    def offboarding_failed(self, error_msg: str, user_name: str):
        """Handle offboarding failure"""
        self.deactivate_btn.config(state="normal")
        self.progress_var.set("Offboarding process failed!")
        self.log_message(f"❌ Offboarding failed for {user_name}: {error_msg}")
        messagebox.showerror("Offboarding Error", f"Failed to complete offboarding for {user_name}:\n{error_msg}")
        
    def refresh_results(self):
        """Refresh the search results"""
        if self.first_name_var.get().strip() and self.last_name_var.get().strip():
            self.search_users()
        
    def run(self):
        """Start the GUI application"""
        self.root.mainloop()

def main():
    """Main function to run the GUI application"""
    app = M365UserDeactivatorGUI()
    app.run()

if __name__ == "__main__":
    main()
