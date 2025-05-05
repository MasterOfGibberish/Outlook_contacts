import sys
import os
import subprocess
import traceback
import tempfile
from datetime import datetime

# Function to check and install required packages
def ensure_packages():
    required_packages = ["pywin32", "pandas", "openpyxl", "tkinter"]
    missing_packages = []
    
    # First try importing each package to see if it's already installed
    for package in required_packages:
        try:
            if package == "pywin32":
                __import__("win32com")
            elif package == "tkinter":
                __import__("tkinter")
            else:
                __import__(package)
        except ImportError:
            missing_packages.append(package)
    
    # If missing packages, install them
    if missing_packages:
        print(f"Installing missing packages: {', '.join(missing_packages)}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing_packages)
            print("Packages installed successfully!")
            
            # A special fix for pywin32 post-install
            if "pywin32" in missing_packages:
                try:
                    # Find pywin32_postinstall.py
                    site_packages = subprocess.check_output([sys.executable, "-c", 
                                 "import site; print(site.getsitepackages()[0])"]).decode().strip()
                    postinstall_script = os.path.join(site_packages, "pywin32_system32", "pywin32_postinstall.py")
                    
                    if os.path.exists(postinstall_script):
                        print("Running pywin32 post-install script...")
                        subprocess.check_call([sys.executable, postinstall_script, "-install"])
                except Exception as e:
                    print(f"Warning: Could not run pywin32 post-install: {e}")
            
            # Restart the script to use newly installed packages
            print("Restarting application with installed packages...")
            os.execv(sys.executable, [sys.executable] + sys.argv)
        except Exception as e:
            print(f"Error installing packages: {e}")
            print("Please run manually: pip install pywin32 pandas openpyxl")
            input("Press Enter to exit...")
            sys.exit(1)

# Make sure required packages are installed first
ensure_packages()

# Now import the packages
import win32com.client
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import re
import threading
import pythoncom  # Import pythoncom for COM initialization

def extract_contacts_thread(progress_window, progress_var, status_var):
    try:
        # Initialize COM in this thread
        pythoncom.CoInitialize()
        
        # Create Outlook application object
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        contacts = []
        signatures_cache = {}  # Cache to store extracted roles from signatures
        total_items_processed = 0
        
        # Update status
        status_var.set("Initializing...")
        progress_var.set(5)
        
        # Function to extract role from email signature or body
        def extract_role_from_body(email_address, sender_name, body_text):
            # Return from cache if we've already processed this email
            if email_address in signatures_cache:
                return signatures_cache[email_address]
                
            # No body text to process
            if not body_text:
                return ""
                
            # Convert HTML to plain text if needed
            if body_text.startswith("<html") or "<body" in body_text:
                # Simple HTML tag removal
                plain_text = re.sub('<[^<]+?>', ' ', body_text)
            else:
                plain_text = body_text
                
            # Get the last few lines where signatures usually appear
            lines = plain_text.splitlines()
            # Focus on the last 15 lines (typical signature length)
            signature_area = "\n".join(lines[-15:]) if len(lines) > 15 else plain_text
            
            # Patterns to identify job titles in signatures
            job_title_patterns = [
                # Pattern for "Name | Title"
                rf"{re.escape(sender_name)}\s*[|\|]\s*([^,\n\|]{3,50})",
                # Pattern for "Title at Company"
                r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,4}(?:\s+at|@)\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3})",
                # Pattern for job title followed by department
                r"([A-Z][a-z]+\s+(?:of|for)\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3})",
                # Patterns for common job titles
                r"((?:Senior|Junior|Chief|Assistant|Associate|Lead|Principal|Director|Manager|Officer|President|CEO|CTO|CFO|COO|VP|Head|Founder|Owner|Specialist|Supervisor|Coordinator|Analyst|Engineer|Developer|Architect|Designer|Consultant|Executive|Administrator|Technician)(?:\s+[A-Z][a-z]+){1,4})",
                # Pattern for roles with "of" construction
                r"((?:Director|Manager|Head|Chief|Officer)\s+of\s+(?:[A-Z][a-z]+\s*){1,5})",
                # Pattern for titles like "Marketing Manager"
                r"((?:Marketing|Sales|Finance|HR|Operations|IT|Product|Software|Network|Data|AI|Business|Project|Program|Customer|Research|Quality|Technical|Support)\s+(?:Manager|Director|Specialist|Analyst|Engineer|Coordinator|Lead|Supervisor|Consultant|Executive))",
                # Pattern for simple title, company format
                r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3}),\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,5})",
            ]
            
            # Look for matches with the patterns
            potential_roles = []
            for pattern in job_title_patterns:
                matches = re.findall(pattern, signature_area)
                if matches:
                    for match in matches:
                        if isinstance(match, tuple):  # Multiple capturing groups
                            for group in match:
                                if group and len(group) > 5 and len(group) < 50:  # Reasonable length for a title
                                    potential_roles.append(group.strip())
                        else:
                            if len(match) > 5 and len(match) < 50:  # Reasonable length for a title
                                potential_roles.append(match.strip())
            
            # If we found roles, use the first one
            role = potential_roles[0] if potential_roles else ""
            
            # Store in cache
            signatures_cache[email_address] = role
            return role
        
        # Get all the folders to scan - expand to more folders to find all contacts
        folders_to_scan = {}
        try:
            folders_to_scan["Sent Items"] = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
        except:
            pass
            
        try:
            folders_to_scan["Inbox"] = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        except:
            pass
            
        try:
            folders_to_scan["Deleted Items"] = namespace.GetDefaultFolder(3)  # 3 = olFolderDeletedItems
        except:
            pass
            
        try:
            folders_to_scan["Drafts"] = namespace.GetDefaultFolder(16)  # 16 = olFolderDrafts
        except:
            pass
            
        try:
            folders_to_scan["Outbox"] = namespace.GetDefaultFolder(4)  # 4 = olFolderOutbox
        except:
            pass
            
        try:
            folders_to_scan["Junk Email"] = namespace.GetDefaultFolder(23)  # 23 = olFolderJunk
        except:
            pass
        
        # Try to access Archive folder if it exists
        try:
            archive_folder = namespace.Stores.Item("Archive").GetRootFolder()
            folders_to_scan["Archive"] = archive_folder
        except:
            pass
        
        # Process Outlook GAL (Global Address List) to match X500 addresses to emails
        def get_exchange_address_mapping():
            exchange_map = {}
            return exchange_map
            
        exchange_map = get_exchange_address_mapping()
        
        # Update progress
        progress_var.set(10)
        
        # Track progress
        folder_count = len(folders_to_scan)
        current_folder = 0
        
        # Function to try to get role/job title from a contact
        def try_get_role(recipient):
            job_title = ""
            try:
                # Method 1: Try to access job title directly
                if hasattr(recipient, 'JobTitle'):
                    job_title = recipient.JobTitle
                
                # Method 2: Try to get from AddressEntry
                if not job_title and hasattr(recipient, 'AddressEntry'):
                    try:
                        addressEntry = recipient.AddressEntry
                        if addressEntry.Type == "EX":  # Exchange user
                            exchangeUser = addressEntry.GetExchangeUser()
                            if exchangeUser and hasattr(exchangeUser, 'JobTitle'):
                                job_title = exchangeUser.JobTitle
                    except:
                        pass
                
                # Method 3: Try to get from GAL property
                if not job_title:
                    try:
                        job_title = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A17001E")  # PR_TITLE
                    except:
                        pass
                
                # Method 4: Try to get from contact item if available
                if not job_title:
                    try:
                        recipient_resolved = namespace.CreateRecipient(recipient.Name)
                        recipient_resolved.Resolve()
                        if recipient_resolved.Resolved:
                            entry = recipient_resolved.AddressEntry
                            if entry.Type == "EX":
                                contact = entry.GetContact()
                                if contact and hasattr(contact, 'JobTitle'):
                                    job_title = contact.JobTitle
                    except:
                        pass
            except:
                pass
            
            return job_title
        
        # Process all folders
        for folder_name, folder in folders_to_scan.items():
            try:
                current_folder += 1
                folder_progress = int(10 + (current_folder / folder_count) * 70)  # 10-80% progress for folders
                progress_var.set(folder_progress)
                status_var.set(f"Scanning {folder_name}...")
                
                # Process all items in the folder
                for item in folder.Items:
                    if item.Class == 43:  # olMailItem
                        total_items_processed += 1
                        
                        # Get the email body for signature analysis
                        email_body = ""
                        try:
                            if hasattr(item, 'Body'):
                                email_body = item.Body
                            elif hasattr(item, 'HTMLBody'):
                                email_body = item.HTMLBody
                        except:
                            pass
                        
                        # Process sender
                        try:
                            if hasattr(item, 'SenderName') and item.SenderName:
                                sender_name = item.SenderName
                                sender_email = None
                                sender_role = ""
                                
                                # Try multiple ways to get the sender email
                                if hasattr(item, 'SenderEmailAddress'):
                                    sender_email = item.SenderEmailAddress
                                    
                                # If it's Exchange format, try alternative methods
                                if sender_email and sender_email.startswith("/o=ExchangeLabs"):
                                    # Try from the email type + address
                                    try:
                                        if hasattr(item, 'SenderEmailType'):
                                            sender_email_alt = item.SenderEmailType + ":" + item.SenderEmailAddress
                                            if "@" in sender_email_alt:
                                                sender_email = sender_email_alt
                                    except:
                                        pass
                                        
                                    # Try extracting from display name
                                    if sender_name and "@" in sender_name:
                                        email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', sender_name)
                                        if email_match:
                                            sender_email = email_match.group(0)
                                            
                                    # Try resolver to get SMTP address
                                    try:
                                        recipient = namespace.CreateRecipient(sender_name)
                                        recipient.Resolve()
                                        if recipient.Resolved:
                                            addressEntry = recipient.AddressEntry
                                            if addressEntry.Type == "EX":
                                                try:
                                                    sender_email = addressEntry.GetExchangeUser().PrimarySmtpAddress
                                                    # Try to get job title/role
                                                    exchangeUser = addressEntry.GetExchangeUser()
                                                    if exchangeUser and hasattr(exchangeUser, 'JobTitle'):
                                                        sender_role = exchangeUser.JobTitle
                                                except:
                                                    pass
                                    except:
                                        pass
                                        
                                    # Last resort: use our Exchange address mapping or extraction
                                    if sender_email.startswith("/o=ExchangeLabs"):
                                        if sender_email in exchange_map:
                                            sender_email = exchange_map[sender_email]
                                        else:
                                            # Extract email from display name - just use name part
                                            name_part = sender_name.split()
                                            if len(name_part) > 0:
                                                domain = "company.com"  # You might need to change this
                                                sender_email = f"{name_part[0].lower()}@{domain}"
                                
                                # Only add if we have a reasonable email now
                                if sender_email and "@" in sender_email:
                                    # Check if we need to extract role from signature
                                    if not sender_role and folder_name != "Sent Items":  # Don't analyze our own signatures
                                        # Try to find a job title in the email signature
                                        signature_role = extract_role_from_body(sender_email.lower(), sender_name, email_body)
                                        if signature_role:
                                            sender_role = signature_role
                                    
                                    # Try to split name into first and last name
                                    name_parts = sender_name.split()
                                    first_name = name_parts[0] if len(name_parts) > 0 else ""
                                    last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
                                    
                                    contacts.append({
                                        "First Name": first_name,
                                        "Last Name": last_name,
                                        "Full Name": sender_name,
                                        "Email": sender_email.lower(),
                                        "Role": sender_role,
                                        "Source": f"{folder_name} (Sender)"
                                    })
                        except Exception as e:
                            # Skip errors silently
                            pass
                            
                        # Process all recipients
                        try:
                            if hasattr(item, 'Recipients'):
                                for recipient in item.Recipients:
                                    try:
                                        name = recipient.Name
                                        email = None
                                        exchange_address = None
                                        role = try_get_role(recipient)  # Try to get role/job title
                                        
                                        # Try multiple methods to get email
                                        # Method 1: Address property
                                        try:
                                            email = recipient.Address
                                            if email and email.startswith("/o=ExchangeLabs"):
                                                exchange_address = email
                                        except:
                                            pass
                                            
                                        # Method 2: SMTP Address property
                                        if email is None or email.startswith("/o=ExchangeLabs"):
                                            try:
                                                email = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                                            except:
                                                pass
                                        
                                        # Method 3: Use AddressEntry to get SMTP address
                                        if email is None or email.startswith("/o=ExchangeLabs"):
                                            try:
                                                if hasattr(recipient, 'AddressEntry'):
                                                    addressEntry = recipient.AddressEntry
                                                    if addressEntry.Type == "EX":  # Exchange user
                                                        try:
                                                            exchangeUser = addressEntry.GetExchangeUser()
                                                            if exchangeUser:
                                                                email = exchangeUser.PrimarySmtpAddress
                                                                # Also try to get job title if we don't have it yet
                                                                if not role and hasattr(exchangeUser, 'JobTitle'):
                                                                    role = exchangeUser.JobTitle
                                                        except:
                                                            pass
                                            except:
                                                pass
                                                
                                        # Method 4: Extract from display name if it contains an email
                                        if (email is None or email.startswith("/o=ExchangeLabs")) and "@" in name:
                                            email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', name)
                                            if email_match:
                                                email = email_match.group(0)
                                                
                                        # Method 5: For Exchange addresses, make a best guess based on name
                                        if email is None or email.startswith("/o=ExchangeLabs"):
                                            # Extract part after the last '-' in the Exchange address
                                            if email and email.startswith("/o=ExchangeLabs"):
                                                name_match = re.search(r'cn=([^-]+)-(.+)$', email)
                                                if name_match:
                                                    extracted_name = name_match.group(2)
                                                    domain = "company.com"  # Change as needed
                                                    email = f"{extracted_name.lower().replace(' ', '.')}@{domain}"
                                                else:
                                                    # Fallback to using the display name
                                                    name_parts = name.replace(',', '').split()
                                                    if len(name_parts) > 0:
                                                        cleaned_name = name_parts[0].lower()
                                                        domain = "company.com"  # Change as needed
                                                        email = f"{cleaned_name}@{domain}"
                                        
                                        # Skip if still no valid email
                                        if email is None and exchange_address is None:
                                            continue
                                            
                                        # Use exchange_address if we don't have a better email
                                        if (email is None or email == "") and exchange_address:
                                            email = exchange_address
                                        
                                        # Try to split name into first and last name
                                        name_parts = name.split()
                                        first_name = name_parts[0] if len(name_parts) > 0 else ""
                                        last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
                                        
                                        # If no role yet and this is in the Inbox, try to extract from signature
                                        if not role and folder_name == "Inbox" and email and "@" in email:
                                            # Check the signature cache first
                                            if email.lower() in signatures_cache:
                                                role = signatures_cache[email.lower()]
                                            # Otherwise try to extract from the message body
                                            elif email_body:
                                                signature_role = extract_role_from_body(email.lower(), name, email_body)
                                                if signature_role:
                                                    role = signature_role
                                        
                                        contacts.append({
                                            "First Name": first_name,
                                            "Last Name": last_name,
                                            "Full Name": name,
                                            "Email": email.lower() if email else "",
                                            "Role": role,
                                            "Source": f"{folder_name} (Recipient)"
                                        })
                                    except Exception as e:
                                        # Skip errors silently
                                        pass
                        except Exception as e:
                            # Skip errors silently for specific item
                            pass
            except Exception as e:
                # Skip this folder and continue with others
                continue
        
        # Update progress for contacts folder processing
        progress_var.set(80)
        status_var.set("Scanning Contacts folder...")
        
        # Additional scan for Contacts folder - this should have the most job title info
        try:
            contacts_folder = namespace.GetDefaultFolder(10)  # 10 = olFolderContacts
            
            for contact_item in contacts_folder.Items:
                if contact_item.Class == 40:  # olContactItem
                    try:
                        if hasattr(contact_item, 'FullName') and hasattr(contact_item, 'Email1Address'):
                            name = contact_item.FullName
                            email = contact_item.Email1Address
                            
                            # Get job title/role if available
                            role = ""
                            if hasattr(contact_item, 'JobTitle'):
                                role = contact_item.JobTitle
                            elif hasattr(contact_item, 'CompanyName'):
                                role = contact_item.CompanyName
                            
                            if email and "@" in email:
                                name_parts = name.split()
                                first_name = contact_item.FirstName if hasattr(contact_item, 'FirstName') else (name_parts[0] if len(name_parts) > 0 else "")
                                last_name = contact_item.LastName if hasattr(contact_item, 'LastName') else (" ".join(name_parts[1:]) if len(name_parts) > 1 else "")
                                
                                contacts.append({
                                    "First Name": first_name,
                                    "Last Name": last_name,
                                    "Full Name": name,
                                    "Email": email.lower(),
                                    "Role": role,
                                    "Source": "Contacts Folder"
                                })
                    except:
                        pass
        except:
            pass
        
        # Update for final processing
        progress_var.set(85)
        status_var.set("Processing contacts...")
        
        # Create DataFrame and ensure all values are strings to avoid type issues
        contacts_df = pd.DataFrame(contacts)
        
        # Skip empty dataframe case
        if len(contacts_df) == 0:
            progress_window.destroy()
            messagebox.showinfo("No Contacts", "No valid contacts found in your mailbox.")
            return False

        # Group by email address and pick the best record (with most information)
        best_contacts = []
        for email, group in contacts_df.groupby('Email'):
            # If any role is available, use that record
            with_role = group[group['Role'].astype(str) != '']
            if len(with_role) > 0:
                best_contacts.append(with_role.iloc[0])
            else:
                # Otherwise use any record
                best_contacts.append(group.iloc[0])
        
        # Create final dataframe
        result_df = pd.DataFrame(best_contacts)
        
        # Sort by name
        result_df = result_df.sort_values(by=['Last Name', 'First Name'])
        
        # Remove the Exchange Address column if it exists
        if 'Exchange Address' in result_df.columns:
            result_df = result_df.drop(columns=['Exchange Address'])
            
        # Update progress
        progress_var.set(95)
        status_var.set("Saving to Excel...")
        
        # Save to Excel
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = os.path.join(desktop_path, f"outlook_contacts_{timestamp}.xlsx")
        
        try:
            result_df.to_excel(file_path, index=False)
        except Exception as excel_error:
            # Try saving to temp directory if desktop fails
            temp_dir = tempfile.gettempdir()
            file_path = os.path.join(temp_dir, f"outlook_contacts_{timestamp}.xlsx")
            result_df.to_excel(file_path, index=False)
        
        # Final update
        progress_var.set(100)
        status_var.set("Complete!")
        
        # Close progress window and show final message
        progress_window.destroy()
        messagebox.showinfo("Success", f"✅ {len(result_df)} unique contacts exported from {total_items_processed} emails\n\nSaved to:\n{file_path}")
        
        # At the end of the function, uninitialize COM:
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        # Show error and close progress window
        try:
            pythoncom.CoUninitialize()  # Make sure to uninitialize even on error
        except:
            pass
        progress_window.destroy()
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        return False

def extract_contacts():
    # Create a progress window
    progress_window = tk.Toplevel()
    progress_window.title("Exporting Contacts")
    progress_window.geometry("400x150")
    progress_window.resizable(False, False)
    progress_window.transient()  # Set as transient window
    progress_window.grab_set()   # Make modal
    
    # Center the window
    progress_window.update_idletasks()
    width = progress_window.winfo_width()
    height = progress_window.winfo_height()
    x = (progress_window.winfo_screenwidth() // 2) - (width // 2)
    y = (progress_window.winfo_screenheight() // 2) - (height // 2)
    progress_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    # Add a frame with padding
    frame = tk.Frame(progress_window, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Add a label
    status_var = tk.StringVar()
    status_var.set("Starting...")
    status_label = tk.Label(frame, textvariable=status_var, font=("Arial", 10))
    status_label.pack(pady=(0, 10))
    
    # Add a progress bar
    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(frame, variable=progress_var, length=300, mode="determinate")
    progress_bar.pack(pady=(0, 10))
    
    # Add a message
    message = tk.Label(frame, text="Please wait while your contacts are being exported...", font=("Arial", 8))
    message.pack()
    
    # Start the extraction in a separate thread
    thread = threading.Thread(target=extract_contacts_thread, args=(progress_window, progress_var, status_var))
    thread.daemon = True
    thread.start()
    
    return True

# Create a simple GUI
def create_gui():
    root = tk.Tk()
    root.title("Outlook Contact Exporter")
    root.geometry("450x250")
    root.resizable(False, False)
    
    # Center the window
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - 450) // 2
    y = (screen_height - 250) // 2
    root.geometry(f"450x250+{x}+{y}")
    
    # Add some padding
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Add heading
    heading = tk.Label(frame, text="Outlook Contact Exporter", font=("Arial", 16, "bold"))
    heading.pack(pady=(0, 10))
    
    # Add description
    description = tk.Label(frame, text="This tool will scan ALL your Outlook folders\nto extract EVERY contact you've ever communicated with.", justify=tk.CENTER)
    description.pack(pady=(0, 5))
    
    # Add warning about time
    warning = tk.Label(frame, text="⚠️ This may take several minutes for large mailboxes", fg="orange", font=("Arial", 9, "italic"))
    warning.pack(pady=(0, 15))
    
    # Add button
    button = tk.Button(frame, text="Export ALL Contacts", command=extract_contacts, 
                      bg="#0078D7", fg="white", font=("Arial", 12), padx=10, pady=5)
    button.pack()
    
    # Add status
    status = tk.Label(frame, text="Will collect ALL contacts from your Outlook", font=("Arial", 8), fg="gray")
    status.pack(pady=(15, 0))
    
    # Add version
    version = tk.Label(frame, text="v1.1.0", font=("Arial", 8), fg="gray")
    version.pack(pady=(5, 0))
    
    root.mainloop()

if __name__ == "__main__":
    try:
        # Make sure required packages are installed
        ensure_packages()
        create_gui()
    except Exception as e:
        # If we get here before tkinter is initialized, we need a basic error message
        try:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
        except:
            print(f"\nERROR: {str(e)}")
            print("Please make sure all required packages are installed:")
            print("pip install pywin32 pandas openpyxl")
            input("\nPress Enter to exit...") 