import win32com.client
import pandas as pd
import os
import sys
import traceback
import pythoncom
import logging
import tempfile
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# Set up logging
log_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "OutlookContactExporter")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, "main_log.txt")
logging.basicConfig(filename=log_file, level=logging.INFO, 
                   format='%(asctime)s - %(levelname)s - %(message)s')

def extract_sent_contacts():
    # Initialize COM in the current thread
    pythoncom.CoInitialize()
    
    try:
        logging.info("Starting contact extraction")
        
        # Try to create Outlook application object
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            logging.info("Connected to Outlook")
        except Exception as e:
            logging.error(f"Failed to connect to Outlook: {e}")
            messagebox.showerror("Error", "Could not connect to Outlook. Please make sure Outlook is installed and running.")
            return False
        
        # Initialize contacts list
        contacts = []
        
        # Try different approaches to get the Sent Items folder
        sent_folder = None
        try:
            # Standard approach
            sent_folder = outlook.GetDefaultFolder(5)  # 5 = olFolderSentMail
            logging.info("Found Sent Items folder using default method")
        except Exception as e:
            logging.warning(f"Could not access Sent Items folder using default method: {e}")
            
            # Try alternative methods
            try:
                # Try to access by name
                for folder in outlook.Folders.Item(1).Folders:
                    if folder.Name.lower() in ["sent items", "sent"]:
                        sent_folder = folder
                        logging.info(f"Found Sent Items folder by name: {folder.Name}")
                        break
            except Exception as e2:
                logging.error(f"Failed to find Sent Items folder by name: {e2}")
        
        if not sent_folder:
            logging.warning("Could not locate Sent Items folder, will try other folders")
        else:
            # Process emails in Sent Items
            logging.info("Processing Sent Items folder")
            processed_count = 0
            
            for item in sent_folder.Items:
                processed_count += 1
                if processed_count % 100 == 0:
                    logging.info(f"Processed {processed_count} emails so far")
                
                if item.Class == 43:  # olMailItem
                    try:
                        # Process all recipients
                        for recipient in item.Recipients:
                            try:
                                name = recipient.Name
                                
                                # Try multiple methods to get email
                                email = None
                                try:
                                    email = recipient.Address
                                except:
                                    pass
                                
                                # Try property accessor if regular approach failed
                                if not email or not "@" in email:
                                    try:
                                        email = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                                    except:
                                        pass
                                
                                # Try to extract from display name if it contains @
                                if (not email or not "@" in email) and "@" in name:
                                    import re
                                    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', name)
                                    if email_match:
                                        email = email_match.group(0)
                                
                                # Only proceed if we have a valid email
                                if email and "@" in email:
                                    # Try to split name into first and last name
                                    name_parts = name.split()
                                    first_name = name_parts[0] if len(name_parts) > 0 else ""
                                    last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
                                    
                                    contacts.append({
                                        "First Name": first_name,
                                        "Last Name": last_name,
                                        "Full Name": name,
                                        "Email": email.lower()  # Normalize to lowercase
                                    })
                            except Exception as rec_err:
                                logging.warning(f"Error processing recipient: {rec_err}")
                                continue
                    except Exception as item_err:
                        logging.warning(f"Error processing email item: {item_err}")
                        continue
        
        # Try to get contacts from Inbox as well
        try:
            inbox = outlook.GetDefaultFolder(6)  # 6 = olFolderInbox
            logging.info("Processing Inbox for additional contacts")
            
            for item in inbox.Items:
                if item.Class == 43:  # olMailItem
                    try:
                        # Get the sender
                        if hasattr(item, 'SenderName') and hasattr(item, 'SenderEmailAddress'):
                            name = item.SenderName
                            email = item.SenderEmailAddress
                            
                            if email and "@" in email:
                                name_parts = name.split()
                                first_name = name_parts[0] if len(name_parts) > 0 else ""
                                last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
                                
                                contacts.append({
                                    "First Name": first_name,
                                    "Last Name": last_name,
                                    "Full Name": name,
                                    "Email": email.lower()
                                })
                    except:
                        continue
        except Exception as inbox_err:
            logging.warning(f"Error accessing Inbox: {inbox_err}")
        
        # Try to get contacts from Contacts folder
        try:
            contacts_folder = outlook.GetDefaultFolder(10)  # 10 = olFolderContacts
            logging.info("Processing Contacts folder")
            
            for contact in contacts_folder.Items:
                try:
                    if hasattr(contact, 'Email1Address') and contact.Email1Address:
                        email = contact.Email1Address
                        name = contact.FullName if hasattr(contact, 'FullName') else ""
                        first_name = contact.FirstName if hasattr(contact, 'FirstName') else ""
                        last_name = contact.LastName if hasattr(contact, 'LastName') else ""
                        
                        contacts.append({
                            "First Name": first_name,
                            "Last Name": last_name,
                            "Full Name": name,
                            "Email": email.lower()
                        })
                except:
                    continue
        except Exception as contacts_err:
            logging.warning(f"Error accessing Contacts folder: {contacts_err}")
        
        # Check if we found any contacts
        if not contacts:
            logging.warning("No contacts found in any folder")
            messagebox.showinfo("No Contacts", "No contacts were found in your Outlook folders.")
            return False
            
        # Remove duplicates based on email address
        logging.info(f"Removing duplicates from {len(contacts)} contacts")
        contacts_df = pd.DataFrame(contacts).drop_duplicates(subset=["Email"])
        logging.info(f"Found {len(contacts_df)} unique contacts")
        
        # Save to Excel with timestamp to avoid overwriting
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        try:
            # Try to save to desktop
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            file_path = os.path.join(desktop_path, f"outlook_contacts_{timestamp}.xlsx")
            contacts_df.to_excel(file_path, index=False)
            logging.info(f"Saved contacts to {file_path}")
        except Exception as excel_err:
            logging.error(f"Error saving to desktop: {excel_err}")
            
            # Fall back to temp directory
            try:
                temp_dir = tempfile.gettempdir()
                file_path = os.path.join(temp_dir, f"outlook_contacts_{timestamp}.xlsx")
                contacts_df.to_excel(file_path, index=False)
                logging.info(f"Saved contacts to temp directory: {file_path}")
            except Exception as temp_err:
                logging.error(f"Error saving to temp directory: {temp_err}")
                messagebox.showerror("Error", f"Could not save contacts file: {str(temp_err)}")
                return False
        
        # Show success message
        messagebox.showinfo("Success", f"✅ {len(contacts_df)} contacts saved to:\n\n{file_path}")
        logging.info("Contact extraction completed successfully")
        return True
    
    except Exception as e:
        logging.error(f"Error in extract_sent_contacts: {e}")
        logging.error(traceback.format_exc())
        messagebox.showerror("Error", f"An error occurred while extracting contacts: {str(e)}")
        return False
    
    finally:
        # Clean up COM
        pythoncom.CoUninitialize()

def show_gui():
    root = tk.Tk()
    root.title("Outlook Contact Exporter")
    root.geometry("400x200")
    root.resizable(False, False)
    
    # Center the window
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - 400) // 2
    y = (screen_height - 200) // 2
    root.geometry(f"400x200+{x}+{y}")
    
    # Add padding
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Add title
    title = tk.Label(frame, text="Outlook Contact Exporter", font=("Arial", 16, "bold"))
    title.pack(pady=(0, 10))
    
    # Add description
    description = tk.Label(frame, text="Export all contacts from your Outlook emails\nto an Excel file", font=("Arial", 10))
    description.pack(pady=(0, 20))
    
    # Add button
    button = tk.Button(frame, text="Extract Contacts", command=extract_sent_contacts, 
                      bg="#007bff", fg="white", font=("Arial", 12), padx=20, pady=5)
    button.pack()
    
    # Add version
    version = tk.Label(frame, text="v1.1.0", font=("Arial", 8), fg="gray")
    version.pack(pady=(20, 0))
    
    root.mainloop()

if __name__ == "__main__":
    try:
        logging.info("Application started")
        show_gui()
    except Exception as e:
        logging.error(f"Unhandled exception: {e}")
        logging.error(traceback.format_exc())
        messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
