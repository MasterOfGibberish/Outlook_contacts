import win32com.client
import sys
import os
from win32com.client import constants
import pythoncom
import pandas as pd
import traceback
import logging
import tempfile
from datetime import datetime

# Set up logging
log_dir = os.path.join(os.path.expanduser("~"), "AppData", "Local", "OutlookContactExporter")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, "addin_log.txt")
logging.basicConfig(filename=log_file, level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

class OutlookAddin:
    _reg_clsid_ = '{E3FF6600-B388-4FCA-9CFA-3A3AAF35726E}'  # Generate a unique GUID
    _reg_progid_ = "OutlookContactExporter.Addin"
    _reg_desc_ = "Outlook Contact Exporter Add-in"
    _com_interfaces_ = ['_IDTExtensibility2', 'IRibbonExtensibility']
    _public_methods_ = ['GetCustomUI', 'OnButtonClick']

    def __init__(self):
        self.application = None
        self.addin_module = None
        logging.info("Add-in initialized")

    def OnConnection(self, application, connectMode, addin, custom):
        try:
            self.application = application
            self.addin_module = addin
            logging.info("Connected to Outlook")
        except Exception as e:
            logging.error(f"Error in OnConnection: {e}")
            logging.error(traceback.format_exc())

    def OnDisconnection(self, Mode, custom):
        try:
            self.application = None
            self.addin_module = None
            logging.info("Disconnected from Outlook")
        except Exception as e:
            logging.error(f"Error in OnDisconnection: {e}")

    def OnAddInsUpdate(self, custom):
        logging.info("Add-ins updated")
        pass

    def OnStartupComplete(self, custom):
        logging.info("Outlook startup complete")
        pass

    def OnBeginShutdown(self, custom):
        logging.info("Outlook shutdown initiated")
        pass

    def GetCustomUI(self, ribbon_id):
        logging.info(f"GetCustomUI called with ribbon_id: {ribbon_id}")
        # Define the ribbon XML
        return '''
        <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
            <ribbon>
                <tabs>
                    <tab id="CustomTab" label="Contact Tools">
                        <group id="ContactGroup" label="Contact Management">
                            <button id="ExportButton" 
                                    label="Save Contacts" 
                                    size="large" 
                                    imageMso="ExportToExcel" 
                                    onAction="OnButtonClick" />
                        </group>
                    </tab>
                </tabs>
            </ribbon>
        </customUI>
        '''

    def OnButtonClick(self, control):
        # This function will be called when the button is clicked
        try:
            logging.info("Save Contacts button clicked")
            self.extract_sent_contacts()
            # Make sure the active explorer is displayed
            try:
                if self.application.ActiveExplorer():
                    self.application.ActiveExplorer().Display()
            except:
                pass
        except Exception as e:
            logging.error(f"Error in OnButtonClick: {e}")
            logging.error(traceback.format_exc())
            try:
                # Try to show error message
                if hasattr(self.application, 'Session') and self.application.Session:
                    draft = self.application.Session.GetDefaultFolder(6).Items.Add()  # 6 = olFolderInbox
                    draft.Body = f"Error in Contact Exporter Add-in: {str(e)}"
                    draft.Display()
            except:
                # If we can't even show the error, output to file
                with open(os.path.join(os.path.expanduser("~"), "Desktop", "outlook_contact_exporter_error.txt"), "w") as f:
                    f.write(f"Error in Contact Exporter Add-in: {str(e)}\n\n")
                    f.write(traceback.format_exc())

    def extract_sent_contacts(self):
        try:
            # Get the Outlook namespace
            outlook = self.application.GetNamespace("MAPI")
            
            # Initialize contacts list and try to get Sent Items folder
            contacts = []
            
            # Try multiple approaches to get the sent folder
            sent_folder = None
            try:
                # Standard approach
                sent_folder = outlook.GetDefaultFolder(5)  # 5 = olFolderSentMail
            except:
                # Alternative approaches
                try:
                    # Try to get folder by name
                    for folder in outlook.Folders.Item(1).Folders:
                        if folder.Name.lower() in ["sent items", "sent"]:
                            sent_folder = folder
                            break
                except:
                    logging.error("Could not locate Sent Items folder by name")
            
            if not sent_folder:
                raise Exception("Could not locate Sent Items folder")
            
            logging.info("Processing sent items folder")
            
            # Process each email in the Sent Items folder
            processed_count = 0
            for item in sent_folder.Items:
                processed_count += 1
                # Process in batches of 100 to avoid long-running operations
                if processed_count % 100 == 0:
                    logging.info(f"Processed {processed_count} emails so far")
                
                if item.Class == 43:  # olMailItem
                    try:
                        for recipient in item.Recipients:
                            try:
                                name = recipient.Name
                                
                                # Try multiple ways to get the email address
                                email = None
                                try:
                                    email = recipient.Address
                                except:
                                    # If Address fails, try other properties
                                    try:
                                        email = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")  # SMTP address
                                    except:
                                        pass
                                
                                # If still no email and name contains @, extract it
                                if (not email or not "@" in email) and "@" in name:
                                    import re
                                    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', name)
                                    if email_match:
                                        email = email_match.group(0)
                                
                                # Only process if we have an email address
                                if email and "@" in email:
                                    # Try to split name into first and last name
                                    name_parts = name.split()
                                    first_name = name_parts[0] if len(name_parts) > 0 else ""
                                    last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
                                    
                                    contacts.append({
                                        "First Name": first_name,
                                        "Last Name": last_name,
                                        "Full Name": name,
                                        "Email": email.lower()  # Normalize email to lowercase
                                    })
                            except Exception as recipient_error:
                                logging.error(f"Error processing recipient: {str(recipient_error)}")
                                # Skip this recipient but continue processing
                                continue
                    except Exception as item_error:
                        logging.error(f"Error processing email item: {str(item_error)}")
                        # Skip this item but continue processing
                        continue
            
            logging.info(f"Finished processing {processed_count} emails")
            
            # If we didn't find any contacts, try to look in other folders
            if len(contacts) == 0:
                logging.info("No contacts found in Sent Items, trying Inbox")
                try:
                    inbox = outlook.GetDefaultFolder(6)  # 6 = olFolderInbox
                    for item in inbox.Items:
                        if item.Class == 43:  # olMailItem
                            try:
                                # Get the sender info
                                if hasattr(item, 'SenderName') and hasattr(item, 'SenderEmailAddress'):
                                    name = item.SenderName
                                    email = item.SenderEmailAddress
                                    
                                    if email and "@" in email:
                                        # Split name into first and last name
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
                except:
                    logging.error("Error processing Inbox")
                
            # Remove duplicates
            if contacts:
                logging.info(f"Removing duplicates from {len(contacts)} contacts")
                contacts_df = pd.DataFrame(contacts).drop_duplicates(subset=['Email'])
                logging.info(f"Found {len(contacts_df)} unique contacts")
                
                # Make sure the directory exists
                desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
                # Include timestamp in filename to avoid overwriting
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_path = os.path.join(desktop_path, f"outlook_contacts_{timestamp}.xlsx")
                
                # Save to Excel
                try:
                    contacts_df.to_excel(file_path, index=False)
                except Exception as excel_error:
                    logging.error(f"Error saving to Excel: {str(excel_error)}")
                    # Try saving to temp directory if desktop fails
                    temp_dir = tempfile.gettempdir()
                    file_path = os.path.join(temp_dir, f"outlook_contacts_{timestamp}.xlsx")
                    contacts_df.to_excel(file_path, index=False)
                
                # Show a message
                try:
                    if hasattr(self.application, 'Session') and self.application.Session:
                        draft = self.application.Session.GetDefaultFolder(6).Items.Add()
                        draft.Body = f"✅ {len(contacts_df)} contacts saved to '{file_path}'"
                        draft.Display()
                except:
                    # If displaying the message fails, write to a text file
                    with open(os.path.join(desktop_path, "contact_export_result.txt"), "w") as f:
                        f.write(f"✅ {len(contacts_df)} contacts saved to '{file_path}'")
            else:
                logging.warning("No contacts found")
                try:
                    if hasattr(self.application, 'Session') and self.application.Session:
                        draft = self.application.Session.GetDefaultFolder(6).Items.Add()
                        draft.Body = "No contacts found in your Outlook folders"
                        draft.Display()
                except:
                    pass
                    
        except Exception as e:
            logging.error(f"Error in extract_sent_contacts: {e}")
            logging.error(traceback.format_exc())
            raise e

# Register the COM server
if __name__ == '__main__':
    try:
        import win32com.server.register
        logging.info("Registering COM server")
        win32com.server.register.UseCommandLine(OutlookAddin)
    except Exception as e:
        logging.error(f"Error in COM server registration: {e}")
        logging.error(traceback.format_exc())
        print(f"Error: {str(e)}")
        sys.exit(1) 