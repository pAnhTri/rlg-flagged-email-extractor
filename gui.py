from tkinter import Tk, ttk, END, filedialog, messagebox
import os
import sys
import calendar
from datetime import date
from utils import (
    is_outlook_installed,
    get_config,
    set_config,
    get_flagged_emails_in_month,
    get_flagged_emails_in_month_pst,
)


def browse_folder(form_entry: ttk.Entry):
    folder_selected = filedialog.askdirectory()
    form_entry.delete(0, END)
    form_entry.insert(0, folder_selected)
    set_config("Folder", "output_folder", folder_selected)


class MainWindow:
    _flagged_emails_in_month = []
    _start_of_month = None
    _end_of_month = None

    def __init__(
        self,
        root,
        title: str = "Rad Law Group. APLC - Email Exporter",
        height: int = 320,
        width: int = 240,
    ):
        self.root = root
        self.root.title(title)
        self.root.geometry(f"{width}x{height}")

        # Set the window icon
        self.set_window_icon()

        # Create main frame
        self.main_frame = ttk.Frame(root, padding=20)
        self.main_frame.pack(expand=True, fill="both")

        # Create loading widgets
        self.create_loading_widgets()

        # Create main widgets (initially hidden)
        self.create_main_widgets()

        # Don't automatically start Outlook check - wait for user to click connect button

    def set_window_icon(self):
        """Set the window icon to the RLG logo"""
        try:
            # Try multiple possible locations for the icon file
            possible_paths = [
                os.path.join(os.path.dirname(__file__), "app.ico"),  # Development
                os.path.join(
                    os.path.dirname(sys.executable), "app.ico"
                ),  # PyInstaller bundled
                "app.ico",  # Current directory
            ]

            icon_path = None
            for path in possible_paths:
                if os.path.exists(path):
                    icon_path = path
                    break

            # Check if the icon file exists
            if icon_path:
                # Set the window icon using iconbitmap for .ico files
                self.root.iconbitmap(icon_path)
                print(f"Window icon set successfully: {icon_path}")
            else:
                print(f"Icon file not found in any of the expected locations")
        except Exception as e:
            print(f"Error setting window icon: {e}")

    def create_loading_widgets(self):
        """Create initial setup widgets"""
        # Welcome label
        self.welcome_label = ttk.Label(
            self.main_frame, text="Email Exporter", font=("Arial", 16, "bold")
        )
        self.welcome_label.pack(pady=(0, 20))

        # Email form frame
        self.email_form_frame = ttk.Frame(self.main_frame)

        # Email form widgets
        self.email_form_label = ttk.Label(
            self.email_form_frame, text="Folder to Extract From"
        )
        self.email_form_label.pack(anchor="w")

        self.email_form_entry = ttk.Entry(self.email_form_frame)
        self.email_form_entry.pack(fill="x", pady=(5, 0))

        # Load default email from config or set empty string
        try:
            default_email = get_config("Email", "primary_email")
        except (FileNotFoundError, KeyError):
            default_email = ""

        self.email_form_entry.insert(0, default_email)

        # Bind the entry to save changes to config
        self.email_form_entry.bind("<FocusOut>", self.save_email_to_config)

        # Date range widgets in the same frame
        self.date_form_label = ttk.Label(self.email_form_frame, text="Date Range")
        self.date_form_label.pack(anchor="w", pady=(15, 0))

        # Start date frame
        self.start_date_frame = ttk.Frame(self.email_form_frame)
        self.start_date_frame.pack(fill="x", pady=(5, 0))

        self.start_date_label = ttk.Label(self.start_date_frame, text="Start Date:")
        self.start_date_label.pack(side="left")

        self.start_date_entry = ttk.Entry(self.start_date_frame, width=15)
        self.start_date_entry.pack(side="right", padx=(10, 0))

        # Default to beginning of current month
        current_date = date.today()
        start_of_month = current_date.replace(day=1)
        self.start_date_entry.insert(0, start_of_month.strftime("%Y-%m-%d"))

        # End date frame
        self.end_date_frame = ttk.Frame(self.email_form_frame)
        self.end_date_frame.pack(fill="x", pady=(5, 0))

        self.end_date_label = ttk.Label(self.end_date_frame, text="End Date:")
        self.end_date_label.pack(side="left")

        self.end_date_entry = ttk.Entry(self.end_date_frame, width=15)
        self.end_date_entry.pack(side="right", padx=(10, 0))

        # Default to end of current month
        _, num_days = calendar.monthrange(current_date.year, current_date.month)
        end_of_month = current_date.replace(day=num_days)
        self.end_date_entry.insert(0, end_of_month.strftime("%Y-%m-%d"))

        self.email_form_frame.pack(pady=10, fill="x")

        # Connect button
        self.connect_button = ttk.Button(
            self.main_frame,
            text="Connect to Outlook",
            command=self.start_outlook_connection,
        )
        self.connect_button.pack(pady=20)

        # Status label
        self.status_label = ttk.Label(self.main_frame, text="", font=("Arial", 10))
        self.status_label.pack(pady=10)

    def create_main_widgets(self):
        """Create main application widgets (initially hidden)"""
        # Welcome label
        self.welcome_label = ttk.Label(
            self.main_frame, text="Email Exporter", font=("Arial", 16, "bold")
        )

        # Status label
        self.connection_status = ttk.Label(
            self.main_frame, text="✓ Outlook connected successfully", foreground="green"
        )

        # Flagged emails count label
        self.flagged_emails_count_label = ttk.Label(
            self.main_frame,
            text="Loading flagged emails...",
        )

        # Form frame
        self.folder_form_frame = ttk.Frame(self.main_frame)

        # Folder form widgets
        self.folder_form_label = ttk.Label(self.folder_form_frame, text="Output Folder")
        self.folder_form_label.pack(anchor="w")

        # Entry and button in a horizontal frame
        self.folder_input_frame = ttk.Frame(self.folder_form_frame)
        self.folder_input_frame.pack(fill="x", pady=(5, 0))

        self.folder_form_entry = ttk.Entry(self.folder_input_frame)
        self.folder_form_entry.pack(side="left", fill="x", expand=True)
        self.folder_form_entry.insert(0, get_config("Folder", "output_folder"))

        self.folder_form_button = ttk.Button(
            self.folder_input_frame,
            text="Browse",
            command=lambda: browse_folder(self.folder_form_entry),
        )
        self.folder_form_button.pack(side="right", padx=(5, 0))

        # Export button
        self.export_button = ttk.Button(
            self.main_frame,
            text="Export Emails",
            command=self.export_emails_with_progress,
        )

        # Quit button
        self.quit_button = ttk.Button(
            self.main_frame, text="Quit", command=self.root.destroy
        )

    def save_email_to_config(self, event=None):
        """Save the folder to config when user finishes editing"""
        folder = self.email_form_entry.get().strip()
        if folder:
            set_config("Email", "primary_email", folder)

    def start_outlook_connection(self):
        """Start the Outlook connection process"""
        # Validate folder input
        folder = self.email_form_entry.get().strip()
        if not folder:
            messagebox.showerror(
                "Error", "Please enter a folder before connecting to Outlook."
            )
            return

        # Validate date inputs
        start_date_str = self.start_date_entry.get().strip()
        end_date_str = self.end_date_entry.get().strip()

        if not start_date_str or not end_date_str:
            messagebox.showerror("Error", "Please enter both start and end dates.")
            return

        # Parse dates
        try:
            start_date = date.fromisoformat(start_date_str)
            end_date = date.fromisoformat(end_date_str)
        except ValueError:
            messagebox.showerror(
                "Error", "Please enter valid dates in YYYY-MM-DD format."
            )
            return

        # Save folder to config
        set_config("Email", "primary_email", folder)

        # Store date values for later use
        self.extract_start_date = start_date
        self.extract_end_date = end_date

        # Hide the email form and connect button
        self.email_form_frame.pack_forget()
        self.connect_button.pack_forget()

        # Show loading widgets
        self.create_loading_widgets_for_connection()

        # Start Outlook check after a short delay
        self.root.after(500, self.check_outlook)

    def create_loading_widgets_for_connection(self):
        """Create loading widgets for Outlook connection"""
        # Loading label
        self.loading_label = ttk.Label(
            self.main_frame,
            text="Checking Outlook connection...",
            font=("Arial", 12),
        )
        self.loading_label.pack(pady=20)

        # Progress bar
        self.progress = ttk.Progressbar(
            self.main_frame, mode="indeterminate", length=200
        )
        self.progress.pack(pady=10)
        self.progress.start()

    def check_outlook(self):
        """Check Outlook connection on main thread"""
        try:
            self.outlook_available = is_outlook_installed()
            if not self.outlook_available:
                self.error_message = "Outlook is not installed or not accessible"
        except Exception as e:
            self.outlook_available = False
            self.error_message = str(e)

        # Show result immediately
        self.show_result()

    def show_result(self):
        """Show the result of the Outlook check"""
        self.progress.stop()

        if self.outlook_available:
            self.loading_label.config(text="Outlook connection successful!")
            self.status_label.config(text="Loading flagged emails...")
            self.root.after(1000, self.load_flagged_emails)
        else:
            self.loading_label.config(text="Outlook connection failed!")
            self.status_label.config(
                text="Please ensure Outlook is installed and running."
            )
            self.root.after(2000, self.show_error)

    def load_flagged_emails(self):
        """Load flagged emails after Outlook connection is confirmed"""
        try:
            # Use stored date values if available, otherwise use default
            if hasattr(self, "extract_start_date") and hasattr(
                self, "extract_end_date"
            ):
                # Use the stored date objects directly
                (
                    self._flagged_emails_in_month,
                    self._start_of_month,
                    self._end_of_month,
                ) = get_flagged_emails_in_month(
                    self.extract_start_date, self.extract_end_date
                )
            else:
                # Use default if no date values are set
                (
                    self._flagged_emails_in_month,
                    self._start_of_month,
                    self._end_of_month,
                ) = get_flagged_emails_in_month()

            self.status_label.config(text="Ready to export")
            self.root.after(500, self.show_main_interface)
        except Exception as e:
            self.status_label.config(text=f"Error loading emails: {str(e)}")
            self.root.after(2000, self.show_main_interface)

    def show_main_interface(self):
        """Hide loading widgets and show main interface"""
        # Hide loading widgets
        self.loading_label.pack_forget()
        self.progress.pack_forget()
        self.status_label.pack_forget()

        # Show main widgets
        self.connection_status.pack(pady=10)
        self.flagged_emails_count_label.config(
            text=f"Total flagged emails between {self._start_of_month.strftime('%d %B %Y')} and {self._end_of_month.strftime('%d %B %Y')}: {len(self._flagged_emails_in_month)}"
        )
        self.flagged_emails_count_label.pack(pady=10)
        self.folder_form_frame.pack(pady=10, fill="x")
        self.export_button.pack(pady=10)
        self.quit_button.pack(pady=5)

    def show_error(self):
        """Show error message and exit"""
        messagebox.showerror(
            "Outlook Error",
            f"Could not connect to Outlook.\n\nError: {self.error_message}\n\nPlease ensure Outlook is installed and running, then restart the application.",
        )
        self.root.destroy()

    def export_emails_with_progress(self):
        """Export flagged emails with progress window"""
        if not self._flagged_emails_in_month:
            messagebox.showinfo("Export", "No flagged emails found for this month.")
            return

        # Create progress window
        progress_window = self.create_progress_window()

        # Start the copy process
        self.copy_emails_with_progress(self._flagged_emails_in_month, progress_window)

    def create_progress_window(self):
        """Create a progress window for the export process"""
        progress_window = Tk()
        progress_window.title("Exporting Emails")
        progress_window.geometry("500x250")
        progress_window.resizable(False, False)

        # Center the window
        progress_window.eval("tk::PlaceWindow . center")

        frame = ttk.Frame(progress_window, padding=20)
        frame.pack(expand=True, fill="both")

        # Progress label
        progress_label = ttk.Label(
            frame, text="Copying flagged emails...", font=("Arial", 12)
        )
        progress_label.pack(pady=(0, 10))

        # Progress bar
        progress_bar = ttk.Progressbar(frame, mode="determinate", length=400)
        progress_bar.pack(pady=10)

        # Status label
        status_label = ttk.Label(frame, text="", font=("Arial", 10))
        status_label.pack(pady=10)

        # Details label
        details_label = ttk.Label(frame, text="", font=("Arial", 9))
        details_label.pack(pady=5)

        # Store references
        progress_window.progress_bar = progress_bar
        progress_window.status_label = status_label
        progress_window.progress_label = progress_label
        progress_window.details_label = details_label

        return progress_window

    def copy_emails_with_progress(self, flagged_emails_in_month, progress_window):
        """Copy emails with progress updates"""
        try:
            # Get the flagged emails PST folder
            flagged_emails_root = get_flagged_emails_in_month_pst(
                self._start_of_month, self._end_of_month
            )

            # Change name
            flagged_emails_root.Name = f"Flagged Emails {self._start_of_month.strftime('%m-%d-%y')} - {self._end_of_month.strftime('%m-%d-%y')}.pst"

            existing_emails_in_store = set()

            # Get existing emails
            for email in flagged_emails_root.Items:
                existing_emails_in_store.add(email.Subject)

            total_emails = len(flagged_emails_in_month)
            copy_count = 0
            skip_count = 0

            # Update progress bar
            progress_window.progress_bar["maximum"] = total_emails

            for i, flagged_email in enumerate(flagged_emails_in_month):
                # Update progress
                progress_window.progress_bar["value"] = i + 1
                progress_window.progress_label.config(
                    text=f"Processing email {i + 1} of {total_emails}"
                )

                # Update status
                if flagged_email.Subject in existing_emails_in_store:
                    progress_window.status_label.config(
                        text=f"Skipping: {flagged_email.Subject[:60]}..."
                    )
                    progress_window.details_label.config(
                        text=f"Already exists in PST folder"
                    )
                    skip_count += 1
                else:
                    progress_window.status_label.config(
                        text=f"Copying: {flagged_email.Subject[:60]}..."
                    )
                    progress_window.details_label.config(
                        text=f"Moving to {flagged_emails_root.Name}"
                    )
                    flagged_email.Copy().Move(flagged_emails_root)
                    copy_count += 1

                # Update the GUI
                progress_window.update()

            # Show completion message
            progress_window.progress_label.config(text="Export completed!")
            progress_window.status_label.config(
                text=f"Copied {copy_count} emails, skipped {skip_count} duplicates"
            )
            progress_window.details_label.config(
                text=f"Total processed: {total_emails} emails"
            )

            # Close progress window after 3 seconds
            progress_window.after(3000, progress_window.destroy)

            # Show completion message
            messagebox.showinfo(
                "Export Complete",
                f"Successfully exported {copy_count} emails to PST.\nSkipped {skip_count} duplicates.\n\nPST file: {flagged_emails_root.Name}",
            )

        except Exception as e:
            progress_window.destroy()
            messagebox.showerror("Export Error", f"Error during export: {str(e)}")


if __name__ == "__main__":
    root = Tk()
    main_window = MainWindow(root, width=640)
    root.mainloop()
