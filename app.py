import os
import win32com.client
from tkinter import Tk, Label, Entry, filedialog, messagebox
from tkinter.ttk import Button, Style, Frame


def download_attachments(email_account, folder_name, save_folder):
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # Access the specified Outlook folder
        outlook_folder = outlook.Folders.Item(email_account).Folders(folder_name)

        msg_count = 0  # Counter for processed messages

        # Process unread emails
        for email in outlook_folder.Items:
            if email.Class == 43 and email.UnRead:  # 43 corresponds to MailItem class
                if email.Attachments.Count > 0:
                    for attachment in email.Attachments:
                        attachment.SaveAsFile(os.path.join(save_folder, attachment.FileName))
                email.UnRead = False  # Mark email as read
                msg_count += 1

        # Notify user of success
        messagebox.showinfo("Success", f"Processed {msg_count} unread emails.\nAttachments saved to:\n{save_folder}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")


def select_folder(entry_field):
    """Helper function to open a folder dialog and set the path."""
    folder_path = filedialog.askdirectory(title="Select Folder")
    if folder_path:
        entry_field.delete(0, "end")  # Clear the field
        entry_field.insert(0, folder_path)  # Set the selected folder


def run_app():
    # Create the main application window
    app = Tk()
    app.title("Outlook Attachment Downloader")
    app.geometry("640x460")  # Adjusted geometry to fit all elements comfortably
    app.resizable(False, False)
    app.configure(bg="#f7f9fc")  # Light background color

    # Style for modern buttons
    style = Style()
    style.configure(
        "TButton",
        font=("Arial", 12, "bold"),
        padding=8,
        background="#4CAF50",  # Green background
        foreground="black",    # White text
        borderwidth=0
    )
    style.map(
        "TButton",
        background=[("active", "#45a049")],  # Slightly darker green when hovered
        foreground=[("active", "black")]
    )

    # Header
    Label(app, text="Outlook Attachment Downloader", font=("Arial", 18, "bold"), bg="#f7f9fc", fg="#333333").pack(pady=20)

    # Form container
    form_frame = Frame(app, style="TFrame", padding=(10, 10))
    form_frame.pack(pady=10, padx=20, fill="x")

    # User Email Field
    Label(form_frame, text="Enter Your Outlook Email Account Name:", font=("Arial", 12), bg="#f7f9fc", fg="#333333").grid(row=0, column=0, sticky="w", pady=5)
    email_entry = Entry(form_frame, font=("Arial", 12), width=45, relief="solid", borderwidth=1)
    email_entry.grid(row=1, column=0, padx=5, pady=5)

    # Folder Name Field
    Label(form_frame, text="Enter the Outlook Folder Name:", font=("Arial", 12), bg="#f7f9fc", fg="#333333").grid(row=2, column=0, sticky="w", pady=5)
    folder_entry = Entry(form_frame, font=("Arial", 12), width=45, relief="solid", borderwidth=1)
    folder_entry.grid(row=3, column=0, padx=5, pady=5)

    # Save Folder Field
    Label(form_frame, text="Select Destination Folder for Attachments:", font=("Arial", 12), bg="#f7f9fc", fg="#333333").grid(row=4, column=0, sticky="w", pady=5)
    save_folder_entry = Entry(form_frame, font=("Arial", 12), width=35, relief="solid", borderwidth=1)
    save_folder_entry.grid(row=5, column=0, padx=5, pady=5, sticky="w")
    Button(form_frame, text="Browse", command=lambda: select_folder(save_folder_entry), style="TButton").grid(row=5, column=1, padx=10, pady=5, sticky="e")

    # Download Button
    Button(
        app,
        text="Download Attachments",
        command=lambda: download_attachments(
            email_entry.get().strip(), folder_entry.get().strip(), save_folder_entry.get().strip()
        ),
        style="TButton"
    ).pack(pady=20)

    # Footer Note
    Label(app, text="Built with ❤️ for Anil Deshmukh", font=("Arial", 10), bg="#f7f9fc", fg="#555555").pack(side="bottom", pady=10)

    # Start the application
    app.mainloop()


if __name__ == "__main__":
    run_app()
