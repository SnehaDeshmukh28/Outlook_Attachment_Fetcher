# Outlook_Attachment_Fetcher

## Overview
**Outlook_Attachment_Fetcher** is a desktop application designed to automate the process of fetching email attachments from a specified Outlook folder. With a clean and modern user interface, users can effortlessly download attachments by providing necessary inputs like their Outlook account name, folder name, and destination folder.

---

## Features
- **Custom Inputs**:
  - Specify your Outlook account name.
  - Input the folder name from which to fetch emails.
  - Choose a destination folder for saving attachments.
- **Automatic Functionality**:
  - Reads all unread emails in the specified folder.
  - Downloads attachments and marks emails as read.
- **User-Friendly Design**:
  - Simple and intuitive interface with modern styling.
  - Clear notifications for success or errors.

---

## Requirements
- **Operating System**: Windows
- **Dependencies**:
  - `pypiwin32` (for Outlook interaction)
  - `tkinter` (for GUI development)

Install dependencies using:
```bash
pip install pypiwin32
```

---

## How to Use
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/YourUsername/Outlook_Attachment_Fetcher.git
   cd Outlook_Attachment_Fetcher
   ```
2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```
3. **Run the Application**:
   Execute the application with:
   ```bash
   python app.py
   ```
4. **Input Details**:
   - Enter your Outlook email account name.
   - Specify the folder name in Outlook (e.g., "Inbox").
   - Select the destination folder for attachments.
5. **Download Attachments**:
   - Click the "Download Attachments" button to fetch and save the email attachments.


## License
This project is licensed under the **MIT License**. Feel free to modify and distribute it as per the terms of the license.