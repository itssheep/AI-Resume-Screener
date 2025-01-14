import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import webbrowser
import os
import json
import pymupdf
import sys
import threading
import re
from datetime import datetime
import pandas as pd
import ai

# Directories and file locations
APP_NAME = "BrightIsle CV Screener"
APPDATA_DIR = os.path.join(os.getenv("APPDATA"), APP_NAME)
CONFIG_FILE = os.path.join(APPDATA_DIR, "config.json")
LOG_FILE = os.path.join(APPDATA_DIR, "log.txt")


def drop(event=None):
    """
    Adds PDF files to the listbox either via a file dialog.
    
    Args:
        event: event (None)
    """

    def insertFile(fileList): 
        for file in fileList:
            if file.endswith('.pdf'):
                if file not in addedFiles: # check for duplicates
                    listbox.insert(tk.END, file)
                    addedFiles.add(file)
                else:
                    messagebox.showinfo("Duplicate File", f"{os.path.basename(file)} is already in the dropbox.")
            else:
                handleError(401)

    addedFiles = set()

    if event is None: # Double click case
        files = filedialog.askopenfilenames(
            title="Select .pdf Files",
            filetypes=[("PDF Files", "*.pdf")],
        )

        if files:
            insertFile(files)
    return


def delete(event=None):
    """
    Deletes selected items from the listbox.

    Args:
        event: Event object (optional).
    """

    items = listbox.curselection()
    for i in reversed(items):
        listbox.delete(i)                
    return


def getPackagedPath(filename):
    """
    Gets the path of a file packaged with the application.

    Args:
        filename (str): Name of the file.

    Returns:
        str: Path to the file.
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    else:
        return os.path.join(os.path.dirname(__file__), filename)


def readme():
    """
    Opens the README.md file in the default web browser.
    """
    readme_path = os.path.join(APPDATA_DIR, "README.md") # Find absolute path of the file
    
    if os.path.exists(readme_path):
        webbrowser.open(f"file://{readme_path}")
    else: # if the file cannot be found
        handleError(410)
    return


def runAI(resume, coverletter, name):
    """
    Runs the AI subprocess for analyzing the resume.

    Args:
        resume (str/None): Plaintext content of the resume.
        coverletter (str/None): Plaintext content of coverletter.

    Returns:
        str: Output from the AI script, or None if an error occurs.
    """

    criteria = userCriteria.get("1.0", tk.END).strip()
    strength = str(strengthSlider.get())

    if not criteria:
        handleError(404)
        return

    if resume is None:
        resume = "None"

    elif coverletter is None:
        coverletter = "None"

    # Ensure API key is set again just in case
        
    if not api_key:
        handleError(402)
        closeApp(root)

    try:
        result = ai.main(name, resume, coverletter, criteria, strength)

        return result
    except Exception as e:
        handleError(999, e)


def gatherPairs():
    """
    Returns a dictionary that maps namekey to resume & coverletter path
    """
    pairs = {}

    for i in range(listbox.size()):
        filePath = listbox.get(i)
        docType, nameKey = parseType(filePath)

        if nameKey not in pairs:
            pairs[nameKey] = {"Resume": None, "CoverLetter": None}

        if docType == "Resume":
            pairs[nameKey]["Resume"] = filePath
        elif docType == "CoverLetter":
            pairs[nameKey]["CoverLetter"] = filePath
        else:
            pass
    
    return pairs


def handleError(err, e=None): 
    """
    Handles application-specific errors and displays messages.

    Args:
        err (int): Error code.
        f (str): Optional filename for file-specific errors.
    """

    errorMessages = {
        400 : "Log file not found. Log function only works on Windows OS.",
        401 : "Only PDF files are allowed.",
        402 : "API key missing in config.json. Click Help for help.",
        403 : "Config file is not a valid JSON. Delete the file and try again.",
        404 : "Criteria was not entered.",
        405 : "The application encountered a network issue. Please check your internet and try again.",
        406 : "Rate limit hit. Please wait before submitting a new request",
        407 : "Quota exceeded. Please contact support",
        408 : "Invalid API key entered, please check your key and try again.",
        409 : "Please add atleast one PDF file to process.",
        410 : "README.MD file not found.",
        411 : f"Failed to process file is the pdf scanned or empty?",
        999 : "Unknown Error, Please contact support",
    }
   
    message = errorMessages.get(err)
    logError(e)
    messagebox.showerror("Error", message)
    

def openLog(): 
    """
    Opens the log file for debugging purposes.
    """

    try:
        file = os.path.join(APPDATA_DIR, "log.txt")
        if os.path.exists(file):
            os.startfile(file)

    except Exception as e: # this is necessary as if its on another OS the file variable line will create an error.
        handleError(400, e)
    return


def checkConfig():
    """
    Ensures the configuration file exists, and validates the API key.

    Returns:
        str: API key 
    """

    os.makedirs(APPDATA_DIR, exist_ok=True)

    if not os.path.exists(CONFIG_FILE):
        # First run: Create the config & log file and prompt user to set up their API key.
        with open(CONFIG_FILE, "w") as file:
            json.dump({"OPENAI_API_KEY": ""}, file)
            
        logFile = os.path.join(APPDATA_DIR, "log.txt")
        with open(logFile, "w") as log:
            log.write("Error log file, for developer use.\n")

        readmeDest = os.path.join(APPDATA_DIR, "README.md") 
        readmeSRC = getPackagedPath("README.md")
        with open(readmeSRC, "r", encoding="utf-8") as src, open(readmeDest, "w", encoding="utf-8") as dest:
            dest.write(src.read())

        messagebox.showinfo("Setup Required", f"Config file created at {CONFIG_FILE}. Please add your OpenAI API key.")
        sys.exit()
    else:
        # Regular validation
        with open(CONFIG_FILE, "r") as file:
            try:
                config = json.load(file)
                api_key = config.get("OPENAI_API_KEY", "").strip()
                if not api_key: # ensure api key is there
                    handleError(402)
                    return 
                return api_key
            except json.JSONDecodeError:
                handleError(403)
                return 


def run():
    """
    Executes the resume screening process for all files in the listbox.
    """
    def processFiles():
        allOut = []
        allPaths = []

        for nameKey, docs in pairs.items():
            
            resumePath = docs.get("Resume")
            coverPath = docs.get("CoverLetter")

            if not resumePath and not coverPath: # if both are somehow missing, skip the entry
                continue

            resumeText = pdfToPlaintext(resumePath) 
            coverText = pdfToPlaintext(coverPath) 

            if (not resumeText or resumeText.strip() == "") and (not coverText or coverText.strip() == ""):
                continue

            try:
                output = runAI(resumeText, coverText, nameKey)
                if output is None:
                    continue

                allOut.append(output)
                allPaths.append((nameKey, resumePath, coverPath))

            except Exception:
                loadingWindow.destroy()
                handleError(999)
                root.destroy()
                return

        loadingWindow.destroy()
        showResultWindow(allOut, allPaths)

    
    if listbox.size() == 0: # Make sure there are files in the listbox to process
        handleError(409)
        return 
    
    loadingWindow = showLoadingBar()
    pairs = gatherPairs()

    threading.Thread(target=processFiles).start()
    

def logError(error):
    """
    Logs an error message to the log file with a timestamp.

    This function appends error details to the log file, including the
    date and time when the error occurred.

    Args:
        error (Exception): The error to log.
    """
    timestamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S") # records the date and time the error was found
    with open(LOG_FILE, "a") as file:
        file.write(f"\n[{timestamp}] Error: {error}\n")
    return

    
def pdfToPlaintext(file_path):
    """
    Converts a PDF file to plaintext.

    Args:
        file_path (str): Path to the PDF file.

    Returns:
        str: Extracted plaintext or None if the PDF cannot be processed.
    """

    text = ""
    
    try:
        if file_path is None: # Ensure file path is provided
            return None
        
        with pymupdf.open(file_path) as document:
            for page in document:
                text += page.get_text()
        
        return text
    
    except Exception as e:
        handleError(411, e)
        return None


def parseType(filePath):
    """
    Given some file path, parses the type of file and matches to the other.
    
    returns:
        docType   "Resume" or "Coverletter"
        nameKey   "name of person" (for dictionary key)
    """
    filename = os.path.basename(filePath)
    base, ext = os.path.splitext(filename)
    parts = base.split("_", 2) # Splits to ["Type", "First-Last", "GetHired"]

    if len(parts) < 2: # If filename is invalid
        return ("Unknown", "Unknown")
    
    docType = parts[0] # Resume or coverletter
    nameKey = parts[1] # Name

    return docType, nameKey


def closeApp(thread):
    """Ensures proper cleanup of application"""
    thread.destroy()
    if thread == root:
        sys.exit(0)


def showLoadingBar():
    """
    Creates and displays a loading bar window.

    Returns:
        tk.Toplevel: The loading bar window.
    """

    loadingWindow = tk.Toplevel(root)
    loadingWindow.title("Processing")
    loadingWindow.geometry("300x100")
    loadingWindow.resizable(False, False)
    loadingWindow.protocol("WM_DELETE_WINDOW", lambda: None)

    try:
        # Get the absolute path to the .ico file
        iconPath = getPackagedPath("icon.ico")
        loadingWindow.iconbitmap(iconPath)
    except Exception:
        pass # Fallback to default tkinter icon if brightisle icon cant be found

    label = tk.Label(loadingWindow, text="Screening in progress. Please wait...")
    label.pack(pady=10)

    progress = ttk.Progressbar(loadingWindow, mode="indeterminate")
    progress.pack(pady=10, padx=10, fill="x")
    progress.start()
    return loadingWindow


def showResultWindow(result, file):
    """
    Displays the results of the resume screening process.

    Args:
        result (list): List of string AI output.
        file (list): List of tuples of file paths that correspond to each output.
    """

    # New window
    resultWindow = tk.Toplevel()
    resultWindow.title("Screening Results")
    resultWindow.geometry("600x400")
    resultWindow.resizable(False, False)
    resultWindow.protocol("WM_DELETE_WINDOW", lambda: closeApp(resultWindow))

    try:
        # Get the absolute path to the .ico file
        iconPath = getPackagedPath("icon.ico")
        resultWindow.iconbitmap(iconPath)
    except Exception:
        pass  # Fallback to default tkinter icon if brightisle icon can't be found

    # Frame to hold listbox and scroller
    frame = tk.Frame(resultWindow)
    frame.pack(padx=10, pady=10, fill="both", expand=True)

    # Scrollbar for the listbox
    scrollbar = tk.Scrollbar(frame, orient="vertical")
    scrollbar.pack(side="right", fill="y")

    # Listbox to display results
    resultsListbox = tk.Listbox(frame, width=80, height=10, yscrollcommand=scrollbar.set)
    resultsListbox.pack(side="left", fill="both", expand=True)

    # Connect scrollbar to listbox
    scrollbar.config(command=resultsListbox.yview)

    # Dictionary to map previews to detailed content
    details_map = {}

    # List to hold results for sorting
    results = []

    # Process results and extract information
    for doctuple, aioutput in zip(file, result):
        name, resumePath, coverPath = doctuple

        # Extract the score using a regex
        score_pattern = re.search(r"Score:\s*(\d+)", aioutput)
        score = int(score_pattern.group(1)) if score_pattern else -1

        # Extract the approval status using a regex
        app_pattern = re.search(r"Rationale:\s*(Approved|Rejected)\.?", aioutput)
        approval = app_pattern.group(1) if app_pattern else "Unknown"

        # Create a preview with the file name, score, and approval status
        preview = f"{name} | {score} | {approval}"

        # Add to results list
        results.append((score, preview, aioutput, resumePath, coverPath))

    # Sort results by score in descending order
    results.sort(key=lambda x: x[0], reverse=True)

    # Populate the listbox and details map
    for score, preview, aioutput, resumePath, coverPath in results:
        details_map[preview] = (aioutput, resumePath, coverPath)  # Include paths in map
        resultsListbox.insert(tk.END, preview)

    # Function to export listbox data to Excel
    def export_to_excel():
        data = []
        for score, preview, aioutput, _, _ in results:
            name, score_str, approval = preview.split(" | ")
            rationale = aioutput.strip()
            data.append([name, score_str, approval, rationale])

        df = pd.DataFrame(data, columns=["Name", "Score", "Approval", "Rationale"])
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Export Successful", f"Results exported to {save_path}")

    # Function to show detailed view with buttons
    def showDetails(event):
        selected = resultsListbox.curselection()
        if selected:
            preview = resultsListbox.get(selected[0])
            detailed, resumePath, coverPath = details_map.get(preview, ("Details not found.", None, None))

            # Create a detailed view window
            detailWindow = tk.Toplevel(resultWindow)
            detailWindow.title("Detailed Result")
            detailWindow.geometry("600x400")
            detailWindow.resizable(True, True)
            detailWindow.protocol("WM_DELETE_WINDOW", lambda: closeApp(detailWindow))
            
            try:
                # Get the absolute path to the .ico file
                iconPath = getPackagedPath("icon.ico")
                detailWindow.iconbitmap(iconPath)
            except Exception:
                pass  # Fallback to default tkinter icon if brightisle icon can't be found

            # Scrollbar for detailed content
            scroll = tk.Scrollbar(detailWindow, orient="vertical")
            scroll.pack(side="right", fill="y")

            # Text widget for detailed content
            detailText = tk.Text(detailWindow, wrap="word", yscrollcommand=scroll.set, bg="white", font=("Arial", 10), height=15)
            detailText.pack(padx=10, pady=10, fill="both", expand=True)
            scroll.config(command=detailText.yview)

            # Insert detailed content into the Text widget
            detailText.insert("1.0", detailed)
            detailText.config(state="disabled")

            # Button frame
            buttonFrame = tk.Frame(detailWindow)
            buttonFrame.pack(pady=10)

            def open_file(filePath, fileType):
                if filePath and os.path.exists(filePath):
                    os.startfile(filePath)
                else:
                    messagebox.showinfo("File Not Found", f"{fileType} file could not be found.")

            # Open Resume button
            openResumeButton = tk.Button(
                buttonFrame, text="Open Resume", command=lambda: open_file(resumePath, "Resume")
            )
            openResumeButton.pack(side="left", padx=10)

            # Open Coverletter button
            openCoverletterButton = tk.Button(
                buttonFrame, text="Open Coverletter", command=lambda: open_file(coverPath, "Coverletter")
            )
            openCoverletterButton.pack(side="left", padx=10)

    # Add Export button at the bottom-right corner of the main results window
    exportButton = tk.Button(resultWindow, text="Export", command=export_to_excel)
    exportButton.pack(side="bottom", anchor="se", padx=10, pady=10)

    # Bind double-click and Enter to show details
    resultsListbox.bind("<Double-1>", showDetails)
    resultsListbox.bind("<Return>", showDetails)


def main():
    """
    Initialize and launch the BrightIsle CV Screener application.

    This function sets up the main GUI window and its components for the 
    application. It performs the following tasks:
    - Verifies the existence and validity of the configuration file (config.json),
      and ensures an API key is provided.
    - Initializes the TkinterDnD root window and configures its layout.
    - Adds widgets for user interaction, including:
        - A text box for entering screening criteria.
        - A listbox for dropping or selecting PDF files to process.
        - A slider to adjust the filter strength for AI screening.
        - Buttons for accessing help, running the screening, and viewing logs.
    - Binds actions such as drag-and-drop, file selection, and file deletion 
      to the appropriate event handlers.

    Notes:
    - If the configuration file is missing or invalid, the user is prompted 
      to set up the necessary API key before proceeding.
    - The application is terminated if the API key is not provided, except 
      during the first run.

    Raises:
        None
    """

    # Allow access to all variables that need to be read in other functions.
    global root, api_key, listbox, userCriteria, strengthSlider 

    # Check for API key    
    api_key = checkConfig()

    os.environ['TKDND_LIBRARY'] = getPackagedPath("tkdnd2.9")

    root = tk.Tk()

    root.title("BrightIsle CV Screener")
    root.geometry("600x400")
    root.resizable(False, False)
    root.protocol("WM_DELETE_WINDOW", lambda: closeApp(root))


    try:
        # Get the absolute path to the .ico file
        iconPath = getPackagedPath("icon.ico")
        root.iconbitmap(iconPath)
    except Exception:
        pass # Fallback to default tkinter icon if brightisle icon cant be found

    # Create a frame around the window to allow widgets to mesh better
    frame = tk.Frame(root, padx=10, pady=10)
    frame.grid(column=0, row=0, sticky="nsew")

    # Label for criteria textbox
    criteriaLabel = tk.Label(frame, text="Screening Criteria")
    criteriaLabel.grid(column=0, row=4, sticky="w")

    # Criteria textbox
    userCriteria = tk.Text(frame, width=45, height=8)
    userCriteria.grid(column=0, row=5, padx=5, pady=5, sticky="w")

    # Add a label for the DnD box
    dropLabel = tk.Label(frame, text="Drop PDF files here")
    dropLabel.grid(column=0, row=2, padx=5, pady=5, sticky="sw")

    # Create listbox to drop resumes into
    listbox = tk.Listbox(frame, width=60, height=8)
    listbox.grid(column=0, row=3, padx=5, pady=5, sticky="sw")

    # Open file explorer when clicked, and delete added files when backspace/delete is pressed on a selected file
    listbox.bind("<Double-1>", lambda event: drop())
    listbox.bind("<Delete>", delete)
    listbox.bind("<BackSpace>", delete)

    
    # Add slider for how rigorous the AI will be
    strengthSlider = tk.Scale(frame, from_=1, to=5, orient="horizontal")
    strengthSlider.grid(column=3, row=5, padx=5, sticky="e")
    strengthSlider.set(3) # Default value

    # Add label for slider
    strengthLabel = tk.Label(frame, text="Filter Strength")
    strengthLabel.place(x=480, y=267) # Place because I can't find a nice grid location for this one. sorry (not sorry)

    # Add help button
    helpButton = tk.Button(frame, text="Help", command=readme)
    helpButton.grid(column=3, row=0, sticky='ne')

    # Add RunAI button
    aiButton = tk.Button(frame, text="Run", command=run)
    aiButton.grid(column=3, row=0, sticky='n')

    # Add Log button
    logButton = tk.Button(frame, text="Log", command=openLog)
    logButton.grid(column=3, row=0, sticky='nw')

    # Configure rows and columns inside the frame
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_rowconfigure(1, weight=1)
    frame.grid_rowconfigure(2, weight=0)
    frame.grid_rowconfigure(3, weight=0)
    frame.grid_rowconfigure(4, weight=0)
    frame.grid_rowconfigure(5, weight=0)
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_columnconfigure(1, weight=1)
    frame.grid_columnconfigure(2, weight=0)
    frame.grid_columnconfigure(3, weight=0)

    # Configure rows and columns inside the window
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # Open the window
    root.mainloop()


if __name__ == "__main__":
    main() 
