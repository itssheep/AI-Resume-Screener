import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from pdfminer.high_level import extract_text
import webbrowser
import os
import json
import sys
import threading
import re
from datetime import datetime
import pandas as pd
import ai


APPNAME = "BrightIsle CV Screener"
APPDATADIR = os.path.join(os.getenv("APPDATA"), APPNAME)
CONFIGFILE = os.path.join(APPDATADIR, "config.json")
LOGFILE = os.path.join(APPDATADIR, "log.txt")


class Window(tk.Tk):
    def __init__(self, title="None", geometry="600x400", close=True, parent=None, isRoot=False):
        super().__init__()
        self.wm_title(title)
        self.wm_geometry(geometry)
        self.wm_resizable(False, False)
        self.parent = parent
        self.isRoot = isRoot

        if close:
            self.wm_protocol("WM_DELETE_WINDOW", self.closeApp)

        else:
            self.wm_protocol("WM_DELETE_WINDOW", lambda: None)

        try:
            #Get the absolute path to the .ico file
            iconPath = getPackagedPath("icon.ico")
            self.iconbitmap(iconPath)
        except Exception:
            pass # Fallback to default tkinter icon if brightisle icon cant be found
    
    def closeApp(self):
        if self.isRoot:
            sys.exit(0)
        elif self.parent:
            self.parent.deiconify()
            self.destroy()
        else:
            self.destroy()

def drop(event=None):
    """
    Adds PDF files to the listbox either via a file dialog.
    
    Args:
        event: event (None)
    """

    # Inserts files into the listbox
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

    if event is None: # If the user double clicks
        files = filedialog.askopenfilenames(
            title="Select .pdf Files",
            filetypes=[("PDF Files", "*.pdf")],
        )

        if files: # If files are valid
            insertFile(files)
    return

def delete(event=None):
    """
    Deletes selected items from the listbox.

    Args:
        event: Event object (optional).
    """

    items = listbox.curselection() # Grab selected file
    for i in reversed(items): 
        listbox.delete(i)                
    return

def readme():
    """
    Opens the README.md file in the default web browser.
    """
    readmePath = os.path.join(APPDATADIR, "README.md") # Get file location
    
    if os.path.exists(readmePath): # If it's found
        webbrowser.open(f"file://{readmePath}")
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
    try:
        criteria = userCriteria.get("1.0", tk.END).strip() # Get criteria from textbox
        strength = str(strengthSlider.get()) # Get strength from slider and turn into a string
        
            
        if not criteria: # Make sure the criteria is not falsy
            handleError(404)
            return

        if resume is None: # Submitting a null value to the ai will cause an error. it needs to be a string
            resume = "None"

        elif coverletter is None: # Same as resume
            coverletter = "None"
            
        if not apiKey: # Ensure api key is set for safety purposes
            handleError(402)
            root.closeApp()

        try:
            result = ai.main(name, resume, coverletter, criteria, strength) # Get result of AI script
            return result
        
        except Exception as e:
            logError(e)
    except Exception as e:
        logError(e)

def gatherPairs():
    """
    Returns a dictionary that maps namekey to resume & coverletter path
    """
    pairs = {}

    for i in range(listbox.size()): 
        filePath = listbox.get(i) # Grab ith file path
        docType, nameKey = parseType(filePath) # unpack tuple into type (resume/coverletter) and name

        if nameKey not in pairs: # If its a new name create a new dictionary entry
            pairs[nameKey] = {"Resume": None, "CoverLetter": None}

        if docType == "Resume": # If the name is in the dictionary edit the resume key to be the file path to the resume
            pairs[nameKey]["Resume"] = filePath
        elif docType == "CoverLetter": # Same as above but with coverletter (capital L is important)
            pairs[nameKey]["CoverLetter"] = filePath
        else: # else a weird glitch happened and we should skip this file
            continue

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
        404 : "Please enter a valid criteria before proceeding.",
        405 : "The application encountered a network issue. Please check your internet and try again.",
        406 : "Rate limit hit. Please wait before submitting a new request",
        407 : "Quota exceeded. Please contact support",
        408 : "Invalid API key entered, please check your key and try again.",
        409 : "Please add atleast one PDF file to process.",
        410 : "README.MD file not found.",
        411 : "Failed to process file, is the pdf scanned or empty?",
        999 : "Unknown Error, Please contact support",
    }
   
    message = errorMessages.get(err)
    logError(e)
    messagebox.showerror("Error", message)

def getPackagedPath(filename):
        """
        Gets the path of a file packaged with the application.

        Args:
            filename (str): Name of the file.

        Returns:
            str: Path to the file.
        """
        if hasattr(sys, '_MEIPASS'): # checks for sys attribute and grabs file packaged with pyinstaller
            return os.path.join(sys._MEIPASS, filename)
        else:
            return os.path.join(os.path.dirname(__file__), filename) # For VSCODE testing

def openLog(): 
    """
    Opens the log file for debugging purposes.
    """

    try:
        file = os.path.join(APPDATADIR, "log.txt") # Find log file in appdata dir
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

    os.makedirs(APPDATADIR, exist_ok=True) # Create appdata directory

    if not os.path.exists(CONFIGFILE):
        # First run: Create the config & log file and prompt user to set up their API key.
        with open(CONFIGFILE, "w") as file:
            json.dump({"OPENAI_API_KEY": ""}, file) # dump our json contents into config
            
        logFile = os.path.join(APPDATADIR, "log.txt") # Create log file
        with open(logFile, "w") as log: # Default write-to log
            log.write("Error log file, for developer use.\n")

        readmeDest = os.path.join(APPDATADIR, "README.md") # readme dest
        readmeSRC = getPackagedPath("README.md") # unpack file
        with open(readmeSRC, "r", encoding="utf-8") as src, open(readmeDest, "w", encoding="utf-8") as dest:
            dest.write(src.read())

        messagebox.showinfo("Setup Required", f"Config file created at {CONFIGFILE}. Please add your OpenAI API key.")
        sys.exit(0)
    else:
        # Regular validation
        with open(CONFIGFILE, "r") as file:
            try:
                config = json.load(file)
                apiKey = config.get("OPENAI_API_KEY", "").strip() # read api key from json
                if not apiKey: # ensure api key is there
                    handleError(402)
                    return 
                return apiKey
            except json.JSONDecodeError: # if json cannot be loaded show an error
                handleError(403)
                return 

def run():
    """
    Executes the resume screening process for all files in the listbox.
    """
    # Helper function to process files
    def processFiles():
        allOut = [] 
        allPaths = []
        for nameKey, docs in pairs.items():
            
            resumePath = docs.get("Resume") # Grab file path (value) from Resume (key) in dict
            coverPath = docs.get("CoverLetter") # Grab file path (value) from CoverLetter (key) in dict

            if not resumePath and not coverPath: # if both are somehow missing, skip the entry
                continue

            resumeText = pdfToPlaintext(resumePath) # Grab the text from the pdfs and turn into string plaintext
            coverText = pdfToPlaintext(coverPath) 
            
            if (not resumeText or resumeText.strip() == "") and (not coverText or coverText.strip() == ""): # Ensure plaintext is not falsy
                continue

            try:
                
                output = runAI(resumeText, coverText, nameKey) # run AI 
                if output is None: # Ensure output is not falsy
                    continue
                
                allOut.append(output) # append output to list
                allPaths.append((nameKey, resumePath, coverPath)) # append path info as tuple
                
            except Exception:
                root.after(0, loadingWindow.stop)
                root.after(0, lambda: handleError(999))
                root.after(0, root.destroy)
                return
        

        root.after(0, loadingWindow.stop) # Destroy loading bar
        root.after(0, lambda: showResultWindow(allOut, allPaths)) # Open results window

    
    if listbox.size() == 0: # Make sure there are files in the listbox to process
        handleError(409)
        return 
    
    if not userCriteria.get("1.0", tk.END).strip():
        handleError(404)
        return
    
    root.withdraw()
    loadingWindow = showLoadingBar() # open loading bar
    pairs = gatherPairs() # create dictionary

    threading.Thread(target=processFiles, daemon=True).start() # New thread for processing so GUI doesn't crash

def logError(error):
    """
    Logs an error message to the log file with a timestamp.

    This function appends error details to the log file, including the
    date and time when the error occurred.

    Args:
        error (Exception): The error to log.
    """
    timestamp = datetime.now().strftime("%d-%m-%Y %H:%M:%S") # records the date and time the error was found
    with open(LOGFILE, "a") as file:
        file.write(f"\n[{timestamp}] Error: {error}\n")
    return
  
def pdfToPlaintext(filePath):
    """
    Converts a PDF file to plaintext.

    Args:
        file_path (str): Path to the PDF file.

    Returns:
        str: Extracted plaintext or None if the PDF cannot be processed.
    """
    try:
        if filePath is None:  # Ensure file path is provided
            return None
        
        # Extract text from the PDF file using pdfminer
        text = extract_text(filePath)
        
        return text.strip() if text else None
    
    except Exception as e:
        handleError(411, e)
        return None

def parseType(filePath):
    """
    Given some file path, parses the type of file and matches to the other.
    
    returns:
        docType   "Resume" or "CoverLetter"
        nameKey   "name of person" (for dictionary key)
    """
    filename = os.path.basename(filePath) 
    base, _ = os.path.splitext(filename) # unpack tuple to grab full filename 
    parts = base.split("_", 2) # Splits to ["Type", "First-Last", "GetHired"]

    if len(parts) < 2: # If filename is invalid handle gracefully
        return ("Unknown", "Unknown")
    
    docType = parts[0] # Resume or coverletter
    nameKey = parts[1] # Name

    if nameKey.count('-') > 1: # If the person has a middle hyphenated name (this took me forever to debug)
        nameParts = nameKey.split('-') # Split by the -
        nameKey = f"{nameParts[0]}-{nameParts[-1]}" # Grab first and last entry (first and last names). 
        # This is not the best solution, i.e for names like Anna o'Keefe it shows it as Anna-O-Keefe which will simplify to Anna Keefe in the program.
        # I cannot figure out a better one. 

    return docType, nameKey

def showLoadingBar(parent=None):
    """
    Creates and displays a loading bar window.

    Returns:
        Window: The loading bar window.
    """

    loadingWindow = Window("Processing", "300x100", close=False, parent=parent)

    label = tk.Label(loadingWindow, text="Processing...") # Processing text
    label.pack(pady=10) # Add y dir padding to label

    progress = ttk.Progressbar(loadingWindow, mode="indeterminate") # Create loading bar
    progress.pack(pady=10, padx=10, fill="x") # Add padding 
    progress.start() # Start loading bar

    def stop():
        progress.stop()
        loadingWindow.destroy()

    loadingWindow.stop = stop
    return loadingWindow

def showResultWindow(result, file):
    """
    Displays the results of the resume screening process.

    Args:
        result (list): List of string AI output.
        file (list): List of tuples of file paths that correspond to each output.
    """
    global resultWindow

    resultWindow = Window("Screening Results", parent=root)
    

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
    detailsMap = {}

    # List to hold results for sorting
    results = []

    # Process results and extract information
    for doctuple, aioutput in zip(file, result):
        name, resumePath, coverPath = doctuple

        # Extract the score using a regex
        scorePattern = re.search(r"Score:\s*(\d+)", aioutput)
        score = int(scorePattern.group(1)) if scorePattern else -1

        # Extract the approval status using a regex
        appPattern = re.search(r"Rationale:\s*(Approved|Rejected)\.?", aioutput)
        approval = appPattern.group(1) if appPattern else "Unknown"

        # Create a preview with the file name, score, and approval status
        preview = f"{name} | {score} | {approval}"

        # Add to results list
        results.append((score, preview, aioutput, resumePath, coverPath))

    # Sort results by score in descending order
    results.sort(key=lambda x: x[0], reverse=True)

    # Populate the listbox and details map
    for score, preview, aioutput, resumePath, coverPath in results:
        detailsMap[preview] = (aioutput, resumePath, coverPath)  # Include paths in map
        resultsListbox.insert(tk.END, preview)

    # Function to export listbox data to Excel
    def export_to_excel():
        data = []
        for _, preview, aioutput, _, _ in results:
            name, scoreStr, approval = preview.split(" | ")
            rationale = aioutput.strip()
            data.append([name, scoreStr, approval, rationale]) # Grab all data from output and add to list

        df = pd.DataFrame(data, columns=["Name", "Score", "Approval", "Rationale"]) # Add to excel file using pandas
        savePath = filedialog.asksaveasfilename( # open file dialog to save
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if savePath: # If the file is saved
            df.to_excel(savePath, index=False)
            messagebox.showinfo("Export Successful", f"Results exported to {savePath}")

    # Function to show detailed view with buttons
    def showDetails(event):
        
        # Helper function to open file on button press.
        def openFile(filePath, fileType):
            if filePath and os.path.exists(filePath):
                os.startfile(filePath)
            else:
                messagebox.showinfo("File Not Found", f"{fileType} file could not be found.")

        selected = resultsListbox.curselection()
        if selected:
            preview = resultsListbox.get(selected[0]) # get current selected item
            detailed, resumePath, coverPath = detailsMap.get(preview, ("Details not found.", None, None)) # get details from dictionary

            # Create a detailed view window
            detailWindow = Window("Detailed Result")

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

            # Open Resume button
            openResumeButton = tk.Button(buttonFrame, text="Open Resume", command=lambda: openFile(resumePath, "Resume"))
            openResumeButton.pack(side="left", padx=10)

            # Open Coverletter button
            openCoverletterButton = tk.Button(buttonFrame, text="Open Coverletter", command=lambda: openFile(coverPath, "Coverletter"))
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
    - Adds widgets for user interaction, including:
        - A text box for entering screening criteria.
        - A listbox for dropping or selecting PDF files to process.
        - A slider to adjust the filter strength for AI screening.
        - Buttons for accessing help, running the screening, and viewing logs.

    Notes:
    - If the configuration file is missing or invalid, the user is prompted 
      to set up the necessary API key before proceeding.
    - The application is terminated if the API key is not provided.

    Raises:
        None
    """

    # Allow access to all variables that need to be read in other functions. They are never written in other functions and im too lazy to modify args for pbv.
    global root, apiKey, listbox, userCriteria, strengthSlider 

    # Check for API key    
    apiKey = checkConfig()

    root = Window("Brightisle CV Screener", isRoot=True)

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
    dropLabel = tk.Label(frame, text="Double click to add PDF files")
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
