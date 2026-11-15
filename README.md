### ROLL CONVERTER





Turns a list of roll numbers into student names, using the student list in

roster.csv. Works on Windows and macOS.

***\*The student list must be obtained from system admin and edited to fit the***

***format specified in "replacing the student list each year" section.***






###### TO RUN - WINDOWS


Double-click "Run" (Run.bat).



If Windows says Python is required, install it from

https://www.python.org/downloads/ and tick "Add python.exe to PATH"

during the setup. Then double-click "Run" again.




###### 

###### TO RUN - MAC

1\. Double-click "Run.command".

2\. If macOS blocks it the first time, right-click "Run.command",

&#x20;  choose "Open", then click "Open" in the dialog.



If the file does not open at all, the terminal needs permission once.

Open Terminal, type "chmod +x " (with a space after it), drag the

Run.command file into the window, press Enter. Then double-click it again.



If Python is missing, macOS offers to install it (click Install, wait,

then double-click Run.command again).



You can also always run it from Terminal:

&#x20;   cd <the folder containing rollconverter.py>

&#x20;   python3 rollconverter.py




###### 

###### HOW TO USE


1\. Type or paste roll numbers in the INPUT box.

&#x20;  Separate them with commas, spaces or new lines.

2\. Names appear in the OUTPUT box as you type.

3\. Click "Copy names" to put them on the clipboard.



Anything that is not a roll number is passed straight through to the output,

so you can paste a note or a heading along with the rolls and it stays

as you typed it.



A roll number that is not in the list stays on screen and is shown in

YELLOW. "Not found" in the small label above tells you how many there are.



Numbers are treated as roll numbers when they have the same number of

digits as the rolls in your list (four digits, for the current list). Any

other number is just treated as ordinary text.




AUTO FORMAT


The switch next to "Copy names" only edits inside each line.



&#x20; ON  (default)  roll numbers sitting next to each other on one line are

&#x20;                separated with commas:

&#x20;                abc abc abc      ->  name, name, name



&#x20; OFF            your own commas and spaces are kept exactly as typed:

&#x20;                abc abc abc      ->  name name name



(Either way, line breaks are never touched and other text is never touched.

A list typed one roll per line comes back one name per line in both modes.)



Copy takes exactly what is shown, either way.



Shortcuts:  Esc = clear input    Ctrl+O / Cmd+O = change student list

&#x20;           Ctrl+Shift+C / Cmd+Shift+C = copy names






###### REPLACING THE STUDENT LIST EACH YEAR


You do not need to edit any file by hand.



1\. Click "Change roster" in the top right.

2\. Pick the new spreadsheet (.xlsx, .csv or .tsv).

3\. Check the summary it shows ("Found 806 students...") and confirm.



The app keeps the chosen list as roster.csv next to the program, so it is

there the next time you open it. The original spreadsheet is never changed.

The previous list is saved as roster.backup.csv in the same folder.



The spreadsheet needs two columns: the roll number, then the student name.

A header row, extra columns, and either column order are all handled. If

roll numbers ever have leading zeros, they are kept as text.






FILES


rollconverter.py   the whole program (Python, standard library only)

roster.csv         the student list currently in use

RollConverter.exe  (optional) standalone Windows app, if built

Run.bat            Windows launcher

Run.command        macOS launcher

README.txt         this file






NOTES FOR WHOEVER MAINTAINS THIS


Check the program loads the list correctly after an edit:

&#x20;   python rollconverter.py --selftest



The student list is looked up in this order: roster.csv next to the

program, then the old "rollnumberswithnames.csv" name (which gets

converted automatically). If the program folder is read-only (for example

inside a macOS .app bundle), roster.csv is kept in the user's application

data folder instead.



The app keeps no settings file: the roster is the only state, and the

"Auto format" switch resets to on each launch. Nothing is written outside

this folder unless the folder itself is read-only.



It starts in well under a second and uses no third-party packages, so

there is nothing extra to install or update.



OPTIONAL: build a standalone app with no Python install required.

&#x20;   pip install pyinstaller

&#x20;   pyinstaller --onefile --windowed --name RollConverter rollconverter.py



That build must be made on the system it is for: run it on Windows to get

RollConverter.exe, run it on a Mac to get RollConverter.app/.









*\~ made by Prayaash Khakholia*

*Academics Captain and ICT Captain '27*

