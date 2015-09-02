# emmekit.vba
One tool kit for reading manipulating and writting INRO EMME files

This tools were developed over the years working with EMME. They were originally in Fortran and I ported them to VBA so I could share with my coleagues... usually with strcit instructions and with me around.

The utilities include, among others:
- Detour/cut/extend (some/all) lines to terminal or corridor
- Create integration fare terminals every first transit lines encounter 
- Import GPX files and compute observed speeds to the links
- Read (some) GTFS feeds

So this is an attempt to bring VBA code to git and document it, working around Excel by exporting the VBA modules, so we can see changes as texts.
This exported modules, which receive the extension ".bas" (".cls" for class modules) are exactly the eddited code plus a single header line.


## If you are a first time user of VBA:

For now you just should download the ".xlsm" file and, after learning the bellow, play with it.

Visual Basic for Applications is basically a simplified Visual Basic interpreter (and run-time) embedded in MS-Office.
It has a very convenient and easy-to-learn-and-use Integrated Development Enviroment (IDE). This means you can write programs inside
 the Excel (and other office files) to automate the use of that application (and others). This programs are called MACROS, as are EMME macros.

Before MS-Office 2007, simple spreadsheet files were all ".xls"  After 2007, they were divided as follows:
- ".xls" became ".xlsb" (b for binary), even though ".xls" and ".xlsb" are internally different, they are both binaries, and:
	- both can have macros embedded (or not)
- ".xlsx" and ".xlsm" are a 'tree of zipped XML files'*, they are usually larger than the equivalent binaries and:
	- ".xlsx" do not have macros:
		- they are called macro free workbooks
		- if you save an ".xlsx" after you created a macro inside it, Excel will throw it away, after giving you proper warning and suggesting other formats to save it.
	- ".xlsm" files are "macro-enabled workbooks"
		- this can have macros embedded (or not)
		
It does not matter which fomat (even if it is called macro-enabled), when you oppen a file with embbeded macros, depending on your 'macro security settings' you may:
	- be asked if you want-to enable macros
	- have macros ignored
	- have macros automatically enabled (this has security concerns, because macros can change things beyond the file scope and beyond the application (Excel scope),
 even when openning the file)

To set this go to "File... Options... Trust Center..." In Trust Center go to Macro Settings...

![Screeen of Excel options](/assets/Screeen_Excel_options.png)
![Screeen of Excel Trsut Center](/assets/Screeen_Excel_options_TrustCenter.png)



* XML are HTML-on-steroids, but still text files: to understand this better you should try to rename an ".xlsx" or a ".docx" to _".zip"_ 
(possible only after unckecking the settings of your Windows Explorer in \[ \] 'Hide extensions of known file types' and disregard the warning 
that doing such a thing can affect how things work), unzip it and explore its contents.



To shift to the IDE in Excel (and shift back to Excel), press <kbd>Alt</kbd>+<kbd>F11</kbd>.

Usually, the uper-left corner will show you the Project Explorer: a typical Windows' folder structure,
 showing existing open workbooks and VBA projects associated with each one. If you don't see it, press  <kbd>Cltrl</kbd> + <kbd>R</kbd>.

There can be code written ''inside'' the workbook (inside each worksheet), to each project you can add forms and modules

![Screenshot_Project_Explorer](/assets/Screeen_Project_Explorer.png)

An exercise to understand how this works:

- First, get the "Developer Ribbon" on your Excell
in File... Options... Customize Ribons... on the right side, check the box ''Developer''

![Screenshot_Excel_options](/assets/Screeen_Excel_options.png)


After while using the VBA IDE you will want to disable: Tools... Options... "Auto Sintax check". When sintax is wrong code lines goes red anyway.

Another useful thing is to click with the right button in th menu bar and select Customize... Edit... and find the commands (buttons) 'Comment blocks' and 'Uncomment blocks' and drag them to the tool bar and the Edit Menu, they are not shown anywhere by default and are very useful.

