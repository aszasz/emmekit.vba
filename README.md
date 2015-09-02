# Emmekit for VBA
One tool kit for reading, manipulating, and writting INRO EMME files

This tools were developed over the years working with EMME. They were originally in Fortran and I ported them to VBA so I could share with my coleagues...
 usually with strict instructions (and with me around).

The utilities are for reading reports and fit them in Excel spreadsheets for further analisys and fitting looks into final documents, but mainly for network 
manipulations -- including previous races data -- that, among others, include:
- Detour/cut/extend (some/all/within a distance) lines to terminal or corridor
- Create integration fare terminals every first transit lines encounter (and closer encounter among routes and centroids)
- Import GPX files and compute observed speeds to the links
- Read GTFS feeds and draw network from it (which required the bellow)
- Map matching (you have a line and a map, and have to decide if that line fits in the map and where... or does it go to a street that is not there).

In 2012, when I saw I needed to clean up the code in order to make it easily re-usable and articulated, I shifted to Haxe, 
under [Jonas](https://github.com/jonasmalacofilho) guidance. While the [EmmeKit for Haxe](https://github.com/jonasmalacofilho/emmekit.hx) is more powerfull also for share,
 using it has a considerably larger barrier, so I was asked by a few coleagues to organize this (so they don't need to bother me anymore to do my magic).

This was prepared within a few days, so I could announce the fufilling of my promise in the INRO Model City São Paulo 2015 (on Sept 1st) in the hope to reach a larger
audience...

... So this is an attempt to bring VBA code to git and document it, working around Excel by exporting the VBA modules, so we can see changes as texts.
This exported modules, which receive the extension ".bas" (".cls" for class modules) are exactly the edited code plus a single header line.

To do so, I organized (thus far only two) tools in a way that one can use without knowing anything about the code, one is very simple and the other
 is very complex (although not as complex as it can be when real applications are made). 

If you learn how this works, you can, you get and idea of
 how the code should to do it, then you can look at the code 
and create your own tools (even if you are a [first time user of VBA](#first-time-user)), and hopefully and share them here.


<a name="play-withorganized-tools"></a>
## Prepared tools:
###Network parameters:
The main gap to cross when you want to be able to manipulate a network, is having a subset of tools that can locate points in the map and calculate distances.
This require learning about how the network is represented:
	- how network coordinates relates to distances

###Lines Mover
Lines mover is a very simple tool to use, that will change transit lines itineraries based on an instruction file.
It requires 3 files for input:
	- emme network output file as exported by EMME with command 2.14
	- emme transit line output file as exported by EMME with command 2.24
	- a prepared file where you inform how transit lines must be changed, following the bellow sintax
		- One instruction per line, words are separated by space
		- first word of instruction is the EMME transit line number (=field line, up to six characters as) to be processed by the following command
			- one line may be processed several times
			- c (or C) as the first word indicates 'comment' and line is ignored (and therefore you never get to process a line whose number is "C" or "c" 
		- second word is the command with one or two characters:
			- first caracter can be:
				- 'X' or 'x' means cut points
				- '+' means add points to go to
				- '-' means remove line
			- the second caracter can be:
				- '>' after
				- '<' before
			- So, compounding we have:
				- 'X>' means delete all points after (if the line passes twice, it is the last pass)
				- 'X<' means delete all points before (if the line passes twice, it is the first pass)
				- '-' (minus) means delete line.
				- '+>' Add points in the END of the line thru the following points (uses shortest path)
				- '+<' Add points in the BEGIN of the line thru (uses shortest path)

and outputs 2 files:
	- emme network input file as exported by EMME with command 2.11: it is a differential file, that adds needed modes for the changes
	- emme transit line input file as exported by EMME with command 2.21: also a differential file
	

###Cones

Creates integration cones, as follows:
    - New nodes will try to have the same number end (above 100,000)
    - will create cones for all stops where a bus bellonging to a group with fare integration stops, with auxiliary mode given bellow and price is placed in ul3
    - uses the 800-900 range for types of the new links
    - will add a walking cost when making cones between two stops




<a name="play-with-code"></a>
## Playing with code:

<a name="overall-workflow"></a>
###Workflow overview

This manipulation tools work as almost all tools:
- Learn what and where are the input and output
- Read input
- Do stuff
- Output 

Usually, the input and/or the output are EMME formatted files. In the "do stuff" part, ideally a program must check if the input is what it
 should be and if it is consistent internally and with other input. According with the proposal of this tools, this is usually only checked;
figuring out what is wrong and solving it is let to the user.

When you are writting your own code, you will addapt it to fix it according to 
specific needs... ideally you should do it as a new and independent tool, so you can use it later. 

<!--- When you do a patch, as I used to
 and that is the main reason why making generic macros are much harder than making specific ones.
--->

### Module "Basic Network"
This module holds the code that deal with network parameters, distance and find tools.

### Module "Emme"
This module holds the functions that deal with reading and writting emme files, querying and changing properties
that are important to write on the output files.


<a name="first-time-user"></a>
## If you are a first time user of VBA:

For now you just should download the ".xlsm" file and, _after learning the bellow_, play with it.

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

To set this go to "File... Options... Trust Center... Trust Center Settings..." In Trust Center go to Macro Settings... and choose the suitable option
(which is likely to be 'Disable all macros with notification')

![Screeen of Excel Options to Trust Center](/assets/Screeen_Excel_options_TrustCenter.png)
![Screeen of Excel options in Trust Center Macros Settings](/assets/Screen_Office_Trust_Center.png)



* XML are HTML-on-steroids, but still text files: to understand this better you should try to rename an ".xlsx" or a ".docx" to _".zip"_ 
(possible only after unckecking the settings of your Windows Explorer in \[ \] 'Hide extensions of known file types' and disregard the warning 
that doing such a thing can affect how things work), unzip it and explore its contents.


An exercise to understand how this works:

- First, get the "Developer Ribbon" on your Excell
in File... Options... Customize Ribons... on the right side, check the box ''Developer''

![Screenshot_Excel_options](/assets/Screeen_Excel_options.png)

- Now you should have a new Ribbon on your Excel (Developer), where you will find the "Record Macro" button...

- Get yourself a new empty file to play with, and hit the button, fill the form about the macro you are about to create and hit "OK": after you do 
Excel will record your moves (and write the code for that in a place where you can change it and call it for execution)
until you hit the record button again (that shall be renamed to stop recordding).
	- Do a few operations: write text, change cell collors, border.
	- Hit stop recording button (that is both on the developer Tool, and (in Excel 2013) appears again on the bottom left, diguised as a square...

![Screen First_Macro](/assets/Screen_First_Macro.png)

![Screeen stop recording](/assets/Screeen_stop_recording.png)

	- To see this work, undo what you did and execute the macro by clicking in the second button of the Developer Ribbon "Macros".
	- To see the code, which you should do after recording a few more macros, use the first button of that Ribbon "Visual Basic", (equivalent 
to <kbd>Alt</kbd>+<kbd>F11</kbd>. You will be in the Visual Basic IDE, as explained bellow




To shift to the IDE in Excel (and shift back to Excel), press <kbd>Alt</kbd>+<kbd>F11</kbd>.

Usually, the uper-left corner will show you the Project Explorer: a typical Windows' folder structure,
 showing existing open workbooks and VBA projects associated with each one. If you don't see it, press  <kbd>Cltrl</kbd> + <kbd>R</kbd>.

There can be code written ''inside'' the workbook (inside each worksheet), to each project you can add forms and modules

![Screenshot_Project_Explorer](/assets/Screeen_Project_Explorer.png)



Well come to a different world.








- to deal with EMME input without problems you must set your Excel and your Windows to use "dots" as decimal separator.
Over the years I've learn that is easier to go back and forth changing this settings, based on your needs, than having your macros to deal with it.
(place settings for Windows on the Start Menu.)

(OBS: VBA does not respect the Excel settings for digits and thousand separator, so make sure you set Windows to 
use . (dot) for decimal digits and , (comma) for thousand separator, so outputs are recognized by EMME.


After while using the VBA IDE you will want to disable: Tools... Options... "Auto Sintax check". When sintax is wrong code lines goes red anyway.

Another useful thing is to click with the right button in th menu bar and select Customize... Edit... and find the commands (buttons) 'Comment blocks' and 'Uncomment blocks' and drag them to the tool bar and the Edit Menu, they are not shown anywhere by default and are very useful.

## Playing with the available prepared and organized tools:
