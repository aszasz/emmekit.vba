# emmekit.vba
One tool kit for reading manipulating and writting INRO EMME files

This tools were developed over the years working with EMME. They were originally in Fortran and I ported them to VBA so I could share with my coleagues... usually with strcit instructions and with me around.

The utilities include, among others:
- Import GPX files and compute observed speeds to the links
- Read (some) GTFS feeds
- Detour/cut/extend (some/all) lines to terminal or corridor (within a radius)
- Create integration fare terminals every first transit lines encounter 

So this is an attempt to bring VBA code to git and document it, working around Excel and exporting the VBA modules, so we can see changes as texts.


## If you are a first time user of VBA:

Visual Basic for Applications is basically a simplified Visual Basic interpreter (and run-time) embedded in MS-Office.
It has a very convenient and easy-to-learn-and-use Integrated Development Enviroment (IDE).
To shift to the IDE in Excel (and shift back to Excel), press <kbd>Alt</kbd>+<kbd>F11</kbd>.

Usually, the uper-left corner will show you the Project Explorer: a typical Windows' folder structure, showing existing open workbooks and VBA projects associated with each one. If you don't see it, press  <kbd>Cltrl</kbd> + <kbd>R</kbd>.

There can be code written ''inside'' the workbook (inside each worksheet), to each project you can add forms and modules

![Screenshot_Project_Explorer](/assets/Screeen_Project_Explorer.png)

An exercise to understand how this works:

- First, get the "Developer Ribbon" on your Excell
in File... Options... Customize Ribons... on the right side, check the box ''Developer''

![Screenshot_Excel_options](/assets/Screeen_Excel_options.png)


After while using the VBA IDE you will want to disable: Tools... Options... "Auto Sintax check". When sintax is wrong code lines goes red anyway.

Another useful thing is to click with the right button in th menu bar and select Customize... Edit... and find the commands (buttons) 'Comment blocks' and 'Uncomment blocks' and drag them to the tool bar and the Edit Menu, they are not shown anywhere by default and are very useful.

