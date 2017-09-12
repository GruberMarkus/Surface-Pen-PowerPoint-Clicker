# Surface Pen PowerPoint Clicker

## WHAT IS IT?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
With this tool, you can use your Surface Pen top button as a clicker replacement to control PowerPoint presentations:
	Single click = cursor right (next slide, next animation)
	Double click = cursor left (previous slide, previous animation)
	Long press = key "b" (black screen)

Outside a PowerPoint full screen presentation, the tool is inactive and the Surface Pen behaves just as without the tool started.


## 32 OR 64 BIT?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Use the 64 bit version whenever possible. It is a requirement if you use the 64 bit Office suite, and it works fine with 32 bit Office installations.
The 32 bit version also works on 64 bit Windows, but has problems with 64 bit Office.


## HOW DO I INSTALL IT?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Just copy the corresponding .exe to a folder of your choice and run it.
If you want it to automatically start at logon:
	Press Win+R
	type "shell:startup"
	An Explorer window pops up
	Place the .exe or a shortcut to it in this folder


## WHAT ABOUT THE LICENSE?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
There is no license attached to this little helper program. Basically, you can do what you want with it.
There is no commercial support available. I support the tool in my spare time, so there is no guarantee here, too.


## HOW CAN I MODIFY IT OR COMPILE IT MYSELF?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
The tool is just a few lines of AutoHotkey code. You can find the source code in the file "Surface Pen PowerPoint.ahk", which can be edited with basically every text editor including Notepad.
To compile it, download and install AutoHotkey from http://autohotkey.org and run the "Convert .ahk to .exe" tool.


## THE PROGRAM DOES NOT WORK
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
**Step 1:** Does the Surface Pen button work after ending "Surface Pen PowerPoint.exe" in task manager?
	No: There is a general problem with the Pen and you should, after the obligatory reboot, delete and re-create the bluetooth connection with the Pen. Then, go to step 1 again.
	Yes: Go to step 2.

**Step 2:** Start "Surface Pen PowerPoint.exe" again. Is there a warning that you try to start a potentially dangerous piece of software?
	No: If there is an icon with a white "H" on green background in the notification area (maybe among the hidden icons), go to step 3.
	Yes: Adjust you SmartScreen filter or you anti-malware software. and go to step 2 again.

**Step 3:** Does the Surface Pen button still work outside PowerPoint?
	No: Go back to step 1. If you come back here regularly, something strange is going on. Please let me know.
	Yes: Go to step 4.


**Step 4:** Start PowerPoint, open a presentation, start presentation mode and try the button. Does it work?
	No: Something strange is going on. Please let me know.
	Yes: You are fine now.
