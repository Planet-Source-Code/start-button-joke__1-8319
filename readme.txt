"Tray & Start Buttons Joke" program - which moves "Start" button or hides task
bar when You place mouse cursor on it.                                         
Program has 2 modes of work - first - joking with Start button and 
the second one is joking with TaskBar.

Program's window is invisible by default and also it is invisible for Windows 9x
Task Manager so if You want to control it manually ou must run one 
more copy of it with parameter "/show"
Example:running joke:             TrayWndJoke.exe
	showing program's window: TrayWndJoke.exe /show

When it's window is visible You can press "Go" or "Stop" buttons
to begin or stop joking and "Set2Def" to set "Start" button to it's default 
position.
Also You can select mode of joke in the Options menu
Set 1 - Start button
Set 2 - TaskBar window
*Note that to apply mode changing You have to press "Go" button.

By default program uses Set 1 but You can redefine this by setting "/set2" 
switch in the command line:
Example: TrayWndJoke.exe /set2
	 runs program with TaskBar joke mode