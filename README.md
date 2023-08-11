# Plant-3D-cable-tray-tool (proof-of-concept, sample code, use at own risk)
1. type "netload" in Plant 3D to load the dll
2. Select the straight tray element and execute the command typing "traysetlength", then follow the commandline input

Note this just works on straight cable trays (recommendation: use type coupling) and it just works when the length parameter is called "L".
For any other length parameter names like e.g. "L1" this script would need to get modified.

Make sure that your straight cable tray has the "Component Designation" parameter set to "Parametric", 
else you risk your cable trays going back to default length if somebody updates them with the specupdate.

Of course you can put your command into the CUI, e.g.: CUI -> Shortcut Menus -> Edit Menu, so it will available by right clicking after selecting the tray.
This command can be created as custom command in the CUI ("create a new command") with the "Macro" entry: ^C^Ctraysetlength

See demo of the code here: https://youtu.be/_au4D_soGBc
(NOTE: the Visual Studio part requires to install <a href="https://www.autodesk.com/developer-network/platform-technologies/autocad/objectarx">ObjectARX</a> and <a href="https://aps.autodesk.com/developer/overview/autocad-plant-3d-and-pid">Plant 3D SDK</a> respectivly)
