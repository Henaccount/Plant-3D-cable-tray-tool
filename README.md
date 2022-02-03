# Plant-3D-cable-tray-tool (proof-of-concept, sample code, use at own risk)
1. type "netload" in Plant 3D to load the dll
2. execute the command typing "traysetlength", then follow the commandline input

Note this just works on straight cable trays (recommendation: use type coupling) and it just works when the length parameter is called "L".
For any other length parameter names like e.g. "L1" this script would need to get modified.

Make sure that your straight cable tray has the "Component Designation" parameter set to "Parametric", 
else you risk your cable trays going back to default length if somebody updates them with the specupdate.


