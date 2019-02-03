# Plant-3D-cable-tray-tool
<pre>
Problem to fix:

Cable tray concept with python based couplings as trays. Length is fixed. Only way to adjust length is by setting properties parameter (with spec part designation is “Custom”) which is not user friendly.

“autotray” subcommand converts all selected pipes into couplings with the needed sizes, and rotates the selected elbows which are headfirst.

 

command

execTrayCommand "tcom=autotray,cspec=CableTraySpec,shortdesc=_cabletraycoup"

action

converts all selected pipes into couplings with the needed sizes, and rotates the selected elbows which are headfirst. The selected parts have to be in the same plain (e.g. horizontal), if there are others (e.g. vertical) they have to be processed separately. In the command sequence you will be asked for the alignment. The spec text is deleted from the properties to avoid length changes from spec update

comment

Doesn’t work if spec is opened by spec editor (readonly error). Use top view for the autotray command, it will connect (general: view from the "up direction" that you will choose, then it will connect)! Can do imperial with imperial spec and metric with metric spec. Width of the tray must be set up in the spec as nominal diameter (size). The outer diameter can be set to the tray height, this will lead to correct "top of pipe" and "bottom of pipe" values in the ortho.

parameters

Same as above, but the nominal diameter will be taken from the pipes

 

 

 

See this article about how do create the dll and how to install it: http://autode.sk/2jYKHJy 
</pre>
