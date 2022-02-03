using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("command: traysetlength")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("command: traysetlength")]
[assembly: AssemblyCopyright("")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("D57DCD91-4BF8-4EC8-B326-0503A212CB82")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
//V1:one point pick instead of two
//V2: can resize the tray from both sides now
//V3: can now also shrink the tray from both sides, user needs to select the tray at the end where shrinking/expanding should take place
//V3: can also use snappoints outside of the tray axis
//V4: change L to L1 as length
//V5: change L1 to L as length
[assembly: AssemblyVersion("0.0.0.5")]
[assembly: AssemblyFileVersion("0.0.0.5")]
