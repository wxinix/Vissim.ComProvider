# Vissim COM Type Provider

Vissim COM type provider is a F# compiler plugin, making Vissim COM type system part of the compiler type system. The added types are automatically visible to user code at compile time, without importing of any type libraries or adding reference assemblies explicilty.

The great thing is, Vissim COM type provider generates Vissim COM types based on dynamic schemas, which are the type libraries registered in Windows Registry. This means, whether you install new Vissim versions, uninstall old versions or whatever – you don’t have to import (or re-import) anything. The Vissim COM type provider will automatically pick up all of the installed up-t-date Vissim COM type libraries from Windows Registry (as its data source) at ***compile time***. Even better is that the same Vissim COM type provider can be used in both F# scripting and F# compiled applications transparently. Check these [F# examples](https://github.com/wxinix/Vissim.ComProvider/tree/master/examples) to learn how to take advantage the convenience and flexibility of Vissim COM type provider.

Using Vissim COM type provider in C# is somewhat tricky since type provider is really an F# thing. However, given this Vissim COM type provider is a generative type provider (with no type erasures),  we can resolve this by creating an extra F# library ```Vissim.ComProvider.TypeLibs``` which include Vissim COM types that can be referenced in C# applications. Check the example [CSharpTest](https://github.com/wxinix/Vissim.ComProvider/tree/master/examples/CSharpTest), illustrating usage with C#.

For more details, read my [blog article](https://blog.wupingxin.net/vissim-com-programming-for-fun-com-type-provider-a-new-way-of-doing-vissim-com-interop/).

## Usage
The following code sample illustrates the usage of Vissim COM type provider. 

### For F# Application
- Add Vissm.ComProvider.dll to the Project Reference, then each and every installed Vissim COM Type Library will be automatically collected by Vissim COM type provider, and added to a namespace called "Vissim.ComProvider". Each type library and its types are handled as two level subtypes.The first level subtype is the name of the type library, e.g., "Vissim Object Library 20.0 64 Bit", while the second level subtype is a combination of VissimVersionNumber and Platform, e.g., "14.0-win64" for Vissim 2020 64 bit. Therefore, the fully qualified types for Vissim COM 2020 64 Bit COM type library would be ``` Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64`` ```. 
- You can freely modify the Vissim COM Type Provider source code to change how fully qualified types are presented. For this repo, the first level subtype just takes the type library's help string as its name, and the second level subtype use a combination of VissimVerNo-Platform. This is a design decision I made which you may not like; feel free to customize for your own convenience.
- You don't have to type the lengthy fully qualified types, use the IDE's IntelliSense to auto-complete, and assign it to an Alias, as in the sample code.

The following example illustrates Vissim.ComProvider in action for F# applications. The computer has three Vissim versions installed, i.e., v5.40 64 bit, v10.0 32 bit, and v2020 64 bit. The Vissim COM Type Provider collects the following three fully-qualified types:
- ```Vissim.ComProvider.``VISSIM_COMServer 5.40 Type Library``.``1.0-win64`` ```
- ```Vissim.ComProvider.``Vissim Object Library 10.0``.``a.0-win32`` ```
- ```Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64`` ```

``` fsharp
open System
open System.Runtime.InteropServices

// Vissim COM Type Provider collects all installed Vissim COM type libraries and make them
// part of the compiler type system. There is no need to explicitly import the type library
// or add references.

// Alias of Vissim version 540 COM Type Lib
type VissimLib_540 = Vissim.ComProvider.``VISSIM_COMServer 5.40 Type Library``.``1.0-win64``

// Alias of Vissim version 100 COM Type Lib
type VissimLib_100 = Vissim.ComProvider.``Vissim Object Library 10.0``.``a.0-win32`` 

// Alias of Vissim version 200 COM Type Lib
type VissimLib_200 = Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64`` 

// The installed Vissim Type Libs are part of the compiler type system. We just alias them with short names.

[<EntryPoint; STAThread>]
let main argv =
    let vissim540 = VissimLib_540.VissimClass() // Create a Vissim 540 COM Object instance
    let vissim100 = VissimLib_100.VissimClass() // Create a Vissim 100 COM Object instance
    let vissim200 = VissimLib_200.VissimClass() // Create a Vissim 200 COM Object instance

    // We can have different versions of Vissim COM objects co-exisit in the same app domain.

    vissim200.LoadNet @"X:\Temp\Urban Intersection Beijing.CN\Intersection Beijing.inpx"
    for linkObj in vissim200.Net.Links do  
        let link = linkObj :?> VissimLib_200.ILink // Downcast, ILink is a subtype under VissimLib_200
        link.AttValue("No") |> string |> printfn "Link No %s"

    let mutable sec = vissim200.Simulation.AttValue("SimSec") |> string
    printfn "Current simulation second is %s" sec
    Console.ReadLine() |> ignore
    0
     
```


### For C# Application

For C# application, add Vissim.ComProvider.TypeLibs.dll to Project Reference. See the following sample code:

``` csharp
using System;
using Vissim.ComProvider.TypeLibs;

namespace CSharpTest
{
    class Program
    {
        public readonly static string ComExampleFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM\";
        public readonly static string NetworkFile = ComExampleFolder + @"Basic Commands\COM Basic Commands.inpx";
        public readonly static string LayoutFile = ComExampleFolder + @"Basic Commands\COM Basic Commands.layx";

        static void Main(string[] args)
        {
            var vissim = new VissimLib200.VissimClass();
            vissim.LoadNet(NetworkFile);
            vissim.Graphics.AttValue["QuickMode"] = true; 
            vissim.Simulation.RunContinuous();
            vissim.Exit();
            Console.ReadLine();
        }
    }
}
```


## Code Style Guidline

Microsoft [coding style guideline for F#](https://docs.microsoft.com/en-us/dotnet/fsharp/style-guide/formatting) is used.

## Maintainer of this Repository
- [@wxinix](https://github.com/wxinix)  Wuping Xin
