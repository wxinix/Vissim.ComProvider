open System
open System.Runtime.InteropServices

// Vissim COM Type Provider collects all installed Vissim COM type libraries and make them
// part of the compiler type system. There is no need to explicitly import the type library
// or add references.

type VissimLib_540 = COM.``VISSIM_COMServer 5.40 Type Library``.``1.0-win64`` // Alias of Vissim version 540 COM Type Lib
type VissimLib_100 = COM.``Vissim Object Library 10.0``.``a.0-win32``         // Alias of Vissim version 100 COM Type Lib
type VissimLib_200 = COM.``Vissim Object Library 20.0 64 Bit``.``14.0-win64`` // Alias of Vissim version 200 COM Type Lib

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
     
