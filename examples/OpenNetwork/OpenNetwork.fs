open System
open System.Runtime.InteropServices

// Vissim COM Type Provider collects all installed Vissim COM type libraries and make them
// part of the compiler type system. There is no need to explicitly import the type library
// or add references.

// Alias to Vissim 2020 COM Type Lib
type VissimLib = COM.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

[<Literal>]
let VissimDemoFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Demo\"

[<EntryPoint; STAThread>]
let main argv =
    // Bind Vissim Demo folder and network file
    let networkFile = VissimDemoFolder + @"Urban Intersection Beijing.CN\Intersection Beijing.inpx"

    let vissim = VissimLib.VissimClass()
    vissim.LoadNet networkFile
    // Do nothing and exit.
    vissim.Exit()
    0
     
