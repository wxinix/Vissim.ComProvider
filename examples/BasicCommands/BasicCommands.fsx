// This script can be directly run from F# interactive or VS Code. Select all the code and
// press Alt + Enter to run the script.

#I "../../output"
#r "Vissim.ComProvider.dll"

open System.IO

// Alias to Vissim 2020 COM Type Lib
type VissimLib = COM.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

[<Literal>]
let VissimComExampleFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM"
let networkFile = VissimComExampleFolder + @"\Basic Commands\COM Basic Commands.inpx"

let vissim = VissimLib.VissimClass()
vissim.LoadNet networkFile
// Do nothing and exit.
vissim.Exit()
