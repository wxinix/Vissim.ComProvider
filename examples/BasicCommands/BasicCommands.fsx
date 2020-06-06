// This script can be directly run from F# interactive or VS Code. Select all the code and
// press Alt + Enter to run the script.

#I "../../output"
#r "Vissim.ComProvider.dll"

// Alias to Vissim 2020 COM Type Lib
type VissimLib = 
    Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

let [<Literal>] ExampleFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM"
let [<Literal>] NetworkFile   = ExampleFolder + @"\Basic Commands\COM Basic Commands.inpx"

let vissim = VissimLib.VissimClass()

vissim.LoadNet NetworkFile
vissim.Graphics.AttValue("QuickMode") <- true
vissim.Simulation.RunContinuous()
vissim.Exit()
