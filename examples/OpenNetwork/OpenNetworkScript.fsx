// This script can be directly run from F# interactive. Select all the code and
// press Alt + Enter to run the script.

#I "../../output"
#r "Vissim.ComProvider.dll"

type Vissim = COM.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

let vissim = Vissim.VissimClass()
vissim.LoadNet @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Demo\Urban Intersection Beijing.CN\\
                 Intersection Beijing.inpx"
