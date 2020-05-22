open System

type Vissim = COM.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

[<EntryPoint; STAThread>]
let main argv =
    let vissim = Vissim.VissimClass()
    vissim.LoadNet @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Demo\Urban Intersection Beijing.CN\Intersection Beijing.inpx"
    let mutable sec = vissim.Simulation.AttValue("SimSec") |> string
    for i in 1..10 do 
        vissim.Simulation.RunSingleStep()
        sec <- vissim.Simulation.AttValue("SimSec") |> string
        printfn "Current simulation second is %s" sec
    0
