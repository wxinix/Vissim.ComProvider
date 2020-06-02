open System
open System.Runtime.InteropServices

// Vissim COM Type Provider collects all installed Vissim COM type libraries and make them
// part of the compiler type system. There is no need to explicitly import the type library
// or add references.

// Alias to Vissim 2020 COM Type Lib
type VissimLib = COM.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

type Object with
    member this.AsArray<'T> () =
        this :?> Object [] |> Array.map(fun o -> o :?> 'T)

[<Literal>]
let VissimComExampleFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM\Basic Commands\"
[<Literal>]
let VissimNetwork = VissimComExampleFolder + @"COM Basic Commands.inpx"
[<Literal>]
let VissimLayout = VissimComExampleFolder + @"COM Basic Commands.layx"

let loadNetwork (vissim: VissimLib.IVissim) =
    vissim.LoadNet VissimNetwork
    VissimNetwork |> printfn "Network \"%s\" loaded." 
    vissim

let loadLayout (vissim: VissimLib.IVissim) =
    vissim.LoadLayout VissimLayout
    VissimLayout |> printfn "Layout \"%s\" loaded."
    vissim

let readLinkName linkNo (vissim: VissimLib.IVissim) =
    printfn "Name of Link[%d]: %s" linkNo (vissim.Net.Links.ItemByKey(linkNo).AttValue("Name") |> string)
    vissim

let setLinkName linkNo linkName (vissim: VissimLib.IVissim) =
    vissim.Net.Links.ItemByKey(linkNo).AttValue("Name") <- linkName
    printfn "Link name \"%s\" set to Link [%d]" linkName linkNo
    vissim

let setSignalControllerProgram scNo programNo (vissim: VissimLib.IVissim) =
    vissim.Net.SignalControllers.ItemByKey(scNo).AttValue("ProgNo") <- programNo
    printfn "Signal controller [%d] has new program No [%d]" scNo programNo
    vissim

let setStaticRouteRelativeFlow decision route flow (vissim: VissimLib.IVissim) =
    vissim.Net.VehicleRoutingDecisionsStatic.ItemByKey(decision).VehRoutSta.ItemByKey(route).AttValue("RelFlow(1)") <- flow
    printfn "Static route [%d] of decision [%d] has new relative flow [%f]" decision route flow
    vissim

let setVehicleInput inputNo newVolume (vissim: VissimLib.IVissim) =
    let input = vissim.Net.VehicleInputs.ItemByKey(inputNo)
    input.AttValue("Volume(1)") <- newVolume
    input.AttValue("Cont(2)") <- false
    input.AttValue("Volume(2)") <- 400
    printfn "Vehicle input [%d] set new volume[1] = %d volume[2] = %d" inputNo newVolume 400
    vissim

let setVehicleComposition vehCompositionNo (vissim: VissimLib.IVissim) =
    let relFlows = vissim.Net.VehicleCompositions.ItemByKey(vehCompositionNo).VehCompRelFlows.GetAll().AsArray<VissimLib.IVehicleCompositionRelativeFlow>()

    relFlows.[0].AttValue("VehType") <- 100
    relFlows.[0].AttValue("DesSpeedDistr") <- 50
    relFlows.[0].AttValue("RelFlow") <- 0.9
    relFlows.[1].AttValue("RelFlow") <- 0.1

    printfn "Vehicle compositions total count [%d], and [%d] has been set." relFlows.Length vehCompositionNo
    vissim

let getLinkMultiAttValues attr (vissim: VissimLib.IVissim) =
    let values = vissim.Net.Links.GetMultiAttValues(attr) :?> Object[,]
    values |> Array2D.iteri(fun i j item -> printfn "\t Link[%2d] Attribute[%d] Value %s" i j (item |> string) )
    vissim

let setLinkMultiAttValues attr values (vissim: VissimLib.IVissim) =
    vissim

[<EntryPoint; STAThread>]
let main argv =        
    let vissim = VissimLib.VissimClass()    
    
    vissim |> loadNetwork 
           |> loadLayout 
           |> readLinkName 1
           |> setLinkName 1 "New Link Name"
           |> setSignalControllerProgram 1 2
           |> setStaticRouteRelativeFlow 1 1 0.6
           |> setVehicleInput 1 600
           |> setVehicleComposition 1
           |> getLinkMultiAttValues "Name" 
           |> ignore
    
    vissim.Exit()
    Console.ReadLine() |> ignore
    0
     
