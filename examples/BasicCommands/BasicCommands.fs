// MIT License
// Copyright (c) Wuping Xin 2020.
//
// Permission is hereby  granted, free of charge, to any  person obtaining a copy
// of this software and associated  documentation files (the "Software"), to deal
// in the Software  without restriction, including without  limitation the rights
// to  use, copy,  modify, merge,  publish, distribute,  sublicense, and/or  sell
// copies  of  the Software,  and  to  permit persons  to  whom  the Software  is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE  IS PROVIDED "AS  IS", WITHOUT WARRANTY  OF ANY KIND,  EXPRESS OR
// IMPLIED,  INCLUDING BUT  NOT  LIMITED TO  THE  WARRANTIES OF  MERCHANTABILITY,
// FITNESS FOR  A PARTICULAR PURPOSE AND  NONINFRINGEMENT. IN NO EVENT  SHALL THE
// AUTHORS  OR COPYRIGHT  HOLDERS  BE  LIABLE FOR  ANY  CLAIM,  DAMAGES OR  OTHER
// LIABILITY, WHETHER IN AN ACTION OF  CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE  OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

open System

// Vissim COM Type Provider collects all installed Vissim COM type libraries and make them
// part of the compiler type system. There is no need to explicitly import the type library
// or add references.

// Alias to Vissim 2020 COM Type Lib
type VissimLib =
    Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

type Object with
    member this.AsArray<'T> () = this :?> Object [] |> Array.map(fun o -> o :?> 'T)

let [<Literal>] ExampleFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM\"
let [<Literal>] LayoutFile    = ExampleFolder + @"Basic Commands\COM Basic Commands.layx"
let [<Literal>] NetworkFile   = ExampleFolder + @"Basic Commands\COM Basic Commands.inpx"

let loadNetwork (vissim: VissimLib.IVissim) =
    vissim.LoadNet NetworkFile
    NetworkFile |> printfn "Network \"%s\" loaded."
    vissim

let loadLayout (vissim: VissimLib.IVissim) =
    vissim.LoadLayout LayoutFile
    LayoutFile |> printfn "Layout \"%s\" loaded."
    vissim

let readLinkName linkNo (vissim: VissimLib.IVissim) =
    printfn "Name of Link[%d]: %s" linkNo (vissim.Net.Links.ItemByKey(linkNo).AttValue("Name") |> string)
    vissim

let setLinkName linkNo linkName (vissim: VissimLib.IVissim) =
    vissim.Net.Links.ItemByKey(linkNo).AttValue("Name") <- linkName
    printfn "Link name \"%s\" set to Link [%d]" linkName linkNo
    vissim

let setSignalControllerProgram scNo progNo (vissim: VissimLib.IVissim) =
    vissim.Net.SignalControllers.ItemByKey(scNo).AttValue("ProgNo") <- progNo
    printfn "Signal controller [%d] has new program No [%d]" scNo progNo
    vissim

let setStaticRouteRelativeFlow decision route flow (vissim: VissimLib.IVissim) =
    vissim.Net.VehicleRoutingDecisionsStatic.ItemByKey(decision).VehRoutSta.ItemByKey(route).AttValue("RelFlow(1)") <- flow
    printfn "Static route [%d] of decision [%d] has new relative flow [%f]" decision route flow
    vissim

let setVehicleInput inputNo newVol (vissim: VissimLib.IVissim) =
    let input = vissim.Net.VehicleInputs.ItemByKey(inputNo)
    input.AttValue("Volume(1)") <- newVol
    input.AttValue(  "Cont(2)") <- false
    input.AttValue("Volume(2)") <- 400
    printfn "Vehicle input [%d] set new volume[1] = %d volume[2] = %d" inputNo newVol 400
    vissim

let setVehicleComposition vehCompNo (vissim: VissimLib.IVissim) =
    let relFlows = vissim.Net.VehicleCompositions.ItemByKey(vehCompNo).VehCompRelFlows.GetAll().AsArray<VissimLib.IVehicleCompositionRelativeFlow>()

    relFlows.[0].AttValue(      "VehType") <- 100
    relFlows.[0].AttValue("DesSpeedDistr") <- 50
    relFlows.[0].AttValue(      "RelFlow") <- 0.9
    relFlows.[1].AttValue(      "RelFlow") <- 0.1

    printfn "Vehicle compositions total count [%d], and [%d] has been set." relFlows.Length vehCompNo
    vissim

let getLinkMultiAttValues attr (vissim: VissimLib.IVissim) =
    let values = vissim.Net.Links.GetMultiAttValues(attr) :?> Object[,]
    values |> Array2D.iteri(fun i j item -> printfn "\t Link[%02d] attr[%d] = %s" i j (item |> string) )
    vissim

let setLinkMultiAttValues attr values (vissim: VissimLib.IVissim) =
    vissim.Net.Links.SetMultiAttValues(attr, values)
    printfn "MultiAttrValues %s set." attr
    vissim

let getLinkMultipleAttributes attrs (vissim: VissimLib.IVissim) =
    let values = vissim.Net.Links.GetMultipleAttributes(attrs) :?> Object[,]
    values |> Array2D.iteri(fun i j item -> printfn "\t Link[%02d] attr[%d] = %s" i j (item |> string) )
    vissim

let setLinkMultipleAttributes attrs values (vissim: VissimLib.IVissim) =
    vissim.Net.Links.SetMultipleAttributes(attrs, values)
    printfn "MultipleAttributes %s set." (attrs.AsArray<string>() |> Array.fold(fun acc elem -> acc + elem + " ") "" )
    vissim

let setLinkAllAttValues attr value add (vissim: VissimLib.IVissim) =
    vissim.Net.Links.SetAllAttValues(attr, value, add) // add meaningful for numbers only
    printfn "SetAllAttValues attr[%s] = [%s] add = [%b]" attr (id<obj> value |> string) add
    vissim

let runSimulation (vissim: VissimLib.IVissim) =
    vissim.Simulation.AttValue(      "RandSeed") <- 42
    vissim.Simulation.AttValue(     "SimPeriod") <- 600
    vissim.Simulation.AttValue(    "SimBreakAt") <- 200
    vissim.Simulation.AttValue("UseMaxSimSpeed") <- true
    vissim.Simulation.AttValue(      "SimSpeed") <- 10
    vissim.Simulation.RunContinuous()
    vissim.Simulation.Stop()
    vissim

let accessSignalsInSimulationRun (vissim: VissimLib.IVissim) =
    vissim.Simulation.AttValue("SimBreakAt") <- 198
    vissim.Simulation.RunContinuous()
    vissim.Net.SignalHeads.ItemByKey(1).AttValue("SigState") |> string |> printfn "Actual state of SignalHead[%d] is: %s" 1
    let sg = vissim.Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1)
    sg.AttValue("SigState") <- "GREEN"  // "ContrByCOM" automatically set true
    vissim.Simulation.AttValue("SimBreakAt") <- 200
    vissim.Simulation.RunContinuous()
    sg.AttValue("ContrByCOM") <- false  // Restore to be false
    vissim

let moveAndDelVehicle (vissim: VissimLib.IVissim) =
    let veh = (vissim.Net.Vehicles.GetAll() :?> Object []).[1] :?> VissimLib.IVehicle
    let vehNo = veh.AttValue("No")
    let moveVeh link lane newLinkPos = veh.MoveToLinkPosition(link, lane, newLinkPos)    
    vehNo |> string |> printfn "VehNo[%s] selected, and to be moved." 
    veh.AttValue("DesSpeed") <- 30
    (1, 1, 70.0) |||> moveVeh
    vissim.Simulation.RunSingleStep()
    vissim.Net.Vehicles.RemoveVehicle(vehNo)
    vissim

let addNewVehicle (vissim: VissimLib.IVissim) =
    let (vehTy, link, lane, linkPos, desSpd, interaction) = (100, 1, 1, 15.0, 53.0, true)
    vissim.Net.Vehicles.AddVehicleAtLinkPosition (vehTy, link, lane, linkPos, desSpd, interaction) |> ignore
    vissim

let takeScreenshots (vissim: VissimLib.IVissim) =
    let camPos = (270.0, 30.0, 15.0, 45.0, 10.0)
    let originalNetworkPos = (250.0, 30.0, 350.0, 135.0)

    let take2DScreenshots =
        vissim.Graphics.CurrentNetworkWindow.ZoomTo originalNetworkPos
        vissim.Graphics.CurrentNetworkWindow.Screenshot("screenshot2D.jpg", 1.0)
        printfn "screenshot2D.jpg saved!"

    let take3DScreenshots =
        vissim.Graphics.CurrentNetworkWindow.AttValue("3D") <- true
        vissim.Graphics.CurrentNetworkWindow.SetCameraPositionAndAngle camPos
        vissim.Graphics.CurrentNetworkWindow.Screenshot("screenshot3D.jpg", 1.0)
        printfn "screenshot3D.jpg saved!"

    let restore2DNetworkPosition =
        vissim.Graphics.CurrentNetworkWindow.AttValue("3D") <- false
        vissim.Graphics.CurrentNetworkWindow.ZoomTo originalNetworkPos
        printfn "Network position restored after taking screenshots!"

    take2DScreenshots
    take3DScreenshots
    restore2DNetworkPosition

    vissim.Simulation.RunContinuous()
    vissim

let retrieveSimResults (vissim: VissimLib.IVissim) =
    let generateSimulationResults =
        vissim.SuspendUpdateGUI()

        for simRun in vissim.Net.SimulationRuns do
            vissim.Net.SimulationRuns.RemoveSimulationRun(simRun :?> VissimLib.ISimulationRun)

        vissim.Simulation.AttValue(     "SimPeriod") <- 600
        vissim.Simulation.AttValue(    "SimBreakAt") <- 0
        vissim.Simulation.AttValue("UseMaxSimSpeed") <- true

        for rs in 1..3 do
            vissim.Simulation.AttValue("RandSeed") <- rs
            vissim.Simulation.RunContinuous()
        
        let simRuns = 
            vissim.Net.SimulationRuns.GetMultipleAttributes(id<obj[]> [| "Timestamp"; "RandSeed"; "SimEnd" |]) :?> Object [,]
        
        for i in 0 .. (Array2D.length1 simRuns) - 1 do
            printfn "%s |%s |%s" (simRuns.[i,0] |> string) (simRuns.[i,1] |> string) (simRuns.[i,2] |> string)

        vissim.ResumeUpdateGUI(false)

    let retrieveTravelTimes =
        let ttMeaNo = 2
        let ttMea = vissim.Net.VehicleTravelTimeMeasurements.ItemByKey(ttMeaNo)
    
        let retrieveTravelTimesAllSimulations =
            let travTm = ttMea.AttValue("TravTm(Avg,Avg,All") :?> double    // sim | time interval | veh class = ALL
            let vehs   = ttMea.AttValue(  "Vehs(Avg,Avg,All") :?> double
            (ttMeaNo, travTm, vehs) |||> printfn "\nTravelTimeMeasurement[%d] | SIMRUN = AVG | INTERVAL = AVG | VCLS = ALL: %5.3f (%5.3f vehs)"         
        
        let retrieveTravelTimesCurrentSimulation =
            let travTm = ttMea.AttValue("TravTm(Current,Max,20") :?> double // sim | time interval | veh class = HGV
            let vehs   = ttMea.AttValue(  "Vehs(Current,Max,20") :?> int
            (ttMeaNo, travTm, vehs) |||> printfn "TravelTimeMeasurement[%d] | SIMRUN = CUR | INTERVAL = MAX | VCLS = HGV: %5.3f (%d vehs)"        
        
        let retrieveTravelTimesBySimulationRunAndInterval =
            let travTm = ttMea.AttValue("TravTm(2,1,All") :?> double        // sim | time interval | veh class = HGV
            let vehs   = ttMea.AttValue(  "Vehs(2,1,All") :?> int
            (ttMeaNo, travTm, vehs) |||> printfn "TravelTimeMeasurement[%d] | SIMRUN = 2 | INTERVAL = 1 | VCLS = ALL: %5.3f (%d vehs)"
   
        retrieveTravelTimesAllSimulations
        retrieveTravelTimesCurrentSimulation
        retrieveTravelTimesBySimulationRunAndInterval
    
    let retrieveDataCollectionMeasurements =
        let dcNo  = 1
        let dcMea = vissim.Net.DataCollectionMeasurements.ItemByKey(dcNo)
        let vehs  = dcMea.AttValue(        "Vehs(Avg,1,All") :?> int
        let speed = dcMea.AttValue(       "Speed(Avg,1,All") :?> double
        let accel = dcMea.AttValue("Acceleration(Avg,1,All") :?> double
        let len   = dcMea.AttValue(      "Length(Avg,1,All") :?> double        
        printfn "\nDataCollection[%d] | SIMRUN = AVG | INTERVAL = 1 | VCLS = ALL: %d (vehs), %5.3f (speed), %5.3f (accel), %5.3f (length)"  dcNo vehs speed accel len 
 
    let retrieveQueueCounterMeasurements =
        let qcNo = 1
        let maxQ = vissim.Net.QueueCounters.ItemByKey(qcNo).AttValue("QLenMax(Avg,Avg") :?> double
        printfn "\nQueueCounter[%d] | SIMRUN = AVG | INTERVAL = ALL: %5.3f (max_queue_length)" qcNo maxQ
   
    generateSimulationResults
    retrieveTravelTimes
    retrieveDataCollectionMeasurements
    retrieveQueueCounterMeasurements
    vissim

[<EntryPoint; STAThread>]
let main argv =
    let vissim = VissimLib.VissimClass()

    printfn "\n***Load Network and Layout Example:"
    vissim |> loadNetwork |> loadLayout |> ignore
    
    printfn "\n***Read and Set Link Name Example:"
    vissim |> readLinkName 1
           |> setLinkName 1 "New Link Name"
           |> readLinkName 1
           |> ignore
    
    printfn "\n***Set Signal Controller Program Example:"
    vissim |> setSignalControllerProgram 1 2 |> ignore
    
    printfn "\n***Set Static Route Relative Flow Example:"
    vissim |> setStaticRouteRelativeFlow 1 1 0.6 |> ignore
    
    printfn "\n***Set Vehicle Input Example:"
    vissim |> setVehicleInput 1 600 |> ignore
    
    printfn "\n***Set Vehicle Composition Example:"
    vissim |> setVehicleComposition 1 |> ignore
    
    printfn "\n***Get and Set Link MultiAttValues Name Example:"
    vissim |> getLinkMultiAttValues "Name"
           |> setLinkMultiAttValues "Name" (array2D<_,obj> [| [| 1; "New Link Name of Link#1" |]
                                                              [| 2; "New Link Name of Link#2" |]
                                                              [| 4; "New Link Name of Link#4" |]  |])
           |> getLinkMultiAttValues "Name"  // Validate the above setLinkMultiAttValues
           |> ignore

    printfn "\n***Get and Set Link Multiple Attributes Name-CostPerKm Example:"
    vissim |> getLinkMultipleAttributes (id<obj[]> [| "Name" ; "CostPerKm" |])
           |> setLinkMultipleAttributes (id<obj[]> [| "Name" ; "CostPerKm" |]) (array2D<_,obj> [| [| "Name1"; 12 |]
                                                                                                  [| "Name2";  7 |]
                                                                                                  [| "Name3";  5 |]
                                                                                                  [| "Name4";  4 |] |])
           |> getLinkMultipleAttributes (id<obj[]> [| "Name"; "CostPerKm" |]) // Validate the above setLinkMultipleAttributes
           |> ignore
    
    printfn "\n***Set Link AllAttValues Name Example:"
    vissim |> setLinkAllAttValues "Name" "All Links have the same name" false |> ignore

    printfn "\n***Set Link AllAttValues CostPerKm Example:"
    vissim |> setLinkAllAttValues "CostPerKm" 5.5 true |> ignore

    vissim.Graphics.AttValue("QuickMode") <- true
    
    printfn "\n***Run Simulation Example:"
    vissim |> runSimulation |> ignore
    
    printfn "\n***Access Signals in Simulation Run Example:"
    vissim |> accessSignalsInSimulationRun |> ignore
    
    vissim.Graphics.AttValue("QuickMode") <- false
    
    printfn "\n***Vehicle Manipulation Example:"
    vissim |> moveAndDelVehicle |> addNewVehicle |> ignore

    vissim.Graphics.AttValue("QuickMode") <- true
    
    printfn "\n***Take Screenshots Example:"
    vissim |> takeScreenshots |> ignore
    
    printfn "\n***Retrieve Simulation Results Example:"
    vissim |> retrieveSimResults |> ignore

    //vissim.SaveNetAs(VissimNetwork)
    //vissim.SaveLayout(VissimLayout)
    vissim.Exit()
    Console.ReadLine() |> ignore
    0

