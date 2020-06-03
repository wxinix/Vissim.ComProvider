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
    COM.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

type Object with
    member this.AsArray<'T> () = this :?> Object [] |> Array.map(fun o -> o :?> 'T)

let [<Literal>] VissimComExampleFolder =
    @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM\"

let [<Literal>] VissimNetwork =
    VissimComExampleFolder + @"Basic Commands\COM Basic Commands.inpx"

let [<Literal>] VissimLayout =
    VissimComExampleFolder  + @"Basic Commands\COM Basic Commands.layx"

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
    input.AttValue("Cont(2)")   <- false
    input.AttValue("Volume(2)") <- 400
    printfn "Vehicle input [%d] set new volume[1] = %d volume[2] = %d" inputNo newVolume 400
    vissim

let setVehicleComposition vehCompositionNo (vissim: VissimLib.IVissim) =
    let relFlows = vissim.Net
                         .VehicleCompositions
                         .ItemByKey(vehCompositionNo)
                         .VehCompRelFlows
                         .GetAll()
                         .AsArray<VissimLib.IVehicleCompositionRelativeFlow>()

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
    vissim.Net.Links.SetMultiAttValues(attr, values)
    printfn "MultiAttrValues %s set." attr
    vissim

let getLinkMultipleAttributes attrs (vissim: VissimLib.IVissim) =
    let values = vissim.Net.Links.GetMultipleAttributes(attrs) :?> Object[,]
    values |> Array2D.iteri(fun i j item -> printfn "\t Link[%2d] Attribute[%d] Value %s" i j (item |> string) )
    vissim

let setLinkMultipleAttributes attrs values (vissim: VissimLib.IVissim) =
    vissim.Net.Links.SetMultipleAttributes(attrs, values)
    printfn "MultipleAttributes %s set." (attrs.AsArray<string>() |> Array.fold(fun acc elem -> acc + elem + " ") "" )
    vissim

let setLinkAllAttValues attr value add (vissim: VissimLib.IVissim) =
    vissim.Net.Links.SetAllAttValues(attr, value, add) // add only meaningful for numbers
    printfn "SetAllAttValues attr[%s] value[%s] add[%b]" attr (value.ToString()) add
    vissim

let runSimulation (vissim: VissimLib.IVissim) =
    vissim.Simulation.AttValue("RandSeed") <- 42
    vissim.Simulation.AttValue("SimPeriod") <- 600
    vissim.Simulation.RunSingleStep()
    vissim.Simulation.AttValue("SimBreakAt") <- 200
    vissim.Simulation.AttValue("UseMaxSimSpeed") <- true
    vissim.Simulation.AttValue("SimSpeed") <- 10
    vissim.Simulation.RunContinuous()
    vissim.Simulation.Stop()
    vissim

let accessSignalsInSimulationRun (vissim: VissimLib.IVissim) =
    vissim.Simulation.AttValue("SimBreakAt") <- 198
    vissim.Simulation.RunContinuous()
    vissim.Net.SignalHeads.ItemByKey(1).AttValue("SigState") |> string |> printfn "Actual state of SignalHead (%d) is: %s" 1
    let sg = vissim.Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1)
    sg.AttValue("SigState") <- "GREEN"  // "ContrByCOM" automatically set true
    vissim.Simulation.AttValue("SimBreakAt") <- 200
    vissim.Simulation.RunContinuous()
    sg.AttValue("ContrByCOM") <- false
    vissim

let loopTest (vissim: VissimLib.IVissim) =
    printfn "\nBeginning loopTest"

    let runVehLoopWithGetAll =
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        let allVehicles = vissim.Net.Vehicles.GetAll().AsArray<VissimLib.IVehicle>()
        for i in 0..allVehicles.Length - 1 do
            let veh = allVehicles.[i]
            let vehNumber   = veh.AttValue("No")      :?> int
            let vehType     = veh.AttValue("VehType") |> string
            let vehSpeed    = veh.AttValue("Speed")   :?> double
            let vehPos      = veh.AttValue("Pos")     :?> double
            let vehLinkLane = veh.AttValue("Lane")    |> string
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t" vehNumber vehType vehSpeed vehPos vehLinkLane
        stopWatch.Stop()
        printfn "Loop all vehicles with GetAll takes %f milliseconds" stopWatch.Elapsed.TotalMilliseconds

    let runVehLoopWithEnumerator =
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        for vehObj in vissim.Net.Vehicles do
            let veh = vehObj :?> VissimLib.IVehicle
            let vehNumber   = veh.AttValue("No")      :?> int
            let vehType     = veh.AttValue("VehType") |> string
            let vehSpeed    = veh.AttValue("Speed")   :?> double
            let vehPos      = veh.AttValue("Pos")     :?> double
            let vehLinkLane = veh.AttValue("Lane")    |> string
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t" vehNumber vehType vehSpeed vehPos vehLinkLane
        stopWatch.Stop()
        printfn "Loop all vehicles with Enumerator takes %f milliseconds" stopWatch.Elapsed.TotalMilliseconds

    let runVehLoopWithIterator =
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        let iter = vissim.Net.Vehicles.Iterator
        while iter.Valid do
            let veh = iter.Item :?> VissimLib.IVehicle
            let vehNumber   = veh.AttValue("No")      :?> int
            let vehType     = veh.AttValue("VehType") |> string
            let vehSpeed    = veh.AttValue("Speed")   :?> double
            let vehPos      = veh.AttValue("Pos")     :?> double
            let vehLinkLane = veh.AttValue("Lane")    |> string
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t" vehNumber vehType vehSpeed vehPos vehLinkLane
            iter.Next()
        stopWatch.Stop()
        printfn "Loop all vehicles with iterator takes %f milliseconds" stopWatch.Elapsed.TotalMilliseconds

    let runVehLoopWithGetMultiAttValues =
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        let vehNumbers   = vissim.Net.Vehicles.GetMultiAttValues("No")      :?> Object [,]
        let vehTypes     = vissim.Net.Vehicles.GetMultiAttValues("VehType") :?> Object [,]
        let vehSpeeds    = vissim.Net.Vehicles.GetMultiAttValues("Speed")   :?> Object [,]
        let vehPositions = vissim.Net.Vehicles.GetMultiAttValues("Pos")     :?> Object [,]
        let vehLinkLanes = vissim.Net.Vehicles.GetMultiAttValues("Lane")    :?> Object [,]
        for i in 0 .. Array2D.length1 vehNumbers - 1 do
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t" (vehNumbers.[i,1] :?> int) (vehTypes.[i,1] |> string) (vehSpeeds.[i,1] :?> double) (vehPositions.[i,1] :?> double) (vehLinkLanes.[i,1] |> string)
        stopWatch.Stop()
        printfn "Loop all vehicles with GetMultiAttValues takes %f milliseconds" stopWatch.Elapsed.TotalMilliseconds

    let runVehLoopWithGetMultipleAttributes =
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        let allVehAttrs = vissim.Net.Vehicles.GetMultipleAttributes(id<obj[]> [| "No";"VehType";"Speed";"Pos"; "Lane" |]) :?> Object[,]
        for i in 0.. Array2D.length1 allVehAttrs - 1 do
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t" (allVehAttrs.[i,0] :?> int) (allVehAttrs.[i,1] |> string) (allVehAttrs.[i,2] :?> double) (allVehAttrs.[i,3] :?> double) (allVehAttrs.[i,4] |> string)
        stopWatch.Stop()
        printfn "Loop all vehicles with GetMultipleAttributes takes %f milliseconds" stopWatch.Elapsed.TotalMilliseconds
    
    runVehLoopWithGetAll
    runVehLoopWithEnumerator
    runVehLoopWithIterator
    runVehLoopWithGetMultiAttValues
    runVehLoopWithGetMultipleAttributes
    vissim

let moveAndDelVehicle (vissim: VissimLib.IVissim) =
    let veh = (vissim.Net.Vehicles.GetAll() :?> Object []).[1] :?> VissimLib.IVehicle
    let vehNum = veh.AttValue("No") |> string
    printfn "Vehicle No[%s] selected, and to be moved." vehNum
    veh.AttValue("DesSpeed") <- 30
    let moveVeh (link: obj) lane (newLinkPos: float) = veh.MoveToLinkPosition(link, lane, newLinkPos)
    (1, 1, 70.0) |||> moveVeh
    vissim.Simulation.RunSingleStep()
    vissim.Net.Vehicles.RemoveVehicle(vehNum)
    vissim

let addNewVehicle (vissim: VissimLib.IVissim) =
    let (vehTy, link, lane, linkPos, desSpd, interaction) = (100, 1, 1, 15.0, 53.0, true)
    vissim.Net.Vehicles.AddVehicleAtLinkPosition (vehTy, link, lane, linkPos, desSpd, interaction) |> ignore
    vissim

let takeScreenshots (vissim: VissimLib.IVissim) =
    let originalNetworkPos = (250.0, 30.0, 350.0, 135.0)
    let camPosition = (270.0, 30.0, 15.0, 45.0, 10.0)

    let take2DScreenshots =
        vissim.Graphics.CurrentNetworkWindow.ZoomTo originalNetworkPos
        vissim.Graphics.CurrentNetworkWindow.Screenshot("screenshot2D.jpg", 1.0)
        printfn "screenshot2D.jpg saved!"

    let take3DScreenshots =
        vissim.Graphics.CurrentNetworkWindow.AttValue("3D") <- true
        vissim.Graphics.CurrentNetworkWindow.SetCameraPositionAndAngle camPosition
        vissim.Graphics.CurrentNetworkWindow.Screenshot("screenshot3D.jpg", 1.0)
        printfn "screenshot3D.jpg saved!"

    let restore2DNetworkPosition =
        vissim.Graphics.CurrentNetworkWindow.AttValue("3D") <- false
        vissim.Graphics.CurrentNetworkWindow.ZoomTo originalNetworkPos
        printfn "Network position restored after taking screenshotsd!"

    take2DScreenshots
    take3DScreenshots
    restore2DNetworkPosition

    vissim.Simulation.RunContinuous()
    vissim

let retrieveSimResults (vissim: VissimLib.IVissim) =
    let generateResultsFromSimRuns =
        vissim.SuspendUpdateGUI()

        for simRun in vissim.Net.SimulationRuns do
            vissim.Net.SimulationRuns.RemoveSimulationRun(simRun :?> VissimLib.ISimulationRun)

        vissim.Simulation.AttValue("SimPeriod") <- 600
        vissim.Simulation.AttValue("SimBreakAt") <- 0
        vissim.Simulation.AttValue("UseMaxSimSpeed") <- true

        for i in 1..3 do
            vissim.Simulation.AttValue("RandSeed") <- i
            vissim.Simulation.RunContinuous()

        let simRuns = vissim.Net.SimulationRuns.GetMultipleAttributes(id<obj[]> [| "Timestamp"; "RandSeed"; "SimEnd" |]) :?> Object [,]

        for i in 0 .. (Array2D.length1 simRuns) - 1 do
            printfn "%s |%s |%s" (simRuns.[i,0] |> string) (simRuns.[i,1] |> string) (simRuns.[i,2] |> string)

        vissim.ResumeUpdateGUI(false)

    let retrieveTravelTimes =
        let ttMeaNumber = 2
        let ttMea = vissim.Net.VehicleTravelTimeMeasurements.ItemByKey(ttMeaNumber)
    
        let retrieveTravelTimesAllSimulations =
            let tt = ttMea.AttValue("TravTm(Avg,Avg,All") :?> double // sim | time interval | veh class
            let noVeh = ttMea.AttValue("Vehs(Avg,Avg,All") :?> double
            printfn "MeaNumber %d Average travel time all time intervals of all simulations of all veh classes: %5.3f (number of vehicles %5.3f)" ttMeaNumber tt noVeh
        
        let retrieveTravelTimesCurrentSimulation =
             let tt = ttMea.AttValue("TravTm(Current,Max,20") :?> double // sim | time interval | veh class = HGV
             let noVeh = ttMea.AttValue("Vehs(Current,Max,20") :?> int
             printfn "MeaNumber %d Max travel time all time intervals of current simulations of veh classes HGV: %5.3f (number of vehicles %d)" ttMeaNumber tt noVeh
        
        let retrieveTravelTimes2ndSimulation1stInterval =
            let tt = ttMea.AttValue("TravTm(2,1,All") :?> double // sim | time interval | veh class = HGV
            let noVeh = ttMea.AttValue("Vehs(2,1,All") :?> int
            printfn "MeaNumber %d travel time of the first time intervals of 2nd simulations of all veh classes: %5.3f (number of vehicles %d)" ttMeaNumber tt noVeh
   
        retrieveTravelTimesAllSimulations
        retrieveTravelTimesCurrentSimulation
        retrieveTravelTimes2ndSimulation1stInterval
    
    let retrieveDataCollectionMeasurements =
        let dcNumber = 1
        let dcMeas = vissim.Net.DataCollectionMeasurements.ItemByKey(dcNumber)
        let noVeh = dcMeas.AttValue("Vehs(Avg,1,All") :?> int
        let spd = dcMeas.AttValue("Speed(Avg,1,All") :?> double
        let accel = dcMeas.AttValue("Acceleration(Avg,1,All") :?> double
        let len = dcMeas.AttValue("Length(Avg,1,All") :?> double
        printfn "Data Collection # %d avg values of all sim runs vehs: %d spd: %f accel: %f len: %f" dcNumber noVeh spd accel len
 
    let retrieveQueueCounters =
        let qcNumber = 1
        let maxQ = vissim.Net.QueueCounters.ItemByKey(qcNumber).AttValue("QLenMax(Avg,Avg") :?> double
        printfn "Avg max queue length of all sim runs and time intervals of q counter %d: %f" qcNumber maxQ
   
    generateResultsFromSimRuns
    retrieveTravelTimes
    retrieveDataCollectionMeasurements
    retrieveQueueCounters
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
    
    printfn "\n***Loop Test Example:"
    vissim |> loopTest |> ignore
        
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

