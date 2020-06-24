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

module Car2X

open System
open System.Collections.Generic

// Vissim COM Type Provider collects all installed Vissim COM type libraries and make them
// part of the compiler type system. There is no need to explicitly import the type library
// or add references.

// Alias to Vissim 2020 COM Type Lib
type VissimLib =
    Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

let [<Literal>] ExampleFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM\"
let [<Literal>] LayoutFile    = ExampleFolder + @"Car2X.PY\Car2X.layx"
let [<Literal>] NetworkFile   = ExampleFolder + @"Car2X.PY\Car2X.inpx"

// The following illustrates how to extend ISimulation interface with customized events.
type SimulationEventArgs =
    { CurStep: int
      TotalSteps: int
      Vissim: VissimLib.IVissim }

type SimulationEvents () =
    let simAfterEnd    = new Event<_>()
    let simAfterStart  = new Event<_>()
    let simBeforeEnd   = new Event<_>()
    let simBeforeStart = new Event<_>()
    let simStepOnEnd   = new Event<_>()
    let simStepOnStart = new Event<_>()

    member _.SimStepOnStart = simStepOnStart
    member _.SimStepOnEnd   = simStepOnEnd
    member _.SimBeforeStart = simBeforeStart
    member _.SimAfterStart  = simAfterStart
    member _.SimBeforeEnd   = simBeforeEnd
    member _.SimAfterEnd    = simAfterEnd

// A store
module SimulationEventStore =
    let TheStore = Dictionary<int, SimulationEvents>()

// Optional type extension for ISimulation interface
type VissimLib.ISimulation with
    // Run with events fired
    member x.Run vissim =
        let totalSteps = x.SimPeriod * x.SimRes
        let simStepOnStart = x.Events.SimStepOnStart
        let simStepOnEnd = x.Events.SimStepOnEnd

        // Firing SimBeforeStart and SimAfterStart events
        x.Events.SimBeforeStart.Trigger({CurStep = 0; TotalSteps = totalSteps; Vissim = vissim})
        x.RunSingleStep()
        x.Events.SimAfterStart.Trigger({CurStep = 1; TotalSteps = totalSteps; Vissim = vissim})
        // Firing SimStepOnStart event
        for step in 2..totalSteps - 1 do
            simStepOnStart.Trigger({CurStep = step; TotalSteps = totalSteps; Vissim = vissim})
            x.RunSingleStep()
            simStepOnEnd.Trigger({CurStep = step; TotalSteps = totalSteps; Vissim = vissim})
        // Firing SimBeforeEnd and SimAfterEndEvent
        x.Events.SimBeforeEnd.Trigger({CurStep = totalSteps - 1; TotalSteps = totalSteps; Vissim = vissim})
        x.RunSingleStep()
        x.Events.SimAfterEnd.Trigger({CurStep = totalSteps; TotalSteps = totalSteps; Vissim = vissim})
        x.Stop()

    member x.SimPeriod
        with get() = x.AttValue("SimPeriod") :?> int
        and set(value: int) = x.AttValue("SimPeriod") <- value

    member x.SimBreakAt
        with get() = x.AttValue("SimBreatAt") :?> int
        and set(value: int) = x.AttValue("SimBreakAt") <- value

    member x.SimRes
        with get() = x.AttValue("SimRes") :?> int
        and set(value: int) = x.AttValue("SimRes") <- value

    member x.Events
        with get(): SimulationEvents =
            match x.GetHashCode() |> SimulationEventStore.TheStore.TryGetValue  with
            | true, events -> events
            | _ -> let events = new SimulationEvents()
                   SimulationEventStore.TheStore.Add(x.GetHashCode(), events)
                   events

// Optional type extension for IScriptContainer interface
type VissimLib.IScriptContainer with
    member x.RemoveAll() =
        for script in x do
            x.RemoveScript(script :?> VissimLib.IScript)

type VissimLib.IVehicle with
    member x.Pos = x.AttValue("Pos") :?> float
    member x.CoordFront = (x.AttValue("CoordFront") |> string).Split ' ' |> Array.map(fun el -> float el)

let asArray arr2d =
    [| for i in 0.. (Array2D.length1 arr2d) - 1 -> arr2d.[i, 0..] |]

[<EntryPoint; STAThread>]
let main argv =
    let vissim = VissimLib.VissimClass()
    vissim.LoadNet NetworkFile
    vissim.Net.Scripts.RemoveAll()     // Remove all currently embedded scripts.
    let simulation = vissim.Simulation // Must save to a local, otherwise, each time vissim.Simulation will return a different reference
    simulation.Events.SimStepOnStart.Publish |> Event.add(
        fun arg ->
            let c2xVehTyNoMsg = "101"
            let c2xVehTyHasActiveMsg = "102"
            let distDistr = 1
            let speedIncident = 80

            let vehIncident =
                let attrsAllVehsInNet = arg.Vissim.Net.Vehicles.GetMultipleAttributes(id<obj[]> [| "RoutDecType";"RoutDecNo";"VehType";"No" |]) :?> Object[,] |> asArray
                let attrVehIncident = attrsAllVehsInNet |> Array.tryFind(
                                         fun a ->
                                             let p = (if isNull(a.[0]) then None else Some(a.[0] |>  string)),
                                                     (if isNull(a.[1]) then None else Some(a.[1] :?> int)),
                                                     a.[2] |> string // vehTy is string                            
                                             match p with
                                             | Some(routDecTy), Some(routDecNo), vehTy ->
                                                   (routDecTy = "PARKING") && (routDecNo = 1) && ((vehTy = c2xVehTyNoMsg) || (vehTy = c2xVehTyHasActiveMsg))
                                             | _ -> false )

                match attrVehIncident with
                | None -> None
                | Some x ->
                       let theVeh = arg.Vissim.Net.Vehicles.ItemByKey(x.[3]) // x.[3] is VehNo
                       theVeh.AttValue("C2X_HasCurrentMessage") <- 1
                       theVeh.AttValue("C2X_MessageOrigin")     <- theVeh.AttValue("Pos")
                       theVeh.AttValue("C2X_SendingMessage")    <- 1
                       Some theVeh
      
            let attrIDs: obj[] = [| "Pos";"VehType";"C2X_HasCurrentMessage";"C2X_MessageOrigin";"C2X_Message";"DesSpeed";"C2X_DesSpeedOld" |]
            // Update C2X vehicles state in range with the incident vehicle
            match vehIncident with
            | None -> ()
            | Some x ->
                  let vehsRecvMsg = arg.Vissim.Net.Vehicles.GetByLocation(x.CoordFront.[0], x.CoordFront.[1], distDistr)
                  let a = vehsRecvMsg.GetMultipleAttributes(attrIDs) :?> Object[,]     
 
                  for i in 0..(Array2D.length1 a) - 1 do
                      let p = a.[i,0] :?> float, a.[i,1] |> string, a.[i,5] :?> float // (pos, vehTy, desSpeed)
                      match p with
                      | (pos, vehTy, desSpeed) when (List.contains vehTy [c2xVehTyHasActiveMsg; c2xVehTyNoMsg]) && (pos < x.Pos) ->
                            a.[i,1] <- id<obj> c2xVehTyHasActiveMsg
                            a.[i,2] <- id<obj> 1
                            a.[i,3] <- id<obj> x.Pos
                            a.[i,4] <- id<obj> "Break down vehicle ahead!"
                            a.[i,5] <- id<obj> speedIncident
                            a.[i,6] <- id<obj> desSpeed
                      | _ -> ()
 
                  vehsRecvMsg.SetMultipleAttributes(attrIDs.[1..], a.[*,1..]) // Send new attribute back to Vissim
            
            let a = arg.Vissim.Net.Vehicles.GetMultipleAttributes(attrIDs) :?> Object[,]

            for i in 0..(Array2D.length1 a) - 1 do
                let p = a.[i, 0] :?> float,
                        (if isNull(a.[i,2]) then None else Some(a.[i,2] :?>   int)),
                        (if isNull(a.[i,3]) then None else Some(a.[i,3] :?> float)),
                        (if isNull(a.[i,6]) then None else Some(a.[i,6] :?> float)) // (pos, hasActiveMsg, posC2x, desSpeedOld)
                match p with
                | (pos, Some(hasActiveMsg), Some(posC2x), Some(desSpeedOld)) when hasActiveMsg = 1 && pos > posC2x ->
                      a.[i,1] <- id<obj> c2xVehTyNoMsg
                      a.[i,2] <- id<obj> 0
                      a.[i,3] <- id<obj> ""
                      a.[i,4] <- id<obj> ""
                      a.[i,5] <- id<obj> desSpeedOld
                      a.[i,6] <- id<obj> ""
                | _ -> ()

            arg.Vissim.Net.Vehicles.SetMultipleAttributes(attrIDs.[1..], a.[*,1..])
        )

    simulation.Run vissim
    Console.ReadLine() |> ignore
    0

