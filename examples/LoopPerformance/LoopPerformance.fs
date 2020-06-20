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

// The COM Type Provider collects all installed Vissim COM type libraries and make
// them part of the compiler type system.There is no need to explicitly import the
// type library or add references.

// Alias to Vissim 2020 COM Type Lib
type VissimLib =
     Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

type Object with
    member this.AsArray<'T> () = this :?> Object [] |> Array.map(fun o -> o :?> 'T)

let [<Literal>] ExampleFolder =
    @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM\"

let [<Literal>] NetworkFile =
    ExampleFolder + @"Basic Commands\COM Basic Commands.inpx"

let [<Literal>] LayoutFile =
    ExampleFolder  + @"Basic Commands\COM Basic Commands.layx"

let loadNetwork (vissim: VissimLib.IVissim) =
    vissim.LoadNet NetworkFile
    NetworkFile |> printfn "Network \"%s\" loaded."
    vissim

let loadLayout (vissim: VissimLib.IVissim) =
    vissim.LoadLayout LayoutFile
    LayoutFile |> printfn "Layout \"%s\" loaded."
    vissim

let loopTest (vissim: VissimLib.IVissim) =
    let stopWatch = System.Diagnostics.Stopwatch.StartNew()

    let arrayLoopTime =
        let allVehicles = vissim.Net.Vehicles.GetAll().AsArray<VissimLib.IVehicle>()
        stopWatch.Stop()
        let timeGetAllElements = stopWatch.ElapsedMilliseconds;
        printfn "----------------------------------------------------------"
        stopWatch.Restart()
        for i in 0..allVehicles.Length - 1 do
            let veh = allVehicles.[i]
            let vehNumber   = veh.AttValue(     "No") :?> int
            let vehType     = veh.AttValue("VehType") |>  string
            let vehSpeed    = veh.AttValue(  "Speed") :?> double
            let vehPos      = veh.AttValue(    "Pos") :?> double
            let vehLinkLane = veh.AttValue(   "Lane") |>  string
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t"
                    vehNumber vehType vehSpeed vehPos vehLinkLane
        stopWatch.Stop()
        timeGetAllElements, stopWatch.ElapsedMilliseconds

    let enumLoopTime =
        stopWatch.Restart()
        for vehObj in vissim.Net.Vehicles do
            let veh = vehObj :?> VissimLib.IVehicle
            let vehNumber   = veh.AttValue(     "No") :?> int
            let vehType     = veh.AttValue("VehType") |>  string
            let vehSpeed    = veh.AttValue(  "Speed") :?> double
            let vehPos      = veh.AttValue(    "Pos") :?> double
            let vehLinkLane = veh.AttValue(   "Lane") |>  string
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t"
                     vehNumber vehType vehSpeed vehPos vehLinkLane
        stopWatch.Stop()
        stopWatch.ElapsedMilliseconds

    let iterLoopTime =
        printfn "----------------------------------------------------------"
        stopWatch.Restart()
        let iter = vissim.Net.Vehicles.Iterator
        while iter.Valid do
            let veh = iter.Item :?> VissimLib.IVehicle
            let vehNumber   = veh.AttValue(     "No") :?> int
            let vehType     = veh.AttValue("VehType") |>  string
            let vehSpeed    = veh.AttValue(  "Speed") :?> double
            let vehPos      = veh.AttValue(    "Pos") :?> double
            let vehLinkLane = veh.AttValue(   "Lane") |>  string
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t"
                    vehNumber vehType vehSpeed vehPos vehLinkLane
            iter.Next()
        stopWatch.Stop()
        stopWatch.ElapsedMilliseconds

    let multiAttValuesLoopTime =
        printfn "----------------------------------------------------------"
        stopWatch.Restart()
        let vehNumbers   = vissim.Net.Vehicles.GetMultiAttValues(     "No") :?> Object [,]
        let vehTypes     = vissim.Net.Vehicles.GetMultiAttValues("VehType") :?> Object [,]
        let vehSpeeds    = vissim.Net.Vehicles.GetMultiAttValues(  "Speed") :?> Object [,]
        let vehPositions = vissim.Net.Vehicles.GetMultiAttValues(    "Pos") :?> Object [,]
        let vehLinkLanes = vissim.Net.Vehicles.GetMultiAttValues(   "Lane") :?> Object [,]
        for i in 0 .. Array2D.length1 vehNumbers - 1 do
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t"
                    (vehNumbers.[i,1]   :?> int   )
                    (vehTypes.[i,1]     |>  string)
                    (vehSpeeds.[i,1]    :?> double)
                    (vehPositions.[i,1] :?> double)
                    (vehLinkLanes.[i,1] |>  string)
        stopWatch.Stop()
        stopWatch.ElapsedMilliseconds

    let multipleAttributesLoopTime =
        printfn "----------------------------------------------------------"
        stopWatch.Restart()
        let allVehAttrs = vissim.Net.Vehicles.GetMultipleAttributes(id<obj[]> [| "No";"VehType";"Speed";"Pos"; "Lane" |]) :?> Object[,]
        for i in 0.. Array2D.length1 allVehAttrs - 1 do
            printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t"
                    (allVehAttrs.[i,0] :?> int   )
                    (allVehAttrs.[i,1] |>  string)
                    (allVehAttrs.[i,2] :?> double)
                    (allVehAttrs.[i,3] :?> double)
                    (allVehAttrs.[i,4] |>  string)
        stopWatch.Stop()
        stopWatch.ElapsedMilliseconds

    let (timeGettingAllElements, timeDoingGetAllLoop) = arrayLoopTime

    printfn "----------------------------------------------------------"
    printfn "%10s \t%10s \t%10s \t%10s \t%18s \t%18s"
            "---"
            "ArrayLoop"
            "EnumLoop"
            "IterLoop"
            "MultiAttValLoop"
            "MultipleAttrsLoop"

    printfn "%10s \t%10s \t%10s \t%10s \t%18s \t%18s" "Time(ms)"
            ((fst(arrayLoopTime) |> string) + "|" + ((fst(arrayLoopTime) + snd(arrayLoopTime)) |> string))
            (enumLoopTime                 |> string)
            (iterLoopTime                 |> string)
            (multiAttValuesLoopTime       |> string)
            (multipleAttributesLoopTime   |> string)

    let baseLineTime = multipleAttributesLoopTime

    printfn "%10s \t%10.3f \t%10.3f \t%10.3f \t%18.3f \t%18.3f"
            "Factor(x1)"
            ((double timeGettingAllElements + double timeDoingGetAllLoop) / double baseLineTime)
            (double enumLoopTime / double baseLineTime)
            (double iterLoopTime / double baseLineTime)
            (double multiAttValuesLoopTime / double baseLineTime)
            (double multipleAttributesLoopTime / double baseLineTime)

    vissim

[<EntryPoint; STAThread>]
let main argv =
    let vissim = VissimLib.VissimClass()

    vissim |> loadNetwork |> loadLayout |> ignore
    vissim.Graphics.AttValue("QuickMode")    <- true
    vissim.Simulation.AttValue("SimBreakAt") <- 200
    vissim.SuspendUpdateGUI()
    vissim.Simulation.RunContinuous()
    printfn "\n***Loop Performance Benchmarking:"
    vissim |> loopTest |> ignore
    vissim.Exit()
    Console.ReadLine() |> ignore
    0

