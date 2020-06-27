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

module Vissim.LoopPerformance.LoopTest

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

let printvi vehInfo =
    let no, ty, spd, pos, lane = vehInfo
    printfn "No %d\t | Type %s\t | Speed %6.3f\t | Pos %10.3f\t | Lane %10s\t" no ty spd pos lane

let printv (veh: VissimLib.IVehicle) =
    let no   = veh.AttValue(     "No") :?> int
    let ty   = veh.AttValue("VehType") |>  string
    let spd  = veh.AttValue(  "Speed") :?> double
    let pos  = veh.AttValue(    "Pos") :?> double
    let lane = veh.AttValue(   "Lane") |>  string
    printvi (no, ty, spd, pos, lane)

let arrayLoop(vissim: VissimLib.IVissim) =
    let allVehs = vissim.Net.Vehicles.GetAll().AsArray<VissimLib.IVehicle>()
    allVehs |> Array.iter(fun veh -> printv veh)

let enumLoop(vissim: VissimLib.IVissim) =
    for veh in vissim.Net.Vehicles do
        printv (veh :?> VissimLib.IVehicle)

let iterLoop(vissim: VissimLib.IVissim) =
    let iter = vissim.Net.Vehicles.Iterator
    while iter.Valid do
        let veh = iter.Item :?> VissimLib.IVehicle
        printv veh
        iter.Next()

let multiAttValuesLoop(vissim: VissimLib.IVissim) =
    let numbers   = vissim.Net.Vehicles.GetMultiAttValues(     "No") :?> Object [,]
    let types     = vissim.Net.Vehicles.GetMultiAttValues("VehType") :?> Object [,]
    let speeds    = vissim.Net.Vehicles.GetMultiAttValues(  "Speed") :?> Object [,]
    let positions = vissim.Net.Vehicles.GetMultiAttValues(    "Pos") :?> Object [,]
    let lanes     = vissim.Net.Vehicles.GetMultiAttValues(   "Lane") :?> Object [,]
    for i in 0 .. Array2D.length1 numbers - 1 do
        printvi (numbers.[i,1] :?> int, types.[i,1] |> string, speeds.[i,1] :?> double, positions.[i,1] :?> double, lanes.[i,1] |> string)

let multiAttrsLoop(vissim: VissimLib.IVissim) =
    let attrs = vissim.Net.Vehicles.GetMultipleAttributes(id<obj[]> [| "No";"VehType";"Speed";"Pos"; "Lane" |]) :?> Object[,]
    for i in 0.. Array2D.length1 attrs - 1 do
        printvi (attrs.[i,0] :?> int, attrs.[i,1] |> string, attrs.[i,2] :?> double, attrs.[i,3] :?> double, attrs.[i,4] |>string)

type LoopTest (simPeriod: uint) =
    let vissim =
        let vissim = VissimLib.VissimClass()
        vissim |> loadNetwork |> loadLayout |> ignore
        vissim.Graphics.AttValue("QuickMode")    <- true
        vissim.Simulation.AttValue("SimBreakAt") <- max simPeriod 60u
        vissim.SuspendUpdateGUI()
        vissim.Simulation.RunContinuous()
        vissim

    let benchmark =
        let stopWatch = System.Diagnostics.Stopwatch.StartNew()
        fun (name, loop) ->
            stopWatch.Restart()
            printfn "\n----------------------------------------------------------"
            printfn " Benchmarking %s" name
            printfn "----------------------------------------------------------"
            do loop vissim
            stopWatch.Stop()
            name, double stopWatch.ElapsedMilliseconds / 1000.0

    member _.run() =
        let results = [ ("ArrayLoop", arrayLoop); ("EnumLoop", enumLoop); ("IterLoop", iterLoop); ("MultiAttValuesLoop", multiAttValuesLoop); ("MultipleAttrsLoop", multiAttrsLoop) ]
                      |> List.map(fun job -> benchmark job)

        printfn "\n----------------------------------------------------------"
        printfn " Loop Performance Benchmarking Summary"
        printfn "----------------------------------------------------------"

        results |> List.iter(fun result -> printfn "%-18s\t%f" (fst result) (snd result))

    interface IDisposable with
        member _.Dispose() = vissim.Exit()