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
open System.IO
open System.Reflection
open System.Runtime.InteropServices

// Vissim COM Type Provider collects all installed Vissim COM type libraries and make them
// part of the compiler type system. There is no need to explicitly import the type library
// or add references.

// Alias to Vissim 2020 COM Type Lib
type VissimLib =
     Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

[<DllImport("Vissim.ComProvider.Utilities.dll", CallingConvention = CallingConvention.StdCall)>]
extern void HideVissim(nativeint unk);

[<DllImport("Vissim.ComProvider.Utilities.dll", CallingConvention = CallingConvention.StdCall)>]
extern void ShowVissim(nativeint unk);

// F# Type Extension to extend IVissim. In C#, you would need Extension Methods (see example "VissimExtension")
type VissimLib.IVissim with
    member this.HideMainWindow () =
        let unk = Marshal.GetIUnknownForObject this
        HideVissim(unk)
        Marshal.Release(unk)

    member this.RestoreMainWindow () =
        let unk = Marshal.GetIUnknownForObject this
        ShowVissim(unk)
        Marshal.Release(unk)

[<EntryPoint; STAThread>]
let main argv =
    let vissim = VissimLib.VissimClass()
    let network =
        let exePath = Uri(Assembly.GetEntryAssembly().GetName().CodeBase).LocalPath
        FileInfo(exePath).Directory.FullName + "\\VissimBenchmark.inpx"
    vissim.LoadNet network
    printfn "Starting benchmarking Vissim on %s" Environment.MachineName
    let simPeriod = vissim.Simulation.AttValue("SimPeriod") :?> int
    let stopWatch = new System.Diagnostics.Stopwatch ()

    let timeRunVissimHiddenGui =
        printfn "\nRunning Vissim with hidden Main Window now..."
        vissim.HideMainWindow () |> ignore
        stopWatch.Restart()
        vissim.Simulation.RunContinuous()
        stopWatch.Stop()
        let time = stopWatch.ElapsedMilliseconds
        vissim.RestoreMainWindow () |> ignore
        let realtimeFactor = (double simPeriod * 1000.0) / double time

        (time, simPeriod, realtimeFactor)
        |||> printfn "Vissim Hidden Mode: takes [%d] ms to complete [%d] sec simulation, realtime factor [%5.3f]"
        "HiddenGui", double time/1000.0, simPeriod, realtimeFactor

    let timeRunVissimTurbo =
        printfn "\nRunning Vissim in Turbo Mode now..."
        vissim.Graphics.AttValue("QuickMode") <- true
        vissim.SuspendUpdateGUI()
        vissim.HideMainWindow () |> ignore
        stopWatch.Restart()
        vissim.Simulation.RunContinuous()
        stopWatch.Stop()
        let time = stopWatch.ElapsedMilliseconds
        vissim.RestoreMainWindow () |> ignore
        vissim.Graphics.AttValue("QuickMode") <- false
        vissim.ResumeUpdateGUI()
        let realtimeFactor = (double simPeriod * 1000.0) / double time

        (time, simPeriod, realtimeFactor)
        |||> printfn "Vissim Turbo Mode: takes [%d] ms to complete [%d] sec simulation, realtime factor [%5.3f]"
        "Turbo", double time/1000.0, simPeriod, realtimeFactor

    let timeRunVissimNormalGui =
        printfn "\nRunning Vissim with normal Main Window now..."
        stopWatch.Restart()
        vissim.Simulation.RunContinuous()
        stopWatch.Stop()
        let time = stopWatch.ElapsedMilliseconds

        let realtimeFactor = (double simPeriod * 1000.0) / double time

        (time, simPeriod, realtimeFactor)
        |||> printfn "Vissim Visual Mode: takes [%d] ms to complete [%d] sec simulation, realtime factor %f"
        "NormalGui", double time/1000.0, simPeriod, realtimeFactor

    vissim.Exit()

    printfn "\n----------------------------------------------------------------------------------------"
    printfn "%-10s\t%-15s\t%-15s\t%-20s" "Mode" "TimeTaken(s)" "SimPeriod(s)" "RealtimeFactor(x1)"
    [timeRunVissimHiddenGui; timeRunVissimTurbo; timeRunVissimNormalGui] |> List.iter(
        fun el ->
            let (mode, time, simPeriod, rtFactor) = el
            printfn "%-10s\t%-15.2f\t%-15d\t%-4.1f" mode time simPeriod rtFactor)
    printfn "----------------------------------------------------------------------------------------\n"

    Console.WriteLine("Benchmarking done. Press any key to exit.") |> ignore
    Console.ReadLine() |> ignore
    0

