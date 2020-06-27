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

module Vissim.Benchmarks.Jobs

open System
open System.IO
open System.Reflection
open System.Runtime.InteropServices

// Alias to Vissim 2020 COM Type Lib
type VissimLib =
    Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

module private SysApi =
    [<DllImport("Vissim.ComProvider.Utilities.dll", CallingConvention = CallingConvention.StdCall)>]
    extern void HideVissim(nativeint unk);

    [<DllImport("Vissim.ComProvider.Utilities.dll", CallingConvention = CallingConvention.StdCall)>]
    extern void ShowVissim(nativeint unk);
    // F# Type Extension to extend IVissim. In C#, you would need Extension Methods (see example "VissimExtension")

type VissimLib.IVissim with
    member this.HideMainWindow () =
        let unk = Marshal.GetIUnknownForObject this
        SysApi.HideVissim(unk)
        Marshal.Release(unk)

    member this.RestoreMainWindow () =
        let unk = Marshal.GetIUnknownForObject this
        SysApi.ShowVissim(unk)
        Marshal.Release(unk)

type VissimLib.ISimulation with
    member x.RunHidden(vissim: VissimLib.IVissim) =
        vissim.HideMainWindow() |> ignore
        x.RunContinuous()
        x.Stop()
        vissim.RestoreMainWindow() |> ignore

    member x.Run(vissim: VissimLib.IVissim) =
        vissim.RestoreMainWindow() |> ignore
        vissim.ResumeUpdateGUI()
        vissim.Graphics.AttValue("QuickMode") <- false
        x.RunContinuous()
        x.Stop()

    member x.RunTurbo(vissim: VissimLib.IVissim) =
        vissim.SuspendUpdateGUI()
        vissim.Graphics.AttValue("QuickMode") <- true
        vissim.HideMainWindow() |> ignore
        x.RunContinuous()
        x.Stop()

    member x.SimPeriod
        with get() = x.AttValue("SimPeriod") :?> int
        and set(value) = x.AttValue("SimPeriod") <- value

module RealtimeFactorBenchmark =
    let private benchmark =
        let vissim = VissimLib.VissimClass()       
        let network =
            let exePath = Uri(Assembly.GetEntryAssembly().GetName().CodeBase).LocalPath
            FileInfo(exePath).Directory.FullName + "\\VissimBenchmarks.inpx"
        vissim.LoadNet network
        
        let simPeriod =
            vissim.Simulation.SimPeriod <- 360
            vissim.Simulation.SimPeriod

        let stopWatch = new System.Diagnostics.Stopwatch()
    
        fun mode onExecute onDone ->
            printfn "\n - Running [%s] mode now..." mode
            stopWatch.Restart()
            onExecute(vissim)
            stopWatch.Stop()
            let time = double stopWatch.ElapsedMilliseconds / 1000.0
            let realtimeFactor = double simPeriod / time
            printfn " - Takes %6.2f seconds" time        
            onDone(vissim)
            (mode, time, simPeriod, realtimeFactor)

    let private print result =
        printfn "\n----------------------------------------------------------------------------------------"
        printfn " %-10s\t%-15s\t%-15s\t%-20s" "Mode" "TimeTaken(s)" "SimPeriod(s)" "RealtimeFactor(x1)"
        result |> List.iter( fun el -> 
                                    let (mode, time, simPeriod, rtFactor) = el
                                    printfn " %-10s\t%-15.2f\t%-15d\t%-4.1f" mode time simPeriod rtFactor )
        printfn "----------------------------------------------------------------------------------------\n"

    let run() =
        printfn "Starting benchmarking Vissim realtime factor on %s" Environment.MachineName
        SystemInfo.print ()

        let timeHidden = benchmark "Hidden"
                                    (fun vsm -> vsm.Simulation.RunHidden vsm)
                                    (fun _ -> ())                            
        let timeTurboo = benchmark "Turboo"
                                    (fun vsm -> vsm.Simulation.RunTurbo vsm)
                                    (fun _ -> ())
        let timeNormal = benchmark "Normal"
                                    (fun vsm -> vsm.Simulation.Run vsm)
                                    (fun vsm -> vsm.Exit())

        print [timeHidden; timeTurboo; timeNormal]
        Console.WriteLine("Benchmarking done. Press any key to exit.") |> ignore
        Console.ReadLine() |> ignore