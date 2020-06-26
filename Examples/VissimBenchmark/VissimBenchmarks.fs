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

module private Vissim.Benchmarks

open System
open System.IO
open System.Reflection
open System.Runtime.InteropServices

module private SysApi =
    [<DllImport("Vissim.ComProvider.Utilities.dll", CallingConvention = CallingConvention.StdCall)>]
    extern void HideVissim(nativeint unk);

    [<DllImport("Vissim.ComProvider.Utilities.dll", CallingConvention = CallingConvention.StdCall)>]
    extern void ShowVissim(nativeint unk);

// Alias to Vissim 2020 COM Type Lib
type VissimTypes =
    Vissim.ComProvider.``Vissim Object Library 20.0 64 Bit``.``14.0-win64``

 // F# Type Extension to extend IVissim. In C#, you would need Extension Methods (see example "VissimExtension")
 type VissimTypes.IVissim with
    member this.HideMainWindow () =
        let unk = Marshal.GetIUnknownForObject this
        SysApi.HideVissim(unk)
        Marshal.Release(unk)

    member this.RestoreMainWindow () =
        let unk = Marshal.GetIUnknownForObject this
        SysApi.ShowVissim(unk)
        Marshal.Release(unk)

type VissimTypes.ISimulation with
    member x.RunHidden(vissim: VissimTypes.IVissim) =
        vissim.HideMainWindow() |> ignore
        x.RunContinuous()
        x.Stop()
        vissim.RestoreMainWindow() |> ignore

    member x.Run() =
        x.RunContinuous()
        x.Stop()

    member x.RunTurbo(vissim: VissimTypes.IVissim) =
        vissim.SuspendUpdateGUI()
        vissim.Graphics.AttValue("QuickMode") <- true
        vissim.HideMainWindow() |> ignore
        x.RunContinuous()
        x.Stop()
        vissim.RestoreMainWindow() |> ignore
        vissim.ResumeUpdateGUI()
        vissim.Graphics.AttValue("QuickMode") <- false

    member x.SimPeriod
        with get() = x.AttValue("SimPeriod") :?> int

module RealtimeFactorBenchmark =
    let private benchmark =
        let vissim = VissimTypes.VissimClass() :> VissimTypes.IVissim            
        let network =
            let exePath = Uri(Assembly.GetEntryAssembly().GetName().CodeBase).LocalPath
            FileInfo(exePath).Directory.FullName + "\\VissimBenchmarks.inpx"
        vissim.LoadNet network
        let simPeriod = vissim.Simulation.SimPeriod
        let stopWatch = new System.Diagnostics.Stopwatch()        
            
        fun mode load exitOnDone ->
            printfn "\n - Running [%s] mode now..." mode
            stopWatch.Restart()
            load(vissim)
            stopWatch.Stop()
            let time = double stopWatch.ElapsedMilliseconds / 1000.0
            let realtimeFactor = double simPeriod / time
            printfn " - Takes %6.2f seconds" time                
            if exitOnDone then vissim.Exit()
            (mode, time, simPeriod, realtimeFactor)

    let run() =
        printfn "Starting benchmarking Vissim realtime factor on %s" Environment.MachineName
        let timeHidden = benchmark "Hidden" (fun vissim -> vissim.Simulation.RunHidden(vissim)) false
        let timeTurboo = benchmark "Turboo" (fun vissim -> vissim.Simulation.RunTurbo(vissim))  false
        let timeNormal = benchmark "Normal" (fun vissim -> vissim.Simulation.Run())             true

        printfn "\n----------------------------------------------------------------------------------------"
        printfn "%-10s\t%-15s\t%-15s\t%-20s" "Mode" "TimeTaken(s)" "SimPeriod(s)" "RealtimeFactor(x1)"

        [timeHidden; timeTurboo; timeNormal] |> List.iter(
            fun el ->
                let (mode, time, simPeriod, rtFactor) = el
                printfn "%-10s\t%-15.2f\t%-15d\t%-4.1f" mode time simPeriod rtFactor)

        printfn "----------------------------------------------------------------------------------------\n"

        Console.WriteLine("Benchmarking done. Press any key to exit.") |> ignore
        Console.ReadLine() |> ignore
        
module Main =
    [<EntryPoint; STAThread>]
    let main argv =
        RealtimeFactorBenchmark.run() |> ignore
        0

