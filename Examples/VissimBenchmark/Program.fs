module Vissim.Benchmarks.Program

open System
open Benchmarks

[<EntryPoint; STAThread>]
let main argv =
    use rtFactorBenchmark = new RealtimeFactorBenchmark(360u)
    rtFactorBenchmark.run()
    Console.WriteLine("\nPress any key to exit.") |> ignore
    Console.ReadLine() |> ignore
    0
