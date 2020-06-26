module Vissim.Benchmarks.Main

open System

[<EntryPoint; STAThread>]
let main argv =
    Vissim.Benchmarks.Jobs.RealtimeFactorBenchmark.run() |> ignore
    0