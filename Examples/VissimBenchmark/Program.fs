namespace Vissim.Benchmarks

open System

module Program =
    [<EntryPoint; STAThread>]
    let main argv =
        Jobs.RealtimeFactorBenchmark.run() |> ignore
        0
