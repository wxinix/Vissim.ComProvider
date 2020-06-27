module Vissim.LoopPerformance.Program

open System
open LoopTest

[<EntryPoint; STAThread>]
let main argv =
    use loopTest = new LoopTest(5u)
    loopTest.run()
    Console.WriteLine("\nPlease enter any key to exit.")
    Console.ReadLine() |> ignore
    0


