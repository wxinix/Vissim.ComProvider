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

module private Vissim.Benchmarks.SystemInfo

open System.Management

let private cpu = new ManagementObjectSearcher( "select * from Win32_Processor"       );
let private gpu = new ManagementObjectSearcher( "select * from Win32_VideoController" );
let private os  = new ManagementObjectSearcher( "select * from Win32_OperatingSystem" );
let private ram = new ManagementObjectSearcher( "select * from Win32_PhysicalMemory"  );

let print() =
    printfn "\n----------------------------------------------------------------------------------------"

    seq { for obj in gpu.Get() do yield (string obj.["Name"]) }
    |> Seq.iteri( fun i el -> 
                        printfn " %-20s\t%s" 
                                ("GPU-" + (string i)) 
                                el )

    seq { for obj in cpu.Get() do yield string obj.["Name"], obj.["NumberOfCores"] :?> uint32, obj.["NumberOfLogicalProcessors"] :?> uint32 }
    |> Seq.iteri( fun i (name, coresCnt, logicalProcessorsCnt) ->
                        printfn " %-20s\t%s, %d cores %d logical processors" 
                                ("CPU-" + (string i))
                                name
                                coresCnt
                                logicalProcessorsCnt )

    seq { for obj in os.Get() do yield string obj.["Caption"], string obj.["Version"] }
    |> Seq.iteri( fun i (osName, osVer) ->
                        printfn " %-20s\t%s %s"
                                ("OS-" + (string i)) 
                                osName
                                osVer)

    seq { for obj in ram.Get() do yield (obj.["Capacity"] :?> uint64) / (1024UL * 1024UL * 1024UL) }
    |> Seq.sum
    |> printfn " %-20s\t%d GB"  "Physical Memory"

    printfn "----------------------------------------------------------------------------------------"


