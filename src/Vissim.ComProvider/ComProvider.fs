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

// The UNLICENSE
// Copyright (c) Microsoft Corporation 2005-2012.
// A license with no conditions whatsoever  which  dedicates works to the  public
// domain. Unlicensed  works, modifications,  and larger works may be distributed
// under different terms and without source code.

namespace FSharp.Interop.ComProvider.Vissim

open System
open System.IO
open System.Reflection
open Microsoft.FSharp.Core.CompilerServices
open ProviderImplementation.ProvidedTypes
open FSharp.Interop.ComProvider.TypeLibInfo
open FSharp.Interop.ComProvider.TypeLibImport

[<TypeProvider>]
type ComProvider(cfg: TypeProviderConfig) as this =
    inherit TypeProviderForNamespaces()

    /// Currently executing assembly.
    let asm = Assembly.GetExecutingAssembly()

    (*
     The TypeLib registry key allows specifying separate type libraries for platform affinity.
     However, the type provider has no way to know what target platform the subject project
     will be compiled to beforehand. Furthermore, the target platform isn't actually known until
     runtime if "Any CPU" is selected. Therefore, we have to leave this as the developer's choice
     in selecting the type library with proper platform affinity at source code level.

     If the type library is for an out-of-process COM automation server, the client application
     and server do not have to have the same CPU platform, that is, a win32 client can communicate
     with a win64 COM server. In this case, we can choose win32 or win64 type library as we like,
     while seting either "Any CPU", "x86", or "x64" as the compiling option for the client application.

     If the type library is for an in-process COM Dll, the client application and the COM Dll must
     have the same target platform. This means, if the COM Dll is 32 bit, the compiling option must
     be consistent, and must be set to x86 CPU as well.
 
    // Depreciated. See above comments.
    let preferredPlatform = 
        if cfg.IsHostedExecution && Environment.Is64BitProcess then "win64" else "win32"
    *)

    /// Takes one of the three values, "win32", "win64", or "*" for both.
    let preferredPlatform = "*"

    // We use nested types as opposed to namespaces for the following reasons:
    // 1. No way with ProvidedTypes API to have sub-namespaces generated on demand.
    // 2. Namespace components cannot contain dots, which are common both in the
    // type library name itself and of course the major.minor version number.
    let types =
        [ for name, libsByName in loadTypeLibs preferredPlatform "Vissim" |> Seq.groupBy (fun lib -> lib.Name) do
            let nameTy = ProvidedTypeDefinition(asm, "COM", name, None) // COM is namespaceName.
            yield nameTy

            for verStr, libsByVer in libsByName |> Seq.groupBy(fun lib -> lib.Version.VersionStr) do
                for lib in libsByVer do
                    let subTy = ProvidedTypeDefinition(TypeContainer.TypeToBeDecided, verStr + "-" + lib.Platform, None)
                    nameTy.AddMember(subTy)
                    subTy.IsErased <- false
                    subTy.AddMembersDelayed <| fun _ ->
                        lib.Pia |> Option.iter (fun pia ->
                            failwithf "Accessing type libraries with Primary Interop Assemblies using the COM Type Provider not supported. Directly referencing the assembly '%s' instead." pia)                           
                        let tempDir = Path.Combine(cfg.TemporaryFolder, "FSharp.Interop.ComProvider", Guid.NewGuid().ToString())
                        Directory.CreateDirectory(tempDir) |> ignore
                        let assemblies = importTypeLib lib.Path tempDir
                        // assemblies |> List.iter(fun asm -> ProvidedAssembly.RegisterGenerated(asm.Location) |> ignore)
                        assemblies |> List.collect(fun asm -> asm.GetTypes() |> Seq.toList) ]

    do
        this.AddNamespace("COM", types)
        // this.RegisterProbingFolder(cfg.TemporaryFolder)

[<TypeProviderAssembly>]
()
