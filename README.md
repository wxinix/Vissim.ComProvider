# Vissim COM Type Provider

Vissim COM type provider is a F# compiler plugin, making Vissim COM type system part of the compiler type system. The added types are automatically visible to user code at compile time, without importing of any type libraries or adding reference assemblies explicilty.

The great thing is, whether you install new Vissim versions, uninstall old versions or whatever – you don’t have to import (or re-import) anything. The Vissim COM type provider will automatically pick up all of the installed Vissim COM type libraries. Another wonderful thing is that the same Vissim COM type provider can be used in both Vissim COM scripting and compiled applications transparently.

For more details, read my [blog article](https://blog.wupingxin.net/vissim-com-programming-for-fun-com-type-provider-a-new-way-of-doing-vissim-com-interop/) 

## Code Style Guidline

Microsoft [coding style guideline for F#](https://docs.microsoft.com/en-us/dotnet/fsharp/style-guide/formatting) is used.

## Maintainer of this Repository
- [@wxinix](https://github.com/wxinix)  Wuping Xin
