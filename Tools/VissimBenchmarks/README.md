# Vissim Benchmarking Tool

This tool benchmarks your workstation in terms of Vissim 2020 realtime factor. It is useful for determining the specs of an optimial workstation to suit simulation needs.

## Usage

Just double click VissimBenchmarks.exe. It will launch Vissim and run three benchmarking scenarios:
- Vissim Hidden Mode, where Vissim main window will be hidden from the desktop while running.
- Vissim Turboo Mode, where Vissim main windwow will be hidden, with QuickMode on, GUI and data list view updates suspended.
- Vissim Normal Mode, where Vissim will just run normal with  2-D animation, GUI and data list view updates.

## Interpretations

The higher the realtime factor, the more powerful your workstation. However, use at your own risk and make your own judgement call. For a decent workstation
in the year 2020, a x20 realtime factor is reasonably expected for the "Turbo" mode. The following screenshot illustrates a sample of benchmarking results:

![Image of Vissim Benchmarking Results](https://github.com/wxinix/Vissim.ComProvider/blob/master/Tools/VissimBenchmarks/Sample.jpg)


## Limitation

This tool only supports Vissim 2020. You can modify/edit the F# [source code](https://github.com/wxinix/Vissim.ComProvider/tree/master/Examples/VissimBenchmark) for other Vissim versions.

Copyright 2020 (c) Wuping Xin