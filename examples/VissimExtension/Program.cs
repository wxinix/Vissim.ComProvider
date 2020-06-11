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

using System;
using System.Runtime.InteropServices;
using Vissim.ComProvider.TypeLibs;

/// <summary>
/// This example illustrates application of LINQ Query Expression on Vissim COM interface that inherits ICollectionBase.
/// </summary>
namespace Vissim.ComProvider.Examples.VissimExtension
{
    /// <summary>
    /// We define VissimExtension class to add two methods HideMainWindow and ShowMainWindow to IVissim interface.
    /// These two methods become (i.e., extends) part of IVissim.
    /// </summary>
    public static class VissimExtension
    {
        [DllImport("Vissim.ComProvider.Utilities.dll", CallingConvention = CallingConvention.StdCall)]
        static extern void HideVissim(IntPtr unk);

        [DllImport("Vissim.ComProvider.Utilities.dll", CallingConvention = CallingConvention.StdCall)]
        static extern void ShowVissim(IntPtr unk);
        
        public static void HideMainWindow(this VissimLib200.IVissim vissim)
        {
            var unk = Marshal.GetIUnknownForObject(vissim);
            HideVissim(unk);
            Marshal.Release(unk);
        }

        public static void RestoreMainWindow(this VissimLib200.IVissim vissim)
        {
            var unk = Marshal.GetIUnknownForObject(vissim);
            ShowVissim(unk);
            Marshal.Release(unk);
        }
    }

    class Program
    {
        static readonly string _exampleFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM\";
        static readonly string _layoutFile    = _exampleFolder + @"Basic Commands\COM Basic Commands.layx";
        static readonly string _networkFile   = _exampleFolder + @"Basic Commands\COM Basic Commands.inpx";
        
        static void Main()
        {
            var vissim = new VissimLib200.VissimClass();
            vissim.LoadNet(_networkFile);
            vissim.LoadLayout(_layoutFile);

            Console.WriteLine("Magic Happens Here! Vissim will disappear!");
            vissim.HideMainWindow();    // This will hide Vissim main window.
            
            vissim.Simulation.AttValue["SimBreakAt"] = 60;
            vissim.Simulation.RunContinuous();
            vissim.RestoreMainWindow(); // This will restore Vissim main window.
            
            Console.WriteLine("Magic Happens again! Vissim will show up!");
            
            vissim.Exit();
            Console.ReadLine();
        }
    }
}
