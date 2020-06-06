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
using Vissim.ComProvider.TypeLibs;

namespace CSharpTest
{
    class Program
    {
        public readonly static string ComExampleFolder = @"C:\Users\Public\Documents\PTV Vision\PTV Vissim 2020\Examples Training\COM\";
        public readonly static string NetworkFile = ComExampleFolder + @"Basic Commands\COM Basic Commands.inpx";
        public readonly static string LayoutFile = ComExampleFolder + @"Basic Commands\COM Basic Commands.layx";

        static void Main(string[] args)
        {
            var vissim = new VissimLib200.VissimClass();
            vissim.LoadNet(NetworkFile);
            vissim.Graphics.AttValue["QuickMode"] = true; 
            vissim.Simulation.RunContinuous();
            vissim.Exit();
            Console.ReadLine();
        }
    }
}
