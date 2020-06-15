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
using System.Linq;
using Vissim.ComProvider.TypeLibs;

/// <summary>
/// This example illustrates application of LINQ Query Expression on Vissim COM interface that inherits ICollectionBase.
/// </summary>
namespace Vissim.ComProvider.Examples.VissimLINQ
{
    // Define a type alias
    using Link = VissimLib200.ILink;

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

            // Use the expressive and powerful LINQ Query Expression on any  Vissim COM interface that  inherits
            // ICollectionBase, for example ILinkContainer. Actually, when  Vissim  COM types are imported  from
            // the type  library, Vissim  ICollectionBase interface is  automatically translated  as  inheriting
            // IEnumerable, because of the special flag [id(0xfffffffc)] specified to its _NewEnum  method. This
            // enables foreach loop and LINQ Query Expression on any Vissim COM interface object that implements
            // ICollectionBase.

            // Alternatively, IConnectionBase.FilteredB() method can also be used to query on conditions, though
            // FilteredBy() is obviously less expressive, powerful, and elegant than the full-fledged LINQ Query
            // Expression.

            var selectedLinks =   from  Link link in vissim.Net.Links
                                 where  (int)link.AttValue["No"] > 10
                                select  link;

            foreach (Link link in selectedLinks) {
                Console.WriteLine("LINQ selected link: {0}", link.AttValue["No"].ToString());
            }

            vissim.Graphics.AttValue["QuickMode"] = true;
            vissim.Simulation.RunContinuous();
            vissim.Exit();
            Console.ReadLine();
        }
    }
}
