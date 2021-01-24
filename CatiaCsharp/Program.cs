using INFITF;
using MECMOD;
using PARTITF;
using ProductStructureTypeLib;
using SPATypeLib;
using NavigatorTypeLib;
using System;
using Microsoft.VisualStudio.OLE.Interop;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualBasic;
using HybridShapeTypeLib;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using System.Diagnostics;

//https://www.codeproject.com/Articles/1195963/Interoping-NET-and-Cplusplus-through-COM
//https://www.maruf.ca/files/caadoc/CAASysTechArticles/CAASysAutomationImpl.htm
//https://limbioliong.wordpress.com/2011/07/16/marshaling-a-safearray-of-managed-structures-by-com-interop-part-2/
namespace CatiaCsharp
{
    using static _Factory;
    using static _CATPart;
    using static _Selection;
    using static Utilities;
    class Program
    {
        [DllImport("NativeLib.dll")]
        public static extern void HelloWorld();

        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GetConsoleWindow();
        private static IntPtr ThisConsole = GetConsoleWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]

        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int HIDE = 0;
        private const int MAXIMIZE = 3;
        private const int MINIMIZE = 6;
        private const int RESTORE = 9;
        static void Main(string[] args)
        {
            ShowWindow(ThisConsole, MINIMIZE);
            InitGameSetUp();
            _Selection.PrintAllElements();
            try
            {
                Snake.CreateSnakes(6);
            }
            catch (Exception ex)
            {
                Utilities.PrintException(ex, ref Globals.errorCount);
                ShowWindow(ThisConsole, RESTORE);
            }
            _Selection.PrintAllElements();
            Console.WriteLine($"Finnished with {Globals.errorCount} errors!");
            Console.ReadLine();
        }
        public static void InitGameSetUp()
        {
            hybridBodyStream = hybridBodies.Add();
            hybridBodyStream.set_Name("hybridBodyStream");
            Hide(hybridBodyStream);
            HybridShapePointCoord HelperOriginPt = CreatePointCoord(new object[] { 0, 0, 50 }, hybridBodyStream);
            Body HelperCube = CreateCube(HelperOriginPt);
            AxisSystem HelperAxisSystem = CreateAxisSystem(HelperOriginPt, "HelperAxisSystem");
            
            HybridShapePointCoord OriginPoint = CreatePointCoord(XYZParse((0, 0)), hybridBodyStream);
            OriginPoint.set_Name("OriginPoint");       
            selection.Clear();
        }

        public static Body CreateCube(Point originPoint)
        {
            Body body = bodies.Add();
            _part.InWorkObject = body;
            var extrude1 = SimpleExtrude(GetCoordinates(originPoint), Globals.PieceLengthDouble, Globals.PieceLengthDouble, hybridBodyStream);
            var extrudeRef1 = GetRefFromObject(extrude1);
            ThickSurface thickSurface = shapeFactory.AddNewThickSurface(extrudeRef1, 0, 0.0, Globals.PieceLengthDouble);
            _part.Update();

            return body;
        }

    }
}