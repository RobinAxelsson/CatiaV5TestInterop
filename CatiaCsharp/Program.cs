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

        //Game-constants//
        public const int PieceLengthInt = 50; //mm
        public static readonly double PieceLengthDouble = (double)PieceLengthInt;
        public const int LengthXPieces = 15;
        public const int LengthYPieces = 10;
        public readonly double BoardXLength = LengthXPieces * PieceLengthDouble;
        public readonly double BoardYWidth = LengthYPieces * PieceLengthDouble;
        public enum GlobalAxis
        {
            X,
            Y,
            Z
        }
        public static int Errors = 0;
        public static Application Catia;
        public static Documents documents;
        public static Windows windows;
        public static Window activeWindow { get; set; }
        //public static ProductDocument productDocument;
        public static PartDocument partDocument;
        public static Part part1;
        public static Selection selection;
        public static HybridShapeFactory hybridShapeFactory;
        public static ShapeFactory shapeFactory;
        public static HybridBodies hybridBodies;
        public static HybridBody hybridBodyStream;
        public static HybridBody hybridBodyResults;
        public static Bodies bodies;
        public static List<Snake> Snakes = new List<Snake>();
        public static HybridShapePointCoord OriginPoint;
        public static AxisSystems axisSystems;
        public static AxisSystem AbsoluteAxisSystem;
        public static AxisSystem HelperAxisSystem;
        public static Body HelperCube;
        public static Point HelperOriginPt;

        static void Main(string[] args)
        {
            ShowWindow(ThisConsole, MINIMIZE);
            InitCatia();
            //VBAScript.MessageBox("Hello World!");
            InitCATpart();
            part1.set_Name("test");
            InitiateCollectionsFactories();
            ResetPartDocument();
            InitGameSetUp();
            PrintAllElements();
            try
            {
                Snake.CreateSnakes(6);
            }
            catch (Exception ex)
            {
                PrintException(ex);
                ShowWindow(ThisConsole, RESTORE);
            }
            PrintAllElements();
            Console.WriteLine($"Finnished with {Errors} errors!");
            Console.ReadLine();
        }
        public static void CreateCube(Body body, object[] originCoord, double length)
        {
            selection.Clear();
            part1.InWorkObject = body;
            var extrude1 = SimpleExtrude(originCoord, length, length, hybridBodyStream);
            var extrudeRef1 = GetRefFromObject(extrude1);
            ThickSurface thickSurface = shapeFactory.AddNewThickSurface(extrudeRef1, 0, 0.0, length);
            part1.Update();
        }
        public static object[] XYZParse((int X, int Y) points, double Z = 0.0)
        {
            return new object[] { points.X * PieceLengthDouble, points.Y * PieceLengthDouble, Z };
        }
        public static (int X, int Y) XYParse(object[] CoordXYZ)
        {
            double X = (double)CoordXYZ[0] / PieceLengthDouble;
            double Y = (double)CoordXYZ[1] / PieceLengthDouble;
            return ((int)X, (int)Y);
        }
        public class Snake
        {
            public int Player { get; }
            public Body pieceBodyLink { get; }
            public Body SnakeBody { get; }
            public HybridShapePointCoord SpawnPoint{ get; }
            public List<(int x, int y)> bodyCoord { get; set; }
            private Snake((int x, int y) spawnPoint, Body firstPiece)
            {
                Snakes.Add(this);
                Player = Snakes.FindIndex(x => x == this) + 1;
                SnakeBody = bodies.Add();
                part1.Update();
                SnakeBody.set_Name($"Player.{Player}");
                SpawnPoint = CreatePointCoord(XYZParse(spawnPoint), hybridBodyStream);
                part1.Update();
                selection.Clear();
                selection.Add(firstPiece);
                selection.Copy();
                selection.PasteSpecial("CATPrtResult");
                part1.Update();

                Body pieceCopy = (Body)part1.InWorkObject;
                part1.InWorkObject = pieceCopy.Shapes.Item(1);
                Translate translate1 = (Translate)shapeFactory.AddNewTranslate2(0.0);
                HybridShapeTranslate hybridTranslate1 = (HybridShapeTranslate)translate1.HybridShape;
                hybridTranslate1.VectorType = 1;
                hybridTranslate1.FirstPoint = GetRefFromObject(OriginPoint);
                hybridTranslate1.SecondPoint = GetRefFromObject(SpawnPoint);
                part1.InWorkObject = hybridTranslate1;
                part1.Update();                
                
                part1.InWorkObject = SnakeBody;
                shapeFactory.AddNewAdd(pieceCopy);
                part1.Update();
                part1.InWorkObject = pieceCopy;

                selection.Clear();
                
            }
            public static void CreateSnakes(int players)
            { 
                Random rand = new Random();
                for (int i = 0; i < players; i++)
                {
                    int X = rand.Next(0, LengthXPieces);
                    int Y = rand.Next(0, LengthYPieces);
                    Snake s = new Snake((X, Y), HelperCube);
                }
            }
        }
        public static void PrintException(Exception ex)
        {
            Errors++;
            Console.WriteLine("Message: " + ex.Message);
            Console.WriteLine("StackTrace: " + ex.StackTrace);
            Console.WriteLine("TargetSite: " + ex.TargetSite);
            Console.WriteLine("Inner Exception: " + ex.InnerException);
        }
        public enum CATPasteType
        {
            CATPrtCont, //"As Specified In Part Document"
            CATPrtResultWithOutLink, //As Result
            CATPrtResult //As result with link
        }
        //public static AnyObject PasteSpecial(AnyObject obj)
        //{
        //    selection.Clear();
        //    selection.Add(obj);
        //    SelectedElement selEl = selection.Item(1);
        //    string type = selEl.Type;
        //    selection.Copy();
        //    AnyObject explicitObj;
        //    if (type.Contains("HybridShape"))
        //    {
        //        if (hybridHome != null)
        //        {
        //            selection.Clear();
        //            selection.Add(hybridHome);
        //        }
        //        else
        //        {
        //            selection.Clear();
        //            selection.Add(hybridBodyResults);
        //        }
        //        selection.PasteSpecial("CATPrtResultWithOutLink");
        //        explicitObj = (AnyObject)selection.Item(1).Value;
        //        selection.Clear();
        //    }
        //    else //Shape assumed
        //    {
        //        selection.PasteSpecial("CATPrtResultWithOutLink");
        //        explicitObj = (AnyObject)selection.Item(1).Value;
        //        selection.Clear();
        //    }
        //    part1.Update();
        //    return explicitObj;
        //}
        // Dim params()
        //CATIA.SystemService.ExecuteScript"Part1.CATPart", catScriptLibraryTypeDocument, "Macro1.catvbs", "CATMain", params
 
        public class VBAScript
        {
            public static void MessageBox(string message)
            {
                string code = System.IO.File.ReadAllText(@"C:\Users\axels\source\repos\CatiaUnitTestProject\CatiaCsharp\VBA-script\messagebox.txt");
                code = code.Replace("{}", $"\"{message}\"");
                object[] parameter = new object[0];
                Catia.SystemService.Evaluate(code, CATScriptLanguage.CATVBALanguage, "CATMain", parameter);
            }
            public static dynamic TestCATMain(string code)
            {
                if (code == null || code == "")
                {
                    code = System.IO.File.ReadAllText(@"C:\Users\axels\source\repos\CatiaUnitTestProject\CatiaCsharp\VBA-script\testcode.txt");
                }
                int iPara = 0;
                while (true)
                {
                    try
                    {
                        Console.WriteLine(iPara);
                        object[] parameter = new object[iPara];
                        var output = Catia.SystemService.Evaluate(code, CATScriptLanguage.CATVBALanguage, "CATMain", parameter);
                        return output;
                    }
                    catch (Exception)
                    {
                        iPara++;
                    }
                }
            }
        }
        public static HybridBody ClearHybridBody1()
        {
            selection.Clear();
            selection.Add(hybridBodyStream);
            selection.Delete();
            hybridBodyStream = hybridBodies.Add();
            return hybridBodyStream;
        }
        public static void CleanHybridBody(AnyObject[] exception, HybridBody hybridBody)
        {            
            selection.Clear();

            if (exception != null)
            {
                foreach (var obj in exception)
                {
                    selection.Add(obj);
                }
                selection.Copy();
            }

            int countHB = hybridBody.HybridBodies.Count;
            while (countHB > 0)
            {
                selection.Clear();
                selection.Add(hybridBodyStream.HybridBodies.Item(1));
                selection.Delete();
                countHB--;
            }
            if (exception != null) selection.Paste();
            part1.Update();
        }

        public static void CastDoubleArray(ref object[] coordinates)
        { 
            if (coordinates.Length != 3)
            {
                throw new Exception("Coordinates length not valid!");
            }
            for (int i = 0; i < 3; i++)
            {
                if (coordinates[i].GetType() == typeof(int))
                {
                    coordinates[i] = (int)coordinates[i] * 1.0;
                }
                else
                {
                    if (coordinates[i].GetType() != typeof(double))
                    {
                        throw new Exception("Coordinates type is not valid!");
                    }
                }
            }
        }
        public static HybridShapePointCoord CreatePointCoord(object[] Coord, HybridBody hybridBody, AxisSystem axisSystem = null)
        {
            CastDoubleArray(ref Coord);
            HybridShapePointCoord point1 = hybridShapeFactory.AddNewPointCoord((double)Coord[0], (double)Coord[1], (double)Coord[2]);
            if (axisSystem != null)
            {
                point1.RefAxisSystem = GetRefFromObject(axisSystem);
            }
            hybridBody.AppendHybridShape(point1);
            part1.Update();
            return point1;
        }
        public static HybridShapeLinePtPt CreatePtPtLine(object[] Coord1, object[] Coord2, HybridBody hybridBody)
        {
            var point1 = CreatePointCoord(Coord1, hybridBody);
            var point2 = CreatePointCoord(Coord2, hybridBody);
            Reference ref1 = GetRefFromObject(point1);
            Reference ref2 = GetRefFromObject(point2);
            HybridShapeLinePtPt line1 = hybridShapeFactory.AddNewLinePtPt(ref1, ref2);
            hybridBody.AppendHybridShape(line1);
            part1.Update();
            return line1;
        }
        public static HybridShapeDirection GetAxisDirection(GlobalAxis axisDir)
        {
            switch (axisDir)
            {
                case Program.GlobalAxis.X:
                    return hybridShapeFactory.AddNewDirection(AbsoluteAxisSystem.XAxisDirection);
                case Program.GlobalAxis.Y:
                    return hybridShapeFactory.AddNewDirection(AbsoluteAxisSystem.YAxisDirection);
                case Program.GlobalAxis.Z:
                    return hybridShapeFactory.AddNewDirection(AbsoluteAxisSystem.ZAxisDirection);
                default:
                    return null;
            }
        }
        public static Reference GetRefFromObject(AnyObject anyObject)
        {
            return part1.CreateReferenceFromObject(anyObject);
        }
        public static HybridShapeLinePtDir CreateLinePtDir(object[] coord, GlobalAxis axisDir, double length, HybridBody hybridBody)
        {
            HybridBody trueHybridBody = hybridBody;
            var point = CreatePointCoord(coord, hybridBody);
            var pointRef = GetRefFromObject(point);
            var dir = GetAxisDirection(axisDir);
            HybridShapeLinePtDir line = hybridShapeFactory.AddNewLinePtDir(pointRef, dir, 0.0, length, false);
            hybridBody.AppendHybridShape(line);
            part1.Update();
            return line;
        }
        public static HybridShapeExtrude SimpleExtrude(object[]OriginPt, double y, double z, HybridBody hybridBody)
        {
            HybridShapeLinePtDir line1 = CreateLinePtDir(OriginPt, GlobalAxis.Z, z, hybridBodyStream);
            Reference lineRef = GetRefFromObject(line1);
            var dir1 = GetAxisDirection(GlobalAxis.Y);
            HybridShapeExtrude extrude1 = hybridShapeFactory.AddNewExtrude(lineRef, y, 0.0, dir1);
            extrude1.SymmetricalExtension = false;
            hybridBody.AppendHybridShape(extrude1);
            part1.Update();
            return extrude1;
        }

        public static void InitCatia()
        {
            try
            {
                Catia = (Application)Marshal.GetActiveObject("Catia.Application");
            }
            catch (Exception)
            {
                Console.WriteLine("Start Catia Manually, then press enter");
                Console.ReadKey();
                //var info = new ProcessStartInfo("cmd.exe")
                //{
                //    Arguments = @"C:\Program Files\Dassault Systemes\B20\win_b64\code\bin\CATSTART.exe"" - run ""CNEXT.exe"" - env CATIA.V5R20.B20 - direnv ""C:\ProgramData\DassaultSystemes\CATEnv"" - nowindow"
                //};
                //Process.Start(info);
                Catia = (Application)Marshal.GetActiveObject("Catia.Application");
            }
            documents = Catia.Documents;
            windows = Catia.Windows;
        }
        public static void InitCATpart()
        {
            Document doc;
            try
            {
                doc = Catia.ActiveDocument;
                var file = new FileInfo(doc.FullName);
                string ext = file.Extension;
                if (ext == ".CATPart")
                {
                    partDocument = (PartDocument)Catia.ActiveDocument;
                }
                else if (ext == ".CATProduct")
                {
                    ProductDocument prodDoc = (ProductDocument)doc;
                    Selection prodSel = prodDoc.Selection;
                    prodSel.Clear();

                    string searchstring = "Type = Part, all";
                    prodSel.Search(ref searchstring);
                    List<Part> parts = new List<Part>();

                    for (int i = 1; i <= prodSel.Count; i++)
                    {
                        try
                        {
                            Part part1 = (Part)prodSel.Item(i).Value;
                            parts.Add(part1);
                            Console.WriteLine(part1.get_Name());
                        }
                        catch (Exception)
                        {
                        }
                    }
                    prodSel.Clear();
                    prodSel.Add(parts[0]);
                    Catia.StartCommand("Open in New Window");
                }
            }
            catch (Exception)
            {
                documents = Catia.Documents;
                partDocument = (PartDocument)documents.Add("Part");
            }         
            
            part1 = partDocument.Part;
            selection = partDocument.Selection;
            activeWindow = Catia.ActiveWindow;
            PrintAppInfo();
        }
        public static void ResetPartDocument()
        {
            int countHBs = hybridBodies.Count;
            while (countHBs > 0)
            {
                selection.Clear();
                selection.Add(hybridBodies.Item(1));
                selection.Delete();
                countHBs--;
            }
            Body newPartBody = bodies.Add();
            part1.MainBody = newPartBody;
            int countBs = bodies.Count;
            int iBs = 1;
            while (countBs > 1)
            {
                selection.Clear();
                Body body = bodies.Item(iBs);
                if (body.get_Name() != part1.MainBody.get_Name())
                {
                    selection.Add(body);
                    selection.Delete();
                    countBs--;
                }
                else
                {
                    iBs++;
                }
            }          

            newPartBody.set_Name("PartBody");
            try
            {
                selection.Clear();
                selection.Add(axisSystems.Item(2));
                selection.Delete();
            }
            catch (Exception ex)
            {
                PrintException(ex);
            }
            part1.Update();
        }
        public static void InitiateCollectionsFactories()
        {
            hybridShapeFactory = (HybridShapeFactory)part1.HybridShapeFactory;
            hybridBodies = part1.HybridBodies;
            bodies = part1.Bodies;
            shapeFactory = (ShapeFactory)part1.ShapeFactory;
            axisSystems = part1.AxisSystems;
            try
            {
                AbsoluteAxisSystem = axisSystems.Item("Absolute Axis System");
            }
            catch (Exception)
            {
                AbsoluteAxisSystem = axisSystems.Item(1);
            }
        }
        public static AxisSystem CreateAxisSystem(Point originPoint, string name)
        {
            var coordRef = new object[3];
            originPoint.GetCoordinates(coordRef);
            var pointXdir = CreatePointCoord(new object[] { (double)coordRef[0] + 1.0, coordRef[1], (double)coordRef[2]}, hybridBodyStream);
            var pointYdir = CreatePointCoord(new object[] { (double)coordRef[0], (double)coordRef[1] + 1.0, (double)coordRef[2] }, hybridBodyStream);
            var pointZdir = CreatePointCoord(new object[] { coordRef[0], coordRef[1], (double)coordRef[2] + 1.0 }, hybridBodyStream);
            AxisSystem axisSystem = axisSystems.Add();
            axisSystem.set_Name(name);
            axisSystem.Type = CATAxisSystemMainType.catAxisSystemStandard;
            axisSystem.OriginType = CATAxisSystemOriginType.catAxisSystemOriginByPoint;
            axisSystem.OriginPoint = GetRefFromObject(originPoint);
            axisSystem.XAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates;
            axisSystem.XAxisDirection = GetRefFromObject(pointXdir);
            axisSystem.YAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates;
            axisSystem.YAxisDirection = GetRefFromObject(pointYdir);
            axisSystem.ZAxisType = CATAxisSystemAxisType.catAxisSystemAxisByCoordinates;
            axisSystem.ZAxisDirection = GetRefFromObject(pointZdir);
            part1.Update();

            return axisSystem;

        }
        public static void InitGameSetUp()
        {
            hybridBodyStream = hybridBodies.Add();
            hybridBodyStream.set_Name("hybridBodyStream");
            Hide(hybridBodyStream);
            HelperOriginPt = CreatePointCoord(new object[] { 0, 0, 50 }, hybridBodyStream);
            HelperCube = CreateCube(HelperOriginPt);
            HelperAxisSystem = CreateAxisSystem(HelperOriginPt, "HelperAxisSystem");
            
            OriginPoint = CreatePointCoord(XYZParse((0, 0)), hybridBodyStream);
            OriginPoint.set_Name("OriginPoint");       
            selection.Clear();
        }
        public static object[] GetCoordinates(Point point)
        {
            object[] coords = new object[3];
            point.GetCoordinates(coords);
            return coords;
        }
        public static Body CreateCube(Point originPoint)
        {
            Body body = bodies.Add();
            part1.InWorkObject = body;
            var extrude1 = SimpleExtrude(GetCoordinates(originPoint), PieceLengthDouble, PieceLengthDouble, hybridBodyStream);
            var extrudeRef1 = GetRefFromObject(extrude1);
            ThickSurface thickSurface = shapeFactory.AddNewThickSurface(extrudeRef1, 0, 0.0, PieceLengthDouble);
            part1.Update();

            return body;
        }
        public static void NewPartDocument()
        {
            PartDocument partDocument = (PartDocument)Catia.Documents.Add("Part");
            partDocument.Activate();
            part1 = partDocument.Part;
            selection = partDocument.Selection;
        }
        public static void PrintAppInfo()
        {
            Window activeWindow = Catia.ActiveWindow;
            Console.WriteLine($"Catia Active Window: {activeWindow.get_Name()}");
            Console.WriteLine($"Catia Active Document: {Catia.ActiveDocument.get_Name()}");

            Console.WriteLine($"Open Windows({windows.Count}):");
            for (int i = 1; i <= windows.Count; i++)
            {
                Console.WriteLine(windows.Item(i).get_Name());
            }
            Console.WriteLine();

            Console.WriteLine($"Views({activeWindow.Viewers.Count})");
            Viewers viewers = activeWindow.Viewers;

            for (int i = 1; i <= activeWindow.Viewers.Count; i++)
            {
                Console.WriteLine(viewers.Item(i).get_Name());
            }
            Console.WriteLine();

            Console.WriteLine($"Documents in use({documents.Count}):");
            for (int i = 1; i <= documents.Count; i++)
            {
                Console.WriteLine(((Document)documents.Item(i)).get_Name());
            }
            Console.WriteLine();
        }
        public static void Hide(params AnyObject[] objs)
        {
            selection.Clear();
            foreach (var obj in objs)
            {
                selection.Add(obj);
            }
            VisPropertySet objsVisProps = selection.VisProperties;
            objsVisProps.SetShow(CatVisPropertyShow.catVisPropertyNoShowAttr);
        }
        public static void PrintAllElements()
        {
            var selElements = SelectElementsByName();
            foreach (var item in selElements)
            {
                AnyObject obj = (AnyObject)item.Value;
                AnyObject parent = (AnyObject)obj.Parent;
                Console.WriteLine($"{obj.get_Name()} (type={item.Type}) (Parent={parent.get_Name()})");
                //string test = ((Collection)item.Parent).
            }
            Console.WriteLine();
        }
        public static List<SelectedElement> SelectElementsByName(string name = "")
        {
            var selElements = new List<SelectedElement>();
            selection.Clear();
            string searchString = $"Name={name}*, all";
            selection.Search(ref searchString);

            for (int i = 1; i <= selection.Count; i++)
            {
                try
                {
                    SelectedElement selElement = selection.Item(i);
                    selElements.Add(selElement);
                }
                catch (Exception ex)
                {
                    PrintException(ex);
                }
            }
            selection.Clear();
            return selElements;
        }
        public class PointCS
        {
            public Type CatType = typeof(Point);
            public object[] Coordinates;
            private Point Value;
            public string Name { get; set; }
            public PointCS()
            {
                string message = "Select point";
                object[] filter = new object[1];
                filter[0] = "Point";
                string status = selection.SelectElement2(filter, message, true); //status normal
                if (status.ToLower() != "normal") throw new Exception();
                Value = ((Point)selection.Item2(1).Value);
                Name = Value.get_Name();
                object[] coord = new object[3];
                Value.GetCoordinates(coord);
                Coordinates = coord;
            }
        }
        private static void RenameSelection(int index, string name)
        {
            int max = selection.Count;
            if (max == 0)
            {
                throw new Exception("Nothing in selection");
            }
            if (!(index <= max && index >= 1))
            {
                Console.WriteLine($"Wrong index of selection, type an new index between {1} and {max}");
                index = int.Parse(Console.ReadLine());
            }
            ((AnyObject)selection.Item(index).Value).set_Name(name);
        }
        private static void newProduct()
        {
            partDocument.Product.Products.AddNewProduct("newProduct");
        }
        //object[] testArr = new object[]
        //{
        //    1.000, 0, 0, 0,0.707,0.707,0,-0.707,0.707,10.000,20.000,30.000                
        //};
        //Move
        //Sometimes it is necessary to move objects around in three-dimensional space via a macro
        //program.There are two main ways to manipulate the position of a “component” (in this case
        //the definition of component is either a child part, child product, or child component):

        // Dim TransformationArray( 11 )
        //'Rotation( 45 degrees around the x axis) components
        //MyMovableObject.Move.Apply TransformationArray
    }
}