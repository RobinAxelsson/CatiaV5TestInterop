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
        
        
       

        static void Main(string[] args)
        {
            ShowWindow(ThisConsole, MINIMIZE);
            _CATPart.Init();
            var _factory = new _Factory();
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
                _part.Update();
                SnakeBody.set_Name($"Player.{Player}");
                SpawnPoint = CreatePointCoord(XYZParse(spawnPoint), hybridBodyStream);
                _part.Update();
                selection.Clear();
                selection.Add(firstPiece);
                selection.Copy();
                selection.PasteSpecial("CATPrtResult");
                _part.Update();

                Body pieceCopy = (Body)_part.InWorkObject;
                _part.InWorkObject = pieceCopy.Shapes.Item(1);
                Translate translate1 = (Translate)shapeFactory.AddNewTranslate2(0.0);
                HybridShapeTranslate hybridTranslate1 = (HybridShapeTranslate)translate1.HybridShape;
                hybridTranslate1.VectorType = 1;
                hybridTranslate1.FirstPoint = GetRefFromObject(OriginPoint);
                hybridTranslate1.SecondPoint = GetRefFromObject(SpawnPoint);
                _part.InWorkObject = hybridTranslate1;
                _part.Update();                
                
                _part.InWorkObject = SnakeBody;
                shapeFactory.AddNewAdd(pieceCopy);
                _part.Update();
                _part.InWorkObject = pieceCopy;

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

        



        public static HybridBody ClearHybridBody1()
        {
            selection.Clear();
            selection.Add(hybridBodyStream);
            selection.Delete();
            hybridBodyStream = hybridBodies.Add();
            return hybridBodyStream;
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

        public static Body CreateCube(Point originPoint)
        {
            Body body = bodies.Add();
            _part.InWorkObject = body;
            var extrude1 = SimpleExtrude(GetCoordinates(originPoint), PieceLengthDouble, PieceLengthDouble, hybridBodyStream);
            var extrudeRef1 = GetRefFromObject(extrude1);
            ThickSurface thickSurface = shapeFactory.AddNewThickSurface(extrudeRef1, 0, 0.0, PieceLengthDouble);
            _part.Update();

            return body;
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