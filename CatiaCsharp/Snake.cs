using HybridShapeTypeLib;
using PARTITF;
using INFITF;
using MECMOD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatiaCsharp
{
    public class Snake
    {
        public static List<Snake> Snakes { get; set; }
        public static Body HelperCube { get; set; }
        public int Player { get; }
        public Body pieceBodyLink { get; }
        public Body SnakeBody { get; }
        public HybridShapePointCoord SpawnPoint { get; }
        public List<(int x, int y)> bodyCoord { get; set; }
        private Snake((int x, int y) spawnPoint, Body firstPiece)
        {
            Selection selection = _CATPart.selection;
            Part _part = _CATPart._part;
            HybridBody hybridBodyStream = _CATPart.hybridBodyStream;
            ShapeFactory shapeFactory = _CATPart.shapeFactory;

            Snakes.Add(this);
            Player = Snakes.FindIndex(x => x == this) + 1;
            SnakeBody = _CATPart.bodies.Add();
            _CATPart._part.Update();
            SnakeBody.set_Name($"Player.{Player}");
            SpawnPoint = _Factory.CreatePointCoord(Utilities.XYZParse(spawnPoint), hybridBodyStream);
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
            hybridTranslate1.FirstPoint = Utilities.GetRefFromObject();
            hybridTranslate1.SecondPoint = Utilities.GetRefFromObject(SpawnPoint);
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
            Body HelperCube = _CATPart.bodies.Add();
            _Factory.CreateCube(HelperCube, new object[] {100, 100, 200}, Globals.PieceLengthDouble);
            Random rand = new Random();
            for (int i = 0; i < players; i++)
            {
                int X = rand.Next(0, Globals.LengthXPieces);
                int Y = rand.Next(0, Globals.LengthYPieces);
                Snake s = new Snake((X, Y), HelperCube);
            }
        }
    }
}
