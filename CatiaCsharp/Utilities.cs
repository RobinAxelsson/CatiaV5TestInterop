using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using PARTITF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatiaCsharp
{
    public static class Utilities
    {

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
        public static object[] GetCoordinates(Point point)
        {
            object[] coords = new object[3];
            point.GetCoordinates(coords);
            return coords;
        }
        public static void PrintException(Exception ex, ref int errorCount)
        {
            errorCount++;
            Console.WriteLine("Message: " + ex.Message);
            Console.WriteLine("StackTrace: " + ex.StackTrace);
            Console.WriteLine("TargetSite: " + ex.TargetSite);
            Console.WriteLine("Inner Exception: " + ex.InnerException);
        }
        public static (int X, int Y) XYParse(object[] CoordXYZ)
        {
            double X = (double)CoordXYZ[0] / Globals.PieceLengthDouble;
            double Y = (double)CoordXYZ[1] / Globals.PieceLengthDouble;
            return ((int)X, (int)Y);
        }
        public static object[] XYZParse((int X, int Y) points, double Z = 0.0)
        {
            return new object[] { points.X * Globals.PieceLengthDouble, points.Y * Globals.PieceLengthDouble, Z };
        }
        public static Reference GetRefFromObject(AnyObject anyObject)
        {
            return _CATPart._part.CreateReferenceFromObject(anyObject);
        }
        public static HybridShapeDirection GetAxisDirection(Globals.Axis axisDir)
        {
            switch (axisDir)
            {
                case Globals.Axis.X:
                    return _CATPart.hybridShapeFactory.AddNewDirection(_CATPart.AbsoluteAxisSystem.XAxisDirection);
                case Globals.Axis.Y:
                    return _CATPart.hybridShapeFactory.AddNewDirection(_CATPart.AbsoluteAxisSystem.YAxisDirection);
                case Globals.Axis.Z:
                    return _CATPart.hybridShapeFactory.AddNewDirection(_CATPart.AbsoluteAxisSystem.ZAxisDirection);
                default:
                    return null;
            }
        }
        public static HybridBody ClearHybridBody(HybridBody hybridBody)
        {
            _CATPart.selection.Clear();
            string name = hybridBody.get_Name();
            _CATPart.selection.Add(hybridBody);
            _CATPart.selection.Delete();
            HybridBody newHybridBody = _CATPart.hybridBodies.Add();
            newHybridBody.set_Name(name);
            return newHybridBody;
        }
        private static void newProduct()
        {
            _CATPart.partDocument.Product.Products.AddNewProduct("newProduct");
        }

    }
}
