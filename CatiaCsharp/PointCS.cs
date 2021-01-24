//using HybridShapeTypeLib;
//using INFITF;
//using MECMOD;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace CatiaCsharp
//{
//    class PointCS
//    {
//        public Type CatType = typeof(Point);
//        public object[] Coordinates;
//        private Point Value;
//        public Part _part { get; set; }
//        public Selection Value { get; set; }
//        public string Name { get; set; }
//        public PointCS(Part part)
//        {
//            _part = part;
//            string message = "Select point";
//            object[] filter = new object[1];
//            filter[0] = "Point";
//            string status = Value.SelectElement2(filter, message, true); //status normal
//            if (status.ToLower() != "normal") throw new Exception();
//            Value = ((Point)Value.Item2(1).Value);
//            Name = Value.get_Name();
//            object[] coord = new object[3];
//            Value.GetCoordinates(coord);
//            Coordinates = coord;
//        }
//    }
//}
