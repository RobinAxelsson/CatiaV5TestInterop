using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Runtime.InteropServices;

namespace TestProjectCatia
{
    [TestClass]
    public static class UnitTest1
    {
        public static INFITF.Application CATIA = (INFITF.Application) Marshal.GetActiveObject("Catia.Application");
        public static PartDocument partDocument = (PartDocument)CATIA.ActiveDocument;
        public static Selection selection = partDocument.Selection;
        public static Part part = partDocument.Part;
        public static HybridShapeFactory hybridShapeFactory = (HybridShapeFactory)part.HybridShapeFactory;
        public static HybridBodies hybridBodies = part.HybridBodies;

        [TestMethod]
        public static void WeldCylinderSweep()
        {
            
        }
    }
}
