using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using PARTITF;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CatiaCsharp
{
    public static class _CATPart
    {
        public static void Init(bool reset = true)
        {
            Catia = (Application)Marshal.GetActiveObject("Catia.Application");
            documents = Catia.Documents;
            windows = Catia.Windows;

            Document document = (Document)Catia.ActiveDocument;
            var file = new FileInfo(document.FullName);
            string ext = file.Extension;
            if (ext == ".CATPart")
            {
                partDocument = (PartDocument)Catia.ActiveDocument;
                _part = partDocument.Part;
                if (reset)
                {
                    ResetPartDocument();
                }
            }
            else
            {
                partDocument = (PartDocument)documents.Add("Part");
                partDocument.Activate();
                _part = partDocument.Part;
            }
            hybridShapeFactory = (HybridShapeFactory)_part.HybridShapeFactory;
            hybridBodies = _part.HybridBodies;
            axisSystems = _part.AxisSystems;
            AbsoluteAxisSystem = _part.AxisSystems.Item(1);
            selection = partDocument.Selection;
            hybridBodyStream = hybridBodies.Add();
            hybridBodyStream.set_Name("hybridBodyStream");
        }
        public static Application Catia { get; private set; }
        public static Documents documents { get; private set; }
        public static Windows windows { get; private set; }
        public static PartDocument partDocument { get; private set; }
        public static Part _part;
        public static Selection selection { get; private set; }
        public static AxisSystems axisSystems { get; private set; }
        public static HybridShapeFactory hybridShapeFactory { get; private set; }
        public static ShapeFactory shapeFactory { get; private set; }
        public static HybridBodies hybridBodies { get; private set; }
        public static AxisSystem AbsoluteAxisSystem { get; set; }
        public static Window activeWindow { get; set; }
        public static HybridBody hybridBodyStream { get; set; }
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
        public static void ResetPartDocument()
        {
            Selection selection = partDocument.Selection;
            HybridBodies hybridBodies = _part.HybridBodies;
            Bodies bodies = _part.Bodies;
            int countHBs = hybridBodies.Count;
            while (countHBs > 0)
            {
                selection.Clear();
                selection.Add(hybridBodies.Item(1));
                selection.Delete();
                countHBs--;
            }
            Body newPartBody = bodies.Add();
            _part.MainBody = newPartBody;
            int countBs = bodies.Count;
            int iBs = 1;
            while (countBs > 1)
            {
                selection.Clear();
                Body body = bodies.Item(iBs);
                if (body.get_Name() != _part.MainBody.get_Name())
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
            int errorCount = 0;
            AxisSystems axisSystems = _part.AxisSystems;
            try
            {
                selection.Clear();
                selection.Add(axisSystems.Item(2));
                selection.Delete();
            }
            catch (Exception ex)
            {
                Utilities.PrintException(ex, ref errorCount);
            }
            _part.Update();
        }
        public static Reference GetRefFromObject(AnyObject anyObject)
        {
            return _part.CreateReferenceFromObject(anyObject);
        }
        public static HybridShapeDirection GetAxisDirection(Globals.Axis axisDir)
        {
            switch (axisDir)
            {
                case Globals.Axis.X:
                    return hybridShapeFactory.AddNewDirection(AbsoluteAxisSystem.XAxisDirection);
                case Globals.Axis.Y:
                    return hybridShapeFactory.AddNewDirection(AbsoluteAxisSystem.YAxisDirection);
                case Globals.Axis.Z:
                    return hybridShapeFactory.AddNewDirection(AbsoluteAxisSystem.ZAxisDirection);
                default:
                    return null;
            }
        }
    }
}
