using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualBasic;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;

namespace NET.FrameworkObjectTable
{
    class Program
    {
        struct RunningObject
        {
            public string name;
            public object o;
        }

        // Returns the contents of the Running Object Table (ROT), where
        // open Microsoft applications and their documents are registered.
        public static List<object> GetRunningObjects()
        {
            // Get the table.
            var res = new List<object>();
            IBindCtx bc;
            CreateBindCtx(0, out bc);
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable runningObjectTable;
            bc.GetRunningObjectTable(out runningObjectTable);
            IEnumMoniker monikerEnumerator;
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            // Enumerate and fill our nice dictionary.
            IMoniker[] monikers = new IMoniker[1];
            IntPtr numFetched = IntPtr.Zero;
            List<string> names = new List<string>();
            List<string> books = new List<string>();
            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                RunningObject running;
                monikers[0].GetDisplayName(bc, null, out running.name);
                runningObjectTable.GetObject(monikers[0], out running.o);
                res.Add(running);
            }
            return res;
        }

        private static void CreateBindCtx(int v, out IBindCtx bc)
        {
            throw new NotImplementedException();
        }

        static void Main(string[] args)
        {
            string cs = File.ReadAllText(@"C:\Users\axels\Google Drive\Code\VStudio_source\repos\CatiaUnitTestProject\Core_test\AccessingCOM.cs");
            cs = cs.Replace(Environment.NewLine, "");
            File.WriteAllText(@"C:\Users\axels\Google Drive\Code\VStudio_source\repos\CatiaUnitTestProject\Core_test\AccessingCOM.cs", cs);
            
            //var o = GetRunningObjects();
            //foreach (RunningObject running in GetRunningObjects())
            //{
            //    if (running.o is Excel.Workbook)
            //        MessageBox.Show("Found " + running.name);
            //}
            //var Catia = (Application)Marshal.GetActiveObject("Catia.Application");
            //Microsoft.Office.Interop.Excel.Application()
            //PartDocument pd = (PartDocument)Catia.ActiveDocument;
            //Part p = pd.Part;
            //p.set_Name("Test");
        }
    }
}