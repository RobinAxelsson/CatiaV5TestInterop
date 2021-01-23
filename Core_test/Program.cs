using INFITF;
using MECMOD;
using PARTITF;
using ProductStructureTypeLib;
using SPATypeLib;
using NavigatorTypeLib;
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualBasic;
using HybridShapeTypeLib;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Reflection;
using System.Runtime.Versioning;
using Westwind.Utilities;

namespace Core_test
{
    class Program
    {
        public static void ComAccessReflectionCoreAnd45Test()
        {
            // this works with both .NET 4.5+ and .NET Core 2.0+

            string progId = "CATIA.Application";
            Type type = Type.GetTypeFromProgID(progId);
            object inst = Activator.CreateInstance(type);


            inst.GetType().InvokeMember("Visible", ReflectionUtils.MemberAccess | BindingFlags.SetProperty, null, inst,
                new object[1]
                {
            true
                });

            inst.GetType().InvokeMember("Navigate", ReflectionUtils.MemberAccess | BindingFlags.InvokeMethod, null,
                inst, new object[]
                {
            "https://markdownmonster.west-wind.com",
                });

            //result = ReflectionUtils.GetPropertyCom(inst, "cAppStartPath");
            bool result = (bool)inst.GetType().InvokeMember("Visible",
                ReflectionUtils.MemberAccess | BindingFlags.GetProperty, null, inst, null);
            Console.WriteLine(result); // path             
        }
        public static object ComAccess(string progId)
        {
            Type type = Type.GetTypeFromProgID(progId);
            object inst = Activator.CreateInstance(type);
            return inst;
        }
        public static string progID = "CATIA.Application";
        public static void Main()
        {
            //ComAccessReflectionCoreAnd45Test();
            //var arr = new string[] { "CATIA.Application"};
            //var objs = AccessingCOM.GetRunningInstances(arr);
            //ComObject obj = objs[0];
            var obj = ComAccess(progID);
            Application catia = (Application)obj;
            PartDocument pd = (PartDocument)catia.ActiveDocument;
            Part p = pd.Part;
            string test = p.get_Name();
            string test1 = p.Bodies.Item(1).get_Name();
            p.Bodies.Item(1).set_Name("test");
            p.set_Name("test");
        }
    ////https://adndevblog.typepad.com/autocad/2013/12/accessing-com-applications-from-the-running-object-table.html

    //[DllImport("ole32.dll")]
    //private static extern void CreateBindCtx(int reserved, out IBindCtx ppbc);

    //[DllImport("ole32.dll")]
    //private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

    //public static void RunningObjectsToText()
    //{
    //    IRunningObjectTable rot;
    //    IEnumMoniker enumMoniker;
    //    int retVal = GetRunningObjectTable(0, out rot);
    //    var info = new List<string>();

    //    if (retVal == 0)
    //    {
    //        rot.EnumRunning(out enumMoniker);

    //        IntPtr fetched = IntPtr.Zero;
    //        IMoniker[] moniker = new IMoniker[1];
    //        while (enumMoniker.Next(1, moniker, fetched) == 0)
    //        {
    //            IBindCtx bindCtx;
    //            CreateBindCtx(0, out bindCtx);
    //            string displayName;
    //            moniker[0].GetDisplayName(bindCtx, null, out displayName);
    //            Guid id;
    //            moniker[0].GetClassID(out id);
    //            Console.WriteLine("Display Name: {0}\nID: {1}", displayName, id);
    //            info.Add(id.ToString());
    //            info.Add(displayName);
    //            info.Add(Environment.NewLine);
    //        }
    //    }
    //    System.IO.File.AppendAllLines(@"C:\Users\axels\Google Drive\Code\VStudio_source\repos\CatiaUnitTestProject\Core_test\IDs.txt", info);
    //}

    ////GetActiveObject("00000303-0000-0000-c000-000000000046");
    //static void Main(string[] args)
    //{
    //    string progId = "CATIA.Application";
    //    var obj = Marshal2.GetActiveObject(progId);
    //    Application catia = (Application)obj;
    //    PartDocument pd = (PartDocument)catia.ActiveDocument;
    //    Part p = pd.Part;
    //    p.set_Name("test");
    //}

    //public static class Marshal2
    //{
    //    internal const String OLEAUT32 = "oleaut32.dll";
    //    internal const String OLE32 = "ole32.dll";

    //    [System.Security.SecurityCritical]  // auto-generated_required
    //    public static Object GetActiveObject(String progID)
    //    {
    //        Object obj = null;
    //        Guid clsid;

    //        // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
    //        // CLSIDFromProgIDEx doesn't exist.
    //        try
    //        {
    //            CLSIDFromProgIDEx(progID, out clsid);
    //        }
    //        //            catch
    //        catch (Exception)
    //        {
    //            CLSIDFromProgID(progID, out clsid);
    //        }

    //        GetActiveObject(ref clsid, IntPtr.Zero, out obj);
    //        return obj;
    //    }

    //    //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
    //    [DllImport(OLE32, PreserveSig = false)]
    //    [ResourceExposure(ResourceScope.None)]
    //    [SuppressUnmanagedCodeSecurity]
    //    [System.Security.SecurityCritical]  // auto-generated
    //    private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

    //    //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
    //    [DllImport(OLE32, PreserveSig = false)]
    //    [ResourceExposure(ResourceScope.None)]
    //    [SuppressUnmanagedCodeSecurity]
    //    [System.Security.SecurityCritical]  // auto-generated
    //    private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

    //    //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
    //    [DllImport(OLEAUT32, PreserveSig = false)]
    //    [ResourceExposure(ResourceScope.None)]
    //    [SuppressUnmanagedCodeSecurity]
    //    [System.Security.SecurityCritical]  // auto-generated
    //    private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);

    //}
    //    public static void ComAccessReflectionCoreAnd45Test()
    //    {
    //        // this works with both .NET 4.5+ and .NET Core 2.0+

    //        string progId = "CATIA.Application";
    //        Type type = Type.GetTypeFromProgID(progId);
    //        object inst = Activator.CreateInstance(type);


    //    inst.GetType().InvokeMember("Visible", ReflectionUtils.MemberAccess | BindingFlags.SetProperty, null, inst,
    //        new object[1]
    //        {
    //    true
    //        });

    //    inst.GetType().InvokeMember("Navigate", ReflectionUtils.MemberAccess | BindingFlags.InvokeMethod, null,
    //        inst, new object[]
    //        {
    //    "https://markdownmonster.west-wind.com",
    //        });

    //    //result = ReflectionUtils.GetPropertyCom(inst, "cAppStartPath");
    //    bool result = (bool)inst.GetType().InvokeMember("Visible",
    //        ReflectionUtils.MemberAccess | BindingFlags.GetProperty, null, inst, null);
    //    Console.WriteLine(result); // path             
    //}


    //Console.WriteLine("\nSample: C# System.Runtime.InteropServices.Marshal.GetActiveObject.cs\n");

    //GetObj(1, "Word.Application");
    //GetObj(2, "Excel.Application");            
    //static void GetObj(int i, String progID)
    //{
    //    Object obj = null;

    //    Console.WriteLine("\n" + i + ") Object obj = GetActiveObject(\"" + progID + "\")");
    //    try
    //    { obj = Marshal.GetActiveObject(progID); }
    //    catch (Exception e)
    //    {
    //        Write2Console("\n   Failure: obj did not get initialized\n" +
    //                      "   Exception = " + e.ToString().Substring(0, 43), 0);
    //    }

    //    if (obj != null)
    //    { Write2Console("\n   Success: obj = " + obj.ToString(), 1); }
    //}

    //static void Write2Console(String s, int color)
    //{
    //    Console.ForegroundColor = color == 1 ? ConsoleColor.Green : ConsoleColor.Red;
    //    Console.WriteLine(s);
    //    Console.ForegroundColor = ConsoleColor.Gray;
    //}


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