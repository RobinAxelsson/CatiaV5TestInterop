using INFITF;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatiaCsharp
{
    public class VBAScript
    {
        public static void MessageBox(string message, Application catia)
        {
            string code = System.IO.File.ReadAllText(@"C:\Users\axels\source\repos\CatiaUnitTestProject\CatiaCsharp\VBA-script\messagebox.txt");
            code = code.Replace("{}", $"\"{message}\"");
            object[] parameter = new object[0];
            catia.SystemService.Evaluate(code, CATScriptLanguage.CATVBALanguage, "CATMain", parameter);
        }
        public static dynamic TestCATMain(string code, Application catia)
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
                    var output = catia.SystemService.Evaluate(code, CATScriptLanguage.CATVBALanguage, "CATMain", parameter);
                    return output;
                }
                catch (Exception)
                {
                    iPara++;
                }
            }
        }
    }
}
