using INFITF;
using MECMOD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatiaCsharp
{
    public static class _Selection
    {
        static _Selection()
        {
            selection = _CATPart.partDocument.Selection;
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
                catch (Exception)
                {
                }
            }
            selection.Clear();
            return selElements;
        }
        public static void RenameSelection(int index, string name)
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
                selection.Add(hybridBody.HybridBodies.Item(1));
                selection.Delete();
                countHB--;
            }
            if (exception != null) selection.Paste();
            _CATPart._part.Update();
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
        //        explicitObj = (AnyObject)selection.Item(1).selection;
        //        selection.Clear();
        //    }
        //    else //Shape assumed
        //    {
        //        selection.PasteSpecial("CATPrtResultWithOutLink");
        //        explicitObj = (AnyObject)selection.Item(1).selection;
        //        selection.Clear();
        //    }
        //    _part.Update();
        //    return explicitObj;
        //}
        // Dim params()
        //CATIA.SystemService.ExecuteScript"_part.CATPart", catScriptLibraryTypeDocument, "Macro1.catvbs", "CATMain", params
    }
}
