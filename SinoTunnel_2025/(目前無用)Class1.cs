using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;

namespace SinoTunnel_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute (ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //place segment

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Application app = doc.Application;
            UIApplication uiapp = new UIApplication(app);
            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(@"C:\Users\exist\OneDrive\桌面\0409\backup\環形樣板.rfa");
            Document edit_doc = edit_uidoc.Document;
            


            List<string> name_list = new List<string> { "A1", "A2", "A3", "B1", "B2", "K" };

            //int segment_type = 1;

            foreach (string name in name_list)
            {
                place_segment(edit_doc, name, 0);
            }
            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;
            edit_doc.SaveAs(@"C:\Users\exist\OneDrive\桌面\0409\環形\正常環形.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc.PathName);
            edit_doc.Close();

            edit_uidoc = uiapp.OpenAndActivateDocument(@"C:\Users\exist\OneDrive\桌面\0409\backup\環形樣板.rfa");
            edit_doc = edit_uidoc.Document;

            foreach (string name in name_list)
            {
                place_segment(edit_doc, name+"b", 0);
            }
            
            edit_doc.SaveAs(@"C:\Users\exist\OneDrive\桌面\0409\環形\正常環形b.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc.PathName);
            edit_doc.Close();

            edit_uidoc = uiapp.OpenAndActivateDocument(@"C:\Users\exist\OneDrive\桌面\0409\backup\環形樣板.rfa");
            edit_doc = edit_uidoc.Document;

            foreach (string name in name_list)
            {
                place_segment(edit_doc, name, 1);   
            }
            edit_doc.SaveAs(@"C:\Users\exist\OneDrive\桌面\0409\環形\異形環形.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc.PathName);
            edit_doc.Close();
            

            return Result.Succeeded;
        }



        public void place_segment(Document edit_doc, string name, int segment_type)
        {
            //segment type normal:0 abnormal:1

            Transaction t1 = new Transaction(edit_doc, "place");
            t1.Start();
            if(segment_type == 0)
            {
                edit_doc.LoadFamily(@"C:\Users\exist\OneDrive\桌面\0409\正常環片\" + name + ".rfa");
            }
            else if(segment_type == 1)
            {
                edit_doc.LoadFamily(@"C:\Users\exist\OneDrive\桌面\0409\異形環片\" + name + ".rfa");
            }

            
            FamilySymbol fam_sym = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name).First();
            fam_sym.Activate();
            FamilyInstance FI = edit_doc.FamilyCreate.NewFamilyInstance(new XYZ(0, 0, 0), fam_sym, StructuralType.NonStructural);

            t1.Commit();
        }
        
    }
}
