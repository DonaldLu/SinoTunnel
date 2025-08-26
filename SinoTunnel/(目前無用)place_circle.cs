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

namespace SinoTunnel
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class place_circle : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            IList<XYZ> pointList = new List<XYZ>();
            pointList.Add(new XYZ(0, 0, 0));
            pointList.Add(new XYZ(1 / 0.3048, 0, 0));
            pointList.Add(new XYZ(2 / 0.3048, 0, 0));
            pointList.Add(new XYZ(3 / 0.3048, 0, 0));
            pointList.Add(new XYZ(4 / 0.3048, 0, 0));
            pointList.Add(new XYZ(5 / 0.3048, 0, 0));
            pointList.Add(new XYZ(6 / 0.3048, 0, 0));
            pointList.Add(new XYZ(7 / 0.3048, 0, 0));
            pointList.Add(new XYZ(8 / 0.3048, 0, 0));
            pointList.Add(new XYZ(9 / 0.3048, 0, 0));
            pointList.Add(new XYZ(10 / 0.3048, 0, 0));

            Family family;
            string name = "正常環形";
            Transaction trans = new Transaction(doc, "load segments");
            trans.Start();

            doc.LoadFamily(@"C:\Users\exist\OneDrive\桌面\0409\環形\" + name +".rfa", out family);
            doc.LoadFamily(@"C:\Users\exist\OneDrive\桌面\0409\環形\" + name+"b" + ".rfa", out family);
            FamilySymbol fam_sym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name).First();
            FamilySymbol fam_sym_b = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name+"b").First();
            fam_sym.Activate();
            fam_sym_b.Activate();
            for (int i = 0; i < pointList.Count; i++)
            {
                if(i%2 == 0)
                {
                    FamilyInstance FI = doc.Create.NewFamilyInstance(pointList[i], fam_sym, StructuralType.NonStructural);
                }
                else
                {
                    FamilyInstance FI = doc.Create.NewFamilyInstance(pointList[i], fam_sym_b, StructuralType.NonStructural);
                }
                
            }

            trans.Commit();
            return Result.Succeeded;
        }
    }
}
