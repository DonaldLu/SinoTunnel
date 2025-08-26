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
using DataObject;
using Autodesk.Revit.UI.Selection;
//再用的走到
namespace SinoTunnel_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class aisle : IExternalCommand
    {
        string path = @"W:\0409\";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Autodesk.Revit.DB.Document document = commandData.Application.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            string doc_path = doc.PathName;
            Application app = doc.Application;
            UIApplication uiapp = new UIApplication(app);


            readfile rf = new readfile();
            rf.read_point();
            rf.read_tunnel_point();

            create_aisle(uiapp, rf.data_list_tunnel, rf.data_list[0], doc_path);

            Transaction trans = new Transaction(doc, "aisle");
            trans.Start();
            doc.LoadFamily(path + "走道\\aisle.rfa");

            FamilySymbol fam_sym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "aisle").First();
            fam_sym.Activate();
            FamilyInstance FI = doc.Create.NewFamilyInstance(new XYZ(0, 0, 0), fam_sym, StructuralType.NonStructural);
            trans.Commit();

            return Result.Succeeded;
        }

        public void create_aisle(UIApplication uiapp, IList<data_object> data_list_tunnel, data_object data, string doc_path)
        {
            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\profile.rfa");
            Document edit_doc = edit_uidoc.Document;

            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;

            Transaction trans = new Transaction(edit_doc, "test");
            trans.Start();

            XYZ point1 = new XYZ((2800 - 374 - 550) / 304.8, (850 - 1700) / 304.8, 0);

            double b2 = calculate_b_of_point_on_circle(new XYZ(0, 0, 0), XYZ.BasisX, point1, 2800 / 304.8);
            XYZ point2 = point1 + b2 * XYZ.BasisX;

            double b3 = calculate_b_of_point_on_circle(new XYZ(0, 0, 0), XYZ.BasisX, new XYZ((2800 - 374 - 550) / 304.8, (-1700) / 304.8, 0), 2800 / 304.8);
            XYZ point3 = new XYZ((2800 - 374 - 550) / 304.8, (-1700) / 304.8, 0) + b3 * XYZ.BasisX;

            double b4 = calculate_b_of_point_on_circle(new XYZ(0, 0, 0), -XYZ.BasisY, point1, 2800 / 304.8);
            XYZ point4 = point1 - b4 * XYZ.BasisY;

            DetailCurve line1 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point1, point2));
            DetailCurve arc = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Arc.Create(point2, point4, point3));
            DetailCurve line2 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point4, point1));

            trans.Commit();
            edit_doc.SaveAs(path + "走道\\" + "profile_L.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc_path);
            edit_doc.Close(false);

            edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\aisle.rfa");
            edit_doc = edit_uidoc.Document;

            trans = new Transaction(edit_doc, "sweep");
            trans.Start();
            Family family;
            edit_doc.LoadFamily(path + "走道\\" + "profile_R.rfa", out family);

            FamilySymbol fam_sym = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "profile_R").First();

            ReferenceArray curves = new ReferenceArray();
            for (int i = 1; i < data_list_tunnel.Count; i++)
            {
                Curve aisle_path = Line.CreateBound(data_list_tunnel[i - 1].start_point, data_list_tunnel[i].start_point);
                ModelCurve path_modelcurve = edit_doc.FamilyCreate.NewModelCurve(aisle_path, Sketch_plain(edit_doc, data_list_tunnel[i - 1].start_point, data_list_tunnel[i].start_point));
                curves.Append(path_modelcurve.GeometryCurve.Reference);
                path_modelcurve.LookupParameter("可見").Set(0);
            }


            SweepProfile sweepProfile = edit_doc.Application.Create.NewFamilySymbolProfile(fam_sym);
            Sweep aisle = edit_doc.FamilyCreate.NewSweep(true, curves, sweepProfile, 0, ProfilePlaneLocation.Start);
            aisle.LookupParameter("角度").Set(Math.PI * 1.5);
            trans.Commit();

            edit_doc.SaveAs(path + "走道\\" + "aisle.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc_path);
            edit_doc.Close(false);

        }

        public SketchPlane Sketch_plain(Document doc, XYZ start, XYZ end)
        {
            SketchPlane sk = null;

            XYZ v = end - start;

            double dxy = Math.Abs(v.X) + Math.Abs(v.Y);

            XYZ w = (dxy > 0.00000001)
              ? XYZ.BasisZ
              : XYZ.BasisY;

            XYZ norm = v.CrossProduct(w).Normalize();

            Plane geomPlane = Plane.CreateByNormalAndOrigin(norm, start);

            sk = SketchPlane.Create(doc, geomPlane);

            return sk;
        }

        public double calculate_b_of_point_on_circle(XYZ center, XYZ vector, XYZ point, double r)
        {
            //-b +- 根號b^2-4ac / 2a

            //TaskDialog.Show("message", "center: " + center.ToString() + " vector : " + vector.ToString() + "point : " + point.ToString());

            double cx = Math.Pow(point.X - center.X, 2);
            double cy = Math.Pow(point.Y - center.Y, 2);
            double cz = Math.Pow(point.Z - center.Z, 2);

            double bx = 2 * (point.X - center.X) * vector.X;
            double by = 2 * (point.Y - center.Y) * vector.Y;
            double bz = 2 * (point.Z - center.Z) * vector.Z;

            double ax = Math.Pow(vector.X, 2);
            double ay = Math.Pow(vector.Y, 2);
            double az = Math.Pow(vector.Z, 2);

            double c = cx + cy + cz - Math.Pow(r, 2);
            double b = bx + by + bz;
            double a = ax + ay + az;

            //TaskDialog.Show("判別式",(Math.Pow(b, 2) - 4 * a * c).ToString());

            double ans = (-b + Math.Pow(Math.Pow(b, 2) - 4 * a * c, 0.5)) / (2 * a);
            if (ans <= 0)
            {
                ans = (-b - Math.Pow(Math.Pow(b, 2) - 4 * a * c, 0.5)) / (2 * a);
            }

            return ans;
        }
    }
}
