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

namespace SinoTunnel
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class Envelope : IExternalEventHandler
    {
        string path;

        public void Execute(UIApplication app)
        {
            path = Form1.path;
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            


            //讀取資料
            readfile rf = new readfile();
            IList<IList<envelope_object>> all_envelopes = new List<IList<envelope_object>>();
            IList<IList<envelope_object>> all_envelopes_third = new List<IList<envelope_object>>();
            rf.read_envelope_1();
            rf.read_envelope_2();
            rf.read_envelope_3_1();
            rf.read_envelope_3_2();
            all_envelopes.Add(rf.envelope_1);
            all_envelopes.Add(rf.envelope_2);
            all_envelopes_third.Add(rf.envelope_3_1);
            all_envelopes_third.Add(rf.envelope_3_2);

            create_envelop(all_envelopes, app, rf, doc);
            create_envelop_third(all_envelopes_third, app, doc);



            TaskDialog.Show("result", "好惹。");
        }

        public void create_envelop(IList<IList<envelope_object>> all_envelopes, UIApplication app, readfile rf, Document doc)
        {
            Transaction t = new Transaction(doc);
            string[] en_class = { "動態", "車輛" };

            int count = 0;
            foreach (IList<envelope_object> envelope in all_envelopes)
            {
                count++;
                foreach (string en in en_class)
                {

                    //開啟包絡線畫包絡線
                    UIDocument edit_uidoc = app.OpenAndActivateDocument(path + "包絡線\\包絡線.rfa");
                    Document edit_doc = edit_uidoc.Document;
                    //動態包絡線
                    Transaction edit_t = new Transaction(edit_doc);
                    edit_t.Start("Create Modelline");
                    //畫每個輪廓的線
                    if (en == "動態")
                    {
                        for (int i = 0; i < envelope.Count; i++)
                        {
                            draw_profile(edit_doc, rf, envelope[i].Dynamic_envelope);
                        }
                    }
                    else if (en == "車輛")
                    {
                        for (int i = 0; i < envelope.Count; i++)
                        {
                            draw_profile(edit_doc, rf, envelope[i].Vehicle_envelope);
                        }
                    }
                    //畫輪廓之間的線
                    //OK我們先不畫
                    //between_profile(edit_doc, rf.envelope, en);

                    edit_t.Commit();

                    //另存起來
                    SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };
                    edit_doc.SaveAs(path + "包絡線\\" + en + count.ToString() + "包絡線.rfa", saveAsOptions);

                    app.OpenAndActivateDocument(doc.PathName);
                    edit_doc.Close();

                    //載入
                    t.Start("載入族群");
                    doc.LoadFamilySymbol(path + "包絡線\\" + en + count.ToString() + "包絡線.rfa", en + count.ToString() + "包絡線");
                    t.Commit();
                    //創建實體
                    t.Start("創建包絡線");
                    FamilySymbol pre_load = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
                    (x => x.Name == en + count.ToString() + "包絡線").First();
                    pre_load.Activate();
                    FamilyInstance object_to_load = doc.Create.NewFamilyInstance(XYZ.Zero, pre_load, StructuralType.NonStructural);
                    t.Commit();
                }
            }
        }

        public void create_envelop_third(IList<IList<envelope_object>> all_envelopes, UIApplication app, Document doc)
        {
            Transaction t = new Transaction(doc);
            string[] en_class = { "第三軌" };

            int count = 0;
            foreach (IList<envelope_object> envelope in all_envelopes)
            {
                count++;
                foreach (string en in en_class)
                {

                    //開啟包絡線畫包絡線
                    UIDocument edit_uidoc = app.OpenAndActivateDocument(path + "包絡線\\包絡線.rfa");
                    Document edit_doc = edit_uidoc.Document;
                    //動態包絡線
                    Transaction edit_t = new Transaction(edit_doc);
                    edit_t.Start("Create Modelline");
                    //畫每個輪廓的線
                    if (en == "第三軌")
                    {
                        for (int i = 0; i < envelope.Count; i++)
                        {
                            draw_profile_third(edit_doc, envelope[i].third_envelope);
                        }
                    }
                    //畫輪廓之間的線
                    //OK我們先不畫
                    //between_profile(edit_doc, rf.envelope, en);

                    edit_t.Commit();

                    //另存起來
                    SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };
                    edit_doc.SaveAs(path + "包絡線\\" + en + count.ToString() + "包絡線.rfa", saveAsOptions);

                    app.OpenAndActivateDocument(doc.PathName);
                    edit_doc.Close();

                    //載入
                    t.Start("載入族群");
                    doc.LoadFamilySymbol(path + "包絡線\\" + en + count.ToString() + "包絡線.rfa", en + count.ToString() + "包絡線");
                    t.Commit();
                    //創建實體
                    t.Start("創建包絡線");
                    FamilySymbol pre_load = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
                    (x => x.Name == en + count.ToString() + "包絡線").First();
                    pre_load.Activate();
                    FamilyInstance object_to_load = doc.Create.NewFamilyInstance(XYZ.Zero, pre_load, StructuralType.NonStructural);
                    t.Commit();
                }
            }
        }

        public SketchPlane Sketch_plain(Document doc, XYZ start, XYZ end)
        {
            SketchPlane sk = null;

            XYZ v = end - start;

            double dxy = Math.Abs(v.X) + Math.Abs(v.Y);

            XYZ w = (dxy > 0.00000001)
              ? XYZ.BasisY
              : XYZ.BasisZ;

            XYZ norm = v.CrossProduct(w).Normalize();

            Plane geomPlane = Plane.CreateByNormalAndOrigin(norm, start);

            sk = SketchPlane.Create(doc, geomPlane);

            return sk;
        }
        void draw_profile(Document edit_doc, readfile rf, List<XYZ> point_list)
        {
            Plane plane = Plane.CreateByThreePoints(point_list[0],
                point_list[6],
                point_list[12]);
            for (int j = 0; j < 12; j++)
            {
                XYZ start_point = point_list[j];
                XYZ end_point = point_list[j + 1];
                SketchPlane skplane = SketchPlane.Create(edit_doc, plane);
                Curve line = Line.CreateBound(start_point, end_point);
                edit_doc.FamilyCreate.NewModelCurve(line, skplane);

            }
        }
        void draw_profile_third(Document edit_doc, List<XYZ> point_list)
        {
            //TaskDialog.Show("message", point_list.Count.ToString());
            Plane plane = Plane.CreateByThreePoints(point_list[0],
                point_list[1],
                point_list[2]);
            //TaskDialog.Show("message", "2.1");
            for (int j = 0; j < 4; j++)
            {
                XYZ start_point = point_list[j];
                XYZ end_point = point_list[j + 1];
                SketchPlane skplane = SketchPlane.Create(edit_doc, plane);
                Curve line = Line.CreateBound(start_point, end_point);
                edit_doc.FamilyCreate.NewModelCurve(line, skplane);

            }
        }
        void between_profile(Document edit_doc, IList<envelope_object> envelope_list, string en)
        {
            for (int i = 0; i < envelope_list.Count - 1; i++)
            {
                for (int j = 0; j < 13; j++)
                {
                    XYZ start = null, end = null;
                    if (en == "動態")
                    {
                        start = envelope_list[i].Dynamic_envelope[j];
                        end = envelope_list[i + 1].Dynamic_envelope[j];
                    }
                    else if (en == "車輛")
                    {
                        start = envelope_list[i].Vehicle_envelope[j];
                        end = envelope_list[i + 1].Vehicle_envelope[j];
                    }
                    Curve line = Line.CreateBound(start, end);
                    edit_doc.FamilyCreate.NewModelCurve(line, Sketch_plain(edit_doc, start, end));
                }
            }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}