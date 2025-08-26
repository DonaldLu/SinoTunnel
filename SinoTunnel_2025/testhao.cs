using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using DataObject;
using System.Windows.Forms;


namespace SinoTunnel_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class testhao : IExternalEventHandler
    {
        string path;

        public void Execute(UIApplication app)
        {
            path = Form1.path;
            Document act_doc = app.ActiveUIDocument.Document;
            Document doc = app.OpenAndActivateDocument(path + "\\螺栓\\螺栓.rfa").Document;
            FamilySymbol symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
                (x => x.Name == "螺栓雛形").First();
            Transaction t = new Transaction(doc);
            //螺栓半徑

            readfile rf = new readfile();
            rf.read_properties();
            double radius = rf.properties.inner_diameter / 2;

            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    t.Start("放螺栓");
                    symbol.Activate();
                    XYZ upoint = new XYZ((250 + j * 500) / 304.8, 0, radius / 304.8);
                    FamilyInstance instance1 = doc.FamilyCreate.NewFamilyInstance(upoint, symbol, StructuralType.NonStructural);

                    t.Commit();
                    t.Start("旋轉");
                    Line a = Line.CreateBound(upoint, upoint + XYZ.BasisZ);
                    if (j == 0)
                    {
                        ElementTransformUtils.RotateElement(doc, instance1.Id, a, -Math.PI / 2);
                    }
                    else
                    {
                        ElementTransformUtils.RotateElement(doc, instance1.Id, a, Math.PI / 2);
                    }
                    Line tunnel_line = Line.CreateBound(XYZ.Zero, XYZ.BasisX);
                    double sida = (-36) * i * Math.PI / 180;
                    ElementTransformUtils.RotateElement(doc, instance1.Id, tunnel_line, sida);
                    t.Commit();
                }
            }
            for (int i = 0; i < 4; i++) //放置的角度
            {
                for (int j = 0; j < 2; j++) //750在放一次
                {
                    for (int k = 0; k < 2; k++) //根據鏡射在一次
                    {
                        t.Start("放螺栓");
                        symbol.Activate();
                        XYZ upoint = new XYZ((250 + j * 500) / 304.8, 0, radius / 304.8);
                        FamilyInstance instance1 = doc.FamilyCreate.NewFamilyInstance(upoint, symbol, StructuralType.NonStructural);

                        t.Commit();
                        t.Start("旋轉");
                        Line a = Line.CreateBound(upoint, upoint + XYZ.BasisZ);
                        if (k == 0)
                        {
                            ElementTransformUtils.RotateElement(doc, instance1.Id, a, Math.PI);
                        }
                        Line tunnel_line = Line.CreateBound(XYZ.Zero, XYZ.BasisX);

                        double sida = (-54 - 72 * i) * Math.PI / 180 + 255.0 / radius;
                        if (k == 1)
                        {
                            sida = (-54 - 72 * i) * Math.PI / 180 - 255.0 / radius;
                        }
                        ElementTransformUtils.RotateElement(doc, instance1.Id, tunnel_line, sida);
                        t.Commit();
                    }
                }
            }

            //K環片
            List<FamilyInstance> K_list = new List<FamilyInstance>();
            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    t.Start("放螺栓");
                    symbol.Activate();
                    XYZ upoint = new XYZ((250 + j * 500) / 304.8, 0, radius / 304.8);
                    FamilyInstance instance1 = doc.FamilyCreate.NewFamilyInstance(upoint, symbol, StructuralType.NonStructural);

                    t.Commit();
                    t.Start("旋轉");
                    Line a = Line.CreateBound(upoint, upoint + XYZ.BasisZ);
                    if (i == 1)
                    {
                        ElementTransformUtils.RotateElement(doc, instance1.Id, a, Math.PI);
                    }
                    Line tunnel_line = Line.CreateBound(XYZ.Zero, XYZ.BasisX);

                    double sida = (10.5) * Math.PI / 180 - 255.0 / radius;
                    if (i == 1)
                    {
                        sida = (25.5) * Math.PI / 180 + 255.0 / radius;
                    }
                    if (j == 1)
                    {
                        if (i == 0)
                        {
                            sida += 100.0 / radius;
                        }
                        else
                        {
                            sida += -100.0 / radius;
                        }
                    }
                    ElementTransformUtils.RotateElement(doc, instance1.Id, tunnel_line, sida);
                    K_list.Add(instance1);
                    t.Commit();
                }
            }
            FamilySymbol K_symbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
                (x => x.Name == "K螺栓大凹槽").First();
            FamilySymbol Ks_symbol = new FilteredElementCollector(doc)
               .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
               (x => x.Name == "K螺栓小凹槽").First();
            t.Start("放螺栓");
            K_symbol.Activate();
            Ks_symbol.Activate();
            XYZ bupoint = new XYZ((250) / 304.8, 0, radius / 304.8);
            XYZ supoint = new XYZ(750 / 304.8, 0, radius / 304.8);
            FamilyInstance binstance1 = doc.FamilyCreate.NewFamilyInstance(bupoint, K_symbol, StructuralType.NonStructural);
            FamilyInstance sinstance1 = doc.FamilyCreate.NewFamilyInstance(supoint, Ks_symbol, StructuralType.NonStructural);
            Line tunnel_lines = Line.CreateBound(XYZ.Zero, XYZ.BasisX);
            double fianl = 18 * Math.PI / 180;
            K_list.Add(binstance1);
            K_list.Add(sinstance1);
            t.Commit();
            t.Start("旋轉");
            ElementTransformUtils.RotateElement(doc, binstance1.Id, tunnel_lines, fianl);
            ElementTransformUtils.RotateElement(doc, sinstance1.Id, tunnel_lines, fianl);
            t.Commit();
            SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };
            doc.SaveAs(path + "\\螺栓\\螺栓_A.rfa", saveAsOptions);

            //開始做B
            t.Start("B環片改動");
            List<FamilyInstance> rotate_ele = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where
                (x => x.Name == "K螺栓大凹槽" || x.Name == "K螺栓小凹槽" || x.Name == "螺栓雛形").ToList();
            List<ElementId> rotate_id = new List<ElementId>();
            foreach (FamilyInstance instance in rotate_ele)
            {
                rotate_id.Add(instance.Id);
            }
            Line mid_line = Line.CreateBound(new XYZ(500 / 304.8, 0, 0), new XYZ(500 / 304.8, 0, 1));
            ElementTransformUtils.RotateElements(doc, rotate_id, mid_line, Math.PI);

            foreach (FamilyInstance a in K_list)
            {
                LocationPoint point = a.Location as LocationPoint;
                if (Math.Round(point.Point.X, 2) == Math.Round(250.0 / 304.8, 2))
                {
                    a.Location.Move(new XYZ(500.0 / 304.8, 0, 0));
                }
                else if (Math.Round(point.Point.X, 2) == Math.Round(750.0 / 304.8, 2))
                {
                    a.Location.Move(new XYZ((-500.0) / 304.8, 0, 0));
                }
            }
            t.Commit();
            doc.SaveAs(path + "\\螺栓\\螺栓_B.rfa", saveAsOptions);
            app.OpenAndActivateDocument(act_doc.PathName);
            Transaction T = new Transaction(act_doc);

            doc.Close();
            T.Start("載入族群");
            act_doc.LoadFamilySymbol(path + "\\螺栓\\螺栓_B.rfa", "螺栓_B");
            act_doc.LoadFamilySymbol(path + "\\螺栓\\螺栓_A.rfa", "螺栓_A");

            T.Commit();
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
