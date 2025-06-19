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
    class shear_rebar : IExternalEventHandler
    {
        public void Execute(UIApplication uiapp)
        {
            readfile rf = new readfile();
            rf.read_tunnel_point();
            rf.read_properties();
            rf.read_rebar_properties();

            Document doc = uiapp.ActiveUIDocument.Document;
            View3D view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().ToList().First();
            Element ele = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList()
                .Where(x => x.Name == "自適應環形00_A" && x.LookupParameter("備註").AsString() == $"{rf.firstStation}").First();
            
            
            Transaction t = new Transaction(doc);
            t.Start("剪力筋測試");
            //A環片剪力筋
            try
            {
                for (int j = 0; j < 3; j++)
                {
                    double dis = 0;
                    double seg_angle = Math.PI / 2.0 + 72.0 * j * Math.PI / 180.0;
                    for (int i = 0; i < rf.rebar.shear_A_distance.Count(); i++)
                    {
                        dis += double.Parse(rf.rebar.shear_A_distance[i]);
                        switch (rf.rebar.shear_A_type[i])
                        {
                            case "1":
                                shear_t1(rf, doc, ele, view3D, dis, seg_angle, 1, "A");
                                break;
                            case "1t":
                                shear_t1(rf, doc, ele, view3D, dis, seg_angle, -1, "A");
                                break;
                            case "2":
                                shear_t2(rf, doc, ele, view3D, dis, seg_angle, 1, "A");
                                break;
                            case "2t":
                                shear_t2(rf, doc, ele, view3D, dis, seg_angle, -1, "A");
                                break;
                            case "3":
                                shear_t3(rf, doc, ele, view3D, dis, seg_angle, "A");
                                break;
                            case "4":
                                shear_t4(rf, doc, ele, view3D, dis, seg_angle, 1, "A");
                                break;
                            case "4t":
                                shear_t4(rf, doc, ele, view3D, dis, seg_angle, -1, "A");
                                break;
                            case "5":
                                shear_t5(rf, doc, ele, view3D, dis, seg_angle, "A");
                                break;
                            case "6":
                                shear_t6(rf, doc, ele, view3D, dis, seg_angle, 1, "A");
                                break;
                            case "7":
                                shear_t7(rf, doc, ele, view3D, dis, seg_angle, "A");
                                break;
                            case "8":
                                shear_t6(rf, doc, ele, view3D, dis, seg_angle, 1, "A");
                                break;
                            case "8t":
                                shear_t6(rf, doc, ele, view3D, dis, seg_angle, -1, "A");
                                break;
                            case "9":
                                shear_t9(rf, doc, ele, view3D, dis, seg_angle, 1, "A");
                                break;
                            case "9t":
                                shear_t9(rf, doc, ele, view3D, dis, seg_angle, -1, "A");
                                break;
                            default:
                                continue;
                        }
                    }
                }
                //B環片剪力筋
                for (int j = 0; j < 2; j++)
                {
                    double dis = 0;
                    double seg_angle = (j == 0) ? 25.5 * Math.PI / 180.0 : 10.5 * Math.PI / 180.0;
                    for (int i = 0; i < rf.rebar.shear_B_distance.Count(); i++)
                    {
                        if (j == 0)
                        {
                            dis += double.Parse(rf.rebar.shear_B_distance[i]);
                        }
                        else
                        {
                            dis -= double.Parse(rf.rebar.shear_B_distance[i]);
                        }
                        switch (rf.rebar.shear_B_type[i])
                        {
                            case "1":
                                shear_t1(rf, doc, ele, view3D, dis, seg_angle, 1, "B");
                                break;
                            case "1t":
                                shear_t1(rf, doc, ele, view3D, dis, seg_angle, -1, "B");
                                break;
                            case "2":
                                shear_t2(rf, doc, ele, view3D, dis, seg_angle, 1, "B");
                                break;
                            case "2t":
                                shear_t2(rf, doc, ele, view3D, dis, seg_angle, -1, "B");
                                break;
                            case "3":
                                shear_t3(rf, doc, ele, view3D, dis, seg_angle, "B");
                                break;
                            case "4":
                                shear_t4(rf, doc, ele, view3D, dis, seg_angle, 1, "B");
                                break;
                            case "4t":
                                shear_t4(rf, doc, ele, view3D, dis, seg_angle, -1, "B");
                                break;
                            case "5":
                                shear_t5(rf, doc, ele, view3D, dis, seg_angle, "B");
                                break;
                            case "6":
                                shear_t6(rf, doc, ele, view3D, dis, seg_angle, 1, "B");
                                break;
                            case "7":
                                shear_t7(rf, doc, ele, view3D, dis, seg_angle, "B");
                                break;
                            case "8":
                                shear_t6(rf, doc, ele, view3D, dis, seg_angle, 1, "B");
                                break;
                            case "8t":
                                shear_t6(rf, doc, ele, view3D, dis, seg_angle, -1, "B");
                                break;
                            case "9":
                                shear_t9(rf, doc, ele, view3D, dis, seg_angle, 1, "B");
                                break;
                            case "9t":
                                shear_t9(rf, doc, ele, view3D, dis, seg_angle, -1, "B");
                                break;
                            default:
                                continue;
                        }
                    }
                }
                //K環片剪力筋
                double dis_K = 0;
                double seg_angle_K = 10.5 * Math.PI / 180.0;
                for (int i = 0; i < rf.rebar.shear_K_distance.Count(); i++)
                {
                    dis_K += double.Parse(rf.rebar.shear_K_distance[i]);

                    switch (rf.rebar.shear_K_type[i])
                    {
                        case "1":
                            shear_t1(rf, doc, ele, view3D, dis_K, seg_angle_K, 1, "K");
                            break;
                        case "1t":
                            shear_t1(rf, doc, ele, view3D, dis_K, seg_angle_K, -1, "K");
                            break;
                        case "2":
                            shear_t2(rf, doc, ele, view3D, dis_K, seg_angle_K, 1, "K");
                            break;
                        case "2t":
                            shear_t2(rf, doc, ele, view3D, dis_K, seg_angle_K, -1, "K");
                            break;
                        case "3":
                            shear_t3(rf, doc, ele, view3D, dis_K, seg_angle_K, "K");
                            break;
                        case "4":
                            shear_t4(rf, doc, ele, view3D, dis_K, seg_angle_K, 1, "K");
                            break;
                        case "4t":
                            shear_t4(rf, doc, ele, view3D, dis_K, seg_angle_K, -1, "K");
                            break;
                        case "5":
                            shear_t5(rf, doc, ele, view3D, dis_K, seg_angle_K, "K");
                            break;
                        case "6":
                            shear_t6(rf, doc, ele, view3D, dis_K, seg_angle_K, 1, "K");
                            break;
                        case "7":
                            shear_t7(rf, doc, ele, view3D, dis_K, seg_angle_K, "K");
                            break;
                        case "8":
                            shear_t6(rf, doc, ele, view3D, dis_K, seg_angle_K, 1, "K");
                            break;
                        case "8t":
                            shear_t6(rf, doc, ele, view3D, dis_K, seg_angle_K, -1, "K");
                            break;
                        case "9":
                            shear_t9(rf, doc, ele, view3D, dis_K, seg_angle_K, 1, "K");
                            break;
                        case "9t":
                            shear_t9(rf, doc, ele, view3D, dis_K, seg_angle_K, -1, "K");
                            break;
                        default:
                            continue;
                    }
                }
            }catch(Exception e) { TaskDialog.Show("error",e.Message + e.StackTrace); }
            //double dis = 0;
            //double seg_angle = 0;
            //shear_t1(rf, doc, ele, view3D, dis, seg_angle, -1);
            TaskDialog.Show("test", "剪力筋測試完畢");
            t.Commit();
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }

        void shear_t1(readfile rf, Document doc, Element ele, View3D view3D, double dis, double seg_angle, int index, string ABK)
        {
            SubTransaction subt = new SubTransaction(doc);

            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (rf.properties.thickness - rf.rebar.main_outer_protect)) / 304.8;
            double side_protect = rf.rebar.main_side_protect / 304.8;

            RebarShape shape = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_T1").First();
            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarHookType hookType = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍地震 - 135 度").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);
            XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
            XYZ start = rf.data_list_tunnel[0].start_point;
            XYZ end = rf.data_list_tunnel[1].start_point;

            XYZ p1 = new XYZ(start.X, start.Y, start.Z + inner_radius);
            p1 = p1 + side_protect * toward.Direction;
            XYZ p2 = new XYZ(end.X, end.Y, end.Z + inner_radius);
            p2 = p2 - side_protect * toward.Direction;
            XYZ p3 = new XYZ(end.X, end.Y, end.Z + outer_radius);
            p3 = p3 - side_protect * toward.Direction;
            XYZ p4 = new XYZ(start.X, start.Y, start.Z + outer_radius);
            p4 = p4 + side_protect * toward.Direction;
            XYZ p5 = (start + end) / 2 + (150 / 304.8) * toward.Direction;
            XYZ p8 = p5.Add(new XYZ(0, 0, inner_radius));
            p5 = p5.Add(new XYZ(0, 0, outer_radius));
            XYZ p7 = (start + end) / 2 - (150 / 304.8) * toward.Direction;
            XYZ p6 = p7.Add(new XYZ(0, 0, outer_radius));
            p7 = p7.Add(new XYZ(0, 0, inner_radius));

            Plane plane = Plane.CreateByThreePoints(p1, p2, p3);
            SketchPlane skp = SketchPlane.Create(doc, plane);

            IList<Curve> curves = new List<Curve>();
            IList<Curve> curves_s = new List<Curve>();

            Curve c = Line.CreateBound(p1, p2);
            Curve c2 = Line.CreateBound(p2, p3);
            Curve c3 = Line.CreateBound(p3, p4);
            Curve c4 = Line.CreateBound(p4, p1);
            curves.Add(c3);
            curves.Add(c4);
            curves.Add(c);
            curves.Add(c2);

            Curve c5 = Line.CreateBound(p5, p6);
            Curve c6 = Line.CreateBound(p6, p7);
            Curve c7 = Line.CreateBound(p7, p8);
            Curve c8 = Line.CreateBound(p8, p5);
            curves_s.Add(c5);
            curves_s.Add(c6);
            curves_s.Add(c7);
            curves_s.Add(c8);

            subt.Start();
            Rebar re = Rebar.CreateFromCurvesAndShape(
                doc, shape, barType13M, hookType, hookType, ele, plane.Normal, curves,
                RebarHookOrientation.Left, RebarHookOrientation.Left);
            Rebar re2 = Rebar.CreateFromCurvesAndShape(
                doc, shape, barType13M, hookType, hookType, ele, plane.Normal, curves_s,
                RebarHookOrientation.Left, RebarHookOrientation.Left);
            subt.Commit();
            subt.Start();
            //re.SetSolidInView(view3D, true);
            //re.SetUnobscuredInView(view3D, true);
            //re2.SetSolidInView(view3D, true);
            //re2.SetUnobscuredInView(view3D, true);
            //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);

            // 培文改：設置鋼筋
            List<Rebar> rebars = new List<Rebar>() { re, re2 };
            foreach (Rebar rebar in rebars) { SetRebar(rebar, view3D, ABK); }

            ICollection<ElementId> id = new List<ElementId>();
            id.Add(re.Id);
            id.Add(re2.Id);
            double steel_rad = re.LookupParameter("鋼筋直徑").AsDouble() * 304.8;
            ElementTransformUtils.RotateElement(doc, re2.Id, toward, (-steel_rad) * index / radius);
            ElementTransformUtils.RotateElements(doc, id, toward, seg_angle + dis / radius);
            subt.Commit();

        }
        void shear_t2(readfile rf, Document doc, Element ele, View3D view3D, double dis, double seg_angle, int index, string ABK)
        {
            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (rf.properties.thickness - rf.rebar.main_outer_protect)) / 304.8;
            double side_protect = rf.rebar.main_side_protect / 304.8;

            RebarShape shape = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_T1").First();
            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarHookType hookType = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍地震 - 135 度").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);
            XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
            XYZ start = rf.data_list_tunnel[0].start_point;
            XYZ end = rf.data_list_tunnel[1].start_point;

            XYZ p1 = new XYZ(start.X, start.Y, start.Z + inner_radius);
            p1 = p1 + side_protect * toward.Direction;
            XYZ p2 = new XYZ(end.X, end.Y, end.Z + inner_radius);
            p2 = p2 - side_protect * toward.Direction;
            XYZ p3 = new XYZ(end.X, end.Y, end.Z + outer_radius);
            p3 = p3 - side_protect * toward.Direction;
            XYZ p4 = new XYZ(start.X, start.Y, start.Z + outer_radius);
            p4 = p4 + side_protect * toward.Direction;
            XYZ p5 = p1 + (430 / 304.8) * toward.Direction;
            XYZ p6 = p4 + (430 / 304.8) * toward.Direction;
            XYZ p7 = p2 - (430 / 304.8) * toward.Direction;
            XYZ p8 = p3 - (430 / 304.8) * toward.Direction;

            Plane plane = Plane.CreateByThreePoints(p1, p2, p3);
            SketchPlane skp = SketchPlane.Create(doc, plane);

            IList<Curve> curves = new List<Curve>();
            IList<Curve> curves_s = new List<Curve>();
            IList<Curve> curves_s2 = new List<Curve>();

            Curve c = Line.CreateBound(p1, p2);
            Curve c2 = Line.CreateBound(p2, p3);
            Curve c3 = Line.CreateBound(p3, p4);
            Curve c4 = Line.CreateBound(p4, p1);
            curves.Add(c3);
            curves.Add(c4);
            curves.Add(c);
            curves.Add(c2);

            Curve c5 = Line.CreateBound(p3, p8);
            Curve c6 = Line.CreateBound(p8, p7);
            Curve c7 = Line.CreateBound(p7, p2);
            Curve c8 = Line.CreateBound(p2, p3);
            curves_s.Add(c5);
            curves_s.Add(c6);
            curves_s.Add(c7);
            curves_s.Add(c8);

            Curve c9 = Line.CreateBound(p4, p6);
            Curve c10 = Line.CreateBound(p6, p5);
            Curve c11 = Line.CreateBound(p5, p1);
            Curve c12 = Line.CreateBound(p1, p4);
            curves_s2.Add(c9);
            curves_s2.Add(c10);
            curves_s2.Add(c11);
            curves_s2.Add(c12);

            Rebar re = Rebar.CreateFromCurvesAndShape(
                doc, shape, barType13M, hookType, hookType, ele, plane.Normal, curves,
                RebarHookOrientation.Left, RebarHookOrientation.Left);
            Rebar re2 = Rebar.CreateFromCurvesAndShape(
                doc, shape, barType13M, hookType, hookType, ele, plane.Normal, curves_s,
                RebarHookOrientation.Left, RebarHookOrientation.Left);
            Rebar re3 = Rebar.CreateFromCurvesAndShape(
                doc, shape, barType13M, hookType, hookType, ele, plane.Normal, curves_s2,
                RebarHookOrientation.Right, RebarHookOrientation.Right);


            //re.SetSolidInView(view3D, true);
            //re.SetUnobscuredInView(view3D, true);
            //re2.SetSolidInView(view3D, true);
            //re2.SetUnobscuredInView(view3D, true);
            //re3.SetSolidInView(view3D, true);
            //re3.SetUnobscuredInView(view3D, true);
            //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);

            // 培文改：設置鋼筋
            List<Rebar> rebars = new List<Rebar>() { re, re2, re3 };
            foreach (Rebar rebar in rebars) { SetRebar(rebar, view3D, ABK); }

            ICollection<ElementId> id = new List<ElementId>();
            id.Add(re.Id);
            id.Add(re2.Id);
            id.Add(re3.Id);
            double steel_rad = re.LookupParameter("鋼筋直徑").AsDouble() * 304.8;
            ElementTransformUtils.RotateElement(doc, re2.Id, toward, (-steel_rad) * index / radius);
            ElementTransformUtils.RotateElement(doc, re3.Id, toward, (-steel_rad) * index / radius);
            ElementTransformUtils.RotateElements(doc, id, toward, seg_angle + dis / radius);
        }
        void shear_t3(readfile rf, Document doc, Element ele, View3D view3D, double dis, double seg_angle, string ABK)
        {
            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (rf.properties.thickness - rf.rebar.main_outer_protect)) / 304.8;
            double side_protect = rf.rebar.main_side_protect / 304.8;

            RebarShape M_00 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_00").First();
            RebarShape M_T1 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_T1").First();
            RebarShape M_03 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_03").First();
            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarBarType barType16M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "16M").First();
            RebarHookType hookType = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍地震 - 135 度").First();
            RebarHookType hookType2 = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍 - 90 度").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);
            XYZ start = rf.data_list_tunnel[0].start_point;
            XYZ end = rf.data_list_tunnel[1].start_point;

            IList<Curve> L1 = new List<Curve>();
            IList<Curve> L2 = new List<Curve>();
            IList<Curve> L3 = new List<Curve>();
            IList<Curve> L4 = new List<Curve>();

            XYZ p1 = new XYZ(start.X, start.Y, start.Z + outer_radius) + side_protect * toward.Direction;
            XYZ p2 = new XYZ(end.X, end.Y, end.Z + outer_radius) - side_protect * toward.Direction;
            Curve c1 = Line.CreateBound(p1, p2);

            L1.Add(c1);

            XYZ p3 = new XYZ(start.X, start.Y, start.Z + outer_radius) + 351.6 / 304.8 * toward.Direction;
            XYZ p4 = new XYZ(end.X, end.Y, end.Z + outer_radius) - 350.6 / 304.8 * toward.Direction;
            XYZ p5 = new XYZ(p4.X, p4.Y, p4.Z - 185.1 / 304.8);
            XYZ p6 = new XYZ(p3.X, p3.Y, p3.Z - 185.1 / 304.8);
            Curve c2 = Line.CreateBound(p3, p4);
            Curve c3 = Line.CreateBound(p4, p5);
            Curve c4 = Line.CreateBound(p5, p6);
            Curve c5 = Line.CreateBound(p6, p3);

            L2.Add(c2);
            L2.Add(c3);
            L2.Add(c4);
            L2.Add(c5);

            XYZ p7 = new XYZ(start.X, start.Y, start.Z + outer_radius - 17.8 / 304.8) + 220.7 / 304.8 * toward.Direction;
            XYZ p8 = new XYZ(p1.X, p1.Y, p1.Z - 17.8 / 304.8);
            XYZ p9 = p8 - new XYZ(0, 0, 167.2 / 304.8);
            XYZ p10 = p9 + 120.0 / 304.8 * toward.Direction;
            XYZ p11 = p10 + new XYZ(0, 0, 149.4 / 304.8) + 49.8 * (149.4 / 136.0) / 304.8 * toward.Direction;
            XYZ p17 = start + 146.0 / 304.8 * toward.Direction - new XYZ(0, 0, 35.6 / 304.8 - outer_radius);
            Curve c6 = Line.CreateBound(p7, p8);
            Curve c7 = Line.CreateBound(p8, p9);
            Curve c8 = Line.CreateBound(p9, p10);
            Curve c9 = Line.CreateBound(p10, p11);
            Curve c14 = Line.CreateBound(p11, p17);

            L3.Add(c6);
            L3.Add(c7);
            L3.Add(c8);
            L3.Add(c9);
            L3.Add(c14);

            XYZ p12 = new XYZ(end.X, end.Y, end.Z + outer_radius - 17.8 / 304.8) - 220.7 / 304.8 * toward.Direction;
            XYZ p13 = new XYZ(p2.X, p2.Y, p2.Z - 17.8 / 304.8);
            XYZ p14 = p13 - new XYZ(0, 0, 167.2 / 304.8);
            XYZ p15 = p14 - 120.0 / 304.8 * toward.Direction;
            XYZ p16 = p15 + new XYZ(0, 0, 149.4 / 304.8) - 49.8 * (149.4 / 136.0) / 304.8 * toward.Direction;
            XYZ p18 = end - 146.0 / 304.8 * toward.Direction - new XYZ(0, 0, 35.6 / 304.8 - outer_radius);
            Curve c10 = Line.CreateBound(p12, p13);
            Curve c11 = Line.CreateBound(p13, p14);
            Curve c12 = Line.CreateBound(p14, p15);
            Curve c13 = Line.CreateBound(p15, p16);
            Curve c15 = Line.CreateBound(p16, p18);

            L4.Add(c10);
            L4.Add(c11);
            L4.Add(c12);
            L4.Add(c13);
            L4.Add(c15);

            Plane plane = Plane.CreateByThreePoints(p8, p2, p3);

            Rebar R1 = Rebar.CreateFromCurvesAndShape(doc, M_00, barType13M, null, null, ele, plane.Normal, L1
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R2 = Rebar.CreateFromCurvesAndShape(doc, M_T1, barType13M, hookType, hookType, ele, plane.Normal, L2
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3_2 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
               , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R4 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L4
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R4_2 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L4
               , RebarHookOrientation.Right, RebarHookOrientation.Right);

            //R1.SetSolidInView(view3D, true);
            //R1.SetUnobscuredInView(view3D, true);
            //R2.SetSolidInView(view3D, true);
            //R2.SetUnobscuredInView(view3D, true);
            //R3.SetSolidInView(view3D, true);
            //R3.SetUnobscuredInView(view3D, true);
            //R4.SetSolidInView(view3D, true);
            //R4.SetUnobscuredInView(view3D, true);
            //R3_2.SetSolidInView(view3D, true);
            //R3_2.SetUnobscuredInView(view3D, true);
            //R4_2.SetSolidInView(view3D, true);
            //R4_2.SetUnobscuredInView(view3D, true);
            //R1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R4.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R4_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);

            // 培文改：設置鋼筋
            List<Rebar> rebars = new List<Rebar>() { R1, R2, R3, R4, R3_2, R4_2 };
            foreach (Rebar rebar in rebars) { SetRebar(rebar, view3D, ABK); }

            ICollection<ElementId> id = new List<ElementId>();
            id.Add(R1.Id);
            id.Add(R2.Id);
            id.Add(R3.Id);
            id.Add(R4.Id);
            id.Add(R3_2.Id);
            id.Add(R4_2.Id);
            ElementTransformUtils.RotateElement(doc, R3.Id, toward, -30.0 / radius);
            ElementTransformUtils.RotateElement(doc, R4.Id, toward, -30.0 / radius);
            ElementTransformUtils.RotateElement(doc, R3_2.Id, toward, 70.0 / radius);
            ElementTransformUtils.RotateElement(doc, R4_2.Id, toward, 70.0 / radius);
            ElementTransformUtils.RotateElements(doc, id, toward, seg_angle + dis / radius);//Math.PI / 2.0 +
        }
        void shear_t4(readfile rf, Document doc, Element ele, View3D view3D, double dis, double seg_angle, int index, string ABK)
        {
            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (rf.properties.thickness - rf.rebar.main_outer_protect)) / 304.8;
            double side_protect = rf.rebar.main_side_protect / 304.8;

            RebarShape shape = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_T1").First();
            RebarShape shape2 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_S4").First();
            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarHookType hookType = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍地震 - 135 度").First();
            RebarHookType hookType2 = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍 - 90 度").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);
            XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
            XYZ start = rf.data_list_tunnel[0].start_point;
            XYZ end = rf.data_list_tunnel[1].start_point;

            XYZ p1 = new XYZ(start.X, start.Y, start.Z + inner_radius) + side_protect * toward.Direction;
            XYZ p2 = new XYZ(end.X, end.Y, end.Z + inner_radius) - side_protect * toward.Direction;
            XYZ p3 = new XYZ(end.X, end.Y, end.Z + outer_radius) - side_protect * toward.Direction;
            XYZ p4 = new XYZ(start.X, start.Y, start.Z + outer_radius) + side_protect * toward.Direction;
            XYZ p5 = p4 + (150 / 304.8) * toward.Direction;
            XYZ p6 = (p1 + (150 / 304.8) * toward.Direction) + new XYZ(0, 0, 60 / 304.8);
            XYZ p7 = p5 + (120 / 304.8) * toward.Direction;
            XYZ p8 = p6 + (120 / 304.8) * toward.Direction;
            XYZ p9 = p3 - ((150 + 120) / 304.8) * toward.Direction;
            XYZ p10 = (p2 - ((150 + 120) / 304.8) * toward.Direction).Add(new XYZ(0, 0, 60 / 304.8));
            XYZ p11 = p3 - (150 / 304.8) * toward.Direction;
            XYZ p12 = (p2 - (150 / 304.8) * toward.Direction).Add(new XYZ(0, 0, 60 / 304.8));

            Plane plane = Plane.CreateByThreePoints(p1, p2, p3);
            SketchPlane skp = SketchPlane.Create(doc, plane);

            IList<Curve> curves = new List<Curve>();
            IList<Curve> curves_s = new List<Curve>();
            IList<Curve> curves_s2 = new List<Curve>();

            Curve c = Line.CreateBound(p1, p2);
            Curve c2 = Line.CreateBound(p2, p3);
            Curve c3 = Line.CreateBound(p3, p4);
            Curve c4 = Line.CreateBound(p4, p1);
            curves.Add(c3);
            curves.Add(c4);
            curves.Add(c);
            curves.Add(c2);

            Curve c5 = Line.CreateBound(p5, p6);
            Curve c6 = Line.CreateBound(p8, p7);
            Arc c7 = Arc.Create(p6, p8, (p1 + (210) / 304.8 * toward.Direction));
            curves_s.Add(c5);
            curves_s.Add(c7);
            curves_s.Add(c6);

            Curve c8 = Line.CreateBound(p9, p10);
            Curve c9 = Line.CreateBound(p12, p11);
            Curve c10 = Arc.Create(p10, p12, (p2 - (210) / 304.8 * toward.Direction));

            curves_s2.Add(c8);
            curves_s2.Add(c10);
            curves_s2.Add(c9);

            Rebar re = Rebar.CreateFromCurvesAndShape(
                doc, shape, barType13M, hookType, hookType, ele, plane.Normal, curves,
                RebarHookOrientation.Left, RebarHookOrientation.Left);
            Rebar re2 = Rebar.CreateFromCurvesAndShape(
                doc, shape2, barType13M, hookType2, hookType2, ele, plane.Normal, curves_s,
                RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar re3 = Rebar.CreateFromCurvesAndShape(
                doc, shape2, barType13M, hookType2, hookType2, ele, plane.Normal, curves_s2,
                RebarHookOrientation.Right, RebarHookOrientation.Right);


            //re.SetSolidInView(view3D, true);
            //re.SetUnobscuredInView(view3D, true);
            //re2.SetSolidInView(view3D, true);
            //re2.SetUnobscuredInView(view3D, true);
            //re3.SetSolidInView(view3D, true);
            //re3.SetUnobscuredInView(view3D, true);
            //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);

            // 培文改：設置鋼筋
            List<Rebar> rebars = new List<Rebar>() { re, re2, re3 };
            foreach (Rebar rebar in rebars) { SetRebar(rebar, view3D, ABK); }

            ICollection<ElementId> id = new List<ElementId>();
            id.Add(re.Id);
            id.Add(re2.Id);
            id.Add(re3.Id);
            double steel_rad = re.LookupParameter("鋼筋直徑").AsDouble() * 304.8;
            ElementTransformUtils.RotateElement(doc, re2.Id, toward, (-steel_rad) * index / radius);
            ElementTransformUtils.RotateElement(doc, re3.Id, toward, (-steel_rad) * index / radius);
            ElementTransformUtils.RotateElements(doc, id, toward, seg_angle + dis / radius);
        }
        void shear_t5(readfile rf, Document doc, Element ele, View3D view3D, double dis, double seg_angle, string ABK)
        {
            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (rf.properties.thickness - rf.rebar.main_outer_protect)) / 304.8;
            double side_protect = rf.rebar.main_side_protect / 304.8;

            RebarShape M_00 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_00").First();
            RebarShape M_T1 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_T1").First();
            RebarShape M_03 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_03").First();
            RebarShape M_S5 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_S5").First();

            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarBarType barType16M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "16M").First();
            RebarHookType hookType = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍地震 - 135 度").First();
            RebarHookType hookType2 = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍 - 90 度").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);
            XYZ start = rf.data_list_tunnel[0].start_point;
            XYZ end = rf.data_list_tunnel[1].start_point;

            IList<Curve> L1 = new List<Curve>();
            IList<Curve> L2 = new List<Curve>();
            IList<Curve> L3 = new List<Curve>();
            IList<Curve> L4 = new List<Curve>();
            IList<Curve> L5 = new List<Curve>();
            IList<Curve> L6 = new List<Curve>();

            XYZ p1 = new XYZ(start.X, start.Y, start.Z + outer_radius) + side_protect * toward.Direction;
            XYZ p2 = new XYZ(end.X, end.Y, end.Z + outer_radius) - side_protect * toward.Direction;
            Curve c1 = Line.CreateBound(p1, p2);

            L1.Add(c1);

            XYZ p3 = new XYZ(start.X, start.Y, start.Z + outer_radius) + 351.6 / 304.8 * toward.Direction;
            XYZ p4 = new XYZ(end.X, end.Y, end.Z + outer_radius) - 350.6 / 304.8 * toward.Direction;
            XYZ p5 = new XYZ(p4.X, p4.Y, p4.Z - 185.1 / 304.8);
            XYZ p6 = new XYZ(p3.X, p3.Y, p3.Z - 185.1 / 304.8);
            Curve c2 = Line.CreateBound(p3, p4);
            Curve c3 = Line.CreateBound(p4, p5);
            Curve c4 = Line.CreateBound(p5, p6);
            Curve c5 = Line.CreateBound(p6, p3);

            L2.Add(c2);
            L2.Add(c3);
            L2.Add(c4);
            L2.Add(c5);

            XYZ p7 = new XYZ(start.X, start.Y, start.Z + outer_radius - 17.8 / 304.8) + 220.7 / 304.8 * toward.Direction;
            XYZ p8 = new XYZ(p1.X, p1.Y, p1.Z - 17.8 / 304.8);
            XYZ p9 = p8 - new XYZ(0, 0, 167.2 / 304.8);
            XYZ p10 = p9 + 120.0 / 304.8 * toward.Direction;
            XYZ p11 = p10 + new XYZ(0, 0, 149.4 / 304.8) + 49.8 * (149.4 / 136.0) / 304.8 * toward.Direction;
            XYZ p17 = start + 146.0 / 304.8 * toward.Direction - new XYZ(0, 0, 35.6 / 304.8 - outer_radius);
            Curve c6 = Line.CreateBound(p7, p8);
            Curve c7 = Line.CreateBound(p8, p9);
            Curve c8 = Line.CreateBound(p9, p10);
            Curve c9 = Line.CreateBound(p10, p11);
            Curve c14 = Line.CreateBound(p11, p17);

            L3.Add(c6);
            L3.Add(c7);
            L3.Add(c8);
            L3.Add(c9);
            L3.Add(c14);

            XYZ p12 = new XYZ(end.X, end.Y, end.Z + outer_radius - 17.8 / 304.8) - 220.7 / 304.8 * toward.Direction;
            XYZ p13 = new XYZ(p2.X, p2.Y, p2.Z - 17.8 / 304.8);
            XYZ p14 = p13 - new XYZ(0, 0, 167.2 / 304.8);
            XYZ p15 = p14 - 120.0 / 304.8 * toward.Direction;
            XYZ p16 = p15 + new XYZ(0, 0, 149.4 / 304.8) - 49.8 * (149.4 / 136.0) / 304.8 * toward.Direction;
            XYZ p18 = end - 146.0 / 304.8 * toward.Direction - new XYZ(0, 0, 35.6 / 304.8 - outer_radius);
            Curve c10 = Line.CreateBound(p12, p13);
            Curve c11 = Line.CreateBound(p13, p14);
            Curve c12 = Line.CreateBound(p14, p15);
            Curve c13 = Line.CreateBound(p15, p16);
            Curve c15 = Line.CreateBound(p16, p18);

            L4.Add(c10);
            L4.Add(c11);
            L4.Add(c12);
            L4.Add(c13);
            L4.Add(c15);

            XYZ p19 = (new XYZ(start.X, start.Y, start.Z + outer_radius) + side_protect * toward.Direction) + (64 / 304.8) * toward.Direction;
            XYZ p20 = ((new XYZ(start.X, start.Y, start.Z + inner_radius) + side_protect * toward.Direction) + (64 / 304.8) * toward.Direction) + new XYZ(0, 0, 50 / 304.8);
            XYZ p21 = p19 + (100 / 304.8) * toward.Direction;
            XYZ p22 = p20 + (100 / 304.8) * toward.Direction;
            XYZ p23 = ((new XYZ(start.X, start.Y, start.Z + inner_radius) + side_protect * toward.Direction) + (114) / 304.8 * toward.Direction);

            Curve c16 = Line.CreateBound(p19, p20);
            Curve c17 = Line.CreateBound(p22, p21);
            Arc c18 = Arc.Create(p20, p22, p23);
            L5.Add(c16);
            L5.Add(c18);
            L5.Add(c17);

            XYZ p24 = p2 - 65 / 304.8 * toward.Direction;
            XYZ p25 = new XYZ(end.X, end.Y, end.Z + inner_radius) - side_protect * toward.Direction - 65 / 304.8 * toward.Direction + new XYZ(0, 0, 50 / 304.8);
            XYZ p26 = p24 - 100 / 304.8 * toward.Direction;
            XYZ p27 = p25 - 100 / 304.8 * toward.Direction;
            XYZ p28 = ((new XYZ(end.X, end.Y, end.Z + inner_radius) - side_protect * toward.Direction) - (115) / 304.8 * toward.Direction);

            Curve c19 = Line.CreateBound(p24, p25);
            Curve c20 = Line.CreateBound(p27, p26);
            Arc c21 = Arc.Create(p25, p27, p28);
            L6.Add(c19);
            L6.Add(c21);
            L6.Add(c20);

            Plane plane = Plane.CreateByThreePoints(p8, p2, p3);

            Rebar R1 = Rebar.CreateFromCurvesAndShape(doc, M_00, barType13M, null, null, ele, plane.Normal, L1
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R2 = Rebar.CreateFromCurvesAndShape(doc, M_T1, barType13M, hookType, hookType, ele, plane.Normal, L2
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R4 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L4
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3_2 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R4_2 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L4
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R5 = Rebar.CreateFromCurvesAndShape(doc, M_S5, barType13M, null, null, ele,
                plane.Normal, L5, RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R6 = Rebar.CreateFromCurvesAndShape(doc, M_S5, barType13M, null, null, ele,
                plane.Normal, L6, RebarHookOrientation.Right, RebarHookOrientation.Right);

            //R1.SetSolidInView(view3D, true);
            //R1.SetUnobscuredInView(view3D, true);
            //R2.SetSolidInView(view3D, true);
            //R2.SetUnobscuredInView(view3D, true);
            //R3.SetSolidInView(view3D, true);
            //R3.SetUnobscuredInView(view3D, true);
            //R4.SetSolidInView(view3D, true);
            //R4.SetUnobscuredInView(view3D, true);
            //R3_2.SetSolidInView(view3D, true);
            //R3_2.SetUnobscuredInView(view3D, true);
            //R4_2.SetSolidInView(view3D, true);
            //R4_2.SetUnobscuredInView(view3D, true);
            //R5.SetSolidInView(view3D, true);
            //R5.SetUnobscuredInView(view3D, true);
            //R6.SetSolidInView(view3D, true);
            //R6.SetUnobscuredInView(view3D, true);
            //R1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R4.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R5.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R6.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R4_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);

            // 培文改：設置鋼筋
            List<Rebar> rebars = new List<Rebar>() { R1, R2, R3, R4, R5, R6, R3_2, R4_2 };
            foreach (Rebar rebar in rebars) { SetRebar(rebar, view3D, ABK); }

            ICollection<ElementId> id = new List<ElementId>();
            id.Add(R1.Id);
            id.Add(R2.Id);
            id.Add(R3.Id);
            id.Add(R4.Id);
            id.Add(R3_2.Id);
            id.Add(R4_2.Id);
            id.Add(R5.Id);
            id.Add(R6.Id);
            R5.Location.Rotate(Line.CreateBound(p23, p23 + XYZ.BasisZ), Math.PI / 2);
            R6.Location.Rotate(Line.CreateBound(p28, p28 + XYZ.BasisZ), Math.PI / 2);
            ElementTransformUtils.RotateElement(doc, R3.Id, toward, 60.0 / radius);
            ElementTransformUtils.RotateElement(doc, R3_2.Id, toward, -60.0 / radius);
            ElementTransformUtils.RotateElement(doc, R4.Id, toward, 60.0 / radius);
            ElementTransformUtils.RotateElement(doc, R4_2.Id, toward, -60.0 / radius);
            ElementTransformUtils.RotateElements(doc, id, toward, seg_angle + dis / radius);//Math.PI / 2.0 +
        }
        void shear_t6(readfile rf, Document doc, Element ele, View3D view3D, double dis, double seg_angle, int index, string ABK)
        {
            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (rf.properties.thickness - rf.rebar.main_outer_protect)) / 304.8;
            double side_protect = rf.rebar.main_side_protect / 304.8;

            RebarShape shape = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_T1").First();
            RebarShape shape2 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_S4").First();
            RebarShape M_00 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_00").First();
            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarHookType hookType = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍地震 - 135 度").First();
            RebarHookType hookType2 = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍 - 90 度").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);
            XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
            XYZ start = rf.data_list_tunnel[0].start_point;
            XYZ end = rf.data_list_tunnel[1].start_point;

            XYZ p1 = new XYZ(start.X, start.Y, start.Z + inner_radius) + side_protect * toward.Direction;
            XYZ p2 = new XYZ(end.X, end.Y, end.Z + inner_radius) - side_protect * toward.Direction;
            XYZ p3 = new XYZ(end.X, end.Y, end.Z + outer_radius) - side_protect * toward.Direction;
            XYZ p4 = new XYZ(start.X, start.Y, start.Z + outer_radius) + side_protect * toward.Direction;
            XYZ p5 = p4 + (150 / 304.8) * toward.Direction;
            XYZ p6 = (p1 + (150 / 304.8) * toward.Direction) + new XYZ(0, 0, 60 / 304.8);
            XYZ p7 = p5 + (120 / 304.8) * toward.Direction;
            XYZ p8 = p6 + (120 / 304.8) * toward.Direction;
            XYZ p9 = p3 - ((150 + 120) / 304.8) * toward.Direction;
            XYZ p10 = (p2 - ((150 + 120) / 304.8) * toward.Direction).Add(new XYZ(0, 0, 60 / 304.8));
            XYZ p11 = p3 - (150 / 304.8) * toward.Direction;
            XYZ p12 = (p2 - (150 / 304.8) * toward.Direction).Add(new XYZ(0, 0, 60 / 304.8));

            Plane plane = Plane.CreateByThreePoints(p1, p2, p3);
            SketchPlane skp = SketchPlane.Create(doc, plane);

            IList<Curve> curves = new List<Curve>();
            IList<Curve> curves_s = new List<Curve>();
            IList<Curve> curves_s2 = new List<Curve>();
            IList<Curve> curves_s3 = new List<Curve>();
            IList<Curve> curves_r1 = new List<Curve>();
            IList<Curve> curves_r2 = new List<Curve>();

            Curve c = Line.CreateBound(p1, p2);
            Curve c2 = Line.CreateBound(p2, p3);
            Curve c3 = Line.CreateBound(p3, p4);
            Curve c4 = Line.CreateBound(p4, p1);
            curves.Add(c3);
            curves.Add(c4);
            curves.Add(c);
            curves.Add(c2);

            Curve c5 = Line.CreateBound(p5, p6);
            Curve c6 = Line.CreateBound(p8, p7);
            Arc c7 = Arc.Create(p6, p8, (p1 + (210) / 304.8 * toward.Direction));
            curves_s.Add(c5);
            curves_s.Add(c7);
            curves_s.Add(c6);

            Curve c8 = Line.CreateBound(p9, p10);
            Curve c9 = Line.CreateBound(p12, p11);
            Curve c10 = Arc.Create(p10, p12, (p2 - (210) / 304.8 * toward.Direction));

            curves_s2.Add(c8);
            curves_s2.Add(c10);
            curves_s2.Add(c9);
            p5 = (start + end) / 2 + (150 / 304.8) * toward.Direction;
            p8 = p5.Add(new XYZ(0, 0, inner_radius));
            p5 = p5.Add(new XYZ(0, 0, outer_radius));
            p7 = (start + end) / 2 - (150 / 304.8) * toward.Direction;
            p6 = p7.Add(new XYZ(0, 0, outer_radius));
            p7 = p7.Add(new XYZ(0, 0, inner_radius));
            c5 = Line.CreateBound(p5, p6);
            Curve cl6 = Line.CreateBound(p6, p7);
            Curve cl7 = Line.CreateBound(p7, p8);
            c8 = Line.CreateBound(p8, p5);

            curves_s3.Add(c5);
            curves_s3.Add(cl6);
            curves_s3.Add(cl7);
            curves_s3.Add(c8);

            //旁邊的補強筋
            XYZ rp1 = start + new XYZ(0, 0, inner_radius) + side_protect * toward.Direction;
            XYZ rp2 = rp1 + 400.0 / 304.8 * toward.Direction;
            XYZ rp3 = end + new XYZ(0, 0, inner_radius) - side_protect * toward.Direction;
            XYZ rp4 = rp3 - 400.0 / 304.8 * toward.Direction;
            Curve crp1 = Line.CreateBound(rp1, rp2);
            Curve crp2 = Line.CreateBound(rp3, rp4);

            curves_r1.Add(crp1);
            curves_r2.Add(crp2);

            Rebar re = Rebar.CreateFromCurvesAndShape(
                doc, shape, barType13M, hookType, hookType, ele, plane.Normal, curves,
                RebarHookOrientation.Left, RebarHookOrientation.Left);
            Rebar re2 = Rebar.CreateFromCurvesAndShape(
                doc, shape2, barType13M, hookType2, hookType2, ele, plane.Normal, curves_s,
                RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar re3 = Rebar.CreateFromCurvesAndShape(
                doc, shape2, barType13M, hookType2, hookType2, ele, plane.Normal, curves_s2,
                RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar re4 = Rebar.CreateFromCurvesAndShape(
               doc, shape, barType13M, hookType, hookType, ele, plane.Normal, curves_s3,
               RebarHookOrientation.Left, RebarHookOrientation.Left);
            //補強筋
            Rebar re5 = Rebar.CreateFromCurvesAndShape(
                doc, M_00, barType13M, null, null, ele, plane.Normal, curves_r1, RebarHookOrientation.Left, RebarHookOrientation.Left);
            Rebar re6 = Rebar.CreateFromCurvesAndShape(
                doc, M_00, barType13M, null, null, ele, plane.Normal, curves_r2, RebarHookOrientation.Left, RebarHookOrientation.Left);

            //re.SetSolidInView(view3D, true);
            //re.SetUnobscuredInView(view3D, true);
            //re2.SetSolidInView(view3D, true);
            //re2.SetUnobscuredInView(view3D, true);
            //re3.SetSolidInView(view3D, true);
            //re3.SetUnobscuredInView(view3D, true);
            //re4.SetSolidInView(view3D, true);
            //re4.SetUnobscuredInView(view3D, true);
            //re5.SetSolidInView(view3D, true);
            //re5.SetUnobscuredInView(view3D, true);
            //re6.SetSolidInView(view3D, true);
            //re6.SetUnobscuredInView(view3D, true);
            //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re4.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re5.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //re6.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);

            // 培文改：設置鋼筋
            List<Rebar> rebars = new List<Rebar>() { re, re2, re3, re4, re5, re6 };
            foreach(Rebar rebar in rebars) { SetRebar(rebar, view3D, ABK); }

            ICollection<ElementId> id = new List<ElementId>();
            id.Add(re.Id);
            id.Add(re2.Id);
            id.Add(re3.Id);
            id.Add(re4.Id);
            id.Add(re5.Id);
            id.Add(re6.Id);
            double steel_rad = re.LookupParameter("鋼筋直徑").AsDouble() * 304.8;
            ElementTransformUtils.RotateElement(doc, re2.Id, toward, 2 * (steel_rad) * index / radius);
            ElementTransformUtils.RotateElement(doc, re3.Id, toward, (-steel_rad) * index / radius);
            ElementTransformUtils.RotateElement(doc, re4.Id, toward, (steel_rad) * index / radius);
            //ElementTransformUtils.RotateElement(doc, re5.Id, toward, (-steel_rad) / radius);
            if (index == 1)
            {
                ElementTransformUtils.RotateElement(doc, re6.Id, toward, (-steel_rad) * index / radius);
            }
            else if (index == -1)
            {
                ElementTransformUtils.RotateElement(doc, re6.Id, toward, 3 * (-steel_rad) * index / radius);
                ElementTransformUtils.RotateElement(doc, re5.Id, toward, 2 * (-steel_rad) * index / radius);
            }
            ElementTransformUtils.RotateElements(doc, id, toward, seg_angle + dis / radius);
        }
        void shear_t7(readfile rf, Document doc, Element ele, View3D view3D, double dis, double seg_angle, string ABK)
        {
            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (rf.properties.thickness - rf.rebar.main_outer_protect)) / 304.8;
            double side_protect = rf.rebar.main_side_protect / 304.8;

            RebarShape M_00 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_00").First();
            RebarShape M_T1 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_T1").First();
            RebarShape M_03 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_03").First();
            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarBarType barType16M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "16M").First();
            RebarHookType hookType = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍地震 - 135 度").First();
            RebarHookType hookType2 = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍 - 90 度").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);
            XYZ start = rf.data_list_tunnel[0].start_point;
            XYZ end = rf.data_list_tunnel[1].start_point;

            IList<Curve> L1 = new List<Curve>();
            IList<Curve> L2 = new List<Curve>();
            IList<Curve> L3 = new List<Curve>();
            IList<Curve> L4 = new List<Curve>();

            XYZ p1 = new XYZ(start.X, start.Y, start.Z + outer_radius) + side_protect * toward.Direction;
            XYZ p2 = new XYZ(end.X, end.Y, end.Z + outer_radius) - side_protect * toward.Direction;
            Curve c1 = Line.CreateBound(p1, p2);

            L1.Add(c1);

            XYZ p3 = new XYZ(start.X, start.Y, start.Z + outer_radius) + 351.6 / 304.8 * toward.Direction;
            XYZ p4 = new XYZ(end.X, end.Y, end.Z + outer_radius) - 350.6 / 304.8 * toward.Direction;
            XYZ p5 = new XYZ(p4.X, p4.Y, p4.Z - 185.1 / 304.8);
            XYZ p6 = new XYZ(p3.X, p3.Y, p3.Z - 185.1 / 304.8);
            Curve c2 = Line.CreateBound(p3, p4);
            Curve c3 = Line.CreateBound(p4, p5);
            Curve c4 = Line.CreateBound(p5, p6);
            Curve c5 = Line.CreateBound(p6, p3);

            L2.Add(c2);
            L2.Add(c3);
            L2.Add(c4);
            L2.Add(c5);

            XYZ p7 = new XYZ(start.X, start.Y, start.Z + outer_radius - 17.8 / 304.8) + 220.7 / 304.8 * toward.Direction;
            XYZ p8 = new XYZ(p1.X, p1.Y, p1.Z - 17.8 / 304.8);
            XYZ p9 = p8 - new XYZ(0, 0, 167.2 / 304.8);
            XYZ p10 = p9 + 120.0 / 304.8 * toward.Direction;
            XYZ p11 = p10 + new XYZ(0, 0, 149.4 / 304.8) + 49.8 * (149.4 / 136.0) / 304.8 * toward.Direction;
            XYZ p17 = start + 146.0 / 304.8 * toward.Direction - new XYZ(0, 0, 35.6 / 304.8 - outer_radius);
            Curve c6 = Line.CreateBound(p7, p8);
            Curve c7 = Line.CreateBound(p8, p9);
            Curve c8 = Line.CreateBound(p9, p10);
            Curve c9 = Line.CreateBound(p10, p11);
            Curve c14 = Line.CreateBound(p11, p17);

            L3.Add(c6);
            L3.Add(c7);
            L3.Add(c8);
            L3.Add(c9);
            L3.Add(c14);

            XYZ p12 = new XYZ(end.X, end.Y, end.Z + outer_radius - 17.8 / 304.8) - 220.7 / 304.8 * toward.Direction;
            XYZ p13 = new XYZ(p2.X, p2.Y, p2.Z - 17.8 / 304.8);
            XYZ p14 = p13 - new XYZ(0, 0, 167.2 / 304.8);
            XYZ p15 = p14 - 120.0 / 304.8 * toward.Direction;
            XYZ p16 = p15 + new XYZ(0, 0, 149.4 / 304.8) - 49.8 * (149.4 / 136.0) / 304.8 * toward.Direction;
            XYZ p18 = end - 146.0 / 304.8 * toward.Direction - new XYZ(0, 0, 35.6 / 304.8 - outer_radius);
            Curve c10 = Line.CreateBound(p12, p13);
            Curve c11 = Line.CreateBound(p13, p14);
            Curve c12 = Line.CreateBound(p14, p15);
            Curve c13 = Line.CreateBound(p15, p16);
            Curve c15 = Line.CreateBound(p16, p18);

            L4.Add(c10);
            L4.Add(c11);
            L4.Add(c12);
            L4.Add(c13);
            L4.Add(c15);

            Plane plane = Plane.CreateByThreePoints(p8, p2, p3);

            Rebar R1 = Rebar.CreateFromCurvesAndShape(doc, M_00, barType13M, null, null, ele, plane.Normal, L1
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R2 = Rebar.CreateFromCurvesAndShape(doc, M_T1, barType13M, hookType, hookType, ele, plane.Normal, L2
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R2a = Rebar.CreateFromCurvesAndShape(doc, M_T1, barType13M, hookType, hookType, ele, plane.Normal, L2
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3a = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R4 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L4
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R4a = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L4
                , RebarHookOrientation.Right, RebarHookOrientation.Right);

            //R1.SetSolidInView(view3D, true);
            //R1.SetUnobscuredInView(view3D, true);
            //R2.SetSolidInView(view3D, true);
            //R2.SetUnobscuredInView(view3D, true);
            //R2a.SetSolidInView(view3D, true);
            //R2a.SetUnobscuredInView(view3D, true);
            //R3.SetSolidInView(view3D, true);
            //R3.SetUnobscuredInView(view3D, true);
            //R4.SetSolidInView(view3D, true);
            //R4.SetUnobscuredInView(view3D, true);
            //R3a.SetSolidInView(view3D, true);
            //R3a.SetUnobscuredInView(view3D, true);
            //R4a.SetSolidInView(view3D, true);
            //R4a.SetUnobscuredInView(view3D, true);
            //R1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R4.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R2a.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3a.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R4a.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);

            // 培文改：設置鋼筋
            List<Rebar> rebars = new List<Rebar>() { R1, R2, R2a, R3, R3a, R4, R4a };
            foreach (Rebar rebar in rebars) { SetRebar(rebar, view3D, ABK); }

            ICollection<ElementId> id = new List<ElementId>();
            id.Add(R1.Id);
            id.Add(R2.Id);
            id.Add(R2a.Id);
            id.Add(R3.Id);
            id.Add(R4.Id);
            id.Add(R3a.Id);
            id.Add(R4a.Id);
            double steel_rad = R1.LookupParameter("鋼筋直徑").AsDouble() * 304.8;
            ElementTransformUtils.RotateElement(doc, R2.Id, toward, 40 / (inner_radius * 304.8));
            ElementTransformUtils.RotateElement(doc, R2a.Id, toward, -30.0 / (inner_radius * 304.8));
            ElementTransformUtils.RotateElement(doc, R3.Id, toward, 90.0 / (inner_radius * 304.8));
            ElementTransformUtils.RotateElement(doc, R3a.Id, toward, -30.0 / (inner_radius * 304.8));
            ElementTransformUtils.RotateElement(doc, R4.Id, toward, 45.0 / (inner_radius * 304.8));
            ElementTransformUtils.RotateElement(doc, R4a.Id, toward, -55.0 / (inner_radius * 304.8));
            ElementTransformUtils.RotateElements(doc, id, toward, seg_angle + dis / radius);//Math.PI / 2.0 +
        }
        void shear_t9(readfile rf, Document doc, Element ele, View3D view3D, double dis, double seg_angle, int index, string ABK)
        {
            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (rf.properties.thickness - rf.rebar.main_outer_protect)) / 304.8;
            double side_protect = rf.rebar.main_side_protect / 304.8;

            RebarShape M_00 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_00").First();
            RebarShape M_T1 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_T1").First();
            RebarShape M_03 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_03").First();
            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarBarType barType16M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "16M").First();
            RebarHookType hookType = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍地震 - 135 度").First();
            RebarHookType hookType2 = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍 - 90 度").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);
            XYZ start = rf.data_list_tunnel[0].start_point;
            XYZ end = rf.data_list_tunnel[1].start_point;

            IList<Curve> L1 = new List<Curve>();
            IList<Curve> L2 = new List<Curve>();
            IList<Curve> L3 = new List<Curve>();
            IList<Curve> L4 = new List<Curve>();

            XYZ p1 = new XYZ(start.X, start.Y, start.Z + outer_radius) + side_protect * toward.Direction;
            XYZ p2 = new XYZ(end.X, end.Y, end.Z + outer_radius) - side_protect * toward.Direction;
            Curve c1 = Line.CreateBound(p1, p2);

            L1.Add(c1);

            XYZ p3 = new XYZ(start.X, start.Y, start.Z + outer_radius) + 351.6 / 304.8 * toward.Direction;
            XYZ p4 = new XYZ(end.X, end.Y, end.Z + outer_radius) - 350.6 / 304.8 * toward.Direction;
            XYZ p5 = new XYZ(p4.X, p4.Y, p4.Z - 185.1 / 304.8);
            XYZ p6 = new XYZ(p3.X, p3.Y, p3.Z - 185.1 / 304.8);
            Curve c2 = Line.CreateBound(p3, p4);
            Curve c3 = Line.CreateBound(p4, p5);
            Curve c4 = Line.CreateBound(p5, p6);
            Curve c5 = Line.CreateBound(p6, p3);

            L2.Add(c2);
            L2.Add(c3);
            L2.Add(c4);
            L2.Add(c5);

            XYZ p7 = new XYZ(start.X, start.Y, start.Z + outer_radius - 17.8 / 304.8) + 220.7 / 304.8 * toward.Direction;
            XYZ p8 = new XYZ(p1.X, p1.Y, p1.Z - 17.8 / 304.8);
            XYZ p9 = p8 - new XYZ(0, 0, 167.2 / 304.8);
            XYZ p10 = p9 + 120.0 / 304.8 * toward.Direction;
            XYZ p11 = p10 + new XYZ(0, 0, 149.4 / 304.8) + 49.8 * (149.4 / 136.0) / 304.8 * toward.Direction;
            XYZ p17 = start + 146.0 / 304.8 * toward.Direction - new XYZ(0, 0, 35.6 / 304.8 - outer_radius);
            Curve c6 = Line.CreateBound(p7, p8);
            Curve c7 = Line.CreateBound(p8, p9);
            Curve c8 = Line.CreateBound(p9, p10);
            Curve c9 = Line.CreateBound(p10, p11);
            Curve c14 = Line.CreateBound(p11, p17);

            L3.Add(c6);
            L3.Add(c7);
            L3.Add(c8);
            L3.Add(c9);
            L3.Add(c14);

            XYZ p12 = new XYZ(end.X, end.Y, end.Z + outer_radius - 17.8 / 304.8) - 220.7 / 304.8 * toward.Direction;
            XYZ p13 = new XYZ(p2.X, p2.Y, p2.Z - 17.8 / 304.8);
            XYZ p14 = p13 - new XYZ(0, 0, 167.2 / 304.8);
            XYZ p15 = p14 - 120.0 / 304.8 * toward.Direction;
            XYZ p16 = p15 + new XYZ(0, 0, 149.4 / 304.8) - 49.8 * (149.4 / 136.0) / 304.8 * toward.Direction;
            XYZ p18 = end - 146.0 / 304.8 * toward.Direction - new XYZ(0, 0, 35.6 / 304.8 - outer_radius);
            Curve c10 = Line.CreateBound(p12, p13);
            Curve c11 = Line.CreateBound(p13, p14);
            Curve c12 = Line.CreateBound(p14, p15);
            Curve c13 = Line.CreateBound(p15, p16);
            Curve c15 = Line.CreateBound(p16, p18);

            L4.Add(c10);
            L4.Add(c11);
            L4.Add(c12);
            L4.Add(c13);
            L4.Add(c15);

            Plane plane = Plane.CreateByThreePoints(p8, p2, p3);

            Rebar R1 = Rebar.CreateFromCurvesAndShape(doc, M_00, barType13M, null, null, ele, plane.Normal, L1
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3_2 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R3_3 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L3
                , RebarHookOrientation.Right, RebarHookOrientation.Right);
            Rebar R4 = Rebar.CreateFromCurvesAndShape(doc, M_03, barType16M, null, null, ele, plane.Normal, L4
                , RebarHookOrientation.Right, RebarHookOrientation.Right);

            //R1.SetSolidInView(view3D, true);
            //R1.SetUnobscuredInView(view3D, true);
            //R3.SetSolidInView(view3D, true);
            //R3.SetUnobscuredInView(view3D, true);
            //R3_2.SetSolidInView(view3D, true);
            //R3_2.SetUnobscuredInView(view3D, true);
            //R3_3.SetSolidInView(view3D, true);
            //R3_3.SetUnobscuredInView(view3D, true);
            //R4.SetSolidInView(view3D, true);
            //R4.SetUnobscuredInView(view3D, true);
            //R1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R3_3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);
            //R4.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(ABK);

            // 培文改：設置鋼筋
            List<Rebar> rebars = new List<Rebar>() { R1, R3, R3_2, R3_3, R4 };
            foreach (Rebar rebar in rebars) { SetRebar(rebar, view3D, ABK); }

            ICollection<ElementId> id = new List<ElementId>();
            id.Add(R1.Id);
            id.Add(R3.Id);
            id.Add(R3_2.Id);
            id.Add(R3_3.Id);
            id.Add(R4.Id);
            double steel_rad = R1.LookupParameter("鋼筋直徑").AsDouble() * 304.8;
            if (index == 1)
            {
                ElementTransformUtils.RotateElement(doc, R4.Id, toward, (-steel_rad) * 2 / radius);
                ElementTransformUtils.RotateElement(doc, R3_3.Id, toward, (-steel_rad) * 2 / radius);
                ElementTransformUtils.RotateElement(doc, R3_2.Id, toward, (-steel_rad - 90.0) / radius);
            }
            else if (index == -1)
            {
                ElementTransformUtils.RotateElement(doc, R3_3.Id, toward, (-steel_rad) * 2 / radius);
                ElementTransformUtils.RotateElement(doc, R3_2.Id, toward, (-steel_rad + 90.0) / radius);
            }
            ElementTransformUtils.RotateElements(doc, id, toward, seg_angle + dis / radius);//Math.PI / 2.0 +
        }
        // 設置鋼筋
        private void SetRebar(Rebar re, View3D view3D, string value)
        {
            //re.SetSolidInView(view3D, true); // 2020
            re.SetUnobscuredInView(view3D, true);
            re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(value);
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
    }
}
