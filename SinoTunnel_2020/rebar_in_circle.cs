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

namespace SinoTunnel_2020
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class rebar_in_circle : IExternalEventHandler
    {
        public void Execute(UIApplication uiapp)
        {
            string path = Form1.path;
            readfile rf = new readfile();
            rf.read_tunnel_point();
            rf.read_properties();
            rf.read_rebar_properties();
            double radius = rf.properties.inner_diameter / 2;
            double inner_radius = (radius + rf.rebar.main_inner_protect) / 304.8;
            double outer_radius = (radius + (250 - rf.rebar.main_outer_protect)) / 304.8;

            Document doc = uiapp.ActiveUIDocument.Document;
            View3D view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().ToList().First();

            Element element = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList()
                .Where(x => x.Name == "自適應環形00_A" && x.LookupParameter("備註").AsString() == $"{rf.firstStation}").First();

            Line toward = Line.CreateBound(rf.data_list_tunnel[0].start_point, rf.data_list_tunnel[1].start_point);

            Transaction t = new Transaction(doc);
            SubTransaction subt = new SubTransaction(doc);
            t.Start("Main Rebar");

            //rebar type&shape
            RebarBarType barType29M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "29M").First();
            RebarBarType barType19M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "19M").First();
            RebarBarType barType13M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "13M").First();
            RebarBarType barType25M = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList().Where(x => x.Name == "25M").First();
            RebarShape M_00 = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "M_00").First();
            RebarShape shape = new FilteredElementCollector(doc).OfClass(typeof(RebarShape)).Cast<RebarShape>().ToList().Where(x => x.Name == "主筋造型").First();

            XYZ mid = rf.data_list_tunnel[0].start_point;
            double side_protect = rf.rebar.main_side_protect;
            double dis_A = 0;
            double dis_B = 0, dis_K = 0;
            for (int i = 0; i < 8; i++)
            {
                dis_A += double.Parse(rf.rebar.main_A_distance[i]);
                dis_B += double.Parse(rf.rebar.main_B_distance[i]);
                dis_K += double.Parse(rf.rebar.main_K_distance[i]);

                //current center point
                XYZ A_center = mid + toward.Direction.Normalize() * (dis_A / 304.8);
                XYZ B_center = mid + toward.Direction.Normalize() * (dis_B / 304.8);
                XYZ K_center = mid + toward.Direction.Normalize() * (dis_K / 304.8);


                //A環片主筋的ARC
                for (int j = 0; j < 3; j++)
                {
                    for (int k = 0; k < 2; k++)
                    {

                        //make plane
                        XYZ planepoint2 = new XYZ(A_center.X - toward.Direction.Y, A_center.Y + toward.Direction.X, A_center.Z);
                        XYZ planepoint3 = A_center + XYZ.BasisZ;
                        Plane plane = Plane.CreateByThreePoints(A_center, planepoint2, planepoint3);
                        RebarBarType barType = (i % 2 == 0) ? barType29M : barType25M;
                        IList<Curve> curves = new List<Curve>();
                        //A2, A2a
                        //double barDiameter = barType13M.BarDiameter + barType.BarDiameter; // 2020
                        double barDiameter = barType13M.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble() + barType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();
                        double Rebar_r = (k == 0) ? (inner_radius + (barDiameter) / 2) : (outer_radius - (barDiameter) / 2); // 培文改
                        //double Rebar_r = (k == 0) ? (inner_radius + (barType13M.BarDiameter + barType.BarDiameter) / 2) : (outer_radius - (barType13M.BarDiameter + barType.BarDiameter) / 2); // 台大
                        Curve c = Arc.Create(plane, Rebar_r, ((-36.0 - 72.0 * j) / 180) * Math.PI + side_protect / radius, ((36.0 - 72.0 * j) / 180) * Math.PI - side_protect / radius);
                        curves.Add(c);

                        subt.Start();


                        XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
                        Rebar re = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                        //re.SetSolidInView(view3D as View3D, true);
                        //re.SetUnobscuredInView(view3D, true);
                        //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("A");

                        SetRebar(doc, re, view3D, "A");

                        subt.Commit();

                    }
                }

                //B環片主筋的ARC

                for (int k = 0; k < 2; k++)
                {
                    //make plane
                    XYZ planepoint2 = new XYZ(B_center.X - toward.Direction.Y, B_center.Y + toward.Direction.X, B_center.Z);
                    XYZ planepoint3 = B_center + XYZ.BasisZ;
                    Plane plane = Plane.CreateByThreePoints(B_center, planepoint2, planepoint3);
                    RebarBarType barType = (i % 2 == 0) ? barType29M : barType25M;
                    IList<Curve> curves = new List<Curve>();
                    IList<Curve> curves2 = new List<Curve>();
                    //B2, B2a
                    //double barDiameter = barType13M.BarDiameter + barType.BarDiameter; // 2020
                    double barDiameter = barType13M.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble() + barType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();
                    double Rebar_r = (k == 0) ? (inner_radius + (barDiameter) / 2) : (outer_radius - (barDiameter) / 2); // 培文改
                    //double Rebar_r = (k == 0) ? (inner_radius + (barType13M.BarDiameter + barType.BarDiameter) / 2) : (outer_radius - (barType13M.BarDiameter + barType.BarDiameter) / 2); // 台大

                    Curve c = Arc.Create(plane, Rebar_r, ((36.0) / 180) * Math.PI + side_protect / radius, ((36.0 + 64.5) / 180) * Math.PI - (side_protect - (100.0) * (dis_B / 1000.0)) / radius);
                    Curve c2 = Arc.Create(plane, Rebar_r, ((180.0 - 64.5) / 180) * Math.PI + (side_protect - (100.0) * (dis_B / 1000.0)) / radius, ((180.0) / 180) * Math.PI - side_protect / radius);
                    curves.Add(c);
                    curves2.Add(c2);

                    subt.Start();


                    XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
                    Rebar re = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                    Rebar re2 = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves2, RebarHookOrientation.Left, RebarHookOrientation.Left);

                    //re.SetSolidInView(view3D as View3D, true);
                    //re.SetUnobscuredInView(view3D, true);
                    //re2.SetSolidInView(view3D as View3D, true);
                    //re2.SetUnobscuredInView(view3D, true);
                    //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");
                    //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");

                    // 培文改：設置鋼筋
                    List<Rebar> rebars = new List<Rebar>() { re, re2 };
                    foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "B"); }

                    subt.Commit();

                }

                //K環片主筋的ARC
                for (int k = 0; k < 2; k++)
                {
                    //make plane
                    XYZ planepoint2 = new XYZ(K_center.X - toward.Direction.Y, K_center.Y + toward.Direction.X, K_center.Z);
                    XYZ planepoint3 = K_center + XYZ.BasisZ;
                    Plane plane = Plane.CreateByThreePoints(K_center, planepoint2, planepoint3);
                    RebarBarType barType = (i % 2 == 0) ? barType29M : barType25M;
                    IList<Curve> curves = new List<Curve>();
                    //K2, K2a
                    //double barDiameter = barType13M.BarDiameter + barType.BarDiameter; // 2020
                    double barDiameter = barType13M.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble() + barType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();
                    double Rebar_r = (k == 0) ? (inner_radius + (barDiameter) / 2) : (outer_radius - (barDiameter) / 2); // 培文改
                    //double Rebar_r = (k == 0) ? (inner_radius + (barType13M.BarDiameter + barType.BarDiameter) / 2) : (outer_radius - (barType13M.BarDiameter + barType.BarDiameter) / 2); // 台大
                    Curve c = Arc.Create(plane, Rebar_r, (((36.0 + 64.5) / 180) * Math.PI + (side_protect + (100.0) * (dis_K / 1000.0)) / radius)
                        , ((180.0 - 64.5) / 180) * Math.PI - (side_protect + (100.0) * (dis_K / 1000.0)) / radius);
                    curves.Add(c);

                    subt.Start();


                    XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
                    Rebar re = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                    //re.SetSolidInView(view3D as View3D, true);
                    //re.SetUnobscuredInView(view3D, true);
                    //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("K");

                    // 培文改：設置鋼筋
                    List<Rebar> rebars = new List<Rebar>() { re };
                    foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "K"); }

                    subt.Commit();
                }
            }

            //換補強筋
            //A環片補強筋(主筋方向)
            dis_A = 0;
            for (int i = 0; i < 2; i++)
            {
                double ba = (i == 0) ? 369.1 : 261.9;
                dis_A += ba;

                for (int j = 0; j < 3; j++)
                {
                    XYZ A_center = mid + toward.Direction.Normalize() * (dis_A / 304.8);
                    //make plane
                    XYZ planepoint2 = new XYZ(A_center.X - toward.Direction.Y, A_center.Y + toward.Direction.X, A_center.Z);
                    XYZ planepoint3 = A_center + XYZ.BasisZ;
                    Plane plane = Plane.CreateByThreePoints(A_center, planepoint2, planepoint3);
                    RebarBarType barType = barType19M;
                    IList<Curve> curves = new List<Curve>();
                    IList<Curve> curves2 = new List<Curve>();
                    //A2, A2a
                    double rad_inner_r = inner_radius * 304.8;
                    // 培文改
                    //double barDiameter = barType13M.BarDiameter + barType.BarDiameter; // 2020
                    double barDiameter = barType13M.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble() + barType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();
                    Curve c = Arc.Create(plane, (inner_radius + (barDiameter) / 2)
                        , ((-36.0 - 72.0 * j) / 180) * Math.PI + side_protect / rad_inner_r,
                         ((-36.0 - 72.0 * j) / 180) * Math.PI - (side_protect - 1056) / rad_inner_r);
                    Curve c2 = Arc.Create(plane, (inner_radius + (barDiameter) / 2),
                        ((36.0 - 72.0 * j) / 180) * Math.PI + (side_protect - 1056) / rad_inner_r,
                        ((36.0 - 72.0 * j) / 180) * Math.PI - side_protect / rad_inner_r);
                    //// 台大
                    //Curve c = Arc.Create(plane, (inner_radius + (barType13M.BarDiameter + barType.BarDiameter) / 2)
                    //    , ((-36.0 - 72.0 * j) / 180) * Math.PI + side_protect / rad_inner_r,
                    //     ((-36.0 - 72.0 * j) / 180) * Math.PI - (side_protect - 1056) / rad_inner_r);
                    //Curve c2 = Arc.Create(plane, (inner_radius + (barType13M.BarDiameter + barType.BarDiameter) / 2),
                    //    ((36.0 - 72.0 * j) / 180) * Math.PI + (side_protect - 1056) / rad_inner_r,
                    //    ((36.0 - 72.0 * j) / 180) * Math.PI - side_protect / rad_inner_r);
                    curves.Add(c);
                    curves2.Add(c2);

                    subt.Start();


                    XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
                    Rebar re = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                    Rebar re2 = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves2, RebarHookOrientation.Left, RebarHookOrientation.Left);
                    //re.SetSolidInView(view3D as View3D, true);
                    //re.SetUnobscuredInView(view3D, true);
                    //re2.SetSolidInView(view3D as View3D, true);
                    //re2.SetUnobscuredInView(view3D, true);
                    //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("A");
                    //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("A");

                    // 培文改：設置鋼筋
                    List<Rebar> rebars = new List<Rebar>() { re, re2 };
                    foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "A"); }

                    subt.Commit();
                }
            }
            dis_A = 0;
            //(剪力筋方向)
            for (int i = 0; i < 4; i++)
            {
                double[] ba = { 73, 320, 400, 191 };
                dis_A += ba[i];

                for (int j = 0; j < 3; j++)
                {
                    XYZ A_center = mid;
                    //make plane
                    XYZ p2 = new XYZ(A_center.X, A_center.Y, A_center.Z + inner_radius) + side_protect / 304.8 * toward.Direction;
                    XYZ p3 = new XYZ(A_center.X, A_center.Y, A_center.Z + inner_radius) + (rf.properties.width - side_protect) / 304.8 * toward.Direction;
                    Plane plane = Plane.CreateByThreePoints(A_center, p2, p3);

                    IList<Curve> curves = new List<Curve>();
                    IList<Curve> curves2 = new List<Curve>();
                    //A2, A2a
                    double Rebar_r = inner_radius;
                    Curve c = Line.CreateBound(p2, p3);
                    //Curve c2 = Arc.Create(plane, Rebar_r, ((36.0 - 72.0 * j) / 180) * Math.PI + (side_protect - 1056) / radius, ((36.0 - 72.0 * j) / 180) * Math.PI - side_protect / radius);
                    curves.Add(c);
                    //curves2.Add(c2);

                    subt.Start();
                    RebarBarType barType = barType13M;
                    RebarShape barshape = M_00;

                    Rebar re = Rebar.CreateFromCurvesAndShape(doc, barshape, barType, null, null, element, plane.Normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                    Rebar re2 = Rebar.CreateFromCurvesAndShape(doc, barshape, barType, null, null, element, plane.Normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                    //re.SetSolidInView(view3D as View3D, true);
                    //re.SetUnobscuredInView(view3D, true);
                    //re2.SetSolidInView(view3D as View3D, true);
                    //re2.SetUnobscuredInView(view3D, true);
                    //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("A");
                    //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("A");

                    // 培文改：設置鋼筋
                    List<Rebar> rebars = new List<Rebar>() { re, re2 };
                    foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "A"); }

                    double rad_inner_r = inner_radius * 304.8;
                    ElementTransformUtils.RotateElement(doc, re.Id, toward, Math.PI / 2.0 + dis_A / rad_inner_r + 72.0 * j * Math.PI / 180.0); //+72.0 * j + dis_A / radius
                    ElementTransformUtils.RotateElement(doc, re2.Id, toward, Math.PI / 2.0 - dis_A / rad_inner_r + 72.0 * (j + 1) * Math.PI / 180.0); //+72.0 * j + dis_A / radius
                    subt.Commit();
                }
            }
            dis_A = 259.0;
            for (int j = 0; j < 3; j++)
            {
                XYZ A_center = mid;
                //make plane
                XYZ p2 = new XYZ(A_center.X, A_center.Y, A_center.Z + outer_radius) + side_protect / 304.8 * toward.Direction;
                XYZ p3 = new XYZ(A_center.X, A_center.Y, A_center.Z + outer_radius) + (rf.properties.width - side_protect) / 304.8 * toward.Direction;
                Plane plane = Plane.CreateByThreePoints(A_center, p2, p3);

                IList<Curve> curves = new List<Curve>();
                IList<Curve> curves2 = new List<Curve>();
                //A2, A2a
                Curve c = Line.CreateBound(p2, p3);
                curves.Add(c);

                subt.Start();

                RebarBarType barType = barType13M;
                RebarShape barshape = M_00;

                Rebar re = Rebar.CreateFromCurvesAndShape(doc, barshape, barType, null, null, element, plane.Normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                Rebar re2 = Rebar.CreateFromCurvesAndShape(doc, barshape, barType, null, null, element, plane.Normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                //re.SetSolidInView(view3D as View3D, true);
                //re.SetUnobscuredInView(view3D, true);
                //re2.SetSolidInView(view3D as View3D, true);
                //re2.SetUnobscuredInView(view3D, true);
                //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("A");
                //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("A");

                // 培文改：設置鋼筋
                List<Rebar> rebars = new List<Rebar>() { re, re2 };
                foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "A"); }

                double rad_outer_r = outer_radius * 304.8;
                ElementTransformUtils.RotateElement(doc, re.Id, toward, Math.PI / 2.0 + dis_A / rad_outer_r + 72.0 * j * Math.PI / 180.0); //+72.0 * j + dis_A / radius
                ElementTransformUtils.RotateElement(doc, re2.Id, toward, Math.PI / 2.0 - dis_A / rad_outer_r + 72.0 * (j + 1) * Math.PI / 180.0); //+72.0 * j + dis_A / radius
                subt.Commit();
            }


            //B環片補強筋(主筋方向)
            dis_B = 0;
            for (int i = 0; i < 2; i++)
            {
                double ba = (i == 0) ? 368.6 : 261.4;
                dis_B += ba;


                XYZ B_center = mid + toward.Direction.Normalize() * (dis_B / 304.8);
                //make plane
                XYZ planepoint2 = new XYZ(B_center.X - toward.Direction.Y, B_center.Y + toward.Direction.X, B_center.Z);
                XYZ planepoint3 = B_center + XYZ.BasisZ;
                Plane plane = Plane.CreateByThreePoints(B_center, planepoint2, planepoint3);
                RebarBarType barType = barType19M;
                IList<Curve> curves = new List<Curve>();
                IList<Curve> curves2 = new List<Curve>();
                IList<Curve> curves3 = new List<Curve>();
                IList<Curve> curves4 = new List<Curve>();
                //A2, A2a
                double rad_inner_r = inner_radius * 304.8;
                // 培文改
                //double barDiameter = barType13M.BarDiameter + barType.BarDiameter; // 2020
                double barDiameter = barType13M.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble() + barType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();
                Curve c = Arc.Create(plane, (inner_radius + (barDiameter) / 2),
                    (180 / 180) * Math.PI - (side_protect + 1060.1) / rad_inner_r,
                                                     (180 / 180) * Math.PI - side_protect / rad_inner_r);
                int index = (i == 0) ? 1 : -1;
                Curve c2 = Arc.Create(plane, (inner_radius + (barDiameter) / 2),
                    (115.5 / 180) * Math.PI + (side_protect + 2.5 * index) / rad_inner_r,
                                                    (115.5 / 180) * Math.PI + (side_protect + 2.5 * index + 1060.1) / rad_inner_r);
                //// 台大
                //Curve c = Arc.Create(plane, (inner_radius + (barType13M.BarDiameter + barType.BarDiameter) / 2),
                //    (180 / 180) * Math.PI - (side_protect + 1060.1) / rad_inner_r,
                //                                     (180 / 180) * Math.PI - side_protect / rad_inner_r);
                //int index = (i == 0) ? 1 : -1;
                //Curve c2 = Arc.Create(plane, (inner_radius + (barType13M.BarDiameter + barType.BarDiameter) / 2),
                //    (115.5 / 180) * Math.PI + (side_protect + 2.5 * index) / rad_inner_r,
                //                                    (115.5 / 180) * Math.PI + (side_protect + 2.5 * index + 1060.1) / rad_inner_r);
                Curve c3 = Arc.Create(plane, inner_radius, (100.5 / 180) * Math.PI - (side_protect + 2.5 * index + 1060.1) / rad_inner_r,
                                                    (100.5 / 180) * Math.PI - (side_protect + 2.5 * index) / rad_inner_r);

                Curve c4 = Arc.Create(plane, inner_radius, (36.0 / 180) * Math.PI + side_protect / rad_inner_r,
                                                     (36.0 / 180) * Math.PI + (side_protect + 1060.1) / rad_inner_r);
                curves.Add(c);
                curves2.Add(c2);
                curves3.Add(c3);
                curves4.Add(c4);

                subt.Start();

                XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
                Rebar re = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                Rebar re2 = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves2, RebarHookOrientation.Left, RebarHookOrientation.Left);
                Rebar re3 = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves3, RebarHookOrientation.Left, RebarHookOrientation.Left);
                Rebar re4 = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves4, RebarHookOrientation.Left, RebarHookOrientation.Left);
                //re.SetSolidInView(view3D as View3D, true);
                //re.SetUnobscuredInView(view3D, true);
                //re2.SetSolidInView(view3D as View3D, true);
                //re2.SetUnobscuredInView(view3D, true);
                //re3.SetSolidInView(view3D as View3D, true);
                //re3.SetUnobscuredInView(view3D, true);
                //re4.SetSolidInView(view3D as View3D, true);
                //re4.SetUnobscuredInView(view3D, true);
                //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");
                //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");
                //re3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");
                //re4.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");

                // 培文改：設置鋼筋
                List<Rebar> rebars = new List<Rebar>() { re, re2, re3, re4 };
                foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "B"); }

                subt.Commit();
            }
            dis_B = 0;
            //B環片補強筋(剪力筋方向)
            for (int i = 0; i < 3; i++)
            {
                double[] ba = { 351, 61, 201 };
                dis_B += ba[i];

                for (int j = 0; j < 2; j++)
                {
                    XYZ B_center = mid;
                    //make plane
                    XYZ p2 = new XYZ(B_center.X, B_center.Y, B_center.Z + inner_radius) + side_protect / 304.8 * toward.Direction;
                    XYZ p3 = new XYZ(B_center.X, B_center.Y, B_center.Z + inner_radius) + (rf.properties.width - side_protect) / 304.8 * toward.Direction;
                    Plane plane = Plane.CreateByThreePoints(B_center, p2, p3);

                    IList<Curve> curves = new List<Curve>();
                    double Rebar_r = inner_radius;
                    Curve c = Line.CreateBound(p2, p3);
                    curves.Add(c);

                    subt.Start();
                    RebarBarType barType = barType13M;
                    RebarShape barshape = M_00;

                    Rebar re = Rebar.CreateFromCurvesAndShape(doc, barshape, barType, null, null, element, plane.Normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                    //re.SetSolidInView(view3D as View3D, true);
                    //re.SetUnobscuredInView(view3D, true);
                    //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");

                    // 培文改：設置鋼筋
                    List<Rebar> rebars = new List<Rebar>() { re };
                    foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "B"); }

                    double rad_inner_r = inner_radius * 304.8;
                    int index = (j == 0) ? 1 : -1;
                    ElementTransformUtils.RotateElement(doc, re.Id, toward, (18.0 + 7.5 * index) / 180 * Math.PI + dis_B / rad_inner_r * index); //+72.0 * j + dis_A / radius
                    subt.Commit();
                }
            }
            dis_B = 0;
            for (int i = 0; i < 4; i++)
            {
                double[] ba = { 73, 320, 400, 191 };
                dis_B += ba[i];

                for (int j = 0; j < 2; j++)
                {
                    XYZ B_center = mid;
                    //make plane
                    XYZ p2 = new XYZ(B_center.X, B_center.Y, B_center.Z + inner_radius) + side_protect / 304.8 * toward.Direction;
                    XYZ p3 = new XYZ(B_center.X, B_center.Y, B_center.Z + inner_radius) + (rf.properties.width - side_protect) / 304.8 * toward.Direction;
                    Plane plane = Plane.CreateByThreePoints(B_center, p2, p3);

                    IList<Curve> curves = new List<Curve>();
                    IList<Curve> curves2 = new List<Curve>();
                    //A2, A2a
                    double Rebar_r = inner_radius;
                    Curve c = Line.CreateBound(p2, p3);
                    //Curve c2 = Arc.Create(plane, Rebar_r, ((36.0 - 72.0 * j) / 180) * Math.PI + (side_protect - 1056) / radius, ((36.0 - 72.0 * j) / 180) * Math.PI - side_protect / radius);
                    curves.Add(c);
                    //curves2.Add(c2);

                    subt.Start();
                    RebarBarType barType = barType13M;
                    RebarShape barshape = M_00;

                    Rebar re = Rebar.CreateFromCurvesAndShape(doc, barshape, barType, null, null, element, plane.Normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                    //re.SetSolidInView(view3D as View3D, true);
                    //re.SetUnobscuredInView(view3D, true);
                    //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");

                    // 培文改：設置鋼筋
                    List<Rebar> rebars = new List<Rebar>() { re };
                    foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "B"); }

                    double rad_inner_r = inner_radius * 304.8;
                    int index = (j == 0) ? 1 : -1;
                    ElementTransformUtils.RotateElement(doc, re.Id, toward, (18.0 + 72 * index) / 180 * Math.PI - dis_B / rad_inner_r * index); //+72.0 * j + dis_A / radius
                    subt.Commit();
                }
            }
            //外側補強筋
            dis_B = 275.0;
            for (int j = 0; j < 2; j++)
            {
                XYZ B_center = mid;
                //make plane
                XYZ p2 = new XYZ(B_center.X, B_center.Y, B_center.Z + outer_radius) + side_protect / 304.8 * toward.Direction;
                XYZ p3 = new XYZ(B_center.X, B_center.Y, B_center.Z + outer_radius) + (rf.properties.width - side_protect) / 304.8 * toward.Direction;
                Plane plane = Plane.CreateByThreePoints(B_center, p2, p3);

                IList<Curve> curves = new List<Curve>();
                //A2, A2a
                Curve c = Line.CreateBound(p2, p3);
                curves.Add(c);

                subt.Start();

                RebarBarType barType = barType13M;
                RebarShape barshape = M_00;

                Rebar re = Rebar.CreateFromCurvesAndShape(doc, barshape, barType, null, null, element, plane.Normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                Rebar re2 = Rebar.CreateFromCurvesAndShape(doc, barshape, barType, null, null, element, plane.Normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                //re.SetSolidInView(view3D as View3D, true);
                //re.SetUnobscuredInView(view3D, true);
                //re2.SetSolidInView(view3D as View3D, true);
                //re2.SetUnobscuredInView(view3D, true);
                //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");
                //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("B");

                // 培文改：設置鋼筋
                List<Rebar> rebars = new List<Rebar>() { re, re2 };
                foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "B"); }

                double rad_outer_r = outer_radius * 304.8;
                int index = (j == 0) ? 1 : -1;
                ElementTransformUtils.RotateElement(doc, re.Id, toward, (18.0 + 72 * index) / 180 * Math.PI - dis_B / rad_outer_r * index); //+72.0 * j + dis_A / radius
                ElementTransformUtils.RotateElement(doc, re2.Id, toward, (18.0 + 7.5 * index) / 180 * Math.PI + dis_B / rad_outer_r * index); //+72.0 * j + dis_A / radius
                subt.Commit();
            }
            //K環片補強筋(主筋方向)
            dis_K = 0;
            for (int i = 0; i < 2; i++)
            {
                double ba = (i == 0) ? 365.0 : 260.0;
                dis_K += ba;

                XYZ K_center = mid + toward.Direction.Normalize() * (dis_K / 304.8);
                //make plane
                XYZ planepoint2 = new XYZ(K_center.X - toward.Direction.Y, K_center.Y + toward.Direction.X, K_center.Z);
                XYZ planepoint3 = K_center + XYZ.BasisZ;
                Plane plane = Plane.CreateByThreePoints(K_center, planepoint2, planepoint3);
                RebarBarType barType = barType19M;

                IList<Curve> curves = new List<Curve>();
                double rad_inner_r = inner_radius * 304.8;

                int index = (i == 0) ? 1 : -1;
                // 培文改
                //double barDiameter = barType13M.BarDiameter + barType.BarDiameter; // 2020
                double barDiameter = barType13M.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble() + barType.get_Parameter(BuiltInParameter.REBAR_BAR_DIAMETER).AsDouble();
                Curve c = Arc.Create(plane, (inner_radius + (barDiameter) / 2)
                    , (100.5 / 180) * Math.PI + (side_protect + dis_K / 10.0) / rad_inner_r
                    , (115.5 / 180) * Math.PI - (side_protect + dis_K / 10.0) / rad_inner_r);//+ (side_protect + 2.5 * index) / rad_inner_r
                //// 台大
                //Curve c = Arc.Create(plane, (inner_radius + (barType13M.BarDiameter + barType.BarDiameter) / 2)
                //    , (100.5 / 180) * Math.PI + (side_protect + dis_K / 10.0) / rad_inner_r
                //    , (115.5 / 180) * Math.PI - (side_protect + dis_K / 10.0) / rad_inner_r);//+ (side_protect + 2.5 * index) / rad_inner_r
                curves.Add(c);
                subt.Start();


                XYZ norm = rf.data_list_tunnel[1].start_point - rf.data_list_tunnel[0].start_point;
                Rebar re = Rebar.CreateFromCurvesAndShape(doc, shape, barType, null, null, element, norm, curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                //re.SetSolidInView(view3D as View3D, true);
                //re.SetUnobscuredInView(view3D, true);
                //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("K");

                // 培文改：設置鋼筋
                List<Rebar> rebars = new List<Rebar>() { re };
                foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "K"); }

                subt.Commit();
            }
            //K環片補強筋(剪力筋方向)

            for (int j = 0; j < 2; j++)
            {

                XYZ K_center = mid;
                //make plane
                XYZ p2 = new XYZ(K_center.X, K_center.Y, K_center.Z + inner_radius) + side_protect / 304.8 * toward.Direction;
                XYZ p3 = new XYZ(K_center.X, K_center.Y, K_center.Z + inner_radius) + (side_protect + 398.0) / 304.8 * toward.Direction;
                XYZ p4 = K_center + new XYZ(0, 0, inner_radius) + (rf.properties.width - side_protect) / 304.8 * toward.Direction;
                XYZ p5 = p4 - 405.0 / 304.8 * toward.Direction;
                Plane plane = Plane.CreateByThreePoints(K_center, p2, p3);

                IList<Curve> curves = new List<Curve>();
                IList<Curve> curves2 = new List<Curve>();
                double Rebar_r = inner_radius;
                Curve c = Line.CreateBound(p2, p3);
                Curve c2 = Line.CreateBound(p4, p5);
                curves.Add(c);
                curves2.Add(c2);

                subt.Start();
                RebarBarType barType = barType13M;
                RebarShape barshape = M_00;

                Rebar re = Rebar.CreateFromCurvesAndShape(
                    doc, barshape, barType, null, null, element, plane.Normal
                    , curves, RebarHookOrientation.Left, RebarHookOrientation.Left);
                Rebar re2 = Rebar.CreateFromCurvesAndShape(
                    doc, barshape, barType, null, null, element, plane.Normal
                    , curves2, RebarHookOrientation.Left, RebarHookOrientation.Left);
                //re.SetSolidInView(view3D as View3D, true);
                //re.SetUnobscuredInView(view3D, true);
                //re2.SetSolidInView(view3D as View3D, true);
                //re2.SetUnobscuredInView(view3D, true);
                //re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("K");
                //re2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("K");

                // 培文改：設置鋼筋
                List<Rebar> rebars = new List<Rebar>() { re, re2 };
                foreach (Rebar rebar in rebars) { SetRebar(doc, rebar, view3D, "K"); }

                double rad_inner_r = inner_radius * 304.8;
                int index = (j == 0) ? 1 : -1;
                ElementTransformUtils.RotateElement(doc, re.Id, toward, (18 + 7.5 * index) / 180 * Math.PI - 96.0 / rad_inner_r * index);
                ElementTransformUtils.RotateElement(doc, re2.Id, toward, (18 + 7.5 * index) / 180 * Math.PI - 135.0 / rad_inner_r * index);
                subt.Commit();
            }

            t.Commit();
        }
        // 設置鋼筋
        private void SetRebar(Document doc, Rebar re, View3D view3D, string value)
        {
            //re.SetSolidInView(view3D as View3D, true); // 2020

            // 2024：假設已有 RebarId 和 View view
            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            // 設定為實心顯示(實心填充 + 透明度為0)
            ogs.SetSurfaceTransparency(0); // 不透明
            ogs.SetSurfaceForegroundPatternId(GetSolidFillPatternId(doc)); // 實心填充
            ogs.SetSurfaceForegroundPatternColor(new Color(0, 0, 0)); // 黑色或其他顏色            
            view3D.SetElementOverrides(re.Id, ogs); // 套用到 View 上的 Rebar 元件

            re.SetUnobscuredInView(view3D, true);
            re.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(value);
        }
        ElementId GetSolidFillPatternId(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill)
                ?.Id ?? ElementId.InvalidElementId;
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

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
