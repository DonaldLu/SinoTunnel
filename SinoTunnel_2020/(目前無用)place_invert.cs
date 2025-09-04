using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using DataObject;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SinoTunnel_2020
{
    //在用的仰拱
    public class MyPreProcessor : IFailuresPreprocessor
    {
        FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            String transactionName = failuresAccessor.GetTransactionName();

            IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();


            if (fmas.Count == 0)
                return FailureProcessingResult.Continue;


            if (transactionName.Equals("EXEMPLE"))
            {
                foreach (FailureMessageAccessor fma in fmas)
                {
                    if (fma.GetSeverity() == FailureSeverity.Error)
                    {
                        failuresAccessor.DeleteAllWarnings();
                        return FailureProcessingResult.ProceedWithRollBack;
                    }
                    else
                    {
                        failuresAccessor.DeleteWarning(fma);
                    }

                }
            }
            else
            {
                foreach (FailureMessageAccessor fma in fmas)
                {
                    failuresAccessor.DeleteAllWarnings();
                }
            }
            return FailureProcessingResult.Continue;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class place_invert : IExternalCommand
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

            //create_ditch(uiapp, doc_path, rf.data_list);

            set_invert_profile_new(uiapp, rf.data_list, rf.data_list_tunnel, doc_path);
            
            
            FamilySymbol fam_sym;
            Transaction trans = new Transaction(document, "load");
            
            trans.Start();
            document.LoadFamily(path + "仰拱\\invert.rfa");

            fam_sym = new FilteredElementCollector(document).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "invert").First();
            fam_sym.Activate();
            FamilyInstance FI = document.Create.NewFamilyInstance(new XYZ(0, 0, 0), fam_sym, StructuralType.NonStructural);
            trans.Commit();

            /*trans.Start();
            document.LoadFamily(path + "仰拱\\ditch.rfa");
            fam_sym = new FilteredElementCollector(document).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "ditch").First();
            fam_sym.Activate();
            FamilyInstance FI_dicth = document.Create.NewFamilyInstance(new XYZ(0, 0, 0), fam_sym, StructuralType.NonStructural);
            trans.Commit();*/


            return Result.Succeeded;
        }

        public void create_ditch_profile(UIApplication uiapp, string doc_path)
        {
            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\profile.rfa");
            Document edit_doc = edit_uidoc.Document;

            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;
            
            double offset = 300.0;
            double width = 900.0;
            double depth = 25.0;
            double displacementY = 350.0;

            XYZ point1 = new XYZ(-(offset+width)/304.8, -displacementY/304.8, 0);
            XYZ point2 = new XYZ(-(offset) / 304.8, -displacementY / 304.8, 0);
            XYZ point3 = point2 - XYZ.BasisY * (depth / 304.8);
            XYZ point4 = point1 - XYZ.BasisY * (depth / 304.8);

            XYZ point5 = new XYZ((offset) / 304.8, -displacementY / 304.8, 0);
            XYZ point6 = new XYZ((offset + width) / 304.8, -displacementY / 304.8, 0);
            XYZ point7 = point6 - XYZ.BasisY * (depth / 304.8);
            XYZ point8 = point5 - XYZ.BasisY * (depth / 304.8);

            Transaction trans = new Transaction(edit_doc, "test");
            trans.Start();

            DetailCurve line1 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point1, point2));
            DetailCurve line2 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point2, point3));
            DetailCurve line3 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point3, point4));
            DetailCurve line4 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point4, point1));
            DetailCurve line5 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point5, point6));
            DetailCurve line6 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point6, point7));
            DetailCurve line7 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point7, point8));
            DetailCurve line8 = edit_doc.FamilyCreate.NewDetailCurve(edit_doc.ActiveView, Line.CreateBound(point8, point5));

            trans.Commit();

            edit_doc.SaveAs(path + "仰拱\\side_ditches_profile.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc_path);
            edit_doc.Close(false);
        }

        public void create_ditch(UIApplication uiapp, string doc_path, IList<data_object> data_list)
        {
            create_ditch_profile(uiapp, doc_path);

            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\ditch.rfa");
            Document edit_doc = edit_uidoc.Document;

            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;

            Transaction trans = new Transaction(edit_doc, "sweep");
            trans.Start();
            edit_doc.LoadFamily(path + "仰拱\\" + "side_ditches_profile.rfa");

            FamilySymbol fam_sym = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "side_ditches_profile").First();

            ReferenceArray curves = new ReferenceArray();
            for (int i = 1; i < data_list.Count; i++)
            {
                Curve aisle_path = Line.CreateBound(data_list[i - 1].start_point, data_list[i].start_point);
                ModelCurve path_modelcurve = edit_doc.FamilyCreate.NewModelCurve(aisle_path, Sketch_plain(edit_doc, data_list[i - 1].start_point, data_list[i].start_point));
                curves.Append(path_modelcurve.GeometryCurve.Reference);
                path_modelcurve.LookupParameter("可見").Set(0);
            }


            SweepProfile sweepProfile = edit_doc.Application.Create.NewFamilySymbolProfile(fam_sym);
            Sweep aisle = edit_doc.FamilyCreate.NewSweep(false, curves, sweepProfile, 0, ProfilePlaneLocation.Start);
            aisle.LookupParameter("角度").Set(Math.PI * 1.5);
            trans.Commit();

            edit_doc.SaveAs(path + "仰拱\\ditch.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc_path);
            edit_doc.Close(false);

            
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


        /*public IList<ModelCurve> create_curve_model_line(Document document, XYZ start, XYZ end, XYZ point_on_plane, XYZ track_point)
        {
            IList<ModelCurve> model_curve_list = new List<ModelCurve>();
            //TaskDialog.Show("message", "start: " + start.ToString() + " end: " + end.ToString() + "point: " + point_on_plane.ToString());
            Curve c = Arc.Create(start, end, point_on_plane);
            Curve line1 = Line.CreateBound(end, track_point);
            Curve line2 = Line.CreateBound(track_point, start);

            //TaskDialog.Show("message", c.GetEndPoint(0).ToString() + "ww" + c.GetEndPoint(1).ToString());

            Transaction trans = new Transaction(document, "test");
            trans.Start();
            Plane p = Plane.CreateByThreePoints(start, end, point_on_plane);
            SketchPlane sp = SketchPlane.Create(document, p);

            ModelCurve mc = document.FamilyCreate.NewModelCurve(c, sp);
            ModelCurve ml1 = document.FamilyCreate.NewModelCurve(line1, sp);
            ModelCurve ml2 = document.FamilyCreate.NewModelCurve(line2, sp);

            //TaskDialog.Show("message", mc.GeometryCurve.GetEndPoint(0).ToString());

            model_curve_list.Add(mc);
            model_curve_list.Add(ml1);
            model_curve_list.Add(ml2);
            

            trans.Commit();

            return model_curve_list;
        }*/

        public Tuple<CurveArray,SketchPlane> create_curve_model_line_new(Document document, XYZ start, XYZ end, XYZ point_on_plane, XYZ track_point)
        {
            CurveArray curve_list = new CurveArray();
            Curve c = Arc.Create(start, end, point_on_plane);
            Curve line1 = Line.CreateBound(end, track_point);
            Curve line2 = Line.CreateBound(track_point, start);
            

            Transaction trans = new Transaction(document, "test");
            trans.Start();
            Plane p = Plane.CreateByThreePoints(start, end, point_on_plane);
            SketchPlane sp = SketchPlane.Create(document, p);
            

            curve_list.Append(c);
            curve_list.Append(line1);
            curve_list.Append(line2);


            trans.Commit();

            Tuple<CurveArray, SketchPlane> value = new Tuple<CurveArray, SketchPlane>(curve_list, sp);

            return value;
        }

        public void set_invert_profile_new(UIApplication uiapp, IList<data_object> data_list, IList<data_object> data_tunnel_list, string doc_path)
        {
            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;

            IList<XYZ> vec_list = new List<XYZ>();

            //IList<Tuple<CurveArray,SketchPlane>> profile_list = new List<Tuple<CurveArray,SketchPlane>>();

            XYZ first_point = data_list[0].start_point;
            XYZ second_point = data_list[1].start_point;

            XYZ direction = new XYZ(second_point.X - first_point.X, second_point.Y - first_point.Y, second_point.Z - first_point.Z);

            double start_angle_horizontal = Math.Atan2(direction.Y, direction.X) * 180 / Math.PI;
            double start_angle_vertical = Math.Atan2(direction.Z, Math.Pow(Math.Pow(direction.X, 2) + Math.Pow(direction.Y, 2), 0.5)) * 180 / Math.PI;

            //vec_list.Append(direction);

            //every profile iteration work
            for (int i = 0; i < data_list.Count; i++)
            {
                UIDocument edit_profile_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\profile.rfa");
                Document edit_profile_doc = edit_profile_uidoc.Document;

                XYZ point = data_list[i].start_point;
                XYZ center = data_tunnel_list[i].start_point;
                Tuple<XYZ, XYZ> vector = random_create_vecXY(direction, start_angle_horizontal, start_angle_vertical);
                XYZ slide_vector1 = vector_normalized(rotate_xyz(direction, point, vector.Item1, Math.Atan2(1, 40)).Normalize());
                XYZ slide_vector2 = vector_normalized(rotate_xyz(direction, point, -vector.Item1, -Math.Atan2(1, 40)).Normalize());

                XYZ this_direction = rotate_by_vh(vector);

                double depth = 400.0;
                XYZ invert_top_point = point - vector.Item2 * (depth / 304.8);

                double rotate_angle_for_profile = -Math.Acos(this_direction.DotProduct(XYZ.BasisX) / (this_direction.GetLength()));
                //TaskDialog.Show("message", rotate_angle_for_profile.ToString());

                double b1 = calculate_b_of_point_on_circle(center, slide_vector1, invert_top_point, 2800 / 304.8);
                double b2 = calculate_b_of_point_on_circle(center, slide_vector2, invert_top_point, 2800 / 304.8);
                double bottom_b = calculate_b_of_point_on_circle(center, -vector.Item2, invert_top_point, 2800 / 304.8);
                

                invert_top_point = new XYZ(0, 0, 0) - vector.Item2 * (400 / 304.8);
                XYZ side_point1 = b1 * slide_vector1 + invert_top_point;
                XYZ side_point2 = b2 * slide_vector2 + invert_top_point;
                XYZ bottom = bottom_b * -vector.Item2 + invert_top_point;

                

                invert_top_point = rotate_xyz(-this_direction.CrossProduct(XYZ.BasisX).Normalize(), new XYZ(0, 0, 0), invert_top_point, rotate_angle_for_profile);
                side_point1 = rotate_xyz(-this_direction.CrossProduct(XYZ.BasisX).Normalize(), new XYZ(0, 0, 0), side_point1, rotate_angle_for_profile);
                side_point2 = rotate_xyz(-this_direction.CrossProduct(XYZ.BasisX).Normalize(), new XYZ(0, 0, 0), side_point2, rotate_angle_for_profile);
                bottom = rotate_xyz(-this_direction.CrossProduct(XYZ.BasisX).Normalize(), new XYZ(0, 0, 0), bottom, rotate_angle_for_profile);


                invert_top_point = new XYZ(-invert_top_point.Y, -invert_top_point.Z, 0);
                side_point1 = new XYZ(-side_point1.Y, -side_point1.Z, 0);
                side_point2 = new XYZ(-side_point2.Y, -side_point2.Z, 0);
                bottom = new XYZ(-bottom.Y, -bottom.Z, 0);

                Transaction t = new Transaction(edit_profile_doc, "draw_proflie");
                t.Start();

                DetailCurve line1 = edit_profile_doc.FamilyCreate.NewDetailCurve(edit_profile_doc.ActiveView, Line.CreateBound(side_point1, invert_top_point));
                DetailCurve line2 = edit_profile_doc.FamilyCreate.NewDetailCurve(edit_profile_doc.ActiveView, Line.CreateBound(invert_top_point, side_point2));
                DetailCurve arc = edit_profile_doc.FamilyCreate.NewDetailCurve(edit_profile_doc.ActiveView, Arc.Create(side_point2, side_point1, bottom));

                t.Commit();

                edit_profile_doc.SaveAs(path + "仰拱\\invert_profile" + i + ".rfa", save_option);
                uiapp.OpenAndActivateDocument(doc_path);
                edit_profile_doc.Close(false);

                //Tuple<CurveArray,SketchPlane> mc_arc = create_curve_model_line_new(edit_doc, side_point1, side_point2, bottom, invert_top_point);

                //profile_list.Add(mc_arc);

                start_angle_horizontal -= data_list[i].horizontal_angle;
                start_angle_vertical -= data_list[i].vertical_angle;

                vec_list.Add(this_direction);
            }

            //TaskDialog.Show("message", "profile done");

            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\_invert.rfa");
            Document edit_doc = edit_uidoc.Document;

            Transaction trans = new Transaction(edit_doc, "blend");
            int count = 0;
            for (int i = 1; i < data_list.Count; i++)
            {
                trans.Start("0");

                FailureHandlingOptions options = trans.GetFailureHandlingOptions();
                MyPreProcessor preproccessor = new MyPreProcessor();
                options.SetClearAfterRollback(true);
                options.SetFailuresPreprocessor(preproccessor);
                trans.SetFailureHandlingOptions(options);

                edit_doc.LoadFamily(path + "仰拱\\invert_profile" + (i - 1) + ".rfa");

                edit_doc.LoadFamily(path + "仰拱\\invert_profile" + (i) + ".rfa");

                FamilySymbol fam_sym_bottom = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "invert_profile" + (i - 1)).First();
                FamilySymbol fam_sym_top = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "invert_profile" + (i)).First();


                Tuple<Curve,SketchPlane> arc_path = create_arc_by_two_vecpoint(data_list[i-1].start_point, vec_list[i-1], data_list[i].start_point, vec_list[i],edit_doc);

                SweepProfile sweepProfile_bottom = edit_doc.Application.Create.NewFamilySymbolProfile(fam_sym_bottom);
                SweepProfile sweepProfile_top = edit_doc.Application.Create.NewFamilySymbolProfile(fam_sym_top);
                
                SweptBlend blend = edit_doc.FamilyCreate.NewSweptBlend(true, arc_path.Item1, arc_path.Item2, sweepProfile_bottom, sweepProfile_top);// (true, profile_list[i].Item1, profile_list[i - 1].Item1, profile_list[i - 1].Item2);
                blend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).Set(Math.PI);
                blend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).Set(Math.PI);

                trans.Commit();
                if (i % 50 == 0)
                {
                    //TaskDialog.Show("message", i.ToString());
                    edit_doc.SaveAs(path + "仰拱\\_invert" + count.ToString() + ".rfa", save_option);
                    uiapp.OpenAndActivateDocument(doc_path);
                    edit_doc.Close(false);

                    count++;

                    edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\_invert.rfa");
                    edit_doc = edit_uidoc.Document;
                }else if (i == data_list.Count - 1)
                {
                    edit_doc.SaveAs(path + "仰拱\\_invert" + count.ToString() + ".rfa", save_option);
                    uiapp.OpenAndActivateDocument(doc_path);
                    edit_doc.Close(false);
                    count++;
                    
                }

            }

            assemble(doc_path, count, uiapp);
            /*edit_doc.SaveAs(path + "仰拱\\_invert.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc_path);
            edit_doc.Close(false);*/

            //assemble(doc_path, count, uiapp);
        }

        public Tuple<Curve,SketchPlane> create_arc_by_two_vecpoint(XYZ start, XYZ start_vec, XYZ end, XYZ end_vec, Document doc)
        {
            double degree = Math.Acos(start_vec.DotProduct(end_vec)/(start_vec.GetLength() * end_vec.GetLength()));
            double radius = (end - start).GetLength()/(2*Math.Sin(degree/2));
            Curve myArc = Line.CreateBound(start, end);

            XYZ midPointOfArc = (start + end) / 2;
            SketchPlane sp = Sketch_plain(doc, start, end);
            //TaskDialog.Show("message", "line" + sp.GetPlane().Normal.ToString());
            if (1/radius > 0.0)
            {
                double sagitta = radius - Math.Sqrt(Math.Pow(radius, 2.0) - Math.Pow((end - start).GetLength() / 2.0, 2.0));
                //TaskDialog.Show("message", radius.ToString());

                XYZ midPointOfChord = (start + end) / 2.0;

                midPointOfArc = midPointOfChord + Transform.CreateRotation(XYZ.BasisZ, Math.PI / 2.0).OfVector((start - end).Normalize().Multiply(sagitta));

                myArc = Arc.Create(start, end, midPointOfArc);

                Plane p = Plane.CreateByThreePoints(start, end, midPointOfArc);
                //TaskDialog.Show("message", "curve:" + p.Normal.ToString());
                sp = SketchPlane.Create(doc, p);

                //doc.FamilyCreate.NewModelCurve(myArc, sp);
            }
            
            
            return new Tuple<Curve, SketchPlane>(myArc, sp);
        }

       

        /*public void set_invert_profile(UIApplication uiapp, IList<data_object> data_list, IList<data_object> data_tunnel_list, string doc_path)
        {
            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\invert.rfa");
            Document edit_doc = edit_uidoc.Document;

            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;
            
            ReferenceArrayArray ref_ar_ar = new ReferenceArrayArray();

            XYZ first_point = data_list[0].start_point;
            XYZ second_point = data_list[1].start_point;

            XYZ direction = new XYZ(second_point.X - first_point.X, second_point.Y - first_point.Y, second_point.Z - first_point.Z);

            double start_angle_horizontal = Math.Atan2(direction.Y, direction.X) * 180 / Math.PI;
            double start_angle_vertical = Math.Atan2(direction.Z, Math.Pow(Math.Pow(direction.X, 2) + Math.Pow(direction.Y, 2), 0.5)) * 180 / Math.PI;

            int count = 0;

            //every profile iteration work
            for(int i  = 0; i < data_list.Count; i++)
            {
                XYZ point = data_list[i].start_point;
                XYZ center = data_tunnel_list[i].start_point;
                Tuple<XYZ, XYZ> vector = random_create_vecXY(direction, start_angle_horizontal, start_angle_vertical);
                XYZ slide_vector1 = rotate_xyz(direction, point, vector.Item1, Math.Atan2(1, 40));
                XYZ slide_vector2 = rotate_xyz(direction, point, -vector.Item1, -Math.Atan2(1, 40));

                XYZ invert_top_point = point - vector.Item2 * (400 / 304.8);

                double b1 = calculate_b_of_point_on_circle(center, slide_vector1, invert_top_point, 2800/ 304.8);
                double b2 = calculate_b_of_point_on_circle(center, slide_vector2, invert_top_point, 2800/ 304.8);
                double bottom_b = calculate_b_of_point_on_circle(center, -vector.Item2, invert_top_point, 2800/ 304.8);


                XYZ side_point1 = b1 * slide_vector1 + invert_top_point;
                XYZ side_point2 = b2 * slide_vector2 + invert_top_point;
                XYZ bottom = bottom_b * -vector.Item2 + invert_top_point;

                
                IList<ModelCurve> mc_arc = create_curve_model_line(edit_doc, side_point1, side_point2, bottom, invert_top_point);
                
                ReferenceArray ref_ar = getObjectRef(mc_arc);
                ref_ar_ar.Append(ref_ar);
                
                start_angle_horizontal -= data_list[i].horizontal_angle;
                start_angle_vertical -= data_list[i].vertical_angle;


                if (i % 100 == 0 && i != 0)
                {
                    Transaction t = new Transaction(edit_doc, "sweep");
                    t.Start();
                    edit_doc.FamilyCreate.NewLoftForm(true, ref_ar_ar).LookupParameter("材料").Set(25);
                    t.Commit();

                    count++;
                    edit_doc.SaveAs(path + "仰拱\\invert" + count + ".rfa", save_option);
                    uiapp.OpenAndActivateDocument(doc_path);
                    edit_doc.Close(false);
                    edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\invert.rfa");
                    edit_doc = edit_uidoc.Document;

                    ref_ar_ar = new ReferenceArrayArray();
                    mc_arc = create_curve_model_line(edit_doc, side_point1, side_point2, bottom, invert_top_point);
                    ref_ar = getObjectRef(mc_arc);
                    ref_ar_ar.Append(ref_ar);
                    
                }
                else if(i == data_list.Count - 1)
                {
                    Transaction t = new Transaction(edit_doc, "sweep");
                    t.Start();
                    edit_doc.FamilyCreate.NewLoftForm(true, ref_ar_ar).LookupParameter("材料").Set(25);
                    t.Commit();

                    //t.Start();
                    //edit_doc.LoadFamily(path + "仰拱\\ditch.rfa");
                    //FamilySymbol fam_sym = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "ditch").First();
                    //fam_sym.Activate();
                    //FamilyInstance FI_dicth = edit_doc.FamilyCreate.NewFamilyInstance(new XYZ(0, 0, 0), fam_sym, StructuralType.NonStructural);

                    //t.Commit();

                    count++;
                    edit_doc.SaveAs(path + "仰拱\\invert" + count + ".rfa", save_option);
                    uiapp.OpenAndActivateDocument(doc_path);
                    edit_doc.Close(false);
                }

            }
            

            assemble(doc_path, count, uiapp);
            
        }*/

        public void assemble(string doc_path, int count, UIApplication uiapp)
        {
            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(path + "backup\\invert.rfa");
            Document edit_doc = edit_uidoc.Document;

            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;

            Family family;
            FamilySymbol fam_sym;
            Transaction trans = new Transaction(edit_doc, "load");


            for(int i = 0; i < count; i++)
            {
                trans.Start();
                edit_doc.LoadFamily(path + "仰拱\\_invert" + (i) + ".rfa", out family);
                fam_sym = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "_invert" + (i)).First();
                fam_sym.Activate();
                FamilyInstance FI = edit_doc.FamilyCreate.NewFamilyInstance(new XYZ(0,0,0), fam_sym, StructuralType.NonStructural);
                trans.Commit();
            }
            

            edit_doc.SaveAs(path + "仰拱\\invert.rfa", save_option);
            uiapp.OpenAndActivateDocument(doc_path);
            edit_doc.Close(false);
            
        }

        public ReferenceArray getObjectRef(IList<ModelCurve> modelcurve_list)
        {
            ReferenceArray ra = new ReferenceArray();
            for (int i = 0; i < modelcurve_list.Count; i++)
            {
                ra.Append(modelcurve_list[i].GeometryCurve.Reference);
            }
            return ra;
        }

        //計算仰拱頂端經過 1/40 相交於圓形的點
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
            if(ans <= 0)
            {
                ans = (-b - Math.Pow(Math.Pow(b, 2) - 4 * a * c, 0.5)) / (2 * a);
            }

            return ans;
        }

        //angle : radius, vector_x : 被轉向量
        public XYZ rotate_xyz(XYZ direction, XYZ point, XYZ vector_x, double angle)
        {
            Transform rot = Transform.CreateRotationAtPoint(direction, angle, point);
            Curve c = Line.CreateBound(point, point + vector_x) as Curve;
            c = c.CreateTransformed(rot);
            XYZ sp = c.GetEndPoint(0);
            XYZ ep = c.GetEndPoint(1);
            XYZ slide_vec = new XYZ(ep.X - sp.X, ep.Y - sp.Y, ep.Z - sp.Z);
            //slide_vec = vector_normalized(slide_vec);

            return slide_vec;
        }

        public XYZ rotate_by_vh(Tuple<XYZ,XYZ> vecs)
        {
            return vecs.Item1.CrossProduct(vecs.Item2);
        }

        //tuple(vec_x, vec_y)
        public Tuple<XYZ, XYZ> random_create_vecXY(XYZ origin_vector, double horizontal_angle, double vertical_angle)
        {
            horizontal_angle = Math.PI * (horizontal_angle / 180);
            vertical_angle = Math.PI * (vertical_angle / 180);
            XYZ x = new XYZ(-Math.Sin(horizontal_angle), Math.Cos(horizontal_angle), 0);
            x = vector_normalized(x);
            XYZ y = vector_cross_product(origin_vector, x);
            y = vector_normalized(y);
            Tuple<XYZ, XYZ> vecs = new Tuple<XYZ, XYZ>(x, y);

            return vecs;
        }

        public XYZ vector_cross_product(XYZ vec1, XYZ vec2)
        {
            double x = vec1.Y * vec2.Z - vec1.Z * vec2.Y;
            double y = vec1.Z * vec2.X - vec1.X * vec2.Z;
            double z = vec1.X * vec2.Y - vec1.Y * vec2.X;
            XYZ vector = new XYZ(x, y, z);
            return vector;
        }

        public XYZ vector_normalized(XYZ vector)
        {
            double X = vector.X;
            double Y = vector.Y;
            double Z = vector.Z;
            double length = Math.Pow((X * X + Y * Y + Z * Z), 0.5);
            vector = new XYZ(X / length, Y / length, Z / length);
            return vector;
        }

    }
}
