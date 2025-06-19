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
using System.Numerics;
using Plane = Autodesk.Revit.DB.Plane;
using System.Windows;

namespace SinoTunnel
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    class Track_bed : IExternalEventHandler
    {
        string path;

        public void Execute(UIApplication app)
        {
            path = Form1.path;
            try
            {
                Document ori_doc = app.ActiveUIDocument.Document;

                readfile rf = new readfile();
                rf.read_point();
                rf.read_point2();
                rf.read_target_station();
                rf.read_target_station2();
                track_bed_properties tb_properties = rf.read_track_bed();

                var list = new List<Tuple<IList<data_object>, setting_station>>();
                list.Add(new Tuple<IList<data_object>, setting_station>(rf.data_list, rf.setting_Station));
                list.Add(new Tuple<IList<data_object>, setting_station>(rf.data_list2, rf.setting_Station2));

                int count = -1;

                // 培文改寫：移除專案中現有的所有"鋼軌基鈑"與"第三軌支架_自"
                using (Transaction trans = new Transaction(ori_doc, "移除鋼軌基鈑與第三軌支架_自"))
                {
                    trans.Start();
                    List<FamilyInstance> foundation_support_symbols = new FilteredElementCollector(ori_doc).OfCategory(BuiltInCategory.OST_GenericModel)
                                                                      .WhereElementIsNotElementType().Where(x => x.Name.Equals("鋼軌基鈑") || x.Name.Equals("第三軌支架_自"))
                                                                      .Cast<FamilyInstance>().ToList();
                    if (foundation_support_symbols.Count > 0) { ori_doc.Delete(foundation_support_symbols.Select(x => x.Id).ToList()); }
                    trans.Commit();
                }

                ///
                foreach (Tuple<IList<data_object>, setting_station> all_list in list)
                {
                    count += 1;
                    IList<data_object> data_list = all_list.Item1;
                    setting_station setting_Station = all_list.Item2;
                    bool isRight = true;
                    int set_filped = 1;
                    if (setting_Station.walk_way_station[0][0].ToString() == "左側") { isRight = false; set_filped = 0; }

                    UIDocument uidoc = app.OpenAndActivateDocument(path + "道床\\道床.rfa");
                    Document doc = uidoc.Document;

                    Transaction big_t = new Transaction(doc);
                    SubTransaction subtran = new SubTransaction(doc);
                    // 關閉警示視窗
                    FailureHandlingOptions options = big_t.GetFailureHandlingOptions();
                    options.SetClearAfterRollback(true);
                    options.SetFailuresPreprocessor(new CloseWarnings());
                    big_t.SetFailureHandlingOptions(options);

                    int station_delta = 0;
                    int station_total = 0;

                    //先對輪廓做參數化sus
                    List<double> angle_list = new List<double>();
                    for (int i = 0; i < data_list.Count; i++) { angle_list.Add(Math.Round(data_list[i].super_high_angle, 10)); }

                    //取得角度種類，以確定要建置幾種輪廓
                    List<double> angle_distinct = angle_list.Distinct().ToList();

                    //如果角度出現的次數大於十，可判斷為是需要建置的輪廓

                    //根據讀入道床樣式建立道床
                    foreach (string[] tb_name in setting_Station.track_bed_station)
                    {
                        track_bed tb = new track_bed(tb_name[0]);

                        if (tb.name == "標準道床_左")
                        {
                            tb.x = tb_properties.standard_center_dis + tb_properties.standard_width;
                            tb.y = -tb_properties.standard_top_height;
                        }
                        else if (tb.name == "平版式道床")
                        {
                            tb.x = tb_properties.flat_width / 2;
                            tb.y = -tb_properties.flat_elevation;
                        }
                        else if (tb.name == "浮動式道床")
                        {
                            tb.x = tb_properties.float_width / 2;
                            tb.y = -tb_properties.float_elevation;
                        }
                        else
                        {
                            TaskDialog.Show("message", "wrong name : " + tb.name);
                            break;
                        }

                        //根據角度不同製造不同輪廓
                        bool de = true;
                        while (de)
                        {

                            for (int i = 0; i < angle_distinct.Count; i++)
                            {
                                big_t.Start("製造輪廓");
                                double angle = angle_distinct[i];
                                string pa_and_ne = (angle >= 0) ? "_超高正" : "_超高負";
                                FamilySymbol original_trackbed_profile = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                                                                         .Cast<FamilySymbol>().ToList().Where(x => x.Name == tb.name + pa_and_ne).First();
                                try
                                {
                                    FamilySymbol new_trackbed_profile = original_trackbed_profile.Duplicate(tb.name + "_angle=" + angle.ToString()) as FamilySymbol;
                                    big_t.Commit();
                                    big_t.Start("調整參數輪廓");
                                    new_trackbed_profile.LookupParameter("超高轉角").SetValueString(angle.ToString());
                                }
                                catch { }

                                big_t.Commit();
                            }


                            Int32.TryParse(tb_name[1], out int start_station);
                            Int32.TryParse(tb_name[2], out int end_station);

                            station_delta = end_station - start_station;

                            if (tb.name != "標準道床_右") { station_total += station_delta; }

                            //來開始掃掠混成
                            big_t.Start("掃掠混成");
                            for (int i = station_total - station_delta + 1; i <= station_total; i++)
                            {
                                bool isSolid = true;
                                subtran.Start();

                                Line t_path = Line.CreateBound(data_list[i].start_point, data_list[i - 1].start_point);

                                FamilySymbol start = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList()
                                                     .Where(x => x.Name == tb.name + "_angle=" + Math.Round(data_list[i - 1].super_high_angle, 10).ToString()).First();

                                FamilySymbol end = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList()
                                                   .Where(x => x.Name == tb.name + "_angle=" + Math.Round(data_list[i].super_high_angle, 10).ToString()).First();

                                //TaskDialog.Show("test", tb.name);
                                if (tb.name == "標準道床_左" || tb.name == "標準道床_右")
                                {
                                    start.LookupParameter("仰拱頂部").Set(tb_properties.standard_top_height / 304.8);
                                    start.LookupParameter("寬度").Set(tb_properties.standard_width / 304.8);
                                    start.LookupParameter("道床凹槽深度").Set(tb_properties.standard_depth / 304.8);
                                    start.LookupParameter("道床與軌道中心距離").Set(tb_properties.standard_center_dis / 304.8);
                                    start.LookupParameter("高程").Set(tb_properties.standard_elevation / 304.8);
                                    end.LookupParameter("仰拱頂部").Set(tb_properties.standard_top_height / 304.8);
                                    end.LookupParameter("寬度").Set(tb_properties.standard_width / 304.8);
                                    end.LookupParameter("道床凹槽深度").Set(tb_properties.standard_depth / 304.8);
                                    end.LookupParameter("道床與軌道中心距離").Set(tb_properties.standard_center_dis / 304.8);
                                    end.LookupParameter("高程").Set(tb_properties.standard_elevation / 304.8);
                                }
                                else if (tb.name == "平版式道床")
                                {
                                    start.LookupParameter("仰拱頂部").Set(tb_properties.flat_top_height / 304.8);
                                    start.LookupParameter("寬度").Set(tb_properties.flat_width / 304.8);
                                    start.LookupParameter("高程").Set(tb_properties.flat_elevation / 304.8);
                                    end.LookupParameter("仰拱頂部").Set(tb_properties.flat_top_height / 304.8);
                                    end.LookupParameter("寬度").Set(tb_properties.flat_width / 304.8);
                                    end.LookupParameter("高程").Set(tb_properties.flat_elevation / 304.8);
                                }
                                else if (tb.name == "浮動式道床")
                                {
                                    start.LookupParameter("仰拱頂部").Set(tb_properties.float_top_height / 304.8);
                                    start.LookupParameter("寬度").Set(tb_properties.float_width / 304.8);
                                    start.LookupParameter("厚度").Set(tb_properties.float_depth / 304.8);
                                    start.LookupParameter("支承墊寬度").Set(tb_properties.suppad_width / 304.8);
                                    start.LookupParameter("支承墊厚度").Set(tb_properties.suppad_depth / 304.8);
                                    start.LookupParameter("高程").Set(tb_properties.float_elevation / 304.8);
                                    end.LookupParameter("仰拱頂部").Set(tb_properties.float_top_height / 304.8);
                                    end.LookupParameter("寬度").Set(tb_properties.float_width / 304.8);
                                    end.LookupParameter("厚度").Set(tb_properties.float_depth / 304.8);
                                    end.LookupParameter("支承墊寬度").Set(tb_properties.suppad_width / 304.8);
                                    end.LookupParameter("支承墊厚度").Set(tb_properties.suppad_depth / 304.8);
                                    end.LookupParameter("高程").Set(tb_properties.float_elevation / 304.8);
                                }
                                else
                                {
                                    TaskDialog.Show("message", "wrong name : " + tb.name);
                                    break;
                                }

                                SweepProfile Profile_top = doc.Application.Create.NewFamilySymbolProfile(start);
                                SweepProfile Profile_bottom = doc.Application.Create.NewFamilySymbolProfile(end);

                                //Plane geometryPlane =  Plane.CreateByThreePoints(path.GetEndPoint(0), path.GetEndPoint(1), (path.GetEndPoint(0) + new XYZ(1,1,1)));

                                SketchPlane sketchPlane = Sketch_plain(doc, data_list[i - 1].start_point, data_list[i].start_point);
                                subtran.Commit();
                                subtran.Start();
                                SweptBlend Blend = doc.FamilyCreate.NewSweptBlend(isSolid, t_path, sketchPlane, Profile_bottom, Profile_top);
                                subtran.Commit();
                                //if (Math.Round(t_path.Direction.Z, 5) != 0)
                                //{
                                //    Blend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString("180");
                                //    Blend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString("180");
                                //}
                                Blend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString("90");
                                Blend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString("90");
                            }
                            big_t.Commit();
                            if (tb.name == "標準道床_左") { tb.name = "標準道床_右"; }
                            else if (tb.name == "標準道床_右" ^ tb.name == "平版式道床" ^ tb.name == "浮動式道床") { de = false; continue; }
                        }
                    }

                    //鋼軌及第三軌部分

                    //EDIT 鋼軌雛型
                    //SubTransaction sub = new SubTransaction(doc);
                    //sub.Start();
                    double gauge = tb_properties.rail_gauge / 2.0; // 鋼軌軌距
                    double face_width = tb_properties.rail_face_width / 2.0; // 鋼軌面寬
                    List<Family> orbit_family_list = new FilteredElementCollector(doc).OfClass(typeof(Family))
                                                     .Cast<Family>().ToList().Where(x => x.Name.Contains("鋼軌雛型")).ToList();
                    foreach (Family orbit_family in orbit_family_list)
                    {
                        if (gauge != 760)
                        {
                            Document orbit_doc = doc.EditFamily(orbit_family);

                            Transaction orbit_t = new Transaction(orbit_doc);
                            orbit_t.Start("根據軌距移動鋼軌");

                            List<ElementId> left_profile = new List<ElementId>();
                            List<ElementId> right_profile = new List<ElementId>();
                            List<CurveElement> detailArc_list = new FilteredElementCollector(orbit_doc)
                                    .OfClass(typeof(CurveElement)).Cast<CurveElement>().Where(x => x.LineStyle.Name == "輪廓").ToList();
                            foreach (CurveElement arc in detailArc_list)
                            {
                                if (arc.get_BoundingBox(orbit_doc.ActiveView).Max.X > 0) { right_profile.Add(arc.Id); }
                                else { left_profile.Add(arc.Id); }
                            }
                            try
                            {
                                //ElementTransformUtils.MoveElements(orbit_doc, left_profile, new XYZ((760.0 - gauge) / 304.8, 0, 0));                                
                                ElementTransformUtils.MoveElements(orbit_doc, left_profile, new XYZ((760.0 - gauge - face_width) / 304.8, 0, 0)); // 培文改寫
                            }
                            catch
                            {
                                //ElementTransformUtils.MoveElements(orbit_doc, right_profile, new XYZ((gauge - 760.0) / 304.8, 0, 0));
                                ElementTransformUtils.MoveElements(orbit_doc, right_profile, new XYZ(-(760.0 - gauge - face_width) / 304.8, 0, 0)); // 培文改寫
                            }

                            orbit_t.Commit();
                            //orbit_doc.Save();

                            orbit_doc.LoadFamily(doc, new FamilyOption());
                            app.OpenAndActivateDocument(doc.PathName);
                            orbit_doc.Close(false);
                        }
                    }

                    //sub.Commit();
                    big_t.Start("Steel Orbit");
                    //掃掠鋼軌及第三軌
                    subtran.Start();
                    FamilySymbol orbit_R_profile = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                                                   .Cast<FamilySymbol>().ToList().Where(x => x.Name == "鋼軌雛型_右").First();
                    FamilySymbol orbit_L_profile = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                                                   .Cast<FamilySymbol>().ToList().Where(x => x.Name == "鋼軌雛型_左").First();
                    SweepProfile orbit_R_sweep_profile = doc.Application.Create.NewFamilySymbolProfile(orbit_R_profile);
                    SweepProfile orbit_L_sweep_profile = doc.Application.Create.NewFamilySymbolProfile(orbit_L_profile);
                    FamilySymbol third_profile = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                                                 .Cast<FamilySymbol>().ToList().Where(x => x.Name == "第三軌_內").First();
                    FamilySymbol third_outer_profile = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                                                       .Cast<FamilySymbol>().ToList().Where(x => x.Name == "第三軌_外").First();
                    //設定第三軌參數
                    third_profile.LookupParameter("高程").SetValueString(tb_properties.third_track_elevation.ToString());
                    third_profile.LookupParameter("鋼軌軌距").SetValueString(gauge.ToString());
                    third_profile.LookupParameter("第三軌與鋼軌距").SetValueString(tb_properties.third_steel_between_dis.ToString());
                    third_outer_profile.LookupParameter("高程").SetValueString(tb_properties.third_track_elevation.ToString());
                    third_outer_profile.LookupParameter("鋼軌軌距").SetValueString(gauge.ToString());
                    third_outer_profile.LookupParameter("第三軌與鋼軌距").SetValueString(tb_properties.third_steel_between_dis.ToString());

                    SweepProfile third_sweep_profile = doc.Application.Create.NewFamilySymbolProfile(third_profile);
                    SweepProfile third_sweep_outer_profile = doc.Application.Create.NewFamilySymbolProfile(third_outer_profile);
                    subtran.Commit();
                    ReferenceArray multi_path = new ReferenceArray();
                    List<ElementId> mc_list = new List<ElementId>();
                    for (int i = 1; i < data_list.Count; i++)
                    {
                        subtran.Start();
                        Line n_single_path = Line.CreateBound(data_list[i - 1].start_point, data_list[i].start_point);
                        SketchPlane temp_plane = Sketch_plain(doc, data_list[i].start_point, data_list[i - 1].start_point);
                        subtran.Commit();
                        subtran.Start();

                        SweptBlend steel_orbit_R_blend = doc.FamilyCreate.NewSweptBlend(true, n_single_path, temp_plane, orbit_R_sweep_profile, orbit_R_sweep_profile);
                        SweptBlend steel_orbit_L_blend = doc.FamilyCreate.NewSweptBlend(true, n_single_path, temp_plane, orbit_L_sweep_profile, orbit_L_sweep_profile);
                        SweptBlend third_rail_blend = doc.FamilyCreate.NewSweptBlend(true, n_single_path, temp_plane, third_sweep_profile, third_sweep_profile);
                        SweptBlend third_rail_outer_blend = doc.FamilyCreate.NewSweptBlend(true, n_single_path, temp_plane, third_sweep_outer_profile, third_sweep_outer_profile);

                        double this_angle = -1 * data_list[i].super_high_angle;
                        double fore_angle = -1 * data_list[i - 1].super_high_angle;

                        // 培文改寫
                        this_angle = 90 + (-1 * data_list[i].super_high_angle);
                        fore_angle = 90 + (-1 * data_list[i - 1].super_high_angle);

                        steel_orbit_R_blend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString(this_angle.ToString());
                        steel_orbit_R_blend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString(fore_angle.ToString());
                        steel_orbit_L_blend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString(this_angle.ToString());
                        steel_orbit_L_blend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString(fore_angle.ToString());
                        third_rail_blend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString(this_angle.ToString());
                        third_rail_blend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString(fore_angle.ToString());
                        third_rail_outer_blend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString(this_angle.ToString());
                        third_rail_outer_blend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString(fore_angle.ToString());
                        subtran.Commit();

                        subtran.Start();

                        //檢查目前里程在第三軌是在左側或右側
                        string target_side = null;

                        int current_station = StationToInt(data_list[i - 1].station);
                        for (int k = 0; k < setting_Station.third_rail_station.Count; k++)
                        {
                            int start_st = int.Parse(setting_Station.third_rail_station[k][1]);
                            int end_st = int.Parse(setting_Station.third_rail_station[k][2]);

                            if (current_station >= start_st && current_station <= end_st) // 培文改：current_station <= end_st
                            {
                                target_side = setting_Station.third_rail_station[k][0];
                            }
                        }
                        //if (target_side == "左側") { set_filped = 0; }
                        //else { set_filped = 1; }

                        // 培文改
                        if (target_side == "右側") { set_filped = 0; }
                        else { set_filped = 1; }

                        third_rail_blend.get_Parameter(BuiltInParameter.PROFILE1_FLIPPED_HOR).Set(set_filped);
                        third_rail_blend.get_Parameter(BuiltInParameter.PROFILE2_FLIPPED_HOR).Set(set_filped);
                        third_rail_outer_blend.get_Parameter(BuiltInParameter.PROFILE1_FLIPPED_HOR).Set(set_filped);
                        third_rail_outer_blend.get_Parameter(BuiltInParameter.PROFILE2_FLIPPED_HOR).Set(set_filped);

                        subtran.Commit();
                    }

                    subtran.Start();

                    doc.Delete(mc_list);

                    //間距擺放基板及支架部分
                    int base_dis = int.Parse(tb_properties.rail_base_dis.ToString());
                    int st_foundation_num = ((data_list.Count - 1) * 1000 / base_dis) + 1;

                    subtran.Commit();
                    big_t.Commit();

                    SaveAsOptions bsaveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };
                    doc.SaveAs(path + "道床\\instance\\道床final_" + count.ToString() + ".rfa", bsaveAsOptions);
                    app.OpenAndActivateDocument(ori_doc.PathName);
                    // 更新專案內Family的參數
                    try { Family family = doc.LoadFamily(ori_doc, new LoadOptions()); }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                    doc.Close();

                    Transaction ori_t = new Transaction(ori_doc);
                    SubTransaction ori_sub_t = new SubTransaction(ori_doc);
                    ori_t.Start("載入族群");
                    try { ori_doc.LoadFamily(path + "道床\\instance\\道床final_" + count.ToString() + ".rfa"); }
                    catch { }

                    FamilySymbol track_bed = new FilteredElementCollector(ori_doc).OfClass(typeof(FamilySymbol))
                                             .Cast<FamilySymbol>().ToList().Where(x => x.Name == "道床final_" + count.ToString()).First();
                    track_bed.Activate();

                    // 如果專案中未放置道床才擺放
                    try
                    {
                        FamilyInstance findIns = new FilteredElementCollector(ori_doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().Where(x => x.Name.Equals(track_bed.Name)).Cast<FamilyInstance>().FirstOrDefault();
                        if (findIns == null) { FamilyInstance object_acr = ori_doc.Create.NewFamilyInstance(XYZ.Zero, track_bed, StructuralType.NonStructural); }
                    }
                    catch (Exception) { FamilyInstance object_acr = ori_doc.Create.NewFamilyInstance(XYZ.Zero, track_bed, StructuralType.NonStructural); }
                    ori_t.Commit();

                    ori_t.Start("第三軌和鋼軌基鈑");

                    FamilySymbol foundation_symbol = new FilteredElementCollector(ori_doc).OfClass(typeof(FamilySymbol))
                                                     .Cast<FamilySymbol>().ToList().Where(x => x.Name == "鋼軌基鈑").First();
                    foundation_symbol.Activate();
                    //設定鋼軌基板參數
                    foundation_symbol.LookupParameter("厚").SetValueString(tb_properties.rail_base_thickness.ToString());
                    foundation_symbol.LookupParameter("寬").SetValueString(tb_properties.rail_base_width.ToString());
                    foundation_symbol.LookupParameter("長").SetValueString(tb_properties.rail_base_length.ToString());
                    foundation_symbol.LookupParameter("鋼軌軌距").SetValueString(gauge.ToString());
                    int start_sta = (Convert.ToInt32(Convert.ToDouble(data_list[0].station.Split('+').Last())) + Int32.Parse(data_list[0].station.Split('+').First()) * 1000) * 1000;
                    int station = start_sta;

                    for (int i = 0; i < st_foundation_num; i++)
                    {
                        try
                        {
                            int b = (station - start_sta) / 1000;
                            int c = station % 1000;
                            Line toward = Line.CreateBound(data_list[b].start_point, data_list[b + 1].start_point);
                            XYZ put_point = data_list[b].start_point + c / 304.8 * toward.Direction;
                            ori_sub_t.Start();
                            FamilyInstance every_foundation = ori_doc.Create.NewFamilyInstance(put_point, foundation_symbol, StructuralType.NonStructural);
                            double toward_slope = toward.Direction.Y / toward.Direction.X;
                            Line rotate_axis = Line.CreateBound(put_point, put_point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(ori_doc, every_foundation.Id, rotate_axis, Math.Atan(toward_slope) + Math.PI / 2);
                            ori_sub_t.Commit();
                            ori_sub_t.Start();
                            //再根據超高轉角旋轉一次
                            double angle = data_list[b].super_high_angle + (data_list[b + 1].super_high_angle - data_list[b].super_high_angle) * (c / 1000.0);
                            every_foundation.Location.Rotate(toward, -Math.PI * angle / 180.0);
                            ori_sub_t.Commit();
                            station += base_dis;
                            //旋轉z軸
                            if (data_list[b + 1].start_point.Z != data_list[b].start_point.Z)
                            {
                                XYZ fa_point = new XYZ(put_point.X + toward.Direction.Y, put_point.Y - toward.Direction.X, put_point.Z);
                                Line sec_rotate_axis = Line.CreateBound(put_point, fa_point);
                                double Z_angle = put_point.Z / data_list[b + 1].start_point.Z;
                                //try { ElementTransformUtils.RotateElement(doc, every_foundation.Id, sec_rotate_axis, Math.Acos(Z_angle)); } catch { }
                            }
                        }
                        catch (Exception e) {  }
                    }

                    int th_support_dis = int.Parse(tb_properties.third_bracket_spacing.ToString());
                    int th_support_num = ((data_list.Count - 1) * 1000 / th_support_dis) + 1;
                    FamilySymbol support_symbol = new FilteredElementCollector(ori_doc).OfClass(typeof(FamilySymbol))
                                                  .Cast<FamilySymbol>().ToList().Where(x => x.Name == "第三軌支架_自").First();
                    support_symbol.Activate();
                    //設定第三軌參數
                    support_symbol.LookupParameter("第三軌與鋼軌距").SetValueString(tb_properties.third_steel_between_dis.ToString());
                    support_symbol.LookupParameter("寬").SetValueString(tb_properties.third_bracket_width.ToString());
                    support_symbol.LookupParameter("長").SetValueString(tb_properties.third_bracket_length.ToString());
                    support_symbol.LookupParameter("與鋼軌距").SetValueString(tb_properties.bracket_steel_between_dis.ToString());
                    support_symbol.LookupParameter("鋼軌軌距").SetValueString(gauge.ToString());
                    station = start_sta;

                    for (int i = 0; i < th_support_num; i++)
                    {
                        try
                        {
                            int b = (station - start_sta) / 1000;
                            int c = station % 1000;
                            ori_sub_t.Start();
                            
                            Line toward = Line.CreateBound(data_list[b].start_point, data_list[b + 1].start_point);
                            XYZ put_point = data_list[b].start_point + c / 304.8 * toward.Direction;
                            FamilyInstance every_support = ori_doc.Create.NewFamilyInstance(put_point, support_symbol, StructuralType.NonStructural);

                            double toward_slope = toward.Direction.Y / toward.Direction.X; // 台大
                            Line rotate_axis = Line.CreateBound(put_point, put_point + XYZ.BasisZ);
                            //ElementTransformUtils.RotateElement(ori_doc, every_support.Id, rotate_axis, Math.Atan(toward_slope) + Math.PI / 2);

                            // 培文改
                            XYZ start = data_list[b].start_point;
                            XYZ end = data_list[b + 1].start_point;
                            double radians = (Math.Atan2((end.Y - start.Y), (end.X - start.X)) * 180.0 / Math.PI) * (Math.PI / 180); // 弧度
                            ElementTransformUtils.RotateElement(ori_doc, every_support.Id, rotate_axis, radians + Math.PI / 2);

                            ori_sub_t.Commit();
                            ori_sub_t.Start();
                            //再根據超高轉角旋轉一次
                            double angle = data_list[b].super_high_angle + (data_list[b + 1].super_high_angle - data_list[b].super_high_angle) * (c / 1000.0);

                            //檢查支架應在左側或右側
                            double for_check_station = station / 1000.0;
                            for (int k = 0; k < setting_Station.third_rail_station.Count; k++)
                            {
                                int start_st = int.Parse(setting_Station.third_rail_station[k][1]);
                                int end_st = int.Parse(setting_Station.third_rail_station[k][2]);

                                //if (for_check_station >= start_st && for_check_station < end_st)
                                //{
                                    if (setting_Station.third_rail_station[k][0] == "右側") { isRight = true; }
                                    else { isRight = false; }
                                //}
                            }

                            // 培文改
                            //double offset = UnitUtils.ConvertToInternalUnits(tb_properties.third_steel_between_dis, DisplayUnitType.DUT_MILLIMETERS); // 2020
                            double offset = UnitUtils.ConvertToInternalUnits(tb_properties.third_steel_between_dis, UnitTypeId.Millimeters); // 2024
                            XYZ normal = XYZ.BasisZ.CrossProduct((end - start).Normalize()); // 垂直的向量
                            Vector vector = new Vector(normal.X, normal.Y); // 方向向量
                            vector = GetVectorOffset(vector, offset);
                            XYZ centerXYZ = new XYZ((start.X + end.X) / 2, (start.Y + end.Y) / 2, (start.Z + end.Z) / 2);
                            XYZ offsetPoint = new XYZ(centerXYZ.X + vector.X, centerXYZ.Y + vector.Y, centerXYZ.Z);
                            string side = GetPointSideOnPlane(data_list[b].start_point, data_list[b + 1].start_point, offsetPoint, XYZ.BasisZ);
                            if (isRight == true && side.Equals("左側")) { every_support.Location.Rotate(rotate_axis, Math.PI); }
                            else if (isRight == false && side.Equals("右側")) { every_support.Location.Rotate(rotate_axis, Math.PI); }
                            every_support.Location.Rotate(toward, -Math.PI * angle / 180.0);
                            //DrawLine(ori_doc, toward);
                            //DrawLine(ori_doc, Line.CreateBound(centerXYZ, offsetPoint));


                            //if (isRight == true) { every_support.Location.Rotate(toward, Math.PI * angle / 180.0); }
                            //else { every_support.Location.Rotate(toward, -Math.PI * angle / 180.0); }
                            //if ((data_list[1].start_point.X - data_list[0].start_point.X) > 0) { every_support.Location.Rotate(rotate_axis, Math.PI); }

                            //every_support.Location.Rotate(rotate_axis, Math.PI);

                            ori_sub_t.Commit();
                            station += th_support_dis;

                            ////旋轉z軸
                            //if (data_list[b + 1].start_point.Z != data_list[b].start_point.Z)
                            //{
                            //    XYZ fa_point = new XYZ(put_point.X + toward.Direction.Y, put_point.Y - toward.Direction.X, put_point.Z);
                            //    Line sec_rotate_axis = Line.CreateBound(put_point, fa_point);
                            //    double Z_angle = put_point.Z / data_list[b + 1].start_point.Z;

                            //    //try { ElementTransformUtils.RotateElement(doc, every_support.Id, sec_rotate_axis, Math.Acos(Z_angle)); } catch { }
                            //}
                        }
                        catch (Exception e) { ori_sub_t.Commit(); }
                    }
                    ori_t.Commit();
                }
                TaskDialog.Show("Revit", "程序處理完畢。");
            }
            catch (Exception e) { TaskDialog.Show("error", e.StackTrace + e.Message); }
        }
        /// <summary>
        /// 辨識放置點在線段的左側或右側
        /// </summary>
        /// <param name="A"></param>
        /// <param name="B"></param>
        /// <param name="P"></param>
        /// <param name="planeNormal"></param>
        /// <returns></returns>
        public string GetPointSideOnPlane(XYZ A, XYZ B, XYZ P, XYZ planeNormal)
        {
            XYZ v = B - A;
            XYZ w = P - A;

            XYZ cross = v.CrossProduct(w);
            double dot = cross.DotProduct(planeNormal);

            if (dot > 1e-9) { return "左側"; }                
            else if (dot < -1e-9) { return "右側"; }
            else { return "在線上"; }
        }
        /// <summary>
        /// 取得向量偏移的距離
        /// </summary>
        /// <param name="vector"></param>
        /// <param name="newLength"></param>
        /// <returns></returns>
        private Vector GetVectorOffset(Vector vector, double newLength)
        {
            double length = Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y); // 計算向量長度
            if (length != 0) { vector = new Vector(vector.X / length * newLength, vector.Y / length * newLength); }
            return vector;
        }

        public class FailureHandler : IFailuresPreprocessor
        {
            public string ErrorMessage { set; get; }
            public string ErrorSeverity { set; get; }

            public FailureHandler()
            {
                ErrorMessage = "";
                ErrorSeverity = "";
            }

            public FailureProcessingResult PreprocessFailures(
              FailuresAccessor failuresAccessor)
            {
                IList<FailureMessageAccessor> failureMessages
                  = failuresAccessor.GetFailureMessages();

                foreach (FailureMessageAccessor
                  failureMessageAccessor in failureMessages)
                {
                    // We're just deleting all of the warning level 
                    // failures and rolling back any others

                    FailureDefinitionId id = failureMessageAccessor
                      .GetFailureDefinitionId();

                    try
                    {
                        ErrorMessage = failureMessageAccessor
                          .GetDescriptionText();
                    }
                    catch
                    {
                        ErrorMessage = "Unknown Error";
                    }

                    try
                    {
                        FailureSeverity failureSeverity
                          = failureMessageAccessor.GetSeverity();

                        ErrorSeverity = failureSeverity.ToString();

                        if (failureSeverity == FailureSeverity.Warning)
                        {
                            failuresAccessor.DeleteWarning(
                              failureMessageAccessor);
                        }
                        else
                        {
                            return FailureProcessingResult
                              .ProceedWithRollBack;
                        }
                    }
                    catch
                    {
                    }
                }
                return FailureProcessingResult.Continue;
            }
        }

        //// 台大
        //public SketchPlane Sketch_plain(Document doc, XYZ start, XYZ end)
        //{
        //    SketchPlane sk = null;
        //    XYZ v = end - start;
        //    double dxy = Math.Abs(v.X) + Math.Abs(v.Y);
        //    XYZ w = (dxy > 0.00000001) ? XYZ.BasisY : XYZ.BasisZ;
        //    XYZ norm = v.CrossProduct(w).Normalize();
        //    Plane geomPlane = Plane.CreateByNormalAndOrigin(norm, start);
        //    sk = SketchPlane.Create(doc, geomPlane);

        //    return sk;
        //}

        /// <summary>
        /// 培文改寫
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public SketchPlane Sketch_plain(Document doc, XYZ start, XYZ end)
        {
            SketchPlane sk = null;
            XYZ v = end - start;
            XYZ w = XYZ.BasisZ;
            XYZ norm = v.CrossProduct(w).Normalize();
            Plane geomPlane = Plane.CreateByNormalAndOrigin(norm, start);
            sk = SketchPlane.Create(doc, geomPlane);

            return sk;
        }

        public SketchPlane write_linear_sketchplane(Document doc, XYZ start, XYZ end)
        {
            SketchPlane sk = null;
            Line toward = Line.CreateBound(start, end);
            XYZ second = new XYZ(start.X + toward.Direction.Y, start.Y - toward.Direction.X, start.Z);
            Plane geomPlane = Plane.CreateByThreePoints(start, end, second);

            sk = SketchPlane.Create(doc, geomPlane);

            return sk;
        }

        public int StationToInt(String station)
        {
            int a = int.Parse(station.Split('+')[0]);
            int b = int.Parse(station.Split('+')[1].Split('.')[0]);
            return a * 1000 + b;
        }
        /// <summary>
        /// 3D視圖中畫模型線
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="curve"></param>
        private void DrawLine(Document doc, Curve curve)
        {
            try
            {
                Line line = Line.CreateBound(curve.Tessellate()[0], curve.Tessellate()[curve.Tessellate().Count - 1]);
                XYZ normal = new XYZ(line.Direction.Z - line.Direction.Y, line.Direction.X - line.Direction.Z, line.Direction.Y - line.Direction.X); // 使用與線不平行的任意向量
                Plane plane = Plane.CreateByNormalAndOrigin(normal, curve.Tessellate()[0]);
                SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
                ModelCurve modelCurve = doc.Create.NewModelCurve(line, sketchPlane);
            }
            catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }

    class FamilyOption : IFamilyLoadOptions
    {
        bool IFamilyLoadOptions.OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }

        bool IFamilyLoadOptions.OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }

    public class track_bed
    {
        public string name;
        public double x;
        public double y;

        public track_bed(string track_bed_name)
        {
            if (track_bed_name == "標準道床")
            {
                name = "標準道床_左";
            }
            else if (track_bed_name == "平版式道床")
            {
                name = track_bed_name;
            }
            else if (track_bed_name == "浮動式道床")
            {
                name = track_bed_name;
            }
            else
            {
                TaskDialog.Show("error", "Wrong track bed name");
            }
        }
    }
    // 更新專案內Family的參數
    public class LoadOptions : IFamilyLoadOptions
    {
        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }
        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = true;
            return true;
        }
    }
    // 關閉警示視窗 
    public class CloseWarnings : IFailuresPreprocessor
    {
        FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            String transactionName = failuresAccessor.GetTransactionName();
            IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();
            if (fmas.Count == 0) { return FailureProcessingResult.Continue; }
            else { failuresAccessor.DeleteAllWarnings(); }
            //if (transactionName.Equals("EXEMPLE"))
            //{
            //    foreach (FailureMessageAccessor fma in fmas)
            //    {
            //        if (fma.GetSeverity() == FailureSeverity.Error)
            //        {
            //            failuresAccessor.DeleteAllWarnings();
            //            return FailureProcessingResult.ProceedWithRollBack;
            //        }
            //        else { failuresAccessor.DeleteWarning(fma); }
            //    }
            //}
            //else
            //{
            //    foreach (FailureMessageAccessor fma in fmas) { failuresAccessor.DeleteAllWarnings(); }
            //}
            return FailureProcessingResult.Continue;
        }
    }
}