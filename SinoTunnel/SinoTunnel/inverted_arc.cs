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
using System.Windows;

namespace SinoTunnel
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    class inverted_arc : IExternalEventHandler
    {
        string path;
        double center_point;

        public void Execute(UIApplication app)
        {
            try
            {
                path = Form1.path;
                Document doc = app.ActiveUIDocument.Document;

                readfile rf = new readfile();
                rf.read_tunnel_point(); // 隧道線形 (DN)
                rf.read_tunnel_point2(); // 隧道線形 (UP)
                rf.read_point(); // 軌道線形 (DN)：座標
                rf.read_point2(); // 軌道線形 (UP)：座標
                rf.read_target_station(); // 軌道線形(DN)：里程起終點
                rf.read_target_station2(); // 軌道線形 (UP)：里程起終點
                rf.read_properties(); // 模型輸入資料

                IList<IList<data_object>> all_data_list_tunnel = new List<IList<data_object>>(); // 隧道線形
                IList<IList<data_object>> all_data_list = new List<IList<data_object>>(); // 軌道線形
                IList<setting_station> all_station_setting = new List<setting_station>(); // 模型輸入資料

                all_data_list_tunnel.Add(rf.data_list_tunnel); // 隧道線形 (DN)
                all_data_list_tunnel.Add(rf.data_list_tunnel2); // 隧道線形 (UP)
                all_data_list.Add(rf.data_list); // 軌道線形 (DN)
                all_data_list.Add(rf.data_list2); // 軌道線形 (UP)
                all_station_setting.Add(rf.setting_Station);
                all_station_setting.Add(rf.setting_Station2);

                center_point = rf.properties.center_point;

                //DrawTrackAndTunnelLines(doc, all_data_list_tunnel, all_data_list); //  畫軌道、隧道的線

                for (int times = 0; times < all_data_list.Count; times++)
                {
                    UIDocument edit_uidoc = app.OpenAndActivateDocument(path + "仰拱\\仰拱.rfa");
                    Document edit_doc = edit_uidoc.Document;
                    Transaction t = new Transaction(edit_doc);
                    // 關閉警示視窗
                    FailureHandlingOptions options = t.GetFailureHandlingOptions();
                    options.SetClearAfterRollback(true);
                    options.SetFailuresPreprocessor(new CloseWarnings());
                    t.SetFailureHandlingOptions(options);

                    int station_delta = 0;
                    int station_total = 0;
                    int gutter_station_delta = 0;
                    int gutter_station_total = 0;

                    //set the depth of the PVC pipe
                    double float_gutter_side_depth = rf.properties.float_top_height + rf.properties.float_gutter_depth + rf.properties.float_gutter_radius - rf.properties.pvc_radius;
                    double standard_gutter_side_depth = rf.properties.flat_top_height + rf.properties.u_depth + rf.properties.u_radius - rf.properties.pvc_radius;
                    int gutter_index = 0;

                    //PVC管熱鍍鋅蓋長
                    double hollow_length = rf.properties.pvc_zn_cover_length / 1000;

                    //every 15 m has one hollow
                    double hollow_gap = rf.properties.pvc_zn_cover_pitch / 1000;

                    //determine the side of the walk way
                    bool isRight = true;
                    if (all_station_setting[times].walk_way_station[0][0] == "右側") { isRight = true; }
                    else if (all_station_setting[times].walk_way_station[0][0] == "左側") { isRight = false; }

                    //先對輪廓做參數化
                    List<double> offset_list = new List<double>();
                    for (int i = 0; i < all_data_list_tunnel[times].Count; i++)
                    {
                        if (times == 1) { offset_list.Add(all_data_list_tunnel[times][i].offset); }
                        else { offset_list.Add(all_data_list_tunnel[times][i].offset); }
                    }

                    //取得偏移量種類，以確定要建置幾種輪廓
                    List<double> offset_distinct = offset_list.Distinct().ToList();

                    //製作不同圓心深度 PVC管排水溝 和 PVC管明溝 的輪廓
                    List<double> gutter_depth = new List<double>();
                    List<double> hollow_depth = new List<double>();
                    string str_temp = "";
                    foreach (string[] gutter in all_station_setting[times].gutter_station)
                    {
                        if (gutter[0] == "PVC管排水溝")
                        {
                            Int32.TryParse(gutter[1], out int gutter_start_station);
                            Int32.TryParse(gutter[2], out int gutter_end_station);
                            gutter_station_delta = gutter_end_station - gutter_start_station;

                            double start_station_depth = 0;
                            double end_station_depth = 0;

                            //determine the PVC pipe will go up or go down
                            if (str_temp == "標準排水溝")
                            {
                                start_station_depth = standard_gutter_side_depth;
                                end_station_depth = float_gutter_side_depth;
                            }
                            else if (str_temp == "浮動式排水溝")
                            {
                                start_station_depth = float_gutter_side_depth;
                                end_station_depth = standard_gutter_side_depth;
                            }

                            double station_depth_delta = end_station_depth - start_station_depth;
                            double n = station_depth_delta / gutter_station_delta;

                            for (double i = -1; i <= gutter_station_delta + 1; i++)
                            {
                                gutter_depth.Add(start_station_depth + i * n);
                                if (i % hollow_gap == 0) { hollow_depth.Add(start_station_depth + i * n); }
                            }
                            List<double> gutter_depth_distinct = gutter_depth.Distinct().ToList();

                            t.Start("輸入參數");

                            FamilySymbol gutter_profile = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                          .Cast<FamilySymbol>().ToList().Where(x => x.Name == gutter[0]).First();
                            input_properties(gutter_profile, rf.properties);

                            FamilySymbol hollow_profile = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                          .Cast<FamilySymbol>().ToList().Where(x => x.Name == "PVC管明溝").First();
                            input_properties(hollow_profile, rf.properties);

                            t.Commit();

                            //根據不同圓心深度製造輪廓
                            t.Start("製造輪廓");
                            for (int i = 0; i < gutter_depth_distinct.Count; i++)
                            {
                                try
                                {
                                    FamilySymbol new_gutter_profile = gutter_profile.Duplicate(gutter[0] + "_PVC管圓心深度=" + gutter_depth_distinct[i].ToString()) as FamilySymbol;
                                    new_gutter_profile.LookupParameter("PVC管圓心深度").SetValueString(gutter_depth_distinct[i].ToString());
                                }
                                catch (Exception e) { }
                            }
                            for (int i = 0; i < hollow_depth.Count; i++)
                            {
                                try
                                {
                                    FamilySymbol new_hollow_profile = hollow_profile.Duplicate("PVC管明溝_PVC管圓心深度=" + hollow_depth[i].ToString()) as FamilySymbol;
                                    new_hollow_profile.LookupParameter("深度").SetValueString(hollow_depth[i].ToString());
                                }
                                catch (Exception e) { }
                            }
                            t.Commit();
                            //輪廓製造完畢
                        }
                        str_temp = gutter[0];
                    }

                    foreach (string[] ia_name in all_station_setting[times].inverted_arc_station) //ia = inverted arc = = (?
                    {
                        string inverted_arc_name = ia_name[0];

                        List<FamilySymbol> invert_profiles = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                             .Cast<FamilySymbol>().ToList().Where(x => x.Name.Contains(inverted_arc_name)).ToList();

                        //根據不同偏移量製造輪廓
                        t.Start("製造輪廓");

                        //先為仰拱寫入試算表之參數
                        foreach (FamilySymbol invert_profile in invert_profiles) { input_properties(invert_profile, rf.properties); }

                        for (int i = 0; i < offset_distinct.Count; i++)
                        {
                            try
                            {
                                FamilySymbol invert_profile_positive = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                                       .Cast<FamilySymbol>().ToList().Where(x => x.Name == inverted_arc_name + "_正偏移量").First();
                                FamilySymbol invert_profile_negative = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                                       .Cast<FamilySymbol>().ToList().Where(x => x.Name == inverted_arc_name + "_負偏移量").First();
                                FamilySymbol new_invert_profile = null;

                                if (offset_distinct[i] >= 0)
                                {
                                    try { new_invert_profile = invert_profile_positive.Duplicate(inverted_arc_name + "_偏移量=" + offset_distinct[i].ToString()) as FamilySymbol; }
                                    catch { }
                                }
                                else
                                {
                                    try { new_invert_profile = invert_profile_negative.Duplicate(inverted_arc_name + "_偏移量=" + offset_distinct[i].ToString()) as FamilySymbol; }
                                    catch { }
                                }
                                try 
                                { 
                                    new_invert_profile.LookupParameter("偏移量").Set((Math.Abs(offset_distinct[i]) / 304.8));
                                    double walkway_edge_to_rail_center_dis = rf.properties.walkway_edge_to_rail_center_dis + offset_distinct[i];
                                    try { new_invert_profile.LookupParameter("走道邊緣與隧道中心距離").SetValueString(walkway_edge_to_rail_center_dis.ToString()); } // 培文改寫
                                    catch (Exception) { new_invert_profile.LookupParameter("走道邊緣與軌道中心距離").SetValueString(walkway_edge_to_rail_center_dis.ToString()); } // 舊版Excel&族群
                                }
                                catch { }
                                //if (times == 1 || 0) { new_invert_profile.LookupParameter("偏移量").Set((Math.Abs(offset_distinct[i]) / 304.8)); }
                                //else { new_invert_profile.LookupParameter("偏移量").Set((offset_distinct[i] - offset_distinct[i]) / 304.8); }
                            }
                            catch (Exception e) { TaskDialog.Show("error", e.Message + e.StackTrace); }
                        }
                        t.Commit();
                        //輪廓製造完畢

                        Int32.TryParse(ia_name[1], out int start_station);
                        Int32.TryParse(ia_name[2], out int end_station);

                        station_delta = end_station - start_station;
                        station_total += station_delta;

                        //開始掃略混成
                        SubTransaction subtran = new SubTransaction(edit_doc);
                        t.Start("掃略混成");
                        for (int i = station_total - station_delta + 1; i <= station_total; i++)
                        {
                            CombinableElementArray elementArray = new CombinableElementArray();

                            //仰拱實體
                            bool isSolid = true;
                            Line path = Line.CreateBound(all_data_list_tunnel[times][i].start_point, all_data_list_tunnel[times][i - 1].start_point);

                            FamilySymbol start_famsy = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                       .ToList().Where(x => x.Name == inverted_arc_name + "_偏移量=" + offset_list[i - 1].ToString()).First();
                            FamilySymbol end_famsy = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                     .ToList().Where(x => x.Name == inverted_arc_name + "_偏移量=" + offset_list[i].ToString()).First();

                            SweepProfile sweepProfile_top = edit_doc.Application.Create.NewFamilySymbolProfile(start_famsy);
                            SweepProfile sweepProfile_bottom = edit_doc.Application.Create.NewFamilySymbolProfile(end_famsy);

                            SketchPlane sketchPlane = Sketch_plain(edit_doc, all_data_list_tunnel[times][i - 1].start_point, all_data_list_tunnel[times][i].start_point);
                            SweptBlend sweptBlend = edit_doc.FamilyCreate.NewSweptBlend(isSolid, path, sketchPlane, sweepProfile_bottom, sweepProfile_top);
                            //if (Math.Round(path.Direction.Z, 5) != 0)
                            //{
                            //    sweptBlend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString("180");
                            //    sweptBlend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString("180");
                            //}
                            // 培文改寫
                            sweptBlend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString("90");
                            sweptBlend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString("90");

                            //開始建立空心元件
                            ReferenceArray multi_path = new ReferenceArray();
                            ReferenceArray multi_path_tunnel = new ReferenceArray();
                            List<ElementId> mc_list = new List<ElementId>();
                            Line single_path = Line.CreateBound(all_data_list[times][i].start_point, all_data_list[times][i - 1].start_point);
                            Line single_path_tunnel = Line.CreateBound(all_data_list_tunnel[times][i].start_point, all_data_list_tunnel[times][i - 1].start_point);

                            XYZ new_start = all_data_list[times][i].start_point - single_path.Length * single_path.Direction;
                            XYZ new_end = all_data_list[times][i - 1].start_point + single_path.Length * single_path.Direction;
                            Line n_single_path = Line.CreateBound(new_start, new_end);
                            ModelCurve m_curve = edit_doc.FamilyCreate.NewModelCurve(n_single_path, Sketch_plain(edit_doc, all_data_list[times][i].start_point, all_data_list[times][i - 1].start_point));
                            mc_list.Add(m_curve.Id);
                            multi_path.Append(m_curve.GeometryCurve.Reference);

                            XYZ new_start_tunnel = all_data_list_tunnel[times][i].start_point - single_path_tunnel.Length * single_path_tunnel.Direction;
                            XYZ new_end_tunnel = all_data_list_tunnel[times][i - 1].start_point + single_path_tunnel.Length * single_path_tunnel.Direction;
                            Line n_single_path_tunnel = Line.CreateBound(new_start_tunnel, new_end_tunnel);
                            ModelCurve m_curve_tunnel = edit_doc.FamilyCreate.NewModelCurve(n_single_path_tunnel, Sketch_plain(edit_doc, all_data_list_tunnel[times][i].start_point, all_data_list_tunnel[times][i - 1].start_point));
                            mc_list.Add(m_curve_tunnel.Id);
                            multi_path_tunnel.Append(m_curve_tunnel.GeometryCurve.Reference);

                            Sweep empty_bed = null;
                            //空心平版式道床
                            if (inverted_arc_name != "浮動式道床仰拱")
                            {
                                FamilySymbol bed_profile = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                           .Cast<FamilySymbol>().ToList().Where(x => x.Name == "平版式道床").First();
                                input_properties(bed_profile, rf.properties);
                                SweepProfile bed_sweep_profile = edit_doc.Application.Create.NewFamilySymbolProfile(bed_profile);
                                empty_bed = edit_doc.FamilyCreate.NewSweep(false, multi_path, bed_sweep_profile, 0, ProfilePlaneLocation.Start);
                                empty_bed.LookupParameter("角度").Set((Math.PI / 180) * -90);
                            }

                            //空心電纜溝槽
                            FamilySymbol powercable_gutter_profile = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                                     .Cast<FamilySymbol>().ToList().Where(x => x.Name == "空心電纜溝槽").First();
                            input_properties(powercable_gutter_profile, rf.properties);
                            // 培文改寫
                            double offset = rf.properties.walkway_edge_to_rail_center_dis/* - offset_list[i - 1]*/; // 走道邊緣與隧道中心距離 - 偏移量
                            try { powercable_gutter_profile.LookupParameter("走道邊緣與隧道中心距離").Set(offset / 304.8); }
                            catch (Exception) { powercable_gutter_profile.LookupParameter("走道邊緣與軌道中心距離").Set(offset / 304.8); }

                            //if (times == 1) { powercable_gutter_profile.LookupParameter("偏移量").Set((offset_distinct[0] - offset_distinct[0]) / 304.8); }
                            //else { powercable_gutter_profile.LookupParameter("偏移量").Set(offset_distinct[0] / 304.8); }

                            SweepProfile powercable_gutter_sweep_profile = edit_doc.Application.Create.NewFamilySymbolProfile(powercable_gutter_profile);
                            //以隧道線形建置
                            //Sweep empty_powercable_gutter = edit_doc.FamilyCreate.NewSweep(false, multi_path_tunnel, powercable_gutter_sweep_profile, 0, ProfilePlaneLocation.Start); // 台大

                            // 培文改寫, 使用CurveArray建立空心電纜溝槽
                            CurveArray curveArray = new CurveArray();
                            Curve curve = Line.CreateBound(new_start_tunnel, new_end_tunnel);
                            curveArray.Append(curve);
                            SketchPlane powercable_gutter_profile_sketchPlane = Sketch_plain(edit_doc, all_data_list_tunnel[times][i].start_point, all_data_list_tunnel[times][i - 1].start_point);
                            Sweep empty_powercable_gutter = edit_doc.FamilyCreate.NewSweep(false, curveArray, powercable_gutter_profile_sketchPlane, powercable_gutter_sweep_profile, 0, ProfilePlaneLocation.Start);

                            empty_powercable_gutter.LookupParameter("角度").Set((Math.PI / 180) * -90);
                            //空心排水溝
                            string gutter_name = all_station_setting[times].gutter_station[gutter_index][0];
                            Int32.TryParse(all_station_setting[times].gutter_station[gutter_index][1], out int gutter_start_station);
                            Int32.TryParse(all_station_setting[times].gutter_station[gutter_index][2], out int gutter_end_station);

                            gutter_station_delta = gutter_end_station - gutter_start_station;

                            //空心PVC管仰拱
                            if (gutter_name == "PVC管排水溝")
                            {
                                FamilySymbol semicircle_gutter_profile = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                                         .Cast<FamilySymbol>().ToList().Where(x => x.Name == "PVC管仰拱").First();
                                input_properties(semicircle_gutter_profile, rf.properties);
                                SweepProfile semicircle_gutter_sweep_profile = edit_doc.Application.Create.NewFamilySymbolProfile(semicircle_gutter_profile);
                                Sweep empty_semicircle_gutter = edit_doc.FamilyCreate.NewSweep(false, multi_path, semicircle_gutter_sweep_profile, 0, ProfilePlaneLocation.Start);

                                elementArray.Append(empty_semicircle_gutter);
                            }

                            //以掃略混成方式建立PVC管排水溝
                            if (gutter_name == "PVC管排水溝")
                            {
                                double start_station_depth = 0;
                                double end_station_depth = 0;

                                //determine the PVC pipe will go up or go down
                                if (all_station_setting[times].gutter_station[gutter_index - 1][0] == "標準排水溝")
                                {
                                    start_station_depth = standard_gutter_side_depth;
                                    end_station_depth = float_gutter_side_depth;
                                }
                                else if (all_station_setting[times].gutter_station[gutter_index - 1][0] == "浮動式排水溝")
                                {
                                    start_station_depth = float_gutter_side_depth;
                                    end_station_depth = standard_gutter_side_depth;
                                }

                                double station_depth_delta = end_station_depth - start_station_depth;
                                double n = station_depth_delta / gutter_station_delta;

                                double gutter_start_station_depth = start_station_depth + n * (i - 2 - gutter_station_total);
                                double gutter_end_station_depth = start_station_depth + n * (i + 1 - gutter_station_total);

                                FamilySymbol gutter_profile_start = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                                    .ToList().Where(x => x.Name == gutter_name + "_PVC管圓心深度=" + gutter_start_station_depth.ToString()).First();

                                FamilySymbol gutter_profile_end = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                                  .ToList().Where(x => x.Name == gutter_name + "_PVC管圓心深度=" + gutter_end_station_depth.ToString()).First();

                                SweepProfile sweepProfile_top2 = edit_doc.Application.Create.NewFamilySymbolProfile(gutter_profile_start);
                                SweepProfile sweepProfile_bottom2 = edit_doc.Application.Create.NewFamilySymbolProfile(gutter_profile_end);

                                SketchPlane sketchPlane2 = Sketch_plain(edit_doc, all_data_list[times][i - 1].start_point, all_data_list[times][i].start_point);
                                SweptBlend empty_gutter = edit_doc.FamilyCreate.NewSweptBlend(false, n_single_path, sketchPlane2, sweepProfile_bottom2, sweepProfile_top2);

                                if (Math.Round(single_path.Direction.Z, 5) != 0)
                                {
                                    empty_gutter.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString("180");
                                    empty_gutter.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString("180");
                                }

                                //PVC管明溝
                                if (i - 1 != gutter_station_total && i != gutter_station_total + gutter_station_delta)
                                {
                                    ReferenceArray multi_path_middle = new ReferenceArray();
                                    List<ElementId> mc_list_middle = new List<ElementId>();

                                    //void hollow for two inverted arc piece
                                    if ((i - gutter_station_total) % hollow_gap == 0)
                                    {
                                        XYZ middle = all_data_list[times][i].start_point;

                                        XYZ new_start_middle = middle - hollow_length * single_path.Direction;
                                        XYZ new_end_middle = middle + hollow_length * single_path.Direction;
                                        Line n_single_path_middle = Line.CreateBound(new_start_middle, new_end_middle);

                                        ModelCurve m_curve_from_middle = edit_doc.FamilyCreate.NewModelCurve(n_single_path_middle, Sketch_plain(edit_doc, all_data_list[times][i].start_point, all_data_list[times][i - 1].start_point));
                                        mc_list_middle.Add(m_curve_from_middle.Id);
                                        multi_path_middle.Append(m_curve_from_middle.GeometryCurve.Reference);

                                        gutter_end_station_depth = start_station_depth + n * (i - gutter_station_total);

                                        FamilySymbol hollow_profile = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                                      .ToList().Where(x => x.Name == "PVC管明溝_PVC管圓心深度=" + gutter_end_station_depth.ToString()).First();
                                        SweepProfile hollow_sweep_profile = edit_doc.Application.Create.NewFamilySymbolProfile(hollow_profile);
                                        Sweep empty_hollow = edit_doc.FamilyCreate.NewSweep(false, multi_path_middle, hollow_sweep_profile, 0, ProfilePlaneLocation.Start);

                                        elementArray.Append(empty_hollow);
                                    }
                                    else if ((i - gutter_station_total) % hollow_gap == 1)
                                    {
                                        XYZ middle = all_data_list[times][i - 1].start_point;

                                        XYZ new_start_middle = middle - hollow_length * single_path.Direction;
                                        XYZ new_end_middle = middle + hollow_length * single_path.Direction;
                                        Line n_single_path_middle = Line.CreateBound(new_start_middle, new_end_middle);

                                        ModelCurve m_curve_from_middle = edit_doc.FamilyCreate.NewModelCurve(n_single_path_middle, Sketch_plain(edit_doc, all_data_list[times][i].start_point, all_data_list[times][i - 1].start_point));
                                        mc_list_middle.Add(m_curve_from_middle.Id);
                                        multi_path_middle.Append(m_curve_from_middle.GeometryCurve.Reference);

                                        gutter_start_station_depth = start_station_depth + n * (i - 1 - gutter_station_total);

                                        FamilySymbol hollow_profile = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                                      .ToList().Where(x => x.Name == "PVC管明溝_PVC管圓心深度=" + gutter_start_station_depth.ToString()).First();
                                        SweepProfile hollow_sweep_profile = edit_doc.Application.Create.NewFamilySymbolProfile(hollow_profile);
                                        Sweep empty_hollow = edit_doc.FamilyCreate.NewSweep(false, multi_path_middle, hollow_sweep_profile, 0, ProfilePlaneLocation.Start);

                                        elementArray.Append(empty_hollow);
                                    }
                                    edit_doc.Delete(mc_list_middle);
                                }
                                elementArray.Append(empty_gutter);
                            }
                            //以掃略方式建立其他種類排水溝
                            else
                            {
                                FamilySymbol gutter_profile = new FilteredElementCollector(edit_doc)
                                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == gutter_name).First();

                                if (gutter_name == "浮動式排水溝") { input_properties(gutter_profile, rf.properties); }
                                else if (gutter_name == "標準排水溝") { input_properties(gutter_profile, rf.properties); }
                                SweepProfile gutter_sweep_profile = edit_doc.Application.Create.NewFamilySymbolProfile(gutter_profile);
                                Sweep empty_gutter = edit_doc.FamilyCreate.NewSweep(false, multi_path, gutter_sweep_profile, 0, ProfilePlaneLocation.Start);
                                empty_gutter.LookupParameter("角度").Set((Math.PI / 180) * -90);

                                elementArray.Append(empty_gutter);

                                //若為標準排水溝，建立標準排水溝蓋板
                                if (gutter_name == "標準排水溝")
                                {
                                    Line path_2 = Line.CreateBound(all_data_list[times][i].start_point, all_data_list[times][i - 1].start_point);
                                    FamilySymbol fms_2 = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                                         .Cast<FamilySymbol>().ToList().Where(x => x.Name == "標準排水溝蓋板").First();
                                    input_properties(fms_2, rf.properties);
                                    SweepProfile sweepProfile_2 = edit_doc.Application.Create.NewFamilySymbolProfile(fms_2);
                                    SketchPlane sketchPlane_2 = Sketch_plain(edit_doc, all_data_list[times][i].start_point, all_data_list[times][i - 1].start_point);
                                    SweptBlend sweptBlend_2 = edit_doc.FamilyCreate.NewSweptBlend(isSolid, single_path, sketchPlane_2, sweepProfile_2, sweepProfile_2);
                                    // 培文改寫
                                    sweptBlend_2.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).Set((Math.PI / 180) * -90);
                                    sweptBlend_2.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).Set((Math.PI / 180) * -90);
                                }
                            }

                            //let gutter to next type of gutter
                            if (i == gutter_station_total + gutter_station_delta)
                            {
                                gutter_index += 1;
                                gutter_station_total += gutter_station_delta;
                            }

                            //set the side of the walk way
                            if (isRight == true)
                            {
                                sweptBlend.get_Parameter(BuiltInParameter.PROFILE1_FLIPPED_HOR).Set(0);
                                sweptBlend.get_Parameter(BuiltInParameter.PROFILE2_FLIPPED_HOR).Set(0);
                                empty_powercable_gutter.get_Parameter(BuiltInParameter.PROFILE_FLIPPED_HOR).Set(0);
                            }
                            else if (isRight == false)
                            {
                                sweptBlend.get_Parameter(BuiltInParameter.PROFILE1_FLIPPED_HOR).Set(1);
                                sweptBlend.get_Parameter(BuiltInParameter.PROFILE2_FLIPPED_HOR).Set(1);
                                empty_powercable_gutter.get_Parameter(BuiltInParameter.PROFILE_FLIPPED_HOR).Set(1);
                            }

                            //combine elements
                            elementArray.Append(sweptBlend);
                            if (inverted_arc_name != "浮動式道床仰拱") { elementArray.Append(empty_bed); }                                
                            elementArray.Append(empty_powercable_gutter);
                            //elementArray.Append(empty_semicircle_gutter);
                            edit_doc.CombineElements(elementArray);
                            edit_doc.Delete(mc_list);
                        }
                        t.Commit();
                    }
                    t.Start("sweep");
                    for (int i = 1; i <= station_total; i++)
                    {
                        bool isSolid = true;
                        Line path = Line.CreateBound(all_data_list_tunnel[times][i].start_point, all_data_list_tunnel[times][i - 1].start_point);

                        FamilySymbol fms = new FilteredElementCollector(edit_doc).OfClass(typeof(FamilySymbol))
                                           .Cast<FamilySymbol>().ToList().Where(x => x.Name == "混凝土蓋板").First();

                        input_properties(fms, rf.properties);
                        // 培文改寫
                        double offset = rf.properties.walkway_edge_to_rail_center_dis/* - offset_list[i - 1]*/; // 走道邊緣與隧道中心距離 - 偏移量
                        try { fms.LookupParameter("走道邊緣與隧道中心距離").Set(offset / 304.8); }
                        catch (Exception) { fms.LookupParameter("走道邊緣與軌道中心距離").Set(offset / 304.8); }
                        //if (times == 1)
                        //{
                        //    fms.LookupParameter("偏移量").Set((offset_distinct[0] - offset_distinct[0]) / 304.8);
                        //}
                        //else
                        //{
                        //    fms.LookupParameter("偏移量").Set(offset_distinct[0] / 304.8);
                        //}

                        SweepProfile sweepProfile = edit_doc.Application.Create.NewFamilySymbolProfile(fms);

                        SketchPlane sketchPlane = Sketch_plain(edit_doc, all_data_list_tunnel[times][i - 1].start_point, all_data_list_tunnel[times][i].start_point);

                        SweptBlend sweptBlend = edit_doc.FamilyCreate.NewSweptBlend(isSolid, path, sketchPlane, sweepProfile, sweepProfile);

                        //if (Math.Round(path.Direction.Z, 5) != 0)
                        //{
                        //    sweptBlend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString("180");
                        //    sweptBlend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString("180");
                        //}
                        // 培文改寫
                        sweptBlend.get_Parameter(BuiltInParameter.PROFILE1_ANGLE).SetValueString("90");
                        sweptBlend.get_Parameter(BuiltInParameter.PROFILE2_ANGLE).SetValueString("90");

                        //set the side of the walk way
                        if (isRight == true)
                        {
                            sweptBlend.get_Parameter(BuiltInParameter.PROFILE1_FLIPPED_HOR).Set(0);
                            sweptBlend.get_Parameter(BuiltInParameter.PROFILE2_FLIPPED_HOR).Set(0);
                        }
                        else if (isRight == false)
                        {
                            sweptBlend.get_Parameter(BuiltInParameter.PROFILE1_FLIPPED_HOR).Set(1);
                            sweptBlend.get_Parameter(BuiltInParameter.PROFILE2_FLIPPED_HOR).Set(1);
                        }
                    }
                    t.Commit();
                    //掃掠完成

                    //另存檔案
                    SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };
                    edit_doc.SaveAs(path + "仰拱\\仰拱final_" + times.ToString() + ".rfa", saveAsOptions);
                    app.OpenAndActivateDocument(doc.PathName);
                    // 更新專案內Family的參數
                    try { Family family = edit_doc.LoadFamily(doc, new LoadOptions()); }
                    catch(Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                    edit_doc.Close();

                    //載入專案內
                    Transaction pro_t = new Transaction(doc);
                    pro_t.SetFailureHandlingOptions(options);
                    pro_t.Start("載入族群");
                    try { doc.LoadFamily(path + "仰拱\\仰拱final_" + times.ToString() + ".rfa"); }
                    catch (Exception e) { TaskDialog.Show("error", e.Message + e.StackTrace); }

                    FamilySymbol invert_arc = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                                              .Cast<FamilySymbol>().ToList().Where(x => x.Name == "仰拱final_" + times.ToString()).First();
                    invert_arc.Activate();

                    // 如果專案中未放置仰拱才擺放
                    try
                    {
                        FamilyInstance findIns = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().Where(x => x.Name.Equals(invert_arc.Name)).Cast<FamilyInstance>().FirstOrDefault();
                        if (findIns == null) { FamilyInstance object_acr = doc.Create.NewFamilyInstance(XYZ.Zero, invert_arc, StructuralType.NonStructural); }
                    }
                    catch(Exception) { FamilyInstance object_acr = doc.Create.NewFamilyInstance(XYZ.Zero, invert_arc, StructuralType.NonStructural); }
                    pro_t.Commit();
                }
            }
            catch (Exception e) { TaskDialog.Show("error", e.Message + e.StackTrace); }
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
            //if (norm.Z > 0) { norm = -norm; }
            Plane geomPlane = Plane.CreateByNormalAndOrigin(norm, start);
            sk = SketchPlane.Create(doc, geomPlane);

            return sk;
        }

        public void input_properties(FamilySymbol fms, properties_object properties)
        {
            string name = fms.Name;
            if (name.Contains("標準道床仰拱") || name.Contains("浮動式道床仰拱"))
            {
                fms.LookupParameter("仰拱洩水斜率").Set(properties.inverted_arc_slope);
                fms.LookupParameter("隧道中心").SetValueString(properties.center_point.ToString());
                string top_ele = (name.Contains("標準")) ? properties.standerd_top_height.ToString()
                    : properties.float_top_height.ToString();
                fms.LookupParameter("仰拱頂部高程").SetValueString(top_ele);
                fms.LookupParameter("隧道內半徑").SetValueString((properties.inner_diameter / 2.0).ToString());
                try { fms.LookupParameter("走道邊緣與隧道中心距離").SetValueString(properties.walkway_edge_to_rail_center_dis.ToString()); } // 培文改寫
                catch (Exception) { fms.LookupParameter("走道邊緣與軌道中心距離").SetValueString(properties.walkway_edge_to_rail_center_dis.ToString()); } // 舊版Excel&族群
                fms.LookupParameter("連結處厚度").SetValueString(properties.connection_thick.ToString());
                fms.LookupParameter("走道頂部高程").SetValueString(properties.walkway_top_elevation.ToString());
                fms.LookupParameter("走道突出底").SetValueString(properties.walkway_protrusion_bottom.ToString());
                fms.LookupParameter("走道突出寬").SetValueString(properties.walkway_protrusion_width.ToString());
                fms.LookupParameter("走道突出深").SetValueString(properties.walkway_protrusion_depth.ToString());
            }
            else if (name == "標準排水溝")
            {
                /*
                fms.LookupParameter("U型圓心深度").SetValueString(ia_properties.u_depth.ToString());
                fms.LookupParameter("U型排水溝半徑").SetValueString(ia_properties.u_radius.ToString());
                fms.LookupParameter("蓋板凹槽寬度").SetValueString(ia_properties.u_groove_width.ToString());
                fms.LookupParameter("蓋板凹槽深度").SetValueString(ia_properties.u_cover_thick.ToString());
                */
                fms.LookupParameter("U形排水溝蓋板寬").Set(properties.u_cover_width / 304.8);
                fms.LookupParameter("U形排水溝蓋板長").Set(properties.u_cover_length / 304.8);
                fms.LookupParameter("U形排水溝蓋板厚").Set(properties.u_cover_thick / 304.8);
                fms.LookupParameter("U形排水溝凹槽寬").Set(properties.u_groove_width / 304.8);
                fms.LookupParameter("U形排水溝深度").Set(properties.u_depth / 304.8);
                fms.LookupParameter("U形排水溝半徑").Set(properties.u_radius / 304.8);
                fms.LookupParameter("U形排水溝長邊距").Set(properties.u_short_side_dis / 304.8);
                fms.LookupParameter("U形排水溝短邊距").Set(properties.float_gutter_depth / 304.8);
                fms.LookupParameter("U形排水溝熱鍍鋅扁鋼厚").Set(properties.u_zn_steel_thick / 304.8);
                fms.LookupParameter("U形排水溝熱鍍鋅扁鋼長間距").Set(properties.u_zn_steel_long_pitch / 304.8);
                fms.LookupParameter("U形排水溝熱鍍鋅扁鋼短間距").Set(properties.u_zn_steel_short_pitch / 304.8);
                fms.LookupParameter("U形排水溝熱鍍鋅間距").Set(properties.u_zn_steel_pitch / 304.8);
                fms.LookupParameter("高程").SetValueString(properties.standerd_top_height.ToString());
            }
            else if (name == "浮動式排水溝")
            {
                fms.LookupParameter("浮動式道床明溝半徑").Set(properties.float_gutter_radius / 304.8);
                fms.LookupParameter("浮動式道床明溝深").Set(properties.float_gutter_depth / 304.8);
                fms.LookupParameter("浮動式道床仰拱頂部高程").SetValueString(properties.float_top_height.ToString());
            }
            else if (name == "PVC管明溝")
            {
                fms.LookupParameter("PVC管半徑").Set(properties.pvc_radius / 304.8);
                fms.LookupParameter("PVC管厚度").Set(properties.pvc_thick / 304.8);
                fms.LookupParameter("PVC管明溝半徑").Set(properties.pvc_gutter_radius / 304.8);
                fms.LookupParameter("PVC管明溝凹槽寬").Set(properties.pvc_gutter_witdh / 304.8);
                fms.LookupParameter("PVC管熱鍍鋅蓋寬").Set(properties.pvc_zn_cover_width / 304.8);
                fms.LookupParameter("PVC管熱鍍鋅蓋長").Set(properties.pvc_zn_cover_length / 304.8);
                fms.LookupParameter("PVC管熱鍍鋅蓋厚").Set(properties.pvc_zn_cover_thick / 304.8);
                fms.LookupParameter("PVC管熱鍍鋅蓋鋼厚").Set(properties.pvc_zn_cover_steel_thick / 304.8);
                fms.LookupParameter("PVC管熱鍍鋅蓋長間距").Set(properties.pvc_zn_cover_long_pitch / 304.8);
                fms.LookupParameter("PVC管熱鍍鋅蓋短間距").Set(properties.pvc_zn_cover_short_pitch / 304.8);
                fms.LookupParameter("PVC管熱鍍鋅蓋間距").Set(properties.pvc_zn_cover_pitch / 304.8);
                fms.LookupParameter("高程").SetValueString(properties.flat_top_height.ToString());
            }
            else if (name == "空心電纜溝槽")
            {
                fms.LookupParameter("走道頂對線形高程").Set(properties.walkway_top_elevation / 304.8);
                try { fms.LookupParameter("走道邊緣與隧道中心距離").Set(properties.walkway_edge_to_rail_center_dis / 304.8); } // 培文改寫
                catch (Exception) { fms.LookupParameter("走道邊緣與軌道中心距離").Set(properties.walkway_edge_to_rail_center_dis / 304.8); } // 舊版Excel&族群
                fms.LookupParameter("走道電纜溝槽與走道距離").Set(properties.cable_distance / 304.8);
                fms.LookupParameter("走道電纜溝槽上底寬").Set(properties.cable_top_width / 304.8);
                fms.LookupParameter("走道電纜溝槽下底寬").Set(properties.cable_bottom_width / 304.8);
                fms.LookupParameter("走道電纜溝槽深度").Set(properties.cable_depth / 304.8);
                fms.LookupParameter("走道電纜溝槽混凝土蓋板厚").Set(properties.cable_cover_thick / 304.8);
                fms.LookupParameter("走道電纜溝槽混凝土蓋板寬").Set(properties.cable_cover_width / 304.8);
                fms.LookupParameter("隧道中心點").Set(center_point / 304.8);
            }
            else if (name == "混凝土蓋板")
            {
                fms.LookupParameter("走道斜率").Set(properties.walkway_slope);
                fms.LookupParameter("走道頂對線形高程").Set(properties.walkway_top_elevation / 304.8);
                try { fms.LookupParameter("走道邊緣與隧道中心距離").Set(properties.walkway_edge_to_rail_center_dis / 304.8); } // 培文改寫
                catch (Exception) { fms.LookupParameter("走道邊緣與軌道中心距離").Set(properties.walkway_edge_to_rail_center_dis / 304.8); } // 舊版Excel&族群
                fms.LookupParameter("走道電纜溝槽混凝土蓋板寬").Set(properties.cable_cover_width / 304.8);
                fms.LookupParameter("走道電纜溝槽混凝土蓋板突出邊距").Set(properties.cable_cover_stick_out_dis / 304.8);
                fms.LookupParameter("走道電纜溝槽混凝土蓋板突出厚").Set(properties.cable_cover_stick_out_thick / 304.8);
                fms.LookupParameter("走道電纜溝槽混凝土蓋板突出寬").Set(properties.cable_cover_stick_out_width / 304.8);
                fms.LookupParameter("走道電纜溝槽混凝土蓋板厚").Set(properties.cable_cover_thick / 304.8);
                fms.LookupParameter("走道電纜溝槽混凝土蓋板長").Set(properties.cable_cover_length / 304.8);
                fms.LookupParameter("隧道中心點").Set(center_point / 304.8);
            }
            else if (name == "標準排水溝蓋板")
            {
                fms.LookupParameter("U形排水溝蓋板寬").Set(properties.u_cover_width / 304.8);
                fms.LookupParameter("U形排水溝蓋板長").Set(properties.u_cover_length / 304.8);
                fms.LookupParameter("U形排水溝蓋板厚").Set(properties.u_cover_thick / 304.8);
                fms.LookupParameter("U形排水溝凹槽寬").Set(properties.u_groove_width / 304.8);
                fms.LookupParameter("U形排水溝深度").Set(properties.u_depth / 304.8);
                fms.LookupParameter("U形排水溝半徑").Set(properties.u_radius / 304.8);
                fms.LookupParameter("U形排水溝長邊距").Set(properties.u_short_side_dis / 304.8);
                fms.LookupParameter("U形排水溝短邊距").Set(properties.float_gutter_depth / 304.8);
                fms.LookupParameter("U形排水溝熱鍍鋅扁鋼厚").Set(properties.u_zn_steel_thick / 304.8);
                fms.LookupParameter("U形排水溝熱鍍鋅扁鋼長間距").Set(properties.u_zn_steel_long_pitch / 304.8);
                fms.LookupParameter("U形排水溝熱鍍鋅扁鋼短間距").Set(properties.u_zn_steel_short_pitch / 304.8);
                fms.LookupParameter("U形排水溝熱鍍鋅間距").Set(properties.u_zn_steel_pitch / 304.8);
                fms.LookupParameter("高程").SetValueString(properties.standerd_top_height.ToString());
            }
            else if (name == "PVC管仰拱")
            {
                fms.LookupParameter("PVC管仰拱半徑").Set(properties.pvc_inverted_arc_radius / 304.8);
            }
            else if (name == "平版式道床")
            {
                fms.LookupParameter("寬度").SetValueString(properties.flat_width.ToString());
                fms.LookupParameter("仰拱頂部").SetValueString(properties.flat_top_height.ToString());
                fms.LookupParameter("高程").SetValueString(properties.flat_elevation.ToString());
            }
            else if (name == "PVC管排水溝")
            {
                fms.LookupParameter("PVC管半徑").Set(properties.pvc_radius / 304.8);
            }
        }
        /// <summary>
        ///  畫軌道、隧道的線
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="all_data_list"></param>
        /// <param name="all_data_list_tunnel"></param>
        private void DrawTrackAndTunnelLines(Document doc, IList<IList<data_object>> all_data_list_tunnel, IList<IList<data_object>> all_data_list)
        {
            List<Curve> curves = new List<Curve>();

            foreach (IList<data_object> data_list_tunnel in all_data_list_tunnel) // 隧道線形
            {
                for (int i = 0; i < data_list_tunnel.Count; i++)
                {
                    try { Line line = Line.CreateBound(data_list_tunnel[i].start_point, data_list_tunnel[i + 1].start_point); curves.Add(line); }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }
            }
            foreach (IList<data_object> data_list in all_data_list) // 軌道線形
            {
                for (int i = 0; i < data_list.Count; i++)
                {
                    try { Line line = Line.CreateBound(data_list[i].start_point, data_list[i + 1].start_point); curves.Add(line); }
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }
            }
            using (Transaction trans = new Transaction(doc, "畫線"))
            {
                trans.Start();
                foreach (Curve curve in curves)
                {
                    try { DrawLine(doc, curve); } // 3D視圖中畫模型線
                    catch (Exception ex) { string error = ex.Message + "\n" + ex.ToString(); }
                }
                trans.Commit();
            }
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
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}