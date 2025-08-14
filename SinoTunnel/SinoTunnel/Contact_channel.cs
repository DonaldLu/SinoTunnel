
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using DataObject;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SinoTunnel
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class Contact_channel : IExternalEventHandler
    {

        string path;
        public void Execute(UIApplication app)
        {
            path = Form1.path;
            try
            {
                Document pro_doc = app.ActiveUIDocument.Document;
                UIDocument uidoc = app.OpenAndActivateDocument(path + "聯絡通道\\通道預覽.rfa"); ;
                Document doc = uidoc.Document;
                Transaction t = new Transaction(doc);
                //// 關閉警示視窗
                //FailureHandlingOptions options = t.GetFailureHandlingOptions();
                //CloseWarnings closeWarnings = new CloseWarnings();
                //options.SetClearAfterRollback(true);
                //options.SetFailuresPreprocessor(closeWarnings);
                //t.SetFailureHandlingOptions(options);

                //設定參數

                //尋找隧道點
                readfile rf = new readfile();
                rf.read_tunnel_point(); //下行線 
                rf.read_tunnel_point2(); //上行線
                contact_channel_properties cc_properties = rf.read_contact_tunnel(); //讀取幾何參數化資訊
                int r = rf.data_list_tunnel[0].id; //起始里程點之id，用於抓取空心隧道之里程點 
                int r2 = rf.data_list_tunnel2[0].id; 

                for (int i = 0; i <= rf.data_list_cd_channel.Count - 1; i++)
                {
                    //軌道模式 (0:平行聯絡通道 1:高差聯絡通道, 2:疊式聯絡通道)
                    int module = cc_properties.tunnel_type[i];

                    //起點&終點座標
                    int channel_p = i;
                    XYZ start_initial = rf.data_list_cd_channel2[channel_p].start_point; //上行點 //如果聯絡通道inverted，改成data_list_cd_channel
                    XYZ end_initial = rf.data_list_cd_channel[channel_p].start_point; //下行點 //如果聯絡通道inverted，改成data_list_cd_channel2

                    int target = rf.data_list_cd_channel[channel_p].id;
                    int target2 = rf.data_list_cd_channel2[channel_p].id; 

                    //delete the previous sweeps
                    Transaction tr = new Transaction(doc);
                    //// 關閉警示視窗
                    //options = tr.GetFailureHandlingOptions();
                    //closeWarnings = new CloseWarnings();
                    //options.SetClearAfterRollback(true);
                    //options.SetFailuresPreprocessor(closeWarnings);
                    //tr.SetFailureHandlingOptions(options);
                    tr.Start("delete the previous sweeps");
                    IList<Sweep> remaining_sweeps = new FilteredElementCollector(doc).
                    OfClass(typeof(Sweep)).Where(x => x.Name.Contains("掃掠")).Cast<Sweep>().ToList();
                    if (remaining_sweeps.Count != 0)
                    {
                        foreach (Sweep sweep in remaining_sweeps)
                        {
                            doc.Delete(sweep.Id);
                        }
                    }
                    tr.Commit();


                    //修改&掃掠輪廓
                    Transaction t_contour = new Transaction(doc);
                    //// 關閉警示視窗
                    //options = t_contour.GetFailureHandlingOptions();
                    //closeWarnings = new CloseWarnings();
                    //options.SetClearAfterRollback(true);
                    //options.SetFailuresPreprocessor(closeWarnings);
                    //t_contour.SetFailureHandlingOptions(options);
                    t_contour.Start("modify the contour");

                    FamilySymbol big_one = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                        .ToList().Where(x => x.Name == "聯絡通道輪廓").First();
                    input_properties(big_one, cc_properties, i);
                    SweepProfile big_one_profile = doc.Application.Create.NewFamilySymbolProfile(big_one);


                    FamilySymbol small_one = new FilteredElementCollector(doc)
                       .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                       .ToList().Where(x => x.Name == "聯絡通道內輪廓").First();
                    input_properties(small_one, cc_properties, i);
                    SweepProfile small_one_profile = doc.Application.Create.NewFamilySymbolProfile(small_one);

                    FamilySymbol door_info = new FilteredElementCollector(doc)
                       .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                       .ToList().Where(x => x.Name == "防火門框").First();
                    input_properties(door_info, cc_properties, i);
                    SweepProfile door_arc_profile = doc.Application.Create.NewFamilySymbolProfile(door_info);

                    FamilySymbol tunnel_fs = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                        .ToList().Where(x => x.Name == "空心隧道").First();
                    input_properties(tunnel_fs, cc_properties, i);
                    SweepProfile tunnel_profile = doc.Application.Create.NewFamilySymbolProfile(tunnel_fs);


                    t_contour.Commit();

                    //隧道底部高程(mm) 971mm level up 
                    double elevation_adjust = big_one.LookupParameter("聯絡通道底部高程").AsDouble();
                    XYZ start = new XYZ(start_initial.X, start_initial.Y, start_initial.Z + elevation_adjust); //modifying the elevaiton of start and end points 
                    XYZ end = new XYZ(end_initial.X, end_initial.Y, end_initial.Z + elevation_adjust);

                    //開孔範圍軸向深度(mm) 500mm
                    double connect_length = big_one.LookupParameter("聯絡通道開孔範圍軸向深度").AsDouble() * 0.3048 * 1000;

                    //防火門厚度(mm) 200mm
                    double door_width = door_info.LookupParameter("防火門厚度").AsDouble() * 0.3048 * 1000;

                    //隧道半徑(mm) 2800mm + 125mm
                    double tunnel_rad = tunnel_fs.LookupParameter("隧道半徑").AsDouble() * 0.3048 * 1000 + 125;


                    ////API輸入
                    //高差聯絡通道層級長度(mm)
                    double head_landing_length = double.NaN;
                    double rear_landing_length = double.NaN;
                    double middle_landing_length = double.NaN;

                    bool mid_exist = false;

                    if (cc_properties.path_levels[i] == 3) //有中間平台
                    {
                        mid_exist = true;
                        head_landing_length = cc_properties.level_one[i]; //前段平台長度(mm) 1000mm
                        middle_landing_length = cc_properties.level_two[i]; //中間段平台長度(mm) 5000mm
                        rear_landing_length = cc_properties.level_three[i]; //後段平台長度(mm) 1000mm
                    }
                    else if (cc_properties.path_levels[i] == 2) //無中間平台
                    {
                        mid_exist = false;
                        head_landing_length = cc_properties.level_one[i];
                        rear_landing_length = cc_properties.level_two[i];
                    }
                    else if (cc_properties.path_levels[i] == 1) //平行聯絡通道
                    {
                        // head_landing_length
                    }

                    //聯絡通道防火門的所在平台編號(0,1,2 or 0,1 or 0) default = 0
                    int door_on_landing = cc_properties.door_levels[i] - 1;

                    //聯絡通道防火門離平台起點之距離(mm) 
                    double door_dis = cc_properties.door_dis[i];



                    //C.疊式聯絡通道 
                    double landing_length = 5000;
                    double circle_rad = 2000;
                    //double tunnel_top = 130000;
                    //double tunnel_bottom = 90000;


                    if (module == 0) //建立高差聯絡通道
                    {
                        //線段部分，先建立modelcurve，再加進一個list刪除他
                        t.Start("建立聯絡通道");
                        Line ch_path = Line.CreateBound(start, end);
                        //與channel path 垂直
                        XYZ test = ch_path.Direction.CrossProduct(XYZ.BasisZ).Normalize();
                        double first_part_length = (tunnel_rad + connect_length) / 1000 / 0.3048; // mm to inch
                        double door_length = (tunnel_rad + connect_length + door_dis) / 1000 / 0.3048;
                        XYZ Node_1 = start + first_part_length * ch_path.Direction;
                        XYZ Node_2 = end - first_part_length * ch_path.Direction;
                        XYZ Node_mid_door = start + door_length * ch_path.Direction;
                        List<ElementId> pre_del_list = new List<ElementId>();
                        Line head = Line.CreateBound(start, Node_1);
                        Line middle = Line.CreateBound(Node_1, Node_2);
                        Line nail = Line.CreateBound(Node_2, end);
                        cc_properties.level_one[i] = middle.Length * 1000 * 0.3048;
                        SketchPlane sketchPlane = Sketch_plain(doc, start, end);
                        ReferenceArray re_head = new ReferenceArray();
                        ReferenceArray re_middle = new ReferenceArray();
                        ReferenceArray re_nail = new ReferenceArray();
                        ModelCurve m_head = doc.FamilyCreate.NewModelCurve(head, sketchPlane);
                        ModelCurve m_middle = doc.FamilyCreate.NewModelCurve(middle, sketchPlane);
                        ModelCurve m_nail = doc.FamilyCreate.NewModelCurve(nail, sketchPlane);
                        re_head.Append(m_head.GeometryCurve.Reference);
                        re_middle.Append(m_middle.GeometryCurve.Reference);
                        re_nail.Append(m_nail.GeometryCurve.Reference);
                        pre_del_list.Add(m_head.Id);
                        pre_del_list.Add(m_middle.Id);
                        pre_del_list.Add(m_nail.Id);

                        Sweep head_part = doc.FamilyCreate.NewSweep(true, re_head, big_one_profile, 0, ProfilePlaneLocation.Start);
                        Sweep middle_part = doc.FamilyCreate.NewSweep(true, re_middle, small_one_profile, 0, ProfilePlaneLocation.Start);
                        Sweep nail_part = doc.FamilyCreate.NewSweep(true, re_nail, big_one_profile, 0, ProfilePlaneLocation.Start);

                        //掃掠空心段
                        //空心段取目標點位及前後兩個點位，取五段作路徑
                        ReferenceArray hollow_dn = new ReferenceArray();
                        ReferenceArray hollow_up = new ReferenceArray();

                        for (int k = (target - r) - 3; k < (target - r) + 4; k++) //excel index has been changed, sketchplane should be modified cuz the slope of tunnel(down)
                        {
                            Line ho_dn = Line.CreateBound(rf.data_list_tunnel[k].start_point, rf.data_list_tunnel[k + 1].start_point);
                            SketchPlane sketchPlane_dn = Sketch_Plane(doc, rf.data_list_tunnel[k].start_point, rf.data_list_tunnel[k + 1].start_point);
                            ModelCurve mho_dn = doc.FamilyCreate.NewModelCurve(ho_dn, sketchPlane_dn);
                            hollow_dn.Append(mho_dn.GeometryCurve.Reference);
                            pre_del_list.Add(mho_dn.Id);
                        }

                        for (int k2 = (target2 - r2) - 3; k2 < (target2 - r2) + 4; k2++) //excel index has been changed, sketchplane should be modified cuz the slope of tunnel(up)
                        {
                            Line ho_up = Line.CreateBound(rf.data_list_tunnel2[k2].start_point, rf.data_list_tunnel2[k2 + 1].start_point);
                            SketchPlane sketchPlane_up = Sketch_Plane(doc, rf.data_list_tunnel2[k2].start_point, rf.data_list_tunnel2[k2 + 1].start_point);
                            ModelCurve mho_up = doc.FamilyCreate.NewModelCurve(ho_up, sketchPlane_up);
                            hollow_up.Append(mho_up.GeometryCurve.Reference);
                            pre_del_list.Add(mho_up.Id);
                        }

                        //建立空心實體
                        Sweep hollow_dn_tunnel = doc.FamilyCreate.NewSweep(false, hollow_dn, tunnel_profile, 0, ProfilePlaneLocation.Start);
                        Sweep hollow_up_tunnel = doc.FamilyCreate.NewSweep(false, hollow_up, tunnel_profile, 0, ProfilePlaneLocation.Start);

                        //切割元件
                        CombinableElementArray elementArray = new CombinableElementArray();
                        elementArray.Append(hollow_up_tunnel); //如果聯絡通道inverted，改成hollow_dn_tunnel
                        elementArray.Append(head_part);
                        doc.CombineElements(elementArray);

                        elementArray.Clear();
                        elementArray.Append(hollow_dn_tunnel); //如果聯絡通道inverted，改成hollow_up_tunnel
                        elementArray.Append(nail_part);
                        doc.CombineElements(elementArray);

                        elementArray.Clear(); //
                        t.Commit();

                        Transaction door_t = new Transaction(doc);
                        //// 關閉警示視窗
                        //options = door_t.GetFailureHandlingOptions();
                        //closeWarnings = new CloseWarnings();
                        //options.SetClearAfterRollback(true);
                        //options.SetFailuresPreprocessor(closeWarnings);
                        //door_t.SetFailureHandlingOptions(options);
                        door_t.Start("加入防火門");

                        XYZ door_mid_point = Node_mid_door;
                        XYZ Node_door = start + first_part_length * ch_path.Direction;
                        XYZ Node_5 = door_mid_point - ch_path.Direction * (door_width / 2) / 1000 / 0.3048;
                        XYZ Node_6 = door_mid_point + ch_path.Direction * (door_width / 2) / 1000 / 0.3048;
                        ReferenceArray door_array = new ReferenceArray();
                        Line line_door = Line.CreateBound(Node_5, Node_6);
                        ModelCurve mho_door = doc.FamilyCreate.NewModelCurve(line_door, sketchPlane);
                        door_array.Append(mho_door.GeometryCurve.Reference);
                        pre_del_list.Add(mho_door.Id);
                        Sweep door_part = doc.FamilyCreate.NewSweep(true, door_array, door_arc_profile, 0, ProfilePlaneLocation.Start);

                        doc.Delete(pre_del_list);
                        door_t.Commit();

                    }

                    if (module == 1) //建立高差聯絡通道
                    {
                        //線段部分，先建立modelcurve，再加進一個list刪除他
                        t.Start("建立高差聯絡通道");

                        Line ch_path = Line.CreateBound(new XYZ(start.X, start.Y, 0), new XYZ(end.X, end.Y, 0));

                        double first_part_length = (tunnel_rad + connect_length) / 1000 / 0.3048; // mm to inch
                        double head_landing_part_length = double.NaN;
                        double rear_landing_part_length = double.NaN;

                        if (start.Z < end.Z) //聯絡通道上下行高程順序對調 
                        {
                            head_landing_part_length = (tunnel_rad + connect_length + head_landing_length) / 1000 / 0.3048;
                            rear_landing_part_length = (tunnel_rad + connect_length + rear_landing_length) / 1000 / 0.3048;
                        }
                        else
                        {
                            rear_landing_part_length = (tunnel_rad + connect_length + head_landing_length) / 1000 / 0.3048;
                            head_landing_part_length = (tunnel_rad + connect_length + rear_landing_length) / 1000 / 0.3048;
                        }

                        List<ElementId> pre_del_list = new List<ElementId>();
                        ReferenceArray re_head = new ReferenceArray();
                        ReferenceArray re_stairs_n_landings = new ReferenceArray();
                        ReferenceArray re_nail = new ReferenceArray();
                        Dictionary<int, (XYZ, XYZ)> sec_of_landing = new Dictionary<int, (XYZ, XYZ)>();

                        if (mid_exist) //有中間平台
                        {
                            XYZ Node_1 = start + first_part_length * ch_path.Direction;
                            XYZ Node_2 = end - first_part_length * ch_path.Direction;
                            XYZ Node_3 = start + head_landing_part_length * ch_path.Direction;
                            XYZ Node_4 = end - rear_landing_part_length * ch_path.Direction;

                            XYZ middle_point = (new XYZ(start.X, start.Y, start.Z) + new XYZ(end.X, end.Y, end.Z)) / 2;
                            double mid_landing_length = middle_landing_length / 1000 / 0.3048;
                            XYZ Node_5 = middle_point - ch_path.Direction * mid_landing_length / 2;
                            XYZ Node_6 = middle_point + ch_path.Direction * mid_landing_length / 2;

                            double sin_A = Math.Abs(Node_3.Z - Node_5.Z) / Line.CreateBound(Node_3, Node_5).ApproximateLength;
                            double sin_B = Math.Abs(Node_6.Z - Node_4.Z) / Line.CreateBound(Node_6, Node_4).ApproximateLength;
                            double first_lim_dis = double.NaN;
                            double second_lim_dis = double.NaN;

                            if (start.Z < end.Z) // 13 low 24 high
                            {
                                first_lim_dis = sin_A * small_one.LookupParameter("聯絡通道上部圓弧半徑").AsDouble();
                                second_lim_dis = sin_B * small_one.LookupParameter("聯絡通道下部垂直長度").AsDouble();

                                if (first_part_length != head_landing_part_length) //head_landing_part_length calibrates
                                {
                                    Node_3 = start + (head_landing_part_length - first_lim_dis) * ch_path.Direction;
                                    head_landing_part_length = head_landing_part_length - first_lim_dis;
                                }

                                if (first_part_length != rear_landing_part_length) //rear_landing_part_length calibrates
                                {
                                    Node_4 = end - (rear_landing_part_length + second_lim_dis) * ch_path.Direction;
                                    rear_landing_part_length = rear_landing_part_length + second_lim_dis;
                                }
                            }
                            else // 13 high 24 low
                            {
                                second_lim_dis = sin_B * small_one.LookupParameter("聯絡通道上部圓弧半徑").AsDouble();
                                first_lim_dis = sin_A * small_one.LookupParameter("聯絡通道下部垂直長度").AsDouble();

                                if (first_part_length != head_landing_part_length) //head_landing_part_length calibrates
                                {
                                    Node_3 = start + (head_landing_part_length + first_lim_dis) * ch_path.Direction;
                                    head_landing_part_length = head_landing_part_length + first_lim_dis;
                                }

                                if (first_part_length != rear_landing_part_length) //rear_landing_part_length calibrates
                                {
                                    Node_4 = end - (rear_landing_part_length - second_lim_dis) * ch_path.Direction;
                                    rear_landing_part_length = rear_landing_part_length - second_lim_dis;
                                }
                            }

                            bool valid_dis = true;
                            while (valid_dis) //tunnel距離調整
                            {
                                sin_A = Math.Abs(Node_3.Z - Node_5.Z) / Line.CreateBound(Node_3, Node_5).ApproximateLength;
                                sin_B = Math.Abs(Node_6.Z - Node_4.Z) / Line.CreateBound(Node_6, Node_4).ApproximateLength;
                                first_lim_dis = double.NaN;
                                second_lim_dis = double.NaN;

                                if (start.Z < end.Z) //聯絡通道上下行高程順序對調
                                {
                                    first_lim_dis = sin_A * small_one.LookupParameter("聯絡通道總直徑").AsDouble();
                                    second_lim_dis = sin_B * (small_one.LookupParameter("聯絡通道下部垂直長度").AsDouble()
                                        + small_one.LookupParameter("聯絡通道下部厚度").AsDouble());
                                }
                                else
                                {
                                    second_lim_dis = sin_B * small_one.LookupParameter("聯絡通道總直徑").AsDouble();
                                    first_lim_dis = sin_A * (small_one.LookupParameter("聯絡通道下部垂直長度").AsDouble()
                                        + small_one.LookupParameter("聯絡通道下部厚度").AsDouble());
                                }

                                if (first_lim_dis + first_part_length > head_landing_part_length) //tunnel前端距離變更 node3
                                {
                                    head_landing_part_length = (first_lim_dis + first_part_length) * 1.001;
                                    Node_3 = start + head_landing_part_length * ch_path.Direction;
                                }

                                if (second_lim_dis + first_part_length > rear_landing_part_length) //tunnel前端距離變更 node4
                                {
                                    rear_landing_part_length = (second_lim_dis + first_part_length) * 1.001;
                                    Node_4 = end - rear_landing_part_length * ch_path.Direction;
                                }

                                if ((first_lim_dis + first_part_length < head_landing_part_length)
                                    && (second_lim_dis + first_part_length < rear_landing_part_length)) //tunnel前後中端距離確認
                                {
                                    /*
                                    TaskDialog.Show("hi", ((head_landing_part_length - first_part_length) * 304.8).ToString() + "\n" 
                                        +(mid_landing_length * 304.8).ToString() + "\n"
                                        + ((rear_landing_part_length - first_part_length) * 304.8).ToString());
                                    */
                                    break;
                                }
                            }

                            Line head = Line.CreateBound(start, Node_1);
                            Line nail = Line.CreateBound(Node_2, end);
                            Line head_landing = Line.CreateBound(Node_1, Node_3);
                            Line nail_landing = Line.CreateBound(Node_4, Node_2);

                            SketchPlane sketchPlane_A = Sketch_plain(doc, start, Node_1);
                            SketchPlane sketchPlane_C = Sketch_plain(doc, Node_2, end);

                            ModelCurve m_head = doc.FamilyCreate.NewModelCurve(head, sketchPlane_A);
                            ModelCurve m_head_landing = doc.FamilyCreate.NewModelCurve(head_landing, sketchPlane_A);
                            ModelCurve m_nail = doc.FamilyCreate.NewModelCurve(nail, sketchPlane_C);
                            ModelCurve m_nail_landing = doc.FamilyCreate.NewModelCurve(nail_landing, sketchPlane_C);

                            re_head.Append(m_head.GeometryCurve.Reference);
                            re_nail.Append(m_nail.GeometryCurve.Reference);
                            re_stairs_n_landings.Append(m_head_landing.GeometryCurve.Reference);
                            re_stairs_n_landings.Append(m_nail_landing.GeometryCurve.Reference);

                            pre_del_list.Add(m_head.Id);
                            pre_del_list.Add(m_head_landing.Id);
                            pre_del_list.Add(m_nail.Id);
                            pre_del_list.Add(m_nail_landing.Id);

                            if (start.Z < end.Z)
                            {
                                sec_of_landing.Add(0, (Node_1, Node_3));
                                sec_of_landing.Add(1, (Node_5, Node_6));
                                sec_of_landing.Add(2, (Node_4, Node_2));
                            }
                            else
                            {
                                sec_of_landing.Add(0, (Node_2, Node_4));
                                sec_of_landing.Add(1, (Node_6, Node_5));
                                sec_of_landing.Add(2, (Node_3, Node_1));
                            }

                            Line mid_landing = Line.CreateBound(Node_5, Node_6);
                            Line head_stair = Line.CreateBound(Node_3, Node_5);
                            Line tail_stair = Line.CreateBound(Node_6, Node_4);
                            SketchPlane sketchPlane_B = Sketch_plain(doc, Node_5, Node_6);
                            SketchPlane sketchPlane_D = Sketch_Plane(doc, Node_3, Node_5);
                            SketchPlane sketchPlane_E = Sketch_Plane(doc, Node_6, Node_4);
                            ModelCurve m_mid_landing = doc.FamilyCreate.NewModelCurve(mid_landing, sketchPlane_B);
                            ModelCurve m_head_stair = doc.FamilyCreate.NewModelCurve(head_stair, sketchPlane_D);
                            ModelCurve m_tail_stair = doc.FamilyCreate.NewModelCurve(tail_stair, sketchPlane_E);
                            re_stairs_n_landings.Append(m_head_stair.GeometryCurve.Reference);
                            re_stairs_n_landings.Append(m_mid_landing.GeometryCurve.Reference);
                            re_stairs_n_landings.Append(m_tail_stair.GeometryCurve.Reference);
                            pre_del_list.Add(m_mid_landing.Id);
                            pre_del_list.Add(m_head_stair.Id);
                            pre_del_list.Add(m_tail_stair.Id);
                        }

                        else //無中間平台
                        {
                            XYZ Node_1 = start + first_part_length * ch_path.Direction;
                            XYZ Node_2 = end - first_part_length * ch_path.Direction;
                            XYZ Node_3 = start + head_landing_part_length * ch_path.Direction;
                            XYZ Node_4 = end - rear_landing_part_length * ch_path.Direction;

                            double sin = Math.Abs(Node_3.Z - Node_4.Z) / Line.CreateBound(Node_3, Node_4).ApproximateLength;
                            double first_lim_dis = double.NaN;
                            double second_lim_dis = double.NaN;

                            if (start.Z < end.Z) // 13 low 24 high
                            {
                                first_lim_dis = sin * small_one.LookupParameter("聯絡通道上部圓弧半徑").AsDouble();
                                second_lim_dis = sin * small_one.LookupParameter("聯絡通道下部垂直長度").AsDouble();

                                if (first_part_length != head_landing_part_length) //head_landing_part_length calibrates
                                {
                                    Node_3 = start + (head_landing_part_length - first_lim_dis) * ch_path.Direction;
                                    head_landing_part_length = head_landing_part_length - first_lim_dis;
                                }

                                if (first_part_length != rear_landing_part_length) //rear_landing_part_length calibrates
                                {
                                    Node_4 = end - (rear_landing_part_length + second_lim_dis) * ch_path.Direction;
                                    rear_landing_part_length = rear_landing_part_length + second_lim_dis;
                                }
                            }
                            else // 13 high 24 low
                            {
                                second_lim_dis = sin * small_one.LookupParameter("聯絡通道上部圓弧半徑").AsDouble();
                                first_lim_dis = sin * small_one.LookupParameter("聯絡通道下部垂直長度").AsDouble();

                                if (first_part_length != head_landing_part_length) //head_landing_part_length calibrates
                                {
                                    Node_3 = start + (head_landing_part_length + first_lim_dis) * ch_path.Direction;
                                    head_landing_part_length = head_landing_part_length + first_lim_dis;
                                }

                                if (first_part_length != rear_landing_part_length) //rear_landing_part_length calibrates
                                {
                                    Node_4 = end - (rear_landing_part_length - second_lim_dis) * ch_path.Direction;
                                    rear_landing_part_length = rear_landing_part_length - second_lim_dis;
                                }
                            }

                            bool valid_dis = true;
                            while (valid_dis) //tunnel距離調整
                            {
                                sin = Math.Abs(Node_3.Z - Node_4.Z) / Line.CreateBound(Node_3, Node_4).ApproximateLength;
                                first_lim_dis = double.NaN;
                                second_lim_dis = double.NaN;

                                if (start.Z < end.Z) //聯絡通道上下行高程順序 
                                {
                                    first_lim_dis = sin * small_one.LookupParameter("聯絡通道總直徑").AsDouble();
                                    second_lim_dis = sin * (small_one.LookupParameter("聯絡通道下部垂直長度").AsDouble()
                                        + small_one.LookupParameter("聯絡通道下部厚度").AsDouble());
                                }
                                else
                                {
                                    second_lim_dis = sin * small_one.LookupParameter("聯絡通道總直徑").AsDouble();
                                    first_lim_dis = sin * (small_one.LookupParameter("聯絡通道下部垂直長度").AsDouble()
                                        + small_one.LookupParameter("聯絡通道下部厚度").AsDouble());
                                }

                                if (first_lim_dis + first_part_length > head_landing_part_length) //tunnel前端距離變更 node3
                                {
                                    head_landing_part_length = (first_lim_dis + first_part_length) * 1.0001;
                                    Node_3 = start + head_landing_part_length * ch_path.Direction;
                                }

                                if (second_lim_dis + first_part_length > rear_landing_part_length) //tunnel後端距離變更 node4
                                {
                                    rear_landing_part_length = (second_lim_dis + first_part_length) * 1.0001;
                                    Node_4 = end - rear_landing_part_length * ch_path.Direction;
                                }
                                if ((first_lim_dis + first_part_length < head_landing_part_length)
                                    && (second_lim_dis + first_part_length < rear_landing_part_length)) //tunnel前後端距離確認
                                {
                                    /*
                                    TaskDialog.Show("hi", ((head_landing_part_length - first_part_length) * 304.8).ToString() + "\n"
                                                                            + ((rear_landing_part_length - first_part_length) * 304.8).ToString() + "\n");
                                    */
                                    break;
                                }
                            }

                            Line head = Line.CreateBound(start, Node_1);
                            Line nail = Line.CreateBound(Node_2, end);
                            Line head_landing = Line.CreateBound(Node_1, Node_3);
                            Line nail_landing = Line.CreateBound(Node_4, Node_2);

                            SketchPlane sketchPlane_A = Sketch_plain(doc, start, Node_1);
                            SketchPlane sketchPlane_C = Sketch_plain(doc, Node_2, end);

                            ModelCurve m_head = doc.FamilyCreate.NewModelCurve(head, sketchPlane_A);
                            ModelCurve m_head_landing = doc.FamilyCreate.NewModelCurve(head_landing, sketchPlane_A);
                            ModelCurve m_nail = doc.FamilyCreate.NewModelCurve(nail, sketchPlane_C);
                            ModelCurve m_nail_landing = doc.FamilyCreate.NewModelCurve(nail_landing, sketchPlane_C);

                            re_head.Append(m_head.GeometryCurve.Reference);
                            re_nail.Append(m_nail.GeometryCurve.Reference);
                            re_stairs_n_landings.Append(m_head_landing.GeometryCurve.Reference);
                            re_stairs_n_landings.Append(m_nail_landing.GeometryCurve.Reference);

                            pre_del_list.Add(m_head.Id);
                            pre_del_list.Add(m_head_landing.Id);
                            pre_del_list.Add(m_nail.Id);
                            pre_del_list.Add(m_nail_landing.Id);


                            if (start.Z < end.Z)
                            {
                                sec_of_landing.Add(0, (Node_1, Node_3));
                                sec_of_landing.Add(1, (Node_4, Node_2));
                            }
                            else
                            {
                                sec_of_landing.Add(0, (Node_2, Node_4));
                                sec_of_landing.Add(1, (Node_3, Node_1));
                            }

                            Line mid_stair = Line.CreateBound(Node_3, Node_4);
                            SketchPlane sketchPlane_F = Sketch_Plane(doc, Node_3, Node_4);
                            ModelCurve m_mid_stair = doc.FamilyCreate.NewModelCurve(mid_stair, sketchPlane_F);
                            re_stairs_n_landings.Append(m_mid_stair.GeometryCurve.Reference);
                            pre_del_list.Add(m_mid_stair.Id);

                        }

                        Sweep head_part = doc.FamilyCreate.NewSweep(true, re_head, big_one_profile, 0, ProfilePlaneLocation.Start);
                        Sweep head_part_landing = doc.FamilyCreate.NewSweep(true, re_stairs_n_landings, small_one_profile, 0, ProfilePlaneLocation.Start);
                        Sweep nail_part = doc.FamilyCreate.NewSweep(true, re_nail, big_one_profile, 0, ProfilePlaneLocation.Start);

                        //掃掠空心段
                        //空心段取目標點位及前後兩個點位，取五段作路徑
                        ReferenceArray hollow_dn = new ReferenceArray();
                        ReferenceArray hollow_up = new ReferenceArray();

                        for (int k = (target - r) - 3; k < (target - r) + 4; k++) //excel index has been changed, sketchplane should be modified cuz the slope of tunnel(down)
                        {
                            Line ho_dn = Line.CreateBound(rf.data_list_tunnel[k].start_point, rf.data_list_tunnel[k + 1].start_point);
                            SketchPlane sketchPlane_dn = Sketch_Plane(doc, rf.data_list_tunnel[k].start_point, rf.data_list_tunnel[k + 1].start_point);
                            ModelCurve mho_dn = doc.FamilyCreate.NewModelCurve(ho_dn, sketchPlane_dn);
                            hollow_dn.Append(mho_dn.GeometryCurve.Reference);
                            pre_del_list.Add(mho_dn.Id);
                        }

                        for (int k2 = (target2 - r2) - 3; k2 < (target2 - r2) + 4; k2++) //excel index has been changed, sketchplane should be modified cuz the slope of tunnel(up)
                        {
                            Line ho_up = Line.CreateBound(rf.data_list_tunnel2[k2].start_point, rf.data_list_tunnel2[k2 + 1].start_point);
                            SketchPlane sketchPlane_up = Sketch_Plane(doc, rf.data_list_tunnel2[k2].start_point, rf.data_list_tunnel2[k2 + 1].start_point);
                            ModelCurve mho_up = doc.FamilyCreate.NewModelCurve(ho_up, sketchPlane_up);
                            hollow_up.Append(mho_up.GeometryCurve.Reference);
                            pre_del_list.Add(mho_up.Id);
                        }

                        //建立空心實體
                        Sweep hollow_dn_tunnel = doc.FamilyCreate.NewSweep(false, hollow_dn, tunnel_profile, 0, ProfilePlaneLocation.Start);
                        Sweep hollow_up_tunnel = doc.FamilyCreate.NewSweep(false, hollow_up, tunnel_profile, 0, ProfilePlaneLocation.Start);

                        //切割元件
                        CombinableElementArray elementArray = new CombinableElementArray();
                        elementArray.Append(hollow_up_tunnel); //如果聯絡通道inverted，改成hollow_dn_tunnel
                        elementArray.Append(head_part);
                        doc.CombineElements(elementArray);

                        elementArray.Clear();
                        elementArray.Append(hollow_dn_tunnel); //如果聯絡通道inverted，改成hollow_up_tunnel
                        elementArray.Append(nail_part);
                        doc.CombineElements(elementArray);

                        elementArray.Clear();
                        t.Commit();

                        Transaction door_t = new Transaction(doc);
                        //// 關閉警示視窗
                        //options = door_t.GetFailureHandlingOptions();
                        //closeWarnings = new CloseWarnings();
                        //options.SetClearAfterRollback(true);
                        //options.SetFailuresPreprocessor(closeWarnings);
                        //door_t.SetFailureHandlingOptions(options);
                        door_t.Start("加入防火門");
                        try
                        {
                            XYZ door_mid_point = null;
                            SketchPlane door_plane = null;

                            if (mid_exist) //有中間平台
                            {
                                switch (door_on_landing)
                                {
                                    case 0:
                                        XYZ head_start = sec_of_landing[0].Item1;
                                        XYZ head_end = sec_of_landing[0].Item2;
                                        door_mid_point = add_door(head_start, head_end, door_dis);
                                        door_plane = Sketch_plain(doc, head_start, head_end);
                                        break;
                                    case 1:
                                        XYZ middle_start = sec_of_landing[1].Item1;
                                        XYZ middle_end = sec_of_landing[1].Item2;
                                        door_mid_point = add_door(middle_start, middle_end, door_dis);
                                        door_plane = Sketch_plain(doc, middle_start, middle_end);
                                        break;
                                    case 2:
                                        XYZ rear_start = sec_of_landing[2].Item1;
                                        XYZ rear_end = sec_of_landing[2].Item2;
                                        door_mid_point = add_door(rear_start, rear_end, door_dis);
                                        door_plane = Sketch_plain(doc, rear_start, rear_end);
                                        break;
                                }
                            }
                            else //無中間平台
                            {
                                switch (door_on_landing)
                                {
                                    case 0:
                                        XYZ head_start = sec_of_landing[0].Item1;
                                        XYZ head_end = sec_of_landing[0].Item2;
                                        door_mid_point = add_door(head_start, head_end, door_dis);
                                        door_plane = Sketch_plain(doc, head_start, head_end);
                                        break;

                                    case 1:
                                        XYZ rear_start = sec_of_landing[1].Item1;
                                        XYZ rear_end = sec_of_landing[1].Item2;
                                        door_mid_point = add_door(rear_start, rear_end, door_dis);
                                        door_plane = Sketch_plain(doc, rear_start, rear_end);
                                        break;
                                }
                            }

                            XYZ Node_7 = door_mid_point - ch_path.Direction * (door_width / 2) / 1000 / 0.3048;
                            XYZ Node_8 = door_mid_point + ch_path.Direction * (door_width / 2) / 1000 / 0.3048;
                            ReferenceArray door_array = new ReferenceArray();
                            Line line_door = Line.CreateBound(Node_7, Node_8);
                            ModelCurve mho_door = doc.FamilyCreate.NewModelCurve(line_door, door_plane);
                            door_array.Append(mho_door.GeometryCurve.Reference);
                            pre_del_list.Add(mho_door.Id);
                            Sweep door_part = doc.FamilyCreate.NewSweep(true, door_array, door_arc_profile, 0, ProfilePlaneLocation.Start);

                        }
                        catch (Exception e)
                        {
                            TaskDialog.Show("error", e.Message + e.StackTrace);
                        }
                        doc.Delete(pre_del_list);
                        door_t.Commit();
                    }

                    if (module == 2) //建立疊式聯絡通道
                    {
                        t.Start("建立疊式聯絡通道");
                        double first_part_length = (tunnel_rad + connect_length) / 1000 / 0.3048; // mm to inch
                        double dis_vir_tunnel = (tunnel_rad + connect_length + landing_length + circle_rad) / 1000 / 0.3048;
                        Line dir_tunnel = Line.CreateBound(rf.data_list_tunnel2[10].start_point, rf.data_list_tunnel2[11].start_point);
                        XYZ dir_connect_tunnel = dir_tunnel.Direction.CrossProduct(new XYZ(0, 0, -1)).Normalize();

                        XYZ Node_1 = start + first_part_length * dir_connect_tunnel;
                        XYZ Node_2 = end + first_part_length * dir_connect_tunnel;
                        XYZ Node_3 = start + dis_vir_tunnel * dir_connect_tunnel;
                        XYZ Node_4 = end + dis_vir_tunnel * dir_connect_tunnel;

                        List<ElementId> pre_del_list = new List<ElementId>();
                        Line up = Line.CreateBound(start, Node_1);
                        Line dn = Line.CreateBound(end, Node_2);
                        Line up_landing = Line.CreateBound(Node_1, Node_3);
                        Line dn_landing = Line.CreateBound(Node_2, Node_4);
                        Line vir_connect_tunnel = Line.CreateBound(Node_3 - new XYZ(0, 0, 8), Node_4 + new XYZ(0, 0, 8));

                        SketchPlane sketchPlane_A = Sketch_plain(doc, start, Node_1);
                        SketchPlane sketchPlane_B = Sketch_plain(doc, end, Node_2);
                        Plane vir_Plane = Plane.CreateByNormalAndOrigin(dir_connect_tunnel, Node_3);
                        SketchPlane sketchPlane_C = SketchPlane.Create(doc, vir_Plane);

                        ReferenceArray re_up = new ReferenceArray();
                        ReferenceArray re_landings_up = new ReferenceArray();
                        ReferenceArray re_landings_dn = new ReferenceArray();
                        ReferenceArray re_dn = new ReferenceArray();
                        ReferenceArray re_vir_connect_tunnel = new ReferenceArray();

                        ModelCurve m_up = doc.FamilyCreate.NewModelCurve(up, sketchPlane_A);
                        ModelCurve m_up_landing = doc.FamilyCreate.NewModelCurve(up_landing, sketchPlane_A);
                        ModelCurve m_dn = doc.FamilyCreate.NewModelCurve(dn, sketchPlane_B);
                        ModelCurve m_dn_landing = doc.FamilyCreate.NewModelCurve(dn_landing, sketchPlane_B);
                        ModelCurve m_vir_connect_tunnel = doc.FamilyCreate.NewModelCurve(vir_connect_tunnel, sketchPlane_C);

                        re_up.Append(m_up.GeometryCurve.Reference);
                        re_dn.Append(m_dn.GeometryCurve.Reference);
                        re_landings_up.Append(m_up_landing.GeometryCurve.Reference);
                        re_landings_dn.Append(m_dn_landing.GeometryCurve.Reference);
                        re_vir_connect_tunnel.Append(m_vir_connect_tunnel.GeometryCurve.Reference);

                        pre_del_list.Add(m_up.Id);
                        pre_del_list.Add(m_up_landing.Id);
                        pre_del_list.Add(m_dn.Id);
                        pre_del_list.Add(m_dn_landing.Id);
                        pre_del_list.Add(m_vir_connect_tunnel.Id);

                        Sweep up_part = doc.FamilyCreate.NewSweep(true, re_up, big_one_profile, 0, ProfilePlaneLocation.Start);
                        Sweep up_part_landing = doc.FamilyCreate.NewSweep(true, re_landings_up, small_one_profile, 0, ProfilePlaneLocation.Start);
                        Sweep dn_part = doc.FamilyCreate.NewSweep(true, re_dn, big_one_profile, 0, ProfilePlaneLocation.Start);
                        Sweep dn_part_landing = doc.FamilyCreate.NewSweep(true, re_landings_dn, small_one_profile, 0, ProfilePlaneLocation.Start);

                        //掃掠空心段
                        //空心段取目標點位及前後兩個點位，取五段作路徑
                        ReferenceArray hollow_dn = new ReferenceArray();
                        ReferenceArray hollow_up = new ReferenceArray();

                        for (int k = target - 2; k < target + 3; k++) //excel index has been changed, sketchplane should be modified cuz the slope of tunnel(up and down)
                        {
                            Line ho_dn = Line.CreateBound(rf.data_list_tunnel[k - 240].start_point, rf.data_list_tunnel[k + 1 - 240].start_point);
                            Line ho_up = Line.CreateBound(rf.data_list_tunnel2[k - 240].start_point, rf.data_list_tunnel2[k + 1 - 240].start_point);
                            SketchPlane sketchPlane_dn = Sketch_Plane(doc, rf.data_list_tunnel[k - 240].start_point, rf.data_list_tunnel[k + 1 - 240].start_point);
                            SketchPlane sketchPlane_up = Sketch_Plane(doc, rf.data_list_tunnel2[k - 240].start_point, rf.data_list_tunnel2[k + 1 - 240].start_point);
                            ModelCurve mho_dn = doc.FamilyCreate.NewModelCurve(ho_dn, sketchPlane_dn);
                            ModelCurve mho_up = doc.FamilyCreate.NewModelCurve(ho_up, sketchPlane_up);
                            hollow_dn.Append(mho_dn.GeometryCurve.Reference);
                            hollow_up.Append(mho_up.GeometryCurve.Reference);
                            pre_del_list.Add(mho_dn.Id);
                            pre_del_list.Add(mho_up.Id);
                        }
                        //建立空心實體
                        Sweep hollow_dn_tunnel = doc.FamilyCreate.NewSweep(false, hollow_dn, tunnel_profile, 0, ProfilePlaneLocation.Start);
                        Sweep hollow_up_tunnel = doc.FamilyCreate.NewSweep(false, hollow_up, tunnel_profile, 0, ProfilePlaneLocation.Start);

                        //建立疊式通道
                        Sweep vir_connect_tunnel_part = doc.FamilyCreate.NewSweep(true, re_vir_connect_tunnel, tunnel_profile, 0, ProfilePlaneLocation.Start);


                        //切割元件
                        CombinableElementArray elementArray = new CombinableElementArray();
                        elementArray.Append(hollow_up_tunnel);
                        elementArray.Append(up_part);
                        doc.CombineElements(elementArray);

                        elementArray.Clear();
                        elementArray.Append(hollow_dn_tunnel);
                        elementArray.Append(dn_part);
                        doc.CombineElements(elementArray);

                        elementArray.Clear();
                        elementArray.Append(up_part_landing);
                        elementArray.Append(dn_part_landing);
                        elementArray.Append(vir_connect_tunnel_part);
                        doc.CombineElements(elementArray);

                        doc.Delete(pre_del_list);
                        t.Commit();
                    }

                    //另存新檔 "聯絡通道0x"
                    SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };
                    doc.SaveAs(path + "聯絡通道\\聯絡通道0" + (channel_p + 1).ToString() + ".rfa", saveAsOptions);
                }

                app.OpenAndActivateDocument(pro_doc.PathName);
                doc.Close();

                Transaction pro_t = new Transaction(pro_doc);
                //// 關閉警示視窗
                //options = pro_t.GetFailureHandlingOptions();
                //closeWarnings = new CloseWarnings();
                //options.SetClearAfterRollback(true);
                //options.SetFailuresPreprocessor(closeWarnings);
                //pro_t.SetFailureHandlingOptions(options);
                pro_t.Start("載入聯絡通道");
                try
                {
                    for (int a = 0; a <= rf.data_list_cd_channel.Count - 1; a++)
                    {
                        pro_doc.LoadFamily(path + "聯絡通道\\聯絡通道0" + (a + 1).ToString() + ".rfa");
                        FamilySymbol contact_tunnel = new FilteredElementCollector(pro_doc)
                            .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
                            (x => x.Name == "聯絡通道0" + (a + 1).ToString()).First();

                        contact_tunnel.Activate();
                        FamilyInstance object_acr = pro_doc.Create.NewFamilyInstance(XYZ.Zero, contact_tunnel, StructuralType.NonStructural);

                        try
                        {
                            object_acr.LookupParameter("聯絡通道層數").Set(cc_properties.path_levels[a]);
                            object_acr.LookupParameter("聯絡通道第一層長度").Set(cc_properties.level_one[a] / 304.8);
                            object_acr.LookupParameter("聯絡通道第二層長度").Set(cc_properties.level_two[a] / 304.8);
                            object_acr.LookupParameter("聯絡通道第三層長度").Set(cc_properties.level_three[a] / 304.8);
                            object_acr.LookupParameter("聯絡通道防火門層數").Set(cc_properties.door_levels[a]);
                            object_acr.LookupParameter("聯絡通道防火門距離").Set(cc_properties.door_dis[a] / 304.8);
                        }
                        catch { continue; }

                    }
                }
                catch (Exception e)
                {
                    TaskDialog.Show("error", e.Message + e.StackTrace);
                }
                pro_t.Commit();

            }
            catch (Exception e)
            {
                TaskDialog.Show("error", e.Message + e.StackTrace);
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

        public SketchPlane Sketch_Plane(Document doc, XYZ start, XYZ end)
        {
            SketchPlane sk = null;
            XYZ v = end - start;
            XYZ vec_on_plane = v.CrossProduct(new XYZ(0, 0, 1));
            XYZ norm = vec_on_plane.CrossProduct(v);
            Plane geomPlane = Plane.CreateByNormalAndOrigin(norm, start);
            sk = SketchPlane.Create(doc, geomPlane);
            return sk;
        }

        public XYZ add_door(XYZ start, XYZ end, double dis)
        {
            XYZ ad = null;
            XYZ vec = end - start;
            ad = dis / Math.Sqrt((vec.X) * (vec.X) + (vec.Y) * (vec.Y)) * vec / 1000 / 0.3048 + start;
            return ad; //vec 如果是inverted要加負號
        }

        public void input_properties(FamilySymbol fms, contact_channel_properties cc_properties, int i)
        {
            string name = fms.Name;

            if (name == "聯絡通道輪廓")
            {
                fms.LookupParameter("聯絡通道上部厚度").Set(cc_properties.up_arc_thickness[i] / 304.8);
                fms.LookupParameter("聯絡通道上部圓弧半徑").Set(cc_properties.up_arc_radius[i] / 304.8);
                fms.LookupParameter("聯絡通道下部厚度").Set(cc_properties.dn_thickness[i] / 304.8);
                fms.LookupParameter("聯絡通道下部垂直長度").Set(cc_properties.dn_height[i] / 304.8);
                fms.LookupParameter("聯絡通道初期支撐厚度").Set(cc_properties.support_thickness[i] / 304.8);
                fms.LookupParameter("聯絡通道底部高程").Set(cc_properties.tunnel_elevation[i] / 304.8);
                fms.LookupParameter("聯絡通道開孔範圍底部增加厚度").Set(cc_properties.hollow_bottom_thickness[i] / 304.8);
                fms.LookupParameter("聯絡通道開孔範圍軸向深度").Set(cc_properties.hollow_depth[i] / 304.8);
                fms.LookupParameter("聯絡通道開孔高度").Set(cc_properties.hollow_height[i] / 304.8);
                fms.LookupParameter("聯絡通道開孔寬度").Set(cc_properties.hollow_width[i] / 304.8);

            }

            else if (name == "聯絡通道內輪廓")
            {
                fms.LookupParameter("聯絡通道總直徑").Set((cc_properties.up_arc_thickness[i] + cc_properties.up_arc_radius[i]) / 304.8);
                fms.LookupParameter("聯絡通道上部圓弧半徑").Set(cc_properties.up_arc_radius[i] / 304.8);
                fms.LookupParameter("聯絡通道下部厚度").Set(cc_properties.dn_thickness[i] / 304.8);
                fms.LookupParameter("聯絡通道下部垂直長度").Set(cc_properties.dn_height[i] / 304.8);
                fms.LookupParameter("聯絡通道初期支撐厚度").Set(cc_properties.support_thickness[i] / 304.8);
                fms.LookupParameter("聯絡通道底部高程").Set(cc_properties.tunnel_elevation[i] / 304.8);
            }

            else if (name == "防火門框")
            {
                fms.LookupParameter("聯絡通道上部厚度").Set(cc_properties.up_arc_thickness[i] / 304.8);
                fms.LookupParameter("聯絡通道上部圓弧半徑").Set(cc_properties.up_arc_radius[i] / 304.8);
                fms.LookupParameter("聯絡通道下部厚度").Set(cc_properties.dn_thickness[i] / 304.8);
                fms.LookupParameter("聯絡通道下部垂直長度").Set(cc_properties.dn_height[i] / 304.8);
                fms.LookupParameter("聯絡通道底部高程").Set(cc_properties.tunnel_elevation[i] / 304.8);
            }
            else if (name == "空心隧道")
            {
                fms.LookupParameter("隧道半徑").Set(cc_properties.tunnel_redius[i] / 304.8);
            }

        }

        //// 關閉警示視窗 
        //public class CloseWarnings : IFailuresPreprocessor
        //{
        //    FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failuresAccessor)
        //    {
        //        String transactionName = failuresAccessor.GetTransactionName();
        //        IList<FailureMessageAccessor> fmas = failuresAccessor.GetFailureMessages();
        //        if (fmas.Count == 0) { return FailureProcessingResult.Continue; }
        //        if (transactionName.Equals("EXEMPLE"))
        //        {
        //            foreach (FailureMessageAccessor fma in fmas)
        //            {
        //                if (fma.GetSeverity() == FailureSeverity.Error)
        //                {
        //                    failuresAccessor.DeleteAllWarnings();
        //                    return FailureProcessingResult.ProceedWithRollBack;
        //                }
        //                else { failuresAccessor.DeleteWarning(fma); }
        //            }
        //        }
        //        else
        //        {
        //            foreach (FailureMessageAccessor fma in fmas) { failuresAccessor.DeleteAllWarnings(); }
        //        }
        //        return FailureProcessingResult.Continue;
        //    }
        //}

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
