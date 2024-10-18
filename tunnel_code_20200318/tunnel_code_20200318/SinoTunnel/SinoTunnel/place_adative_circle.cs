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

namespace SinoTunnel
{
    //在用的環片
    struct segment_para
    {
        public double displacement_angle, displacement_width, r1, r2, horizontal_angle, vertical_angle, width;
        public Dictionary<string, double> angle_dic;
        public Dictionary<string, double> rotation_dic;

    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class place_adative_circle : IExternalEventHandler
    {
        string path;//Form1.path;

        public void Execute(UIApplication uiapp)
        {
            path = Form1.path;
            try
            {
                // 讀取revit檔案
                Autodesk.Revit.DB.Document document = uiapp.ActiveUIDocument.Document;
                UIDocument uidoc = new UIDocument(document);
                Autodesk.Revit.DB.Document doc = uidoc.Document;
                string doc_path = doc.PathName;
                Application app = doc.Application;
                
                // 建立改動檔案交易
                Transaction t = new Transaction(doc);

                // 讀取excel幾何資料
                readfile rf = new readfile();
                rf.read_tunnel_point();
                rf.read_tunnel_point2();
                rf.read_properties();
                List<string> ch_channel_points = rf.cd_channel_points();
                rf.data_list_tunnel.RemoveAt(rf.data_list_tunnel.Count() - 1);
                rf.data_list_tunnel2.RemoveAt(rf.data_list_tunnel2.Count() - 1);

                // 建立環形元件
                circle_family_create(rf.data_list_tunnel, doc_path, uiapp, rf.properties, ch_channel_points);
                circle_family_create(rf.data_list_tunnel2, doc_path, uiapp, rf.properties, ch_channel_points);
                
                // 匯入環形元件至.rvt檔案
                t.Start("Activate FamilySymbol");
                List<FamilySymbol> familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name.Contains("螺栓")).ToList();
                foreach (FamilySymbol symbol in familySymbol)
                    symbol.Activate();
                t.Commit();
                
                // 放置環形元件至.rvt檔案
                // 上行與下行
                set_tunnel(doc, rf.data_list_tunnel, rf.properties, ch_channel_points);
                set_tunnel(doc, rf.data_list_tunnel2, rf.properties, ch_channel_points);
            }
            catch(Exception e) { TaskDialog.Show("error", e.StackTrace + e.Message); }

        }


        // 放置環形元件至.rvt檔案
        public void set_tunnel(Document doc, IList<data_object> data_list, properties_object properties, List<string> cd_channel_points)
        {
            Family family;
            string name = "自適應環形";
            // 建立繪入環形族群檔交易
            Transaction trans = new Transaction(doc, "load segments");
            trans.Start();

            // 讀取初始兩點並計算隧道方向
            XYZ first_point = data_list[0].start_point;
            XYZ second_point = data_list[1].start_point;
            XYZ direction = new XYZ(second_point.X - first_point.X, second_point.Y - first_point.Y, second_point.Z - first_point.Z);
            
            // 計算隧道起始方向的水平與垂直轉角
            double start_angle_horizontal = Math.Atan2(direction.Y, direction.X);
            double start_angle_vertical = Math.Atan2(direction.Z, Math.Pow(Math.Pow(direction.X, 2) + Math.Pow(direction.Y, 2), 0.5));

            // 依照每一個里程的點位建立及放置環形
            for (int i = 0; i < data_list.Count; i++)
            {
                // 讀取該次點位的下一個里程的點位
                string next_station = "nnn";
                try
                {
                    next_station = data_list[i + 1].station;
                }
                catch {  }

                // 依照剛還片位置建立鋼環片
                if (cd_channel_points.Contains(data_list[i].station))
                {
                    doc.LoadFamily(path + "自適應環形\\instance\\" + name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A_S_1" + ".rfa", out family);
                    doc.LoadFamily(path + "自適應環形\\instance\\" + name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B_S_1" + ".rfa", out family);
                }
                else if(cd_channel_points.Contains(next_station))
                {
                    doc.LoadFamily(path + "自適應環形\\instance\\" + name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A_S_2" + ".rfa", out family);
                    doc.LoadFamily(path + "自適應環形\\instance\\" + name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B_S_2" + ".rfa", out family);
                }
                else
                {
                    doc.LoadFamily(path + "自適應環形\\instance\\" + name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A" + ".rfa", out family);
                    doc.LoadFamily(path + "自適應環形\\instance\\" + name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B" + ".rfa", out family);
                }
                trans.Commit();

                trans.Start();
                FamilySymbol fam_sym;
                FamilySymbol bolt= null;
                bool is_spacial_segment = false;

                // 判斷偶數基數決定放入環片軸向的旋轉角度，讀入對應放置的環形
                if (i % 2 == 0)
                {
                    if(cd_channel_points.Contains(data_list[i].station))
                    {
                        is_spacial_segment = true;
                        fam_sym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A_S_1").First();
                        //bolt = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "螺栓_A").First();
                    }
                    else if (cd_channel_points.Contains(next_station))
                    {
                        is_spacial_segment = true;
                        fam_sym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A_S_2").First();
                        //bolt = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "螺栓_A").First();
                    }
                    else
                    {
                        fam_sym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A").First();
                        if (Form1.bolt)
                            bolt = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "螺栓_A").First();
                    }
                }
                else
                {
                    if (cd_channel_points.Contains(data_list[i].station))
                    {
                        is_spacial_segment = true;
                        fam_sym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B_S_1").First();
                        //bolt = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "螺栓_B").First();
                    }
                    else if (cd_channel_points.Contains(next_station))
                    {
                        is_spacial_segment = true;
                        fam_sym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B_S_2").First();
                        //bolt = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "螺栓_B").First();
                    }
                    else
                    {
                        fam_sym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == name + data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B").First();
                        if (Form1.bolt)
                            bolt = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == "螺栓_B").First();
                    }

                }

                // 激活專案檔中族群
                fam_sym.Activate();
                // 放置環片
                FamilyInstance FI = doc.Create.NewFamilyInstance(data_list[i].start_point, fam_sym, StructuralType.UnknownFraming);

                // 寫入環片實體幾何以及非幾何資訊
                FI.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(data_list[i].station.ToString());
                FI.LookupParameter("編號").Set(properties.id);
                FI.LookupParameter("Type_K環片插入型式").Set(properties.type_K_insert_type);
                FI.LookupParameter("隧道內徑").Set(properties.inner_diameter / 304800);
                FI.LookupParameter("環片厚度").Set(properties.thickness / 304800);
                FI.LookupParameter("環片寬度").Set(properties.width / 304800);
                FI.LookupParameter("環圈交錯角度").Set(properties.displacement_angle * Math.PI / 180);
                FI.LookupParameter("Type_A數量").Set(properties.Type_A_q);
                FI.LookupParameter("Type_B數量").Set(properties.Type_B_q);
                FI.LookupParameter("Type_K數量").Set(properties.Type_K_q);
                FI.LookupParameter("type_K內縮量").Set(properties.displacement / 304800);
                FI.LookupParameter("type_A角度").Set(properties.Type_A_1 * Math.PI / 180);
                FI.LookupParameter("type_B角度").Set(properties.Type_B_1_head * Math.PI / 180);
                FI.LookupParameter("type_K角度").Set(properties.Type_K_1_head * Math.PI / 180);

                // 旋轉環片至正確角度 先轉水平角度再轉垂直角度
                Line axis = Line.CreateBound(data_list[i].start_point, new XYZ(data_list[i].start_point.X, data_list[i].start_point.Y, data_list[i].start_point.Z + 1));          
                FI.Location.Rotate(axis, start_angle_horizontal);
                axis = Line.CreateBound(data_list[i].start_point, new XYZ(data_list[i].start_point.X - Math.Sin(start_angle_horizontal), data_list[i].start_point.Y + Math.Cos(start_angle_horizontal), data_list[i].start_point.Z));
                FI.Location.Rotate(axis, start_angle_vertical);

                // 建立正常環片
                if (!is_spacial_segment && Form1.bolt)
                {
                    FamilyInstance boltFI = doc.Create.NewFamilyInstance(data_list[i].start_point, bolt, StructuralType.NonStructural);
                    axis = Line.CreateBound(data_list[i].start_point, new XYZ(data_list[i].start_point.X, data_list[i].start_point.Y, data_list[i].start_point.Z + 1));
                    boltFI.Location.Rotate(axis, start_angle_horizontal);
                    axis = Line.CreateBound(data_list[i].start_point, new XYZ(data_list[i].start_point.X - Math.Sin(start_angle_horizontal), data_list[i].start_point.Y + Math.Cos(start_angle_horizontal), data_list[i].start_point.Z));
                    boltFI.Location.Rotate(axis, start_angle_vertical);

                    foreach (ElementId id in FI.GetSubComponentIds())
                    {
                        InstanceVoidCutUtils.AddInstanceVoidCut(doc, doc.GetElement(id), boltFI);
                    }
                }

                // 紀錄該點位水平及垂直旋轉角度
                start_angle_horizontal -= (data_list[i].horizontal_angle / 180) * Math.PI;
                start_angle_vertical -= (data_list[i].vertical_angle / 180) * Math.PI;

            }

            trans.Commit();
        }

        // 建立環形族群檔
        public IList<string> circle_family_create(IList<data_object> data_list, string doc_path, UIApplication uiapp, properties_object properties, List<string> cd_channel_points)
        {
            //create different kinds of circle family
            List<string> name_list = new List<string> { "K", "B2" }; // 一般環片
            List<string> spacial_name_list = new List<string> { "S", "A3", "A2", "A1", "B1" }; // 剛環片
            for (int i = properties.Type_A_q; i > 0; i--)
            {
                name_list.Add("A" + i.ToString());
            }
            name_list.Add("B1");

            IList<string> done_list = new List<string>(); // 紀錄完成的環形，避免重複建立浪費運算時間

            segment_para segment_a = new segment_para();
            segment_para segment_b = new segment_para();
            segment_para spacial_segment_a = new segment_para();
            segment_para spacial_segment_b = new segment_para();

            // 依照點位建立分別的環形
            for (int i = 0; i < data_list.Count; i++)
            {
                // 判斷是否是已經建立過的環形元件
                bool isCreate = false;
                bool isCreate_S_1 = false;
                bool isCreate_S_2 = false;
                for (int j = 0; j < done_list.Count; j++)
                {
                    string e_name = data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString();
                    if (done_list[j] == e_name)
                    {
                        isCreate = true;
                    }
                    if(done_list[j] == e_name + "S_1")
                    {
                        isCreate_S_1 = true;
                    }
                    if (done_list[j] == e_name + "S_2")
                    {
                        isCreate_S_2 = true;
                    }
                }

                // 紀錄該里程下一個的下一個里程點位
                string next_station = "nnn";
                try
                {
                    next_station = data_list[i + 1].station;
                }
                catch { }


                // 判斷是剛環片或是正常環片
                if (cd_channel_points.Contains(data_list[i].station))
                {
                    if (isCreate_S_1 == false)
                    {
                        //for spacial segment
                        spacial_segment_a = set_spacial_segmet_para(spacial_segment_a, properties.displacement_angle, properties, data_list[i].vertical_angle, data_list[i].vertical_angle);
                        spacial_segment_b = set_spacial_segmet_para(spacial_segment_b, properties.displacement_angle, properties, data_list[i].vertical_angle, data_list[i].vertical_angle);

                        create_adative_circle(doc_path, uiapp, spacial_segment_a, spacial_name_list, data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A_S_1", false);
                        create_adative_circle(doc_path, uiapp, spacial_segment_b, spacial_name_list, data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B_S_1", false);

                        done_list.Add(data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "S_1");
                    }
                }
                else if(cd_channel_points.Contains(next_station))
                {
                    if (isCreate_S_2 == false)
                    {
                        //for spacial segment
                        spacial_segment_a = set_spacial_segmet_para(spacial_segment_a, properties.displacement_angle, properties, data_list[i].vertical_angle, data_list[i].vertical_angle);
                        spacial_segment_b = set_spacial_segmet_para(spacial_segment_b, properties.displacement_angle, properties, data_list[i].vertical_angle, data_list[i].vertical_angle);

                        create_adative_circle(doc_path, uiapp, spacial_segment_a, spacial_name_list, data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A_S_2", true);
                        create_adative_circle(doc_path, uiapp, spacial_segment_b, spacial_name_list, data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B_S_2", true);

                        done_list.Add(data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "S_2");
                    }
                }
                else
                {
                    if (isCreate == false)
                    {
                        segment_a = set_segmet_para(segment_a, properties.displacement_angle, properties, data_list[i].horizontal_angle, data_list[i].vertical_angle);
                        segment_b = set_segmet_para(segment_a, -properties.displacement_angle, properties, data_list[i].horizontal_angle, data_list[i].vertical_angle);
                        create_adative_circle(doc_path, uiapp, segment_a, name_list, data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_A", false);
                        create_adative_circle(doc_path, uiapp, segment_b, name_list, data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString() + "_B", false);

                        done_list.Add(data_list[i].horizontal_angle.ToString() + data_list[i].vertical_angle.ToString());
                    }
                }

                
            }

            return done_list;
        }



        // 讀取並計算環片參數至儲存類別（正常環片）
        public segment_para set_segmet_para(segment_para segment_Para, double displacement_angle, properties_object properties, double horizontal_angel, double vertical_angle)
        {
            segment_Para.displacement_angle = displacement_angle / 2;
            segment_Para.angle_dic = new Dictionary<string, double>();
            segment_Para.angle_dic.Add("K", properties.Type_K_1_head / 2);
            segment_Para.angle_dic.Add("B1", properties.Type_B_1_head / 2);
            segment_Para.angle_dic.Add("B2", properties.Type_B_2_head / 2);
            for( int i = 0; i < properties.Type_A_q; i++)
            {
                segment_Para.angle_dic.Add("A"+(i+1).ToString(), properties.Type_A_1 / 2);
            }
            //segment_Para.angle_dic.Add("A1", properties.Type_A_1 / 2);
            //segment_Para.angle_dic.Add("A2", properties.Type_A_2 / 2);
            //segment_Para.angle_dic.Add("A3", properties.Type_A_3 / 2);
            /*segment_Para.rotation_dic = new Dictionary<string, double>();
            segment_Para.rotation_dic.Add("K", segment_Para.displacement_angle + 90);
            segment_Para.rotation_dic.Add("B2", segment_Para.rotation_dic["K"] + segment_Para.angle_dic["K"] + segment_Para.angle_dic["B2"]);
            segment_Para.rotation_dic.Add("A3", segment_Para.rotation_dic["B2"] + segment_Para.angle_dic["B2"] + segment_Para.angle_dic["A3"]);
            segment_Para.rotation_dic.Add("A2", segment_Para.rotation_dic["A3"] + segment_Para.angle_dic["A3"] + segment_Para.angle_dic["A2"]);
            segment_Para.rotation_dic.Add("A1", segment_Para.rotation_dic["A2"] + segment_Para.angle_dic["A2"] + segment_Para.angle_dic["A1"]);
            segment_Para.rotation_dic.Add("B1", segment_Para.rotation_dic["A1"] + segment_Para.angle_dic["A1"] + segment_Para.angle_dic["B1"]);*/

            segment_Para.rotation_dic = new Dictionary<string, double>();
            segment_Para.rotation_dic.Add("K", segment_Para.displacement_angle + 90);
            segment_Para.rotation_dic.Add("B2", segment_Para.rotation_dic["K"] + segment_Para.angle_dic["K"] + segment_Para.angle_dic["B2"]);
            segment_Para.rotation_dic.Add("A" + properties.Type_A_q.ToString(), segment_Para.rotation_dic["B2"] + segment_Para.angle_dic["B2"] + segment_Para.angle_dic["A" + properties.Type_A_q.ToString()]);
            for(int i = properties.Type_A_q-1; i > 0; i--)
            {
                segment_Para.rotation_dic.Add("A" + i.ToString(), segment_Para.rotation_dic["A" + (i + 1).ToString()] + segment_Para.angle_dic["A" + (i + 1).ToString()] + segment_Para.angle_dic["A" + i.ToString()]);
            }
            segment_Para.rotation_dic.Add("B1", segment_Para.rotation_dic["A1"] + segment_Para.angle_dic["A1"] + segment_Para.angle_dic["B1"]);

            segment_Para.displacement_width = properties.displacement;
            segment_Para.r2 = properties.inner_diameter;
            segment_Para.r1 = properties.inner_diameter + properties.thickness * 2;
            segment_Para.horizontal_angle = horizontal_angel;
            segment_Para.vertical_angle = vertical_angle;
            segment_Para.width = properties.width;

            return segment_Para;
        }

        // 讀取並計算環片參數至儲存類別（鋼環片）
        public segment_para set_spacial_segmet_para(segment_para segment_Para, double displacement_angle, properties_object properties, double horizontal_angel, double vertical_angle)
        {
            segment_Para.displacement_angle = displacement_angle / 2;
            segment_Para.angle_dic = new Dictionary<string, double>();
            segment_Para.angle_dic.Add("S", properties.Type_S_1 / 2);
            segment_Para.angle_dic.Add("B1", properties.Type_B_1_head / 2);
            segment_Para.angle_dic.Add("A1", properties.Type_A_1 / 2);
            segment_Para.angle_dic.Add("A2", properties.Type_A_2 / 2);
            segment_Para.angle_dic.Add("A3", properties.Type_A_3 / 2);
            segment_Para.rotation_dic = new Dictionary<string, double>();
            segment_Para.rotation_dic.Add("S", segment_Para.displacement_angle + 90);
            segment_Para.rotation_dic.Add("A3", segment_Para.rotation_dic["S"] + segment_Para.angle_dic["S"] + segment_Para.angle_dic["A3"]);
            segment_Para.rotation_dic.Add("A2", segment_Para.rotation_dic["A3"] + segment_Para.angle_dic["A3"] + segment_Para.angle_dic["A2"]);
            segment_Para.rotation_dic.Add("A1", segment_Para.rotation_dic["A2"] + segment_Para.angle_dic["A2"] + segment_Para.angle_dic["A1"]);
            segment_Para.rotation_dic.Add("B1", segment_Para.rotation_dic["A1"] + segment_Para.angle_dic["A1"] + segment_Para.angle_dic["B1"]);

            segment_Para.displacement_width = properties.displacement;
            segment_Para.r2 = properties.inner_diameter;
            segment_Para.r1 = properties.inner_diameter + properties.thickness * 2;
            segment_Para.horizontal_angle = horizontal_angel;
            segment_Para.vertical_angle = vertical_angle;
            segment_Para.width = properties.width;

            return segment_Para;
        }

        // 建立環形
        public void create_adative_circle(string doc_path, UIApplication uiapp, segment_para segment, List<string> name_list, string tail, bool is_spacial_segment)
        {
            //TaskDialog.Show("test", path + "自適應環形\\BACKUP環形樣板adative.rfa");
            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(path + "自適應環形\\BACKUP環形樣板adative.rfa");
            Document edit_doc = edit_uidoc.Document;

            create_tunnel(edit_doc, segment, name_list, is_spacial_segment);

            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;
            edit_doc.SaveAs(path + "自適應環形\\instance\\自適應環形" + tail + ".rfa", save_option);

            uiapp.OpenAndActivateDocument(doc_path);
            edit_doc.Close(false);

        }

        // 建立環形
        public void create_tunnel(Document doc, segment_para segment_Para, List<string> namelist, bool is_spacial_segment)
        {
            ICollection<FamilySymbol> fam_sym_list = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            FamilySymbol segment = (from x in fam_sym_list where x.Name == "segment" select x).First();

            // 讀取環片水平以及垂直旋轉角度
            segment_Para.horizontal_angle = -segment_Para.horizontal_angle;
            segment_Para.vertical_angle = -segment_Para.vertical_angle;

            // 依照水平以及垂直旋轉角度建立向量
            XYZ start = new XYZ(0, 0, 0);
            XYZ end = new XYZ((segment_Para.width / 304.8), 0, 0);
            XYZ vector1 = new XYZ(1, 0, 0);
            XYZ vector2 = rotate_vector(segment_Para.horizontal_angle, segment_Para.vertical_angle);

            // 以上述向量為法向計算平面之X、Y向量
            Tuple<XYZ, XYZ> vecs1 = random_create_vecXY(vector1, 0, 0);
            Tuple<XYZ, XYZ> vecs2 = random_create_vecXY(vector2, segment_Para.horizontal_angle, segment_Para.vertical_angle);
            
            // 自適應環形點位
            IList<IList<XYZ>> circle_point_list = new List<IList<XYZ>>();

            using (Transaction t = new Transaction(doc, "segment"))
            {
                // 分四次計算第一、二面內圈、外圈點位 儲存到circle_point_list，依此建立環形。
                t.Start();
                try
                {
                    // 判斷是建立鋼環片或正常環片
                    if (!is_spacial_segment)
                    {
                        circle_point_list.Add(create_circle_point(start, segment_Para.r1, vecs1, segment_Para, namelist, 1));
                        circle_point_list.Add(create_circle_point(start, segment_Para.r2, vecs1, segment_Para, namelist, 1));
                        circle_point_list.Add(create_circle_point(end, segment_Para.r1, vecs2, segment_Para, namelist, 2));
                        circle_point_list.Add(create_circle_point(end, segment_Para.r2, vecs2, segment_Para, namelist, 2));
                        create_adaptive_component_instance(doc, segment, point_arrange(circle_point_list));
                    }
                    else
                    {
                        circle_point_list.Add(create_circle_point(start, segment_Para.r1, vecs1, segment_Para, namelist, 2));
                        circle_point_list.Add(create_circle_point(start, segment_Para.r2, vecs1, segment_Para, namelist, 2));
                        circle_point_list.Add(create_circle_point(end, segment_Para.r1, vecs2, segment_Para, namelist, 1));
                        circle_point_list.Add(create_circle_point(end, segment_Para.r2, vecs2, segment_Para, namelist, 1));
                        create_adaptive_component_instance(doc, segment, point_arrange(circle_point_list));
                    }
                }
                catch(Exception e)
                {
                    TaskDialog.Show("error", e.StackTrace + e.Message);
                }

                t.Commit();
            }
        }

        // 旋轉向量小工具
        public XYZ rotate_vector(double rotate_angle_h, double rotate_angle_v)
        {
            rotate_angle_h = (rotate_angle_h / 180) * Math.PI;
            rotate_angle_v = (rotate_angle_v / 180) * Math.PI;

            XYZ vector = new XYZ(Math.Cos(rotate_angle_v) * Math.Cos(rotate_angle_h), Math.Cos(rotate_angle_v) * Math.Sin(rotate_angle_h), Math.Sin(rotate_angle_v));

            return vector;
        }


        // 環形點位小工具，讓點位能符合建立環形內環片建立的方法。
        public IList<IList<XYZ>> point_arrange(IList<IList<XYZ>> point_list)
        {
            IList<IList<XYZ>> point_lists = new List<IList<XYZ>>();
            for (int i = 0; i < point_list.Count; i++)
            {
                //TaskDialog.Show("message", i.ToString() + "," + point_list[i].Count.ToString());
                for (int j = 0; j < point_list[i].Count; j++)
                {
                    try
                    {
                        point_lists[(j) / 3].Add(point_list[i][j]);
                    }
                    catch
                    {
                        //TaskDialog.Show("message", point_lists.Count.ToString());
                        point_lists.Add(new List<XYZ>());
                        point_lists[(j) / 3].Add(point_list[i][j]);
                    }
                }
            }

            return point_lists;
        }


        // 建立環形內環片
        public void create_adaptive_component_instance(Document document, FamilySymbol symbol, IList<IList<XYZ>> point_list)
        {
            /*// Create a new instance of an adaptive component family 取得自適應點
            FamilyInstance instance1 = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(document, symbol);
            FamilyInstance instance2 = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(document, symbol);
            FamilyInstance instance3 = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(document, symbol);
            FamilyInstance instance4 = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(document, symbol);
            FamilyInstance instance5 = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(document, symbol);

            // Get the placement points of this instance 儲存自適應點
            IList<IList<ElementId>> place_point_ids_list = new List<IList<ElementId>>();
            place_point_ids_list.Add(AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(instance1));
            place_point_ids_list.Add(AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(instance2));
            place_point_ids_list.Add(AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(instance3));
            place_point_ids_list.Add(AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(instance4));
            place_point_ids_list.Add(AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(instance5));

            if (point_list.Count > 5)
            {
                FamilyInstance instance6 = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(document, symbol);
                place_point_ids_list.Add(AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(instance6));
            }*/


            // Get the placement points of this instance 儲存自適應點
            IList<IList<ElementId>> place_point_ids_list = new List<IList<ElementId>>();
            for (int i = 0; i < point_list.Count; i++)
            {
                // Create a new instance of an adaptive component family 取得自適應點
                FamilyInstance instance = AdaptiveComponentInstanceUtils.CreateAdaptiveComponentInstance(document, symbol);
                place_point_ids_list.Add(AdaptiveComponentInstanceUtils.GetInstancePlacementPointElementRefIds(instance));
            }

            ReferencePoint point;
            // Set the position of each placement point
            for (int i = 0; i < place_point_ids_list.Count; i++)
            {
                for (int j = 0; j < place_point_ids_list[i].Count; j++)
                {
                    point = document.GetElement(place_point_ids_list[i][j]) as ReferencePoint;
                    point.Position = point_list[i][j];
                }
            }
        }

        // 計算環形內自適應環片各個點位
        public IList<XYZ> sub_create_circle_point(XYZ center_point, double r, Tuple<XYZ, XYZ> vecs, segment_para segment_Para, string name, bool isFirst, int sequence)
        {
            //(displacement_width / (r * 2 * pi)) * 360° + 小半弧長
            // sequence = 1:環形第一面 2:環形第二面
            IList<XYZ> circle_point_list = new List<XYZ>();
            //Arc circle = null;
            double startAngle = 0;
            double endAngle = 0;
            if (sequence == 1)
            {
                startAngle = Math.PI * ((segment_Para.rotation_dic[name] - segment_Para.angle_dic[name]) / 180);
                endAngle = Math.PI * ((segment_Para.rotation_dic[name]) / 180);
                
            }
            else if (sequence == 2)
            {
                if (name == "K" || name == "S")
                {
                    startAngle = Math.PI * ((segment_Para.rotation_dic[name] - segment_Para.angle_dic[name]) / 180) + (segment_Para.displacement_width / (r / 2));
                    endAngle = Math.PI * ((segment_Para.rotation_dic[name]) / 180);
                }
                else if (name == "B1")
                {
                    startAngle = Math.PI * ((segment_Para.rotation_dic[name] - segment_Para.angle_dic[name]) / 180);
                    endAngle = Math.PI * ((segment_Para.rotation_dic[name]) / 180);
                }
                else if (name == "B2")
                {
                    startAngle = Math.PI * ((segment_Para.rotation_dic[name] - segment_Para.angle_dic[name]) / 180) - (segment_Para.displacement_width / (r / 2));
                    endAngle = Math.PI * ((segment_Para.rotation_dic[name]) / 180);
                }
                else
                {
                    startAngle = Math.PI * ((segment_Para.rotation_dic[name] - segment_Para.angle_dic[name]) / 180);
                    endAngle = Math.PI * ((segment_Para.rotation_dic[name]) / 180);
                }
            }
            Arc circle = Arc.Create(center_point, r / (304.8 * 2), startAngle, endAngle, vecs.Item1, vecs.Item2);
            XYZ start_point = circle.GetEndPoint(0);
            XYZ end_point = circle.GetEndPoint(1);
            if (!isFirst)
            {
                circle_point_list.Add(start_point);
            }
            circle_point_list.Add(start_point);
            circle_point_list.Add(end_point);
            circle.Dispose();


            return circle_point_list;
        }
        
        // 結合兩個lists
        public IList<XYZ> append_Ilist(IList<XYZ> ilistA, IList<XYZ> ilistB)
        {
            IList<XYZ> point_list = new List<XYZ>();
            for (int i = 0; i < ilistA.Count; i++)
            {
                point_list.Add(ilistA[i]);
            }
            for (int i = 0; i < ilistB.Count; i++)
            {
                point_list.Add(ilistB[i]);
            }

            return point_list;
        }

        // 計算環形內自適應環片各個點位
        public IList<XYZ> create_circle_point(XYZ center_point, double r, Tuple<XYZ, XYZ> vecs, segment_para segment_Para, List<string> name_list, int sequence)
        {
            IList<XYZ> circle_point_list = new List<XYZ>();
            bool isFirst = true;
            foreach (string name in name_list)
            {
                //circle_point_list = sub_create_circle_point(center_point, r, vecs, segment_Para, name);
                circle_point_list = append_Ilist(circle_point_list, sub_create_circle_point(center_point, r, vecs, segment_Para, name, isFirst, sequence));
                isFirst = false;
            }
            circle_point_list.Add(circle_point_list[0]);

            return circle_point_list;
        }

        // 兩點之間的單位方向向量
        public XYZ create_vector_by_two_points(XYZ point1, XYZ point2)
        {
            XYZ vector = new XYZ(point2.X - point1.X, point2.Y - point1.Y, point2.Z - point1.Z);
            vector = vector_normalized(vector);

            return vector;
        }

        // 隨機建立某一個法向量的另外兩個互相垂直之向量
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

        // 兩向量外基
        public XYZ vector_cross_product(XYZ vec1, XYZ vec2)
        {
            double x = vec1.Y * vec2.Z - vec1.Z * vec2.Y;
            double y = vec1.Z * vec2.X - vec1.X * vec2.Z;
            double z = vec1.X * vec2.Y - vec1.Y * vec2.X;
            XYZ vector = new XYZ(x, y, z);
            return vector;
        }

        // 正規化
        public XYZ vector_normalized(XYZ vector)
        {
            double X = vector.X;
            double Y = vector.Y;
            double Z = vector.Z;
            double length = Math.Pow((X * X + Y * Y + Z * Z), 0.5);
            vector = new XYZ(X / length, Y / length, Z / length);
            return vector;
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }

    }
}
