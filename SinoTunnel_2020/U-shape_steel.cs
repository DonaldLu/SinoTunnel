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
    class U_shape_steel : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

            //讀取資料
            readfile rf = new readfile();
            rf.read_tunnel_point();
            rf.read_tunnel_point2();
            rf.read_target_station();
            rf.read_target_station2();
            rf.read_properties();
            IList<IList<data_object>> all_data_list_tunnel = new List<IList<data_object>>();
            all_data_list_tunnel.Add(rf.data_list_tunnel);
            all_data_list_tunnel.Add(rf.data_list_tunnel2);

            IList<IList<string[]>> walk_way_list = new List<IList<string[]>>();

            walk_way_list.Add(rf.setting_Station.walk_way_station);
            walk_way_list.Add(rf.setting_Station2.walk_way_station);

            List<string[]> walk_way = new List<string[]>();
            walk_way.AddRange(rf.setting_Station.walk_way_station);
            walk_way.AddRange(rf.setting_Station2.walk_way_station);

            all_data_list_tunnel = cut_point_list(all_data_list_tunnel, walk_way_list);

            //int start = 0; // 起始里程
            // 編輯 U型槽鋼
            string two_or_one = (rf.properties.U_steel_aisle_top == rf.properties.center_point + (rf.properties.inner_diameter / 2)) ? "U型槽鋼for_one" : "U型槽鋼";

            // 匯入族群檔
            FamilySymbol Usteel = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
                (x => x.Name == two_or_one + "v2").First();
            Transaction U_t1 = new Transaction(doc);
            // 寫入族群參數
            U_t1.Start("設定參數");
            Usteel.LookupParameter("隧道內徑").SetValueString((rf.properties.inner_diameter / 2).ToString());
            Usteel.LookupParameter("T/R").SetValueString(rf.properties.center_point.ToString());
            Usteel.LookupParameter("走道側底").SetValueString(rf.properties.U_steel_aisle_bottom.ToString());
            Usteel.LookupParameter("非走道側底").SetValueString(rf.properties.U_steel_nonaisle_bottom.ToString());
            try
            {
                Usteel.LookupParameter("走道側頂").SetValueString(rf.properties.U_steel_aisle_top.ToString());
                Usteel.LookupParameter("非走道側頂").SetValueString(rf.properties.U_steel_nonaisle_top.ToString());
            }
            catch
            {
                TaskDialog.Show("U形槽鋼", "單一槽鋼無頂部高程");
            }

            U_t1.Commit();
            // 擺放Ｕ形槽鋼
            FamilySymbol steel = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
                (x => x.Name == two_or_one).First();
            Transaction t = new Transaction(doc);
            t.Start("放置U型槽鋼");
            steel.Activate();
            int count = 0;
            foreach (IList<data_object> data_list_tunnel in all_data_list_tunnel)
            {
                count++;
                for (int i = 0; i < data_list_tunnel.Count - 1; i++)
                {
                    Line toward = Line.CreateBound(data_list_tunnel[i].start_point, data_list_tunnel[i + 1].start_point);
                    XYZ put_point = (0.61 / 0.3048) * toward.Direction + data_list_tunnel[i].start_point;
                    FamilyInstance every_steel = doc.Create.NewFamilyInstance(put_point, steel, StructuralType.NonStructural);
                    double toward_slope = toward.Direction.Y / toward.Direction.X;
                    Line rotate_axis = Line.CreateBound(put_point, put_point + XYZ.BasisZ);
                    double toward_angle = Math.Atan(toward_slope);

                    // 依照Ｕ形槽鋼走道側與非走道側改變Ｕ形槽鋼位置
                    bool turn = false;
                    if (walk_way[count - 1][0] == "右側")
                        turn = true;
                    if (turn)
                    {
                        toward_angle += Math.PI;
                    }
                    if (toward.Direction.X < 0)
                    {
                        toward_angle += Math.PI;
                    }
                    ElementTransformUtils.RotateElement(doc, every_steel.Id, rotate_axis, toward_angle + Math.PI / 2);

                    // 旋轉z軸
                    if (data_list_tunnel[i + 1].start_point.Z != data_list_tunnel[i].start_point.Z)
                    {
                        XYZ fa_point = new XYZ(put_point.X + toward.Direction.Y, put_point.Y - toward.Direction.X, put_point.Z);
                        Line sec_rotate_axis = Line.CreateBound(put_point, fa_point);
                        double Z_angle = Math.Asin((data_list_tunnel[i + 1].start_point.Z - data_list_tunnel[i].start_point.Z) / toward.Length);
                        ElementTransformUtils.RotateElement(doc, every_steel.Id, sec_rotate_axis, Z_angle);
                    }
                }
            }
            t.Commit();
            TaskDialog.Show("U型槽鋼", "U型槽鋼放置完畢。");
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }

        // 依照走道方向不同分割里程點位，給後續建立管線與附掛設施能依照走道側與非走到側建置
        public IList<IList<data_object>> cut_point_list(IList<IList<data_object>> all_data_list_tunnel, IList<IList<string[]>> walk_way_list)
        {
            IList<IList<data_object>> output = new List<IList<data_object>>();
            for (int i = 0; i < walk_way_list.Count; i++)
            {
                for (int j = 0; j < walk_way_list[i].Count; j++)
                {
                    IList<data_object> temp = new List<data_object>();
                    bool isAdd = false;
                    for (int k = 0; k < all_data_list_tunnel[i].Count; k++)
                    {

                        if (to_station(all_data_list_tunnel[i][k].station) == walk_way_list[i][j][1])
                        {
                            isAdd = true;
                        }
                        else if (to_station(all_data_list_tunnel[i][k].station) == walk_way_list[i][j][2])
                        {
                            isAdd = false;
                            temp.Add(all_data_list_tunnel[i][k]);
                            output.Add(temp);
                            break;
                        }
                        if (isAdd)
                        {
                            temp.Add(all_data_list_tunnel[i][k]);
                        }
                    }
                }
            }

            return output;
        }

        // 里程點的文字轉成有效浮點數
        public string to_station(string station)
        {
            string temp = station.Split('+')[0] + station.Split('+')[1];
            temp = temp.Split('.')[0];
            return Convert.ToInt32(temp).ToString();
        }
    }
}
