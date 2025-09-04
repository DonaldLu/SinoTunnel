using Autodesk.Revit.DB;
using DataObject;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace SinoTunnel_2020
{
    class readfile
    {
        string path = Form1.path;
        //public static string path = @"C:\Users\User\Desktop\潛盾隧道\潛盾隧道\"; 

        public IList<data_object> data_list = new List<data_object>();
        public IList<data_object> data_list2 = new List<data_object>();
        public IList<data_object> data_list_tunnel = new List<data_object>();
        public IList<data_object> data_list_tunnel2 = new List<data_object>();
        public IList<data_object> data_list_cd_channel = new List<data_object>(); //隧道下行線里程
        public IList<data_object> data_list_cd_channel2 = new List<data_object>(); //隧道上行線里程
        public properties_object properties = new properties_object();
        public rebar_properties rebar = new rebar_properties();
        public IList<envelope_object> envelope_1 = new List<envelope_object>();
        public IList<envelope_object> envelope_2 = new List<envelope_object>();
        public IList<envelope_object> envelope_3_1 = new List<envelope_object>();
        public IList<envelope_object> envelope_3_2 = new List<envelope_object>();
        public setting_station setting_Station = new setting_station();
        public setting_station setting_Station2 = new setting_station();

        public string firstStation;

        //string file_name = "20191120-萬大線BIM資料格式(隧道)_0120_0150.xlsx";
        string file_name = Form1.name;
        public List<Tuple<string, string>> tunnel_endpts_info() //隧道頭尾里程資料讀取***
        {
            List<Tuple<string, string>> miles_pts = new List<Tuple<string, string>>();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);

            foreach (string up_n_dn in new List<string> { "軌道線形 (UP)", "軌道線形 (DN)" })
            {
                Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[up_n_dn];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int row_count = xlRange.Rows.Count;
                for(int r=1; r<=row_count;r++)
                {
                    if (GetCell(xlRange, r, 2).Value2 == null)
                    //if (xlRange.Cells[r, 2].Value2 == null)
                    {
                        row_count = r-1;
                        break;
                    }
                }
                string start_pt = GetCell(xlRange, 2, 2).Value2.ToString();
                string end_pt = GetCell(xlRange, row_count, 2).Value2.ToString();
                //string start_pt = xlRange.Cells[2, 2].Value2.ToString();
                //string end_pt = xlRange.Cells[row_count, 2].Value2.ToString();
                miles_pts.Add(Tuple.Create(start_pt, end_pt));
            }
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
            return miles_pts;
        }

        public List<Dictionary<string, string>> Attached_facilities_info()
        {
            List<Dictionary<string, string>> facilities_info = new List<Dictionary<string, string>>();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["附掛設施"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                Dictionary<string, string> one_info = new Dictionary<string, string>();
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    if (GetCell(xlRange, i, j).Value2 != null)
                    { one_info.Add(GetCell(xlRange, 1, j).Value2.ToString(), GetCell(xlRange, i, j).Value2.ToString()); }
                    else
                    { one_info.Add(GetCell(xlRange, 1, j).Value2.ToString(), "0"); }
                    //if (xlRange.Cells[i, j].Value2 != null)
                    //{ one_info.Add(xlRange.Cells[1, j].Value2.ToString(), xlRange.Cells[i, j].Value2.ToString()); }
                    //else
                    //{ one_info.Add(xlRange.Cells[1, j].Value2.ToString(), "0"); }
                }
                facilities_info.Add(one_info);
            }
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
            return facilities_info;
        }

        public List<Tuple<string, Tuple<string, string>>> miles_info() //仰拱、道床和排水設施里程資料讀取
        {
            List<Tuple<string, Tuple<string, string>>> miles_pts = new List<Tuple<string, Tuple<string, string>>>();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);

            foreach (string up_n_dn in new List<string> { "軌道線形 (UP)", "軌道線形 (DN)" })
            {
                Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[up_n_dn];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                Excel.Range row_Range = (Excel.Range)xlWorksheet.Rows[1];

                foreach (string item in new List<string> { "仰拱", "道床", "排水溝" })
                {
                    Tuple<int, int> pos = FindAddress(item, row_Range);
                    int i = 1;
                    do
                    {                        
                        string item_name = GetCell(xlRange, pos.Item1 + i, pos.Item2).Value2.ToString();
                        string start_pts = GetCell(xlRange, pos.Item1 + i, pos.Item2 + 1).Value2.ToString();
                        string end_pts = GetCell(xlRange, pos.Item1 + i, pos.Item2 + 2).Value2.ToString();
                        //string item_name = xlRange.Cells[pos.Item1 + i, pos.Item2].Value2.ToString();
                        //string start_pts = xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2.ToString();
                        //string end_pts = xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2.ToString();
                        Tuple<string, string> tu_start_n_end = Tuple.Create(start_pts, end_pts);
                        miles_pts.Add(Tuple.Create(item_name, tu_start_n_end));
                        i++;
                    } 
                    while (GetCell(xlRange, pos.Item1 + i, pos.Item2 + 1).Value2 != null);
                    //while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

                }
            }
            xlWorkbook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
            return miles_pts;
        }

        public List<string> cd_channel_points() //需創建聯絡通道之里程
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["模型輸入資料"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            List<string> channel_pts = new List<string>();

            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                if (GetCell(xlRange, i, 1).Value2.ToString() == "聯絡通道里程_上行")
                //if (xlRange.Cells[i, 1].Value2.ToString() == "聯絡通道里程_上行")
                {
                    Excel.Range row_of_cd_channel = (Excel.Range)xlRange.Rows[i];
                    for (int j = 2; j <= row_of_cd_channel.Columns.Count ; j++)
                    {
                        if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                        //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            channel_pts.Add(GetCell(xlRange, i, j).Value2.ToString());
                            //channel_pts.Add(xlRange.Cells[i, j].Value2.ToString());
                        }
                    }
                    break;
                }
            }

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
            return channel_pts;

        }

        public List<string> cd_channel_points_dn() /////
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["模型輸入資料"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            List<string> channel_pts = new List<string>();

            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                try
                {                    
                    if (GetCell(xlRange, i, 1).Value2.ToString() == "聯絡通道里程_下行")
                    //if (xlRange.Cells[i, 1].Value2.ToString() == "聯絡通道里程_下行")
                    {
                        Excel.Range row_of_cd_channel = (Excel.Range)xlRange.Rows[i];
                        for (int j = 2; j <= row_of_cd_channel.Columns.Count; j++)
                        {
                            if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                            //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            {
                                channel_pts.Add(GetCell(xlRange, i, j).Value2.ToString());
                                //channel_pts.Add(xlRange.Cells[i, j].Value2.ToString());
                            }
                        }
                        break;
                    }
                }
                catch { }
            }

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
            return channel_pts;
        }
        
        public void read_tunnel_point()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["隧道線形 (DN)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            List<string> chs_pts = cd_channel_points_dn();

            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                double X = 0, Y = 0, Z = 0;
                data_object data = new data_object();
                if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
                //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    //write the value to the console
                    if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) { Int32.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.id); }
                        else if (j == 2) { data.station = GetCell(xlRange, i, j).Value2.ToString(); }
                        else if (j == 3) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out X); X = X / 0.3048; }
                        else if (j == 4) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out Y); Y = Y / 0.3048; }
                        else if (j == 5) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out Z); Z = Z / 0.3048; }
                        else if (j == 7) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.horizontal_angle); }
                        else if (j == 9) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.vertical_angle); }
                        else if (j == 10) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.offset); }
                        else if (j == 11) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.super_high_angle); }
                        //if (j == 1) { Int32.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.id); }
                        //else if (j == 2) { data.station = xlRange.Cells[i, j].Value2.ToString(); }
                        //else if (j == 3) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out X); X = X / 0.3048; }
                        //else if (j == 4) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out Y); Y = Y / 0.3048; }
                        //else if (j == 5) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out Z); Z = Z / 0.3048; }
                        //else if (j == 7) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.horizontal_angle); }
                        //else if (j == 9) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.vertical_angle); }
                        //else if (j == 10) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.offset); }
                        //else if (j == 11) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.super_high_angle); }
                    }

                    //add useful things here!   
                }
                if (i == 2) firstStation = data.station;
                data.start_point = new XYZ(X, Y, Z);
                data_list_tunnel.Add(data);
                foreach (string ele in chs_pts)
                {
                    if (data.station == ele) { data_list_cd_channel.Add(data); }
                }
            }
            string message = "";
            foreach (data_object f in data_list_cd_channel)
            {
                message += f.station;
                message += "\n";
            }
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
        }

        public void read_tunnel_point2()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["隧道線形 (UP)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            List<string> chs_pts = cd_channel_points();

            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                double X = 0, Y = 0, Z = 0;
                data_object data = new data_object();
                if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
                //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    //write the value to the console
                    if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) { Int32.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.id); }
                        else if (j == 2) { data.station = GetCell(xlRange, i, j).Value2.ToString(); }
                        else if (j == 3) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out X); X = X / 0.3048; }
                        else if (j == 4) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out Y); Y = Y / 0.3048; }
                        else if (j == 5) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out Z); Z = Z / 0.3048; }
                        else if (j == 7) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.horizontal_angle); }
                        else if (j == 9) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.vertical_angle); }
                        else if (j == 10) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.offset); }
                        else if (j == 11) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.super_high_angle); }
                        //if (j == 1) { Int32.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.id); }
                        //else if (j == 2) { data.station = xlRange.Cells[i, j].Value2.ToString(); }
                        //else if (j == 3) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out X); X = X / 0.3048; }
                        //else if (j == 4) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out Y); Y = Y / 0.3048; }
                        //else if (j == 5) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out Z); Z = Z / 0.3048; }
                        //else if (j == 7) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.horizontal_angle); }
                        //else if (j == 9) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.vertical_angle); }
                        //else if (j == 10) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.offset); }
                        //else if (j == 11) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.super_high_angle); }
                    }

                    //add useful things here!   
                }
                data.start_point = new XYZ(X, Y, Z);
                data_list_tunnel2.Add(data);
                foreach (string ele in chs_pts)
                {
                    if (data.station == ele) { data_list_cd_channel2.Add(data); }
                }
            }
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(xlApp);
        }

        public contact_channel_properties read_contact_tunnel()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["模型輸入資料"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            contact_channel_properties setting = new contact_channel_properties();

            //起始聯絡通道里程index, 注意聯絡通道里程_上行的書寫形式: ex 0+244 & 0+244.00
            int start_index = FindAddress(data_list_cd_channel[0].station.ToString(), xlRange).Item2;

            for (int j = start_index; j <= data_list_cd_channel.Count + (start_index - 1); j++)
            {
                for (int i = 2; i <= xlRange.Rows.Count; i++)
                {
                    string name = "";
                    bool success; //default
                    double double_temp = double.NaN;
                    int int_temp = 0;
                    try { name = GetCell(xlRange, i, 1).Value2.ToString(); }
                    //try { name = xlRange.Cells[i, 1].Value2.ToString(); }
                    catch { continue; }

                    switch (name)
                    {
                        case "隧道內徑":
                            success = double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out double_temp);
                            setting.tunnel_redius.Add(double_temp / 2);
                            break;

                        case "聯絡通道型式":
                            string temp = GetCell(xlRange, i, j).Value2.ToString();
                            //string temp = xlRange.Cells[i, j].Value2.ToString();
                            if (temp == "平行")
                                setting.tunnel_type.Add(0);
                            else if (temp == "高差")
                                setting.tunnel_type.Add(1);
                            break;

                        case "聯絡通道底部高程":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.tunnel_elevation.Add(double_temp);
                            break;

                        case "聯絡通道上部圓弧半徑":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.up_arc_radius.Add(double_temp);
                            break;

                        case "聯絡通道上部厚度":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.up_arc_thickness.Add(double_temp);
                            break;

                        case "聯絡通道下部垂直長度":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.dn_height.Add(double_temp);
                            break;

                        case "聯絡通道下部厚度":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.dn_thickness.Add(double_temp);
                            break;

                        case "聯絡通道開孔寬度":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.hollow_width.Add(double_temp);
                            break;

                        case "聯絡通道開孔高度":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.hollow_height.Add(double_temp);
                            break;

                        case "聯絡通道開孔範圍軸向深度":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.hollow_depth.Add(double_temp);
                            break;

                        case "聯絡通道開孔範圍底部增加厚度":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.hollow_bottom_thickness.Add(double_temp);
                            break;

                        case "聯絡通道初期支撐厚度":
                            success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                            //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                            setting.support_thickness.Add(double_temp);
                            break;

                        case "聯絡通道層數":
                            if (GetCell(xlRange, i, j).Value2 != null && GetCell(xlRange, i, j) != null)
                            //if (xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j] != null)
                            {
                                success = Int32.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out int_temp);
                                //success = Int32.TryParse(xlRange.Cells[i, j].Value2.ToString(), out int_temp);
                                if (success == true && int_temp > 1)
                                    setting.path_levels.Add(int_temp); //高差                           
                                else if (success == true && int_temp == 1)
                                    setting.path_levels.Add(int_temp); //平行       
                            }
                            else
                            {
                                setting.path_levels.Add(1); //平行 
                                break;
                            }
                            break;

                        case "聯絡通道第一層長度":
                            if (GetCell(xlRange, i, j).Value2 != null && GetCell(xlRange, i, j) != null)
                            //if (xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j] != null)
                            {
                                success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                                //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                                if (success)
                                    setting.level_one.Add(double_temp);
                            }
                            else
                            {
                                setting.level_one.Add(0);
                                break;
                            }
                            break;

                        case "聯絡通道第二層長度":
                            if (GetCell(xlRange, i, j).Value2 != null && GetCell(xlRange, i, j) != null)
                            //if (xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j] != null)
                            {
                                success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                                //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                                if (success)
                                    setting.level_two.Add(double_temp);
                            }
                            else
                            {
                                setting.level_two.Add(0);
                                break;
                            }
                            break;

                        case "聯絡通道第三層長度":
                            if (GetCell(xlRange, i, j).Value2 != null && GetCell(xlRange, i, j) != null)
                            //if (xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j] != null)
                            {
                                success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                                //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                                if (success)
                                    setting.level_three.Add(double_temp);
                            }
                            else
                            {
                                setting.level_three.Add(0);
                                break;
                            }
                            break;

                        case "聯絡通道防火門層數":
                            if (GetCell(xlRange, i, j).Value2 != null && GetCell(xlRange, i, j) != null)
                            //if (xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j] != null)
                            {
                                success = Int32.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out int_temp);
                                //success = Int32.TryParse(xlRange.Cells[i, j].Value2.ToString(), out int_temp);
                                if (success)
                                    setting.door_levels.Add(int_temp);
                            }
                            else
                            {
                                setting.door_levels.Add(1); //Default為樓層1
                                break;
                            }
                            break;

                        case "聯絡通道防火門距離":
                            if (GetCell(xlRange, i, j).Value2 != null && GetCell(xlRange, i, j) != null)
                            //if (xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j] != null)
                            {
                                success = double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out double_temp);
                                //success = double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out double_temp);
                                if (success)
                                    setting.door_dis.Add(double_temp);
                            }
                            else
                            {
                                setting.door_dis.Add(0);
                                break;
                            }
                            break;

                        default:
                            break;
                    }
                }
            }
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
            return setting;
        }

        public void read_rebar_properties()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["配筋資料"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            //外側及內側邊緣
            double.TryParse(GetCell(xlRange, 5, 2).Value2.ToString(), out rebar.main_inner_protect);
            double.TryParse(GetCell(xlRange, 4, 2).Value2.ToString(), out rebar.main_outer_protect);
            //主筋徑向邊緣
            double.TryParse(GetCell(xlRange, 3, 2).Value2.ToString(), out rebar.main_side_protect);

            ////外側及內側邊緣
            //double.TryParse(xlRange.Cells[5, 2].Value2.ToString(), out rebar.main_inner_protect);
            //double.TryParse(xlRange.Cells[4, 2].Value2.ToString(), out rebar.main_outer_protect);
            ////主筋徑向邊緣
            //double.TryParse(xlRange.Cells[3, 2].Value2.ToString(), out rebar.main_side_protect);

            for (int i = 6; i <= 26; i++)
            {
                string name = "";
                try
                {
                    name = GetCell(xlRange, 6, i).Value2.ToString();
                    //name = xlRange.Cells[6, i].Value2.ToString();
                }
                catch { continue; }
                switch (name)
                {
                    case "主筋A環間距":
                        for (int j = 0; j < 8; j++)
                        {
                            rebar.main_A_distance.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.main_A_distance.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    case "主筋B環間距":
                        for (int j = 0; j < 8; j++)
                        {
                            rebar.main_B_distance.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.main_B_distance.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    case "主筋K環間距":
                        for (int j = 0; j < 8; j++)
                        {
                            rebar.main_K_distance.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.main_K_distance.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    case "剪力筋A環間距":
                        for (int j = 0; j < 22; j++)
                        {
                            rebar.shear_A_distance.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.shear_A_distance.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    case "剪力筋A環type":
                        for (int j = 0; j < 22; j++)
                        {
                            rebar.shear_A_type.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.shear_A_type.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    case "剪力筋B環間距":
                        for (int j = 0; j < 20; j++)
                        {
                            rebar.shear_B_distance.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.shear_B_distance.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    case "剪力筋B環type":
                        for (int j = 0; j < 20; j++)
                        {
                            rebar.shear_B_type.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.shear_B_type.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    case "剪力筋K環間距":
                        for (int j = 0; j < 4; j++)
                        {
                            rebar.shear_K_distance.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.shear_K_distance.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    case "剪力筋K環type":
                        for (int j = 0; j < 4; j++)
                        {
                            rebar.shear_K_type.Add(GetCell(xlRange, j + 7, i).Value2.ToString());
                            //rebar.shear_K_type.Add(xlRange.Cells[j + 7, i].Value2.ToString());
                        }
                        break;
                    default:
                        continue;
                }
            }

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }

        public void read_point()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["軌道線形 (DN)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                double X = 0, Y = 0, Z = 0;
                data_object data = new data_object();
                if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
                //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    //write the value to the console
                    if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) { Int32.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.id); }
                        else if (j == 2) { data.station = GetCell(xlRange, i, j).Value2.ToString(); }
                        else if (j == 3) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out X); X = X / 0.3048; }
                        else if (j == 4) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out Y); Y = Y / 0.3048; }
                        else if (j == 5) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out Z); Z = Z / 0.3048; }
                        else if (j == 6) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.super_high_angle); }
                        //if (j == 1) { Int32.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.id); }
                        //else if (j == 2) { data.station = xlRange.Cells[i, j].Value2.ToString(); }
                        //else if (j == 3) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out X); X = X / 0.3048; }
                        //else if (j == 4) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out Y); Y = Y / 0.3048; }
                        //else if (j == 5) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out Z); Z = Z / 0.3048; }
                        //else if (j == 6) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.super_high_angle); }
                    }
                    //else if (j == 9) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.vertical_angle); }

                    //add useful things here!   
                }
                data.start_point = new XYZ(X, Y, Z);
                data_list.Add(data);
            }
            
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(xlApp);
        }

        public void read_point2()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["軌道線形 (UP)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                double X = 0, Y = 0, Z = 0;
                data_object data = new data_object();
                if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
                //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
                for (int j = 1; j <= xlRange.Columns.Count; j++)
                {
                    //write the value to the console
                    if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) { Int32.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.id); }
                        else if (j == 2) { data.station = GetCell(xlRange, i, j).Value2.ToString(); }
                        else if (j == 3) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out X); X = X / 0.3048; }
                        else if (j == 4) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out Y); Y = Y / 0.3048; }
                        else if (j == 5) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out Z); Z = Z / 0.3048; }
                        else if (j == 6) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out data.super_high_angle); }
                        //if (j == 1) { Int32.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.id); }
                        //else if (j == 2) { data.station = xlRange.Cells[i, j].Value2.ToString(); }
                        //else if (j == 3) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out X); X = X / 0.3048; }
                        //else if (j == 4) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out Y); Y = Y / 0.3048; }
                        //else if (j == 5) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out Z); Z = Z / 0.3048; }
                        //else if (j == 6) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.super_high_angle); }
                    }
                    //else if (j == 9) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out data.vertical_angle); }

                    //add useful things here!   
                }
                data.start_point = new XYZ(X, Y, Z);
                data_list2.Add(data);
            }

            xlWorkbook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(xlApp);
        }

        public void read_properties()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["模型輸入資料"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                string name = "";
                try { name = GetCell(xlRange, i, 1).Value2.ToString(); }
                //try { name = xlRange.Cells[i, 1].Value2.ToString(); }
                catch { continue; }
                switch (name)
                {
                    case "編號":
                        properties.id = GetCell(xlRange, i, 2).Value2.ToString();
                        //properties.id = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "隧道內徑":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.inner_diameter);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.inner_diameter);
                        break;
                    case "隧道中心點":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.center_point);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.center_point);
                        break;
                    case "環片寬度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.width);
                        break;
                    case "環片厚度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.thickness);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.thickness);
                        break;
                    case "A環片數量":
                        int.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_A_q);
                        //int.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_A_q);
                        break;
                    case "B環片數量":
                        int.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_B_q);
                        //int.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_B_q);
                        break;
                    case "K環片數量":
                        int.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_K_q);
                        //int.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_K_q);
                        break;
                    case "環圈交錯角度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.displacement_angle);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.displacement_angle);
                        break;
                    //case "Type_K環片插入型式":
                    //    int.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out properties.type_K_insert_type);
                    //    TaskDialog.Show("message", properties.type_K_insert_type.ToString());
                    //    break;
                    case "A環片角度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_A_1);
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_A_2);
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_A_3);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_A_1);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_A_2);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_A_3);
                        break;
                    case "鋼環片S環片角度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_S_1);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_S_1);
                        break;
                    //case "Type_A-2(盾面)":
                    //    double.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out properties.Type_A_2);
                    //    break;
                    //case "Type_A-3(盾面)":
                    //    double.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out properties.Type_A_3);
                    //    break;
                    //case "Type_B-1(盾面)":
                    //    double.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out properties.Type_B_1_head);
                    //    break;
                    //case "Type_B-2(盾面)":
                    //    double.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out properties.Type_B_2_head);
                    //    break;
                    case "K環片角度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_K_1_head);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_K_1_head);
                        break;
                    case "B環片角度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_B_1_head);
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.Type_B_2_head);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_B_1_head);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.Type_B_2_head);
                        break;
                    //case "Type_B-2(尾面)":
                    //    double.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out properties.Type_B_2_tail);
                    //    break;
                    //case "Type_K(尾面)":
                    //    double.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out properties.Type_K_1_tail);
                    //    break;
                    case "K環單邊內縮量":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.displacement);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.displacement);
                        break;
                    case "U形槽鋼走道側底":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.U_steel_aisle_bottom);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.U_steel_aisle_bottom);
                        break;
                    case "U形槽鋼走道側頂":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.U_steel_aisle_top);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.U_steel_aisle_top);
                        break;
                    case "U形槽鋼非走道側底":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.U_steel_nonaisle_bottom);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.U_steel_nonaisle_bottom);
                        break;
                    case "U形槽鋼非走道側頂":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.U_steel_nonaisle_top);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.U_steel_nonaisle_top);
                        break;

                    //仰拱
                    case "仰拱洩水斜率":
                        string temp = GetCell(xlRange, i, 2).Value2.ToString();
                        //string temp = xlRange.Cells[i, 2].Value2.ToString();
                        double temp2 = Int32.Parse(temp.Split('/')[1]);
                        properties.inverted_arc_slope = (Math.Atan(1 / temp2) * 180 / Math.PI).ToString();
                        break;
                    
                    //U形排水溝(標準排水溝)
                    case "U形排水溝蓋板寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_cover_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_cover_width);
                        break;
                    case "U形排水溝蓋板長":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_cover_length);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_cover_length);
                        break;
                    case "U形排水溝蓋板厚":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_cover_thick);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_cover_thick);
                        break;
                    case "U形排水溝凹槽寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_groove_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_groove_width);
                        break;
                    case "U形排水溝深度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_depth);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_depth);
                        break;
                    case "U形排水溝半徑":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_radius);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_radius);
                        break;
                    case "U形排水溝長邊距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_long_side_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_long_side_dis);
                        break;
                    case "U形排水溝短邊距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_short_side_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_short_side_dis);
                        break;
                    case "U形排水溝熱鍍鋅扁鋼厚":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_zn_steel_thick);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_zn_steel_thick);
                        break;
                    case "U形排水溝熱鍍鋅扁鋼長間距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_zn_steel_long_pitch);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_zn_steel_long_pitch);
                        break;
                    case "U形排水溝熱鍍鋅扁鋼短間距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_zn_steel_short_pitch);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_zn_steel_short_pitch);
                        break;
                    case "U形排水溝熱鍍鋅間距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.u_zn_steel_pitch);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.u_zn_steel_pitch);
                        break;

                    //PVC管排水溝
                    case "PVC管半徑":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_radius);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_radius);
                        break;
                    case "PVC管厚度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_thick);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_thick);
                        break;
                    case "PVC管明溝深度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_gutter_depth);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_gutter_depth);
                        break;
                    case "PVC管明溝半徑":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_gutter_radius);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_gutter_radius);
                        break;
                    case "PVC管仰拱半徑":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_inverted_arc_radius);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_inverted_arc_radius);
                        break;
                    case "PVC管明溝凹槽寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_gutter_witdh);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_gutter_witdh);
                        break;
                    case "PVC管熱鍍鋅蓋寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_zn_cover_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_zn_cover_width);
                        break;
                    case "PVC管熱鍍鋅蓋長":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_zn_cover_length);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_zn_cover_length);
                        break;
                    case "PVC管熱鍍鋅蓋厚":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_zn_cover_thick);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_zn_cover_thick);
                        break;
                    case "PVC管熱鍍鋅蓋鋼厚":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_zn_cover_steel_thick);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_zn_cover_steel_thick);
                        break;
                    case "PVC管熱鍍鋅蓋長間距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_zn_cover_long_pitch);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_zn_cover_long_pitch);
                        break;
                    case "PVC管熱鍍鋅蓋短間距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_zn_cover_short_pitch);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_zn_cover_short_pitch);
                        break;
                    case "PVC管熱鍍鋅蓋間距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.pvc_zn_cover_pitch);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.pvc_zn_cover_pitch);
                        break;
                    
                    //浮動式道床排水溝
                    case "浮動式道床明溝半徑":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.float_gutter_radius);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.float_gutter_radius);
                        break;
                    case "浮動式道床明溝深":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.float_gutter_depth);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.float_gutter_depth);
                        break;
                    
                    //走道
                    case "走道頂對線形高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.walkway_top_elevation);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.walkway_top_elevation);
                        break;
                    case "走道邊緣與隧道中心距離":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.walkway_edge_to_rail_center_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.walkway_edge_to_rail_center_dis);
                        break;
                    case "走道邊緣與軌道中心距離": // 舊版Excel
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.walkway_edge_to_rail_center_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.walkway_edge_to_rail_center_dis);
                        break;
                    case "走道需求寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.walkway_requirement_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.walkway_requirement_width);
                        break;
                    case "連結處厚度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.connection_thick);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.connection_thick);
                        break;
                    case "走道斜率":
                        string temp3 = GetCell(xlRange, i, 2).Value2.ToString();
                        //string temp3 = xlRange.Cells[i, 2].Value2.ToString();
                        double temp4 = Int32.Parse(temp3.Split('/')[1]);
                        properties.inverted_arc_slope = (Math.Atan(1 / temp4) * 180 / Math.PI).ToString();
                        break;
                    case "走道突出寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.walkway_protrusion_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.walkway_protrusion_width);
                        break;
                    case "走道突出深":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.walkway_protrusion_depth);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.walkway_protrusion_depth);
                        break;
                    case "走道突出底":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.walkway_protrusion_bottom);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.walkway_protrusion_bottom);
                        break;
                    
                    //電纜溝槽
                    case "走道電纜溝槽與走道距離":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_distance);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_distance);
                        break;
                    case "走道電纜溝槽上底寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_top_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_top_width);
                        break;
                    case "走道電纜溝槽下底寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_bottom_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_bottom_width);
                        break;
                    case "走道電纜溝槽深度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_depth);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_depth);
                        break;

                    //電纜溝槽混凝土蓋板
                    case "走道電纜溝槽混凝土蓋板寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_cover_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_cover_width);
                        break;
                    case "走道電纜溝槽混凝土蓋板突出邊距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_cover_stick_out_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_cover_stick_out_dis);
                        break;
                    case "走道電纜溝槽混凝土蓋板突出厚":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_cover_stick_out_thick);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_cover_stick_out_thick);
                        break;
                    case "走道電纜溝槽混凝土蓋板突出寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_cover_stick_out_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_cover_stick_out_width);
                        break;
                    case "走道電纜溝槽混凝土蓋板厚":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_cover_thick);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_cover_thick);
                        break;
                    case "走道電纜溝槽混凝土蓋板長":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.cable_cover_length);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.cable_cover_length);
                        break;

                    //混凝土強度
                    case "仰拱混凝土強度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.inverted_arc_concrete_strength);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.inverted_arc_concrete_strength);
                        break;
                    case "走道混凝土強度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.walk_way_concrete_strength);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.walk_way_concrete_strength);
                        break;




                    //平板式道床
                    case "平板式道床仰拱頂部高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.flat_top_height);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.flat_top_height);
                        break;
                    case "平板式道床高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.flat_elevation);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.flat_elevation);
                        break;
                    case "平板式道床寬度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.flat_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.flat_width);
                        break;

                    //標準道床
                    case "標準道床仰拱頂部高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.standerd_top_height);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.standerd_top_height);
                        break;

                    //浮動式道床
                    case "浮動式道床仰拱頂部高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out properties.float_top_height);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out properties.float_top_height);
                        break;

                    default:
                        continue;
                }
            }

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(xlApp);
        }

        public void read_envelope_1()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["包絡線幾何 (DN)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            // 培文改, 不限制13個點位
            envelope_object envelope_data = new envelope_object();
            for (int i = 3; i <= xlRange.Rows.Count; i++)
            {
                if (GetCell(xlRange, i, 7).Value2 != null) { envelope_data = new envelope_object(); }
                //if (xlRange.Cells[i, 7].Value2 != null) { envelope_data = new envelope_object(); }
                double mX = 0, mY = 0, mZ = 0;
                double cX = 0, cY = 0, cZ = 0;
                if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
                //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
                for (int j = 1; j <= 6; j++)
                {
                    if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mX); mX = mX / 0.3048; }
                        else if (j == 2) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mY); mY = mY / 0.3048; }
                        else if (j == 3) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
                        else if (j == 4) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out cX); cX = cX / 0.3048; }
                        else if (j == 5) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out cY); cY = cY / 0.3048; }
                        else if (j == 6) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out cZ); cZ = cZ / 0.3048; }
                        //if (j == 1) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mX); mX = mX / 0.3048; }
                        //else if (j == 2) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mY); mY = mY / 0.3048; }
                        //else if (j == 3) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
                        //else if (j == 4) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out cX); cX = cX / 0.3048; }
                        //else if (j == 5) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out cY); cY = cY / 0.3048; }
                        //else if (j == 6) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out cZ); cZ = cZ / 0.3048; }
                    }
                }
                XYZ m = new XYZ(mX, mY, mZ);
                XYZ c = new XYZ(cX, cY, cZ);
                envelope_data.Dynamic_envelope.Add(m);
                envelope_data.Vehicle_envelope.Add(c);
                if (GetCell(xlRange, i + 1, 7).Value2 != null || i == xlRange.Rows.Count) { envelope_1.Add(envelope_data); }
                //if (xlRange.Cells[i + 1, 7].Value2 != null || i == xlRange.Rows.Count) { envelope_1.Add(envelope_data); }
            }

            //// 台大
            //for (int i = 3; i <= xlRange.Rows.Count; i = i + 13)
            //{
            //    double mX = 0, mY = 0, mZ = 0;
            //    double cX = 0, cY = 0, cZ = 0;
            //    envelope_object data = new envelope_object();
            //    if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null)
            //    //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null)
            //    {
            //        break;
            //    }
            //    List<XYZ> move_point = new List<XYZ>();
            //    List<XYZ> car_point = new List<XYZ>();
            //    for (int k = 0; k < 13; k++)
            //    {
            //        for (int j = 1; j <= xlRange.Columns.Count; j++)
            //        {
            //            //write the value to the console
            //            if (GetCell(xlRange, i + k, j) != null && GetCell(xlRange, i + k, j).Value2 != null)
            //                //if (xlRange.Cells[i + k, j] != null && xlRange.Cells[i + k, j].Value2 != null)
            //                if (j == 1)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mX);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mX);
            //                    mX = mX / 0.3048;
            //                }
            //                else if (j == 2)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mY);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mY);
            //                    mY = mY / 0.3048;
            //                }
            //                else if (j == 3)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mZ);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mZ);
            //                    mZ = mZ / 0.3048;
            //                }
            //                else if (j == 4)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out cX);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out cX);
            //                    cX = cX / 0.3048;
            //                }
            //                else if (j == 5)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out cY);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out cY);
            //                    cY = cY / 0.3048;
            //                }
            //                else if (j == 6)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out cZ);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out cZ);
            //                    cZ = cZ / 0.3048;
            //                }

            //            //add useful things here!   
            //        }
            //        XYZ m = new XYZ(mX, mY, mZ);
            //        XYZ c = new XYZ(cX, cY, cZ);
            //        data.Dynamic_envelope.Add(m);
            //        data.Vehicle_envelope.Add(c);
            //    }
            //    envelope_1.Add(data);
            //}

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }

        public void read_envelope_2()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["包絡線幾何 (UP)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            // 培文改, 不限制13個點位
            envelope_object envelope_data = new envelope_object();
            for (int i = 3; i <= xlRange.Rows.Count; i++)
            {
                if (GetCell(xlRange, i, 7).Value2 != null) { envelope_data = new envelope_object(); }
                //if (xlRange.Cells[i, 7].Value2 != null) { envelope_data = new envelope_object(); }
                double mX = 0, mY = 0, mZ = 0;
                double cX = 0, cY = 0, cZ = 0;
                if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
                //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
                for (int j = 1; j <= 6; j++)
                {
                    if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mX); mX = mX / 0.3048; }
                        else if (j == 2) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mY); mY = mY / 0.3048; }
                        else if (j == 3) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
                        else if (j == 4) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out cX); cX = cX / 0.3048; }
                        else if (j == 5) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out cY); cY = cY / 0.3048; }
                        else if (j == 6) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out cZ); cZ = cZ / 0.3048; }
                        //if (j == 1) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mX); mX = mX / 0.3048; }
                        //else if (j == 2) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mY); mY = mY / 0.3048; }
                        //else if (j == 3) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
                        //else if (j == 4) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out cX); cX = cX / 0.3048; }
                        //else if (j == 5) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out cY); cY = cY / 0.3048; }
                        //else if (j == 6) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out cZ); cZ = cZ / 0.3048; }
                    }
                }
                XYZ m = new XYZ(mX, mY, mZ);
                XYZ c = new XYZ(cX, cY, cZ);
                envelope_data.Dynamic_envelope.Add(m);
                envelope_data.Vehicle_envelope.Add(c);
                if (GetCell(xlRange, i + 1, 7).Value2 != null || i == xlRange.Rows.Count) { envelope_2.Add(envelope_data); }
                //if (xlRange.Cells[i + 1, 7].Value2 != null || i == xlRange.Rows.Count) { envelope_2.Add(envelope_data); }
            }

            //// 台大
            //for (int i = 3; i <= xlRange.Rows.Count; i = i + 13)
            //{
            //    double mX = 0, mY = 0, mZ = 0;
            //    double cX = 0, cY = 0, cZ = 0;
            //    envelope_object data = new envelope_object();
            //    if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null)
            //    //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null)
            //    {
            //        break;
            //    }
            //    List<XYZ> move_point = new List<XYZ>();
            //    List<XYZ> car_point = new List<XYZ>();
            //    for (int k = 0; k < 13; k++)
            //    {

            //        for (int j = 1; j <= xlRange.Columns.Count; j++)
            //        {

            //            //write the value to the console
            //            if (GetCell(xlRange, i + k, j) != null && GetCell(xlRange, i + k, j).Value2 != null)
            //            //if (xlRange.Cells[i + k, j] != null && xlRange.Cells[i + k, j].Value2 != null)
            //                if (j == 1)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mX);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mX);
            //                    mX = mX / 0.3048;
            //                }
            //                else if (j == 2)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mY);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mY);
            //                    mY = mY / 0.3048;
            //                }
            //                else if (j == 3)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mZ);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mZ);
            //                    mZ = mZ / 0.3048;
            //                }
            //                else if (j == 4)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out cX);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out cX);
            //                    cX = cX / 0.3048;
            //                }
            //                else if (j == 5)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out cY);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out cY);
            //                    cY = cY / 0.3048;
            //                }
            //                else if (j == 6)
            //                {
            //                    double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out cZ);
            //                    //double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out cZ);
            //                    cZ = cZ / 0.3048;
            //                }

            //            //add useful things here!   
            //        }
            //        XYZ m = new XYZ(mX, mY, mZ);
            //        XYZ c = new XYZ(cX, cY, cZ);
            //        data.Dynamic_envelope.Add(m);
            //        data.Vehicle_envelope.Add(c);
            //    }
            //    envelope_2.Add(data);
            //}

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }

        public void read_envelope_3_1()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["第三軌包絡線 (UP)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            envelope_object envelope_data = new envelope_object();
            for (int i = 3; i <= xlRange.Rows.Count; i++)
            {
                if (GetCell(xlRange, i, 4).Value2 != null) { envelope_data = new envelope_object(); }
                //if (xlRange.Cells[i, 4].Value2 != null) { envelope_data = new envelope_object(); }
                double mX = 0, mY = 0, mZ = 0;
                if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
                //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
                for (int j = 1; j <= 3; j++)
                {
                    if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mX); mX = mX / 0.3048; }
                        else if (j == 2) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mY); mY = mY / 0.3048; }
                        else if (j == 3) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
                        //if (j == 1) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mX); mX = mX / 0.3048; }
                        //else if (j == 2) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mY); mY = mY / 0.3048; }
                        //else if (j == 3) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
                    }
                }
                XYZ m = new XYZ(mX, mY, mZ);
                envelope_data.third_envelope.Add(m);
                if (GetCell(xlRange, i + 1, 4).Value2 != null || i == xlRange.Rows.Count) { envelope_3_1.Add(envelope_data); }
                //if (xlRange.Cells[i + 1, 4].Value2 != null || i == xlRange.Rows.Count) { envelope_3_1.Add(envelope_data); }
            }

            //// 台大
            //for (int i = 3; i <= xlRange.Rows.Count; i = i + 5)
            //{
            //    double mX = 0, mY = 0, mZ = 0;
            //    envelope_object data = new envelope_object();
            //    if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
            //    //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
            //    for (int k = 0; k < 5; k++)
            //    {
            //        for (int j = 1; j <= xlRange.Columns.Count; j++)
            //        {
            //            //write the value to the console
            //            if (GetCell(xlRange, i + k, j) != null && GetCell(xlRange, i + k, j).Value2 != null)
            //            //if (xlRange.Cells[i + k, j] != null && xlRange.Cells[i + k, j].Value2 != null)
            //            {
            //                if (j == 1) { double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mX); mX = mX / 0.3048; }
            //                else if (j == 2) { double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mY); mY = mY / 0.3048; }
            //                else if (j == 3) { double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
            //                //if (j == 1) { double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mX); mX = mX / 0.3048; }
            //                //else if (j == 2) { double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mY); mY = mY / 0.3048; }
            //                //else if (j == 3) { double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
            //                //add useful things here!
            //            }
            //        }
            //        XYZ m = new XYZ(mX, mY, mZ);
            //        data.third_envelope.Add(m);
            //    }
            //    envelope_3_1.Add(data);
            //}

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }

        public void read_envelope_3_2()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["第三軌包絡線 (DN)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            envelope_object envelope_data = new envelope_object();
            for (int i = 3; i <= xlRange.Rows.Count; i++)
            {
                if (GetCell(xlRange, i, 4).Value2 != null) { envelope_data = new envelope_object(); }
                //if (xlRange.Cells[i, 4].Value2 != null) { envelope_data = new envelope_object(); }
                double mX = 0, mY = 0, mZ = 0;
                if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
                //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
                for (int j = 1; j <= 3; j++)
                {
                    if (GetCell(xlRange, i, j) != null && GetCell(xlRange, i, j).Value2 != null)
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mX); mX = mX / 0.3048; }
                        else if (j == 2) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mY); mY = mY / 0.3048; }
                        else if (j == 3) { double.TryParse(GetCell(xlRange, i, j).Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
                        //if (j == 1) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mX); mX = mX / 0.3048; }
                        //else if (j == 2) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mY); mY = mY / 0.3048; }
                        //else if (j == 3) { double.TryParse(xlRange.Cells[i, j].Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
                    }
                }
                XYZ m = new XYZ(mX, mY, mZ);
                envelope_data.third_envelope.Add(m);
                if (GetCell(xlRange, i + 1, 4).Value2 != null || i == xlRange.Rows.Count) { envelope_3_2.Add(envelope_data); }
                //if (xlRange.Cells[i + 1, 4].Value2 != null || i == xlRange.Rows.Count) { envelope_3_2.Add(envelope_data); }
            }

            //// 台大
            //for (int i = 3; i <= xlRange.Rows.Count; i = i + 5)
            //{
            //    double mX = 0, mY = 0, mZ = 0;
            //    envelope_object data = new envelope_object();
            //    if (GetCell(xlRange, i, 1) == null || GetCell(xlRange, i, 1).Value2 == null) { break; }
            //    //if (xlRange.Cells[i, 1] == null || xlRange.Cells[i, 1].Value2 == null) { break; }
            //    for (int k = 0; k < 5; k++)
            //    {
            //        for (int j = 1; j <= xlRange.Columns.Count; j++)
            //        {
            //            //write the value to the console
            //            if (GetCell(xlRange, i + k, j) != null && GetCell(xlRange, i + k, j).Value2 != null)
            //            //if (xlRange.Cells[i + k, j] != null && xlRange.Cells[i + k, j].Value2 != null)
            //            {
            //                if (j == 1) { double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mX); mX = mX / 0.3048; }
            //                else if (j == 2) { double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mY); mY = mY / 0.3048; }
            //                else if (j == 3) { double.TryParse(GetCell(xlRange, i + k, j).Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
            //                //if (j == 1) { double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mX); mX = mX / 0.3048; }
            //                //else if (j == 2) { double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mY); mY = mY / 0.3048; }
            //                //else if (j == 3) { double.TryParse(xlRange.Cells[i + k, j].Value2.ToString(), out mZ); mZ = mZ / 0.3048; }
            //                //add useful things here!
            //            }
            //        }
            //        XYZ m = new XYZ(mX, mY, mZ);
            //        data.third_envelope.Add(m);
            //    }
            //    envelope_3_2.Add(data);
            //}

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }

        public void read_target_station()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["軌道線形 (DN)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            Excel.Range xlRange_partial = (Excel.Range)xlRange.Columns[2];

            Excel.Range start_range = null;
            Excel.Range end_range = null;

            string start_station = "";
            string end_station = "";

            for (int j = 7; j <= xlRange.Columns.Count; j = j + 3)
            {
                if (GetCell(xlRange, 1, j).Value2 != null && GetCell(xlRange, 1, j).Value2 != "")
                //if (xlRange.Cells[1, j].Value2 != null && xlRange.Cells[1, j].Value2 != "")
                {
                    for (int i = 2; i <= xlRange.Rows.Count; i++)
                    {
                        if (GetCell(xlRange, i, j).Value2 != null && GetCell(xlRange, i, j).Value2 != "")
                        //if (xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j].Value2 != "")
                        {
                            //transfer the station number into series number
                            start_range = xlRange_partial.Find(GetCell(xlRange, i, j + 1).Value2.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                            end_range = xlRange_partial.Find(GetCell(xlRange, i, j + 2).Value2.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);

                            start_station = start_range.Offset[0, -1].Value2.ToString();
                            end_station = end_range.Offset[0, -1].Value2.ToString();

                            string name = GetCell(xlRange, i, j).Value2.ToString();
                            string[] a = new string[] { GetCell(xlRange, i, j).Value2.ToString(), start_station, end_station };
                            //start_range = xlRange_partial.Find(xlRange.Cells[i, j + 1].Value2.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                            //end_range = xlRange_partial.Find(xlRange.Cells[i, j + 2].Value2.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);

                            //start_station = start_range.Offset[0, -1].Value2.ToString();
                            //end_station = end_range.Offset[0, -1].Value2.ToString();

                            //string name = xlRange.Cells[i, j].Value2.ToString();
                            //string[] a = new string[] { xlRange.Cells[i, j].Value2.ToString(), start_station, end_station };

                            if (name.Contains("仰拱"))
                            {
                                setting_Station.inverted_arc_station.Add(a);
                            }
                            else if (name.Contains("道床") && name.Contains("仰拱") == false)
                            {
                                setting_Station.track_bed_station.Add(a);
                            }
                            else if (name.Contains("排水溝"))
                            {
                                setting_Station.gutter_station.Add(a);
                            }
                            else if (name.Contains("側"))
                            {
                                if (GetCell(xlRange, 1, j).Value2.ToString() == "走道")
                                //if(xlRange.Cells[1, j].Value2.ToString() == "走道")
                                {
                                    setting_Station.walk_way_station.Add(a);
                                }
                                else if (GetCell(xlRange, 1, j).Value2.ToString() == "第三軌")
                                //else if(xlRange.Cells[1, j].Value2.ToString() == "第三軌")
                                {
                                    // 培文改
                                    start_station = GetCell(xlRange, i, j + 1).Value2.ToString();
                                    end_station = GetCell(xlRange, i, j + 2).Value2.ToString();
                                    //start_station = xlRange.Cells[i, j+1].Value2.ToString();
                                    //end_station = xlRange.Cells[i, j+2].Value2.ToString();
                                    int start = StationToInt(start_station);
                                    int end = StationToInt(end_station);
                                    a = new string[] { name, start.ToString(), end.ToString() };

                                    setting_Station.third_rail_station.Add(a);
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }


            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(xlApp);
        }

        public void read_target_station2()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["軌道線形 (UP)"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            Excel.Range xlRange_partial = (Excel.Range)xlRange.Columns[2];

            Excel.Range start_range = null;
            Excel.Range end_range = null;

            string start_station = "";
            string end_station = "";

            for (int j = 7; j <= xlRange.Columns.Count; j = j + 3)
            {
                if (GetCell(xlRange, 1, j).Value2 != null && GetCell(xlRange, 1, j).Value2 != "")
                //if (xlRange.Cells[1, j].Value2 != null && xlRange.Cells[1, j].Value2 != "")
                {
                    for (int i = 2; i <= xlRange.Rows.Count; i++)
                    {
                        if (GetCell(xlRange, i, j).Value2 != null && GetCell(xlRange, i, j).Value2 != "")
                        //if (xlRange.Cells[i, j].Value2 != null && xlRange.Cells[i, j].Value2 != "")
                        {
                            //transfer the station number into series number
                            start_range = xlRange_partial.Find(GetCell(xlRange, i, j + 1).Value2.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                            end_range = xlRange_partial.Find(GetCell(xlRange, i, j + 2).Value2.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                            start_station = start_range.Offset[0, -1].Value2.ToString();
                            end_station = end_range.Offset[0, -1].Value2.ToString();

                            string name = GetCell(xlRange, i, j).Value2.ToString();
                            string[] a = new string[] { GetCell(xlRange, i, j).Value2.ToString(), start_station, end_station };
                            //start_range = xlRange_partial.Find(xlRange.Cells[i, j + 1].Value2.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                            //end_range = xlRange_partial.Find(xlRange.Cells[i, j + 2].Value2.ToString(), Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                            //start_station = start_range.Offset[0, -1].Value2.ToString();
                            //end_station = end_range.Offset[0, -1].Value2.ToString();

                            //string name = xlRange.Cells[i, j].Value2.ToString();
                            //string[] a = new string[] { xlRange.Cells[i, j].Value2.ToString(), start_station, end_station };

                            if (name.Contains("仰拱"))
                            {
                                setting_Station2.inverted_arc_station.Add(a);
                            }
                            else if (name.Contains("道床") && name.Contains("仰拱") == false)
                            {
                                setting_Station2.track_bed_station.Add(a);
                            }
                            else if (name.Contains("排水溝"))
                            {
                                setting_Station2.gutter_station.Add(a);
                            }
                            else if (name.Contains("側"))
                            {
                                if (GetCell(xlRange, 1, j).Value2.ToString() == "走道")
                                //if (xlRange.Cells[1, j].Value2.ToString() == "走道")
                                {
                                    setting_Station2.walk_way_station.Add(a);
                                }
                                else if (GetCell(xlRange, 1, j).Value2.ToString() == "第三軌")
                                //else if (xlRange.Cells[1, j].Value2.ToString() == "第三軌")
                                {
                                    // 培文改
                                    start_station = GetCell(xlRange, i, j + 1).Value2.ToString();
                                    end_station = GetCell(xlRange, i, j + 2).Value2.ToString();
                                    //start_station = xlRange.Cells[i, j + 1].Value2.ToString();
                                    //end_station = xlRange.Cells[i, j + 2].Value2.ToString();
                                    int start = StationToInt(start_station);
                                    int end = StationToInt(end_station);
                                    a = new string[] { name, start.ToString(), end.ToString() };

                                    setting_Station2.third_rail_station.Add(a);
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }

            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(xlApp);
        }

        //讀道床資訊
        public track_bed_properties read_track_bed()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path + file_name);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets["模型輸入資料"];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            track_bed_properties setting = new track_bed_properties();

            for (int i = 2; i <= xlRange.Rows.Count; i++)
            {
                string name = "";
                try
                {
                    name = GetCell(xlRange, i, 1).Value2.ToString();
                    //name = xlRange.Cells[i, 1].Value2.ToString();
                }
                catch { continue; }
                switch (name)
                {
                    case "標準道床仰拱頂部高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.standard_top_height);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.standard_top_height);
                        break;
                    case "標準道床凹槽與軌道中心距離":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.standard_center_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.standard_center_dis);
                        //setting.standard_center_dis = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "標準道床凹槽水平寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.standard_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.standard_width);
                        //setting.standard_width = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "標準道床凹槽深度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.standard_depth);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.standard_depth);
                        //setting.standard_depth = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "標準道床高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.standard_elevation);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.standard_elevation);
                        //setting.standard_elevation = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "浮動式道床仰拱頂部高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.float_top_height);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.float_top_height);
                        //setting.float_top_height = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "浮動式道床高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.float_elevation);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.float_elevation);
                        //setting.float_elevation = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "浮動式道床寬度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.float_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.float_width);
                        //setting.float_width = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "浮動式道床厚度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.float_depth);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.float_depth);
                        //setting.float_depth = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "浮動式道床支承墊長":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.suppad_length);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.suppad_length);
                        //setting.suppad_length = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "浮動式道床支承墊寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.suppad_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.suppad_width);
                        //setting.suppad_width = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "浮動式道床支承墊厚":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.suppad_depth);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.suppad_depth);
                        //setting.suppad_depth = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "浮動式道床支承墊與邊緣距離":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.suppad_side_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.suppad_side_dis);
                        //setting.suppad_side_dis = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "平板式道床仰拱頂部高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.flat_top_height);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.flat_top_height);
                        //setting.flat_top_height = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "平板式道床高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.flat_elevation);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.flat_elevation);
                        //setting.flat_elevation = xlRange.Cells[i, 2].Value2.ToString();
                        break;
                    case "平板式道床寬度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.flat_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.flat_width);
                        //setting.flat_width = xlRange.Cells[i, 2].Value2.ToString();
                        break;

                    //第三軌&鋼軌
                    case "鋼軌軌距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.rail_gauge);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.rail_gauge);
                        break;
                    case "鋼軌面寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.rail_face_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.rail_face_width);
                        break;
                    case "鋼軌基鈑長":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.rail_base_length);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.rail_base_length);
                        break;
                    case "鋼軌基鈑寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.rail_base_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.rail_base_width);
                        break;
                    case "鋼軌基鈑厚":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.rail_base_thickness);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.rail_base_thickness);
                        break;
                    case "鋼軌基鈑傾斜度":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.rail_base_slope);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.rail_base_slope);
                        break;
                    case "鋼軌基鈑間距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.rail_base_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.rail_base_dis);
                        break;
                    case "第三軌與鋼軌距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.third_steel_between_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.third_steel_between_dis);
                        break;
                    case "第三軌高程":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.third_track_elevation);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.third_track_elevation);
                        break;
                    case "第三軌支架間距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.third_bracket_spacing);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.third_bracket_spacing);
                        break;
                    case "第三軌支架與鋼軌距":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.bracket_steel_between_dis);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.bracket_steel_between_dis);
                        break;
                    case "第三軌支架長":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.third_bracket_length);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.third_bracket_length);
                        break;
                    case "第三軌支架寬":
                        double.TryParse(GetCell(xlRange, i, 2).Value2.ToString(), out setting.third_bracket_width);
                        //double.TryParse(xlRange.Cells[i, 2].Value2.ToString(), out setting.third_bracket_width);
                        break;
                }

            }
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
            return setting;
        }

        public Tuple<int, int> FindAddress(string name, Excel.Range xlRange) //取得excel中的指定名稱cell
        {
            Excel.Range address;
            address = xlRange.Find(name, MatchCase: true);
            if (address == null)
            {
                return null;
            }
            else
            {
                var pos = Tuple.Create(address.Row, address.Column);
                return pos;
            }
        }

        public int StationToInt(String station)
        {
            int a = int.Parse(station.Split('+')[0]);
            int b = int.Parse(station.Split('+')[1].Split('.')[0]);
            return a * 1000 + b;
        }
        /// <summary>
        /// 培文改：.NET Core 8.0 需要使用dynamic來取得Excel的Cell值
        /// </summary>
        /// <param name="xlRange"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private dynamic GetCell(Excel.Range xlRange, int row, int col)
        {
            return xlRange.Cells[row, col];
        }
    }
}