using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using DataObject;

namespace SinoTunnel_2020
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class Counting : IExternalEventHandler
    {
        string path;

        public void Execute(UIApplication app)
        {
            try
            {

                path = Form1.path;
                Document document = app.ActiveUIDocument.Document;
                UIDocument uidoc = new UIDocument(document);
                Document doc = uidoc.Document;
                Transaction t = new Transaction(doc);
                readfile rf = new readfile();
                miles_data data = new miles_data();

                List<Tuple<string, string>> tunnel_endpts = rf.tunnel_endpts_info();
                List<Tuple<string, Tuple<string, string>>> miles_info = rf.miles_info();

                //開始數量計算

                ////PART ONE:鑽掘隧道里程計算
                foreach (Tuple<string, string> endpt in tunnel_endpts)
                {
                    string pt_s = endpt.Item1;
                    string pt_e = endpt.Item2;
                    string[] pt_s_trim = pt_s.Split(new char[1] { '+' });
                    string[] pt_e_trim = pt_e.Split(new char[1] { '+' });
                    double.TryParse(string.Concat(pt_s_trim).ToString(), out double val_pt_s);
                    double.TryParse(string.Concat(pt_e_trim).ToString(), out double val_pt_e);
                    double diff = val_pt_e - val_pt_s;
                    string concat = "+(" + val_pt_e.ToString() + "-" + val_pt_s.ToString() + ")";

                    data.sum_miles += diff;
                    if (data.sum_miles_cal == null)
                        data.sum_miles_cal += concat.Split(new char[1] { '+' }).Last();
                    else
                        data.sum_miles_cal += concat;
                }
                //里程扣除(假設20m&22m)
                /*
                data.sum_miles -= 20;
                data.sum_miles -= 22;
                data.sum_miles_cal += "-20";
                data.sum_miles_cal += "-22";
                */

                foreach (Tuple<string, Tuple<string, string>> name_pt_set in miles_info) //仰拱、道床和排水設施里程計算
                {

                    Tuple<string, string> pt = name_pt_set.Item2;
                    string pt_s = pt.Item1;
                    string pt_e = pt.Item2;
                    string[] pt_s_trim = pt_s.Split(new char[1] { '+' });
                    string[] pt_e_trim = pt_e.Split(new char[1] { '+' });
                    double.TryParse(string.Concat(pt_s_trim).ToString(), out double val_pt_s);
                    double.TryParse(string.Concat(pt_e_trim).ToString(), out double val_pt_e);
                    double diff = val_pt_e - val_pt_s;
                    string concat = "+(" + val_pt_e.ToString() + "-" + val_pt_s.ToString() + ")";

                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "浮動式道床仰拱")
                    {
                        data.inverted_floating += diff;
                        if (data.inverted_floating_cal == null)
                            data.inverted_floating_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.inverted_floating_cal += concat;
                    }

                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "標準道床仰拱")
                    {
                        data.inverted_standard += diff;
                        if (data.inverted_standard_cal == null)
                            data.inverted_standard_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.inverted_standard_cal += concat;
                    }

                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "平板式道床仰拱")
                    {
                        data.inverted_flat += diff;
                        if (data.inverted_flat_cal == null)
                            data.inverted_flat_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.inverted_flat_cal += concat;

                    }
                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "浮動式道床")
                    {
                        data.track_bed_floating += diff;
                        if (data.track_bed_floating_cal == null)
                            data.track_bed_floating_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.track_bed_floating_cal += concat;

                    }
                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "標準道床")
                    {
                        data.track_bed_standard += diff;
                        if (data.track_bed_standard_cal == null)
                            data.track_bed_standard_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.track_bed_standard_cal += concat;

                    }
                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "平板式道床")
                    {
                        data.track_bed_flat += diff;
                        if (data.track_bed_flat_cal == null)
                            data.track_bed_flat_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.track_bed_flat_cal += concat;

                    }
                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "浮動式排水溝")
                    {
                        data.drainage_floating += diff;
                        if (data.drainage_floating_cal == null)
                            data.drainage_floating_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.drainage_floating_cal += concat;

                    }
                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "PVC管排水溝")
                    {
                        data.drainage_standard += diff;
                        if (data.drainage_standard_cal == null)
                            data.drainage_standard_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.drainage_standard_cal += concat;

                    }
                    if (name_pt_set.Item1 != null && name_pt_set.Item1 == "U形排水溝")
                    {
                        data.drainage_flat += diff;
                        if (data.drainage_flat_cal == null)
                            data.drainage_flat_cal += concat.Split(new char[1] { '+' }).Last();
                        else
                            data.drainage_flat_cal += concat;

                    }
                }

                ////PART TWO:從專案檔中抓取仰拱、走道及排水幾何&非幾何資訊

                Family family = null;
                Document familyDoc = null;

                IList<FamilyInstance> famin_inverted_arcs = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance)).Where(x => x.Name.Contains("仰拱final_0")).Cast<FamilyInstance>().ToList();

                double stn_volume = double.NaN; //標準仰拱體積 m3
                double flt_volume = double.NaN; //浮動式仰拱體積 
                string str_stn_volume = string.Empty;
                string str_flt_volume = string.Empty;

                double stn_arc_temp = double.NaN; //標準仰拱模板面積 m2
                double flt_arc_temp = double.NaN; //浮動式仰拱模板面積 m2

                double stn_path_volume = double.NaN; //標準仰拱走道體積 m3 
                double flt_path_volume = double.NaN; //浮動式仰拱走道體積 m3
                string str_stn_path_vol = string.Empty;
                string str_flt_path_vol = string.Empty;

                double stn_path_temp = double.NaN; //標準仰拱走道模板面積 m2
                double flt_path_temp = double.NaN; //浮動式仰拱走道模板面積 m2
                string str_stn_path = string.Empty;
                string str_flt_path = string.Empty;

                double gutter_temp = double.NaN;  //標準型道床排水模板 m2
                string str_gutter_temp = string.Empty;

                double U_shape_steel_volume = double.NaN; //槽鋼體積 m3

                double stn_sidewalk_lid = 1 / 0.5 * 100;
                string str_stn_sidewalk_lid = "1 / 0.5 * 100";
                double flt_sidewalk_lid = 1 / 0.5 * 100;
                string str_flt_sidewalk_lid = "1 / 0.5 * 100";

                foreach (FamilyInstance famin_inverted_arc in famin_inverted_arcs)
                {
                    family = famin_inverted_arc.Symbol.Family;
                    familyDoc = document.EditFamily(family); //switch to family doc of inverted arcs
                }

                if (null != familyDoc && familyDoc.IsFamilyDocument == true)
                {
                    IList<SweptBlend> swept_arcs = new FilteredElementCollector(familyDoc).OfClass(typeof(SweptBlend)).Cast<SweptBlend>().ToList();
                    List<Tuple<SweptBlend, int>> stn_arcs = new List<Tuple<SweptBlend, int>>();
                    List<Tuple<SweptBlend, int>> flt_arcs = new List<Tuple<SweptBlend, int>>();
                    List<SweptBlend> sidewalks = new List<SweptBlend>();

                    int i = 0;
                    foreach (SweptBlend ele in swept_arcs)
                    {
                        Tuple<SweptBlend, int> ele_index = Tuple.Create(ele, i);
                        if (ele.TopProfileSymbol.Profile.Name.Contains("標準道床"))
                        {
                            stn_arcs.Add(ele_index);
                            i++;
                        }
                        else if (ele.TopProfileSymbol.Profile.Name.Contains("浮動式道床"))
                        {
                            flt_arcs.Add(ele_index);
                            i++;
                        }
                        else if (ele.TopProfileSymbol.Profile.Name.Contains("混凝土蓋板"))
                        {
                            sidewalks.Add(ele);
                            i++;
                        }
                    }

                    Options options = new Options();
                    List<Solid> solid_arcs = new List<Solid>();
                    foreach (FamilyInstance famin_inverted_arc in famin_inverted_arcs)
                    {
                        GeometryElement geoElement = famin_inverted_arc.get_Geometry(options);
                        foreach (GeometryObject geoObject in geoElement)
                        {
                            GeometryInstance instance = geoObject as GeometryInstance;
                            if (null != instance)
                            {
                                foreach (GeometryObject o in instance.SymbolGeometry)
                                {
                                    Solid solid = o as Solid;
                                    if (null == solid || solid.Volume < 0.8)
                                        continue;

                                    solid_arcs.Add(solid);
                                }
                            }
                        }
                    }
                    solid_arcs.Reverse(); //familyinstance中的geoelements 和 family中的geoelements one by one


                    ////標準型、浮動式道床仰拱體積抓取: m3

                    //標準排水溝:
                    FamilySymbol contour_gutter = new FilteredElementCollector(familyDoc).OfClass(typeof(FamilySymbol)).
                        Cast<FamilySymbol>().ToList().Where(x => x.Name == "標準排水溝").First();
                    double depth_u = contour_gutter.LookupParameter("U形排水溝深度").AsDouble() * 2;
                    double semi_u = contour_gutter.LookupParameter("U形排水溝半徑").AsDouble() * 3.14;
                    gutter_temp = (depth_u + semi_u) * 0.3048 * 100;
                    str_gutter_temp = "(" + (contour_gutter.LookupParameter("U形排水溝深度").AsDouble() * 0.3048).ToString() + "*2+" +
                        (contour_gutter.LookupParameter("U形排水溝半徑").AsDouble() * 0.3048 * 2).ToString() + "*3.14/2)*100";

                    double sta_gutter_area = (contour_gutter.LookupParameter("U形排水溝深度").AsDouble() - contour_gutter.LookupParameter("U形排水溝蓋板厚").AsDouble())
                        * contour_gutter.LookupParameter("U形排水溝半徑").AsDouble() * 2
                        + Math.Pow(contour_gutter.LookupParameter("U形排水溝半徑").AsDouble(), 2) * 3.14 / 2;

                    //PVC管仰拱:
                    FamilySymbol pvc_s = new FilteredElementCollector(familyDoc).OfClass(typeof(FamilySymbol)).
                        Cast<FamilySymbol>().ToList().Where(x => x.Name == "PVC管仰拱").First();
                    double pvc_s_area = Math.Pow(pvc_s.LookupParameter("PVC管仰拱半徑").AsDouble(), 2) * 3.14 / 2;

                    //PVC管排水溝:
                    FamilySymbol pvc_l = new FilteredElementCollector(familyDoc).OfClass(typeof(FamilySymbol)).
                        Cast<FamilySymbol>().ToList().Where(x => x.Name == "PVC管排水溝").First();
                    double pvc_l_area = Math.Pow(pvc_l.LookupParameter("PVC管半徑").AsDouble(), 2) * 3.14;

                    //平板式道床: 
                    FamilySymbol track_bed = new FilteredElementCollector(familyDoc).OfClass(typeof(FamilySymbol)).
                        Cast<FamilySymbol>().ToList().Where(x => x.Name == "平版式道床").First();
                    double bed_groove = track_bed.LookupParameter("道床凹槽深度").AsDouble() *
                        (track_bed.LookupParameter("寬度").AsDouble() - track_bed.LookupParameter("道床與軌道中心距離").AsDouble() * 2);

                    //浮動式排水溝:
                    FamilySymbol flt_gutter = new FilteredElementCollector(familyDoc).OfClass(typeof(FamilySymbol)).
                        Cast<FamilySymbol>().ToList().Where(x => x.Name == "浮動式排水溝").First();
                    double flt_gutter_area = Math.Pow(flt_gutter.LookupParameter("浮動式道床明溝半徑").AsDouble(), 2) * 3.14 / 2
                        + flt_gutter.LookupParameter("浮動式道床明溝半徑").AsDouble() * 2 * flt_gutter.LookupParameter("浮動式道床明溝深").AsDouble();


                    //標準道床仰拱_正偏移量:
                    FamilySymbol sta_contour = new FilteredElementCollector(familyDoc).OfClass(typeof(FamilySymbol)).
                        Cast<FamilySymbol>().ToList().Where(x => x.Name == "標準道床仰拱_正偏移量").First();
                    double sta_peri_s = sta_contour.LookupParameter("隧道內半徑").AsDouble() - sta_contour.LookupParameter("連結處厚度").AsDouble();
                    double sta_peri_l = sta_contour.LookupParameter("隧道內半徑").AsDouble();
                    double sta_contour_h = sta_contour.LookupParameter("隧道中心").AsDouble() + sta_contour.LookupParameter("仰拱頂部高程").AsDouble();
                    double sta_ang_a = Math.Acos(sta_contour_h / sta_peri_s);
                    double sta_ang_b = Math.Acos(sta_contour_h / sta_peri_l);

                    double sta_tri_area = 0.5 * sta_peri_s * sta_contour_h * Math.Sin(sta_ang_a) + 0.5 * sta_peri_l * sta_contour_h * Math.Sin(sta_ang_b);
                    double sta_arc_area = Math.Pow(sta_peri_l, 2) * 3.14 * (sta_ang_a + sta_ang_b) / 6.28 - sta_tri_area;


                    //浮動式道床仰拱_正偏移量:
                    FamilySymbol flt_contour = new FilteredElementCollector(familyDoc).OfClass(typeof(FamilySymbol)).
                        Cast<FamilySymbol>().ToList().Where(x => x.Name == "浮動式道床仰拱_正偏移量").First();

                    double flt_peri_s = flt_contour.LookupParameter("隧道內半徑").AsDouble() - flt_contour.LookupParameter("連結處厚度").AsDouble();
                    double flt_peri_l = flt_contour.LookupParameter("隧道內半徑").AsDouble();
                    double flt_contour_h = flt_contour.LookupParameter("隧道中心").AsDouble() + flt_contour.LookupParameter("仰拱頂部高程").AsDouble();
                    double flt_ang_a = Math.Acos(flt_contour_h / flt_peri_s);
                    double flt_ang_b = Math.Acos(flt_contour_h / flt_peri_l);

                    double flt_tri_area = 0.5 * flt_peri_s * flt_contour_h * Math.Sin(flt_ang_a) + 0.5 * flt_peri_l * flt_contour_h * Math.Sin(flt_ang_b);
                    double flt_arc_area = Math.Pow(flt_peri_l, 2) * 3.14 * (flt_ang_a + flt_ang_b) / 6.28 - flt_tri_area;


                    //////標準&浮動式走道模板抓取
                    //string len_add_detail = "";

                    List<List<Tuple<SweptBlend, int>>> arcs_sort = new List<List<Tuple<SweptBlend, int>>>();

                    if (stn_arcs.Count != 0)
                    {
                        arcs_sort.Add(stn_arcs);
                    }

                    if (flt_arcs.Count != 0)
                    {
                        arcs_sort.Add(flt_arcs);
                    }

                    double len_add = 0;
                    string str_len_add = string.Empty;

                    List<double> arc_temps = new List<double>();
                    List<Tuple<double, string>> path_temp = new List<Tuple<double, string>>();
                    List<Tuple<double, string>> path_vol = new List<Tuple<double, string>>();

                    foreach (List<Tuple<SweptBlend, int>> arcs in arcs_sort) //取得仰拱截面
                    {
                        Face max_face = null;
                        double max_area = 0;

                        foreach (Face f in solid_arcs[arcs.First().Item2].Faces)
                        {
                            PlanarFace planar_f = f as PlanarFace;

                            if (null == planar_f || planar_f.Area == 0)
                                continue;

                            else if (max_face == null)
                            {
                                max_area = planar_f.Area;
                                max_face = planar_f;
                                continue;
                            }

                            if (planar_f.Area > max_area && Math.Round(planar_f.FaceNormal.Z) == 0)
                            {
                                max_area = planar_f.Area;
                                max_face = planar_f;
                            }
                        }
                        CurveLoop eArray = max_face.GetEdgesAsCurveLoops().ElementAt(0);
                        //TaskDialog.Show("Revit", max_area.ToString());


                        Arc max_arc = null;
                        double max_arc_len = 0;

                        if (eArray != null) //取得最大弧線arc
                        {
                            foreach (Curve curve in eArray)
                            {
                                Arc arc = curve as Arc;
                                if (arc == null || arc.ApproximateLength == 0)
                                    continue;
                                else
                                {
                                    if (arc.ApproximateLength > max_arc_len)
                                    {
                                        max_arc_len = arc.ApproximateLength;
                                        max_arc = arc;
                                    }
                                }
                            }
                        }
                        //TaskDialog.Show("Revit", max_arc_len.ToString());

                        double up_bound = 0;
                        double dn_bound = 0;
                        if (max_arc.GetEndPoint(1).Z > max_arc.GetEndPoint(0).Z)
                        {
                            up_bound = max_arc.GetEndPoint(1).Z;
                            dn_bound = max_arc.GetEndPoint(0).Z;
                        }
                        else
                        {
                            up_bound = max_arc.GetEndPoint(0).Z;
                            dn_bound = max_arc.GetEndPoint(1).Z;
                        }

                        double fix_bound = (up_bound - dn_bound) * 0.03; //Shrink the range of bbox to collect the contour of template
                        dn_bound = dn_bound + fix_bound;

                        //TaskDialog.Show("Revit", max_arc.GetEndPoint(1).Z.ToString() +  max_arc.GetEndPoint(0).Z.ToString());

                        if (eArray != null) //抓取模板輪廓並求出長度
                        {
                            double arc_temp = 0;
                            foreach (Curve curve in eArray)
                            {
                                Arc c = curve as Arc;
                                if (c == null || c.ApproximateLength == 0)
                                    continue;

                                else
                                {
                                    XYZ origin = c.GetEndPoint(0);
                                    XYZ end = c.GetEndPoint(1);

                                    if ((origin.Z < up_bound && origin.Z > dn_bound) || (end.Z > dn_bound && end.Z < up_bound)) //to collection the lines that intersect or inside the bbox
                                    {
                                        if (Math.Round(c.Length) != Math.Round(max_arc.Length))
                                        {
                                            //len_add_detail += Math.Round(c.ApproximateLength * 0.3048, 3).ToString() + origin.ToString() + end.ToString() + "\n";
                                            len_add += c.ApproximateLength;
                                            if (str_len_add == string.Empty)
                                                str_len_add += (Math.Round(c.ApproximateLength * 0.3048, 3).ToString());
                                            else
                                                str_len_add += ("+" + Math.Round(c.ApproximateLength * 0.3048, 3).ToString());

                                            arc_temp = (max_arc.Radius - c.Radius) * 0.3048; //得到仰拱模板面積
                                        }
                                    }
                                }

                            }
                            arc_temps.Add(arc_temp);


                            foreach (Curve curve in eArray)
                            {
                                Line l = curve as Line;
                                if (l == null || l.ApproximateLength == 0)
                                    continue;

                                else
                                {
                                    XYZ origin = l.GetEndPoint(0);
                                    XYZ end = l.GetEndPoint(1);

                                    if ((origin.Z < up_bound && origin.Z > dn_bound) || (end.Z > dn_bound && end.Z < up_bound)) //to collection the lines that intersect or inside the bbox
                                    {
                                        //len_add_detail += Math.Round(l.ApproximateLength * 0.3048, 3).ToString() + origin.ToString() + end.ToString() + "\n";
                                        len_add += l.ApproximateLength;
                                        if (str_len_add == string.Empty)
                                            str_len_add += (Math.Round(l.ApproximateLength * 0.3048, 3).ToString());
                                        else
                                            str_len_add += ("+" + Math.Round(l.ApproximateLength * 0.3048, 3).ToString());
                                    }
                                }
                            }
                        }
                        path_temp.Add(Tuple.Create(len_add * 0.3048 * 100, "(" + str_len_add + ")*100"));
                        len_add = 0;
                        str_len_add = string.Empty;
                    }

                    if (stn_arcs.Count != 0)
                    {
                        stn_volume = Math.Round((sta_arc_area - bed_groove - sta_gutter_area) * 100 * 0.3048 * 0.3048, 1); // 標準型 & 平板型道床仰拱體積:
                        str_stn_volume = (stn_volume / 100).ToString() + "*100";
                        //double stn_volume_pvc = Math.Round((sta_arc_area - bed_groove - pvc_s_area - pvc_l_area) * 100 * 0.3048 * 0.3048, 1); //PVC排水溝
                        stn_path_volume = Math.Round((solid_arcs[stn_arcs.First().Item2].Volume * Math.Pow(0.3048, 3) * 100) - stn_volume, 1); //標準仰拱走道體積 m3
                        str_stn_path_vol = "(" + Math.Round(solid_arcs[stn_arcs.First().Item2].Volume * Math.Pow(0.3048, 3), 3).ToString() + "-" + (stn_volume / 100).ToString() + ")*100";
                        stn_arc_temp = Math.Round(arc_temps[0], 3); //標準仰拱模板面積 m2
                        stn_path_temp = Math.Round(path_temp[0].Item1, 1); //標準走道模板面積 m2
                        str_stn_path = path_temp[0].Item2;
                    }
                    else
                    {
                        stn_volume = 0;
                        stn_path_volume = 0;
                        stn_arc_temp = 0;
                        stn_path_temp = 0;
                        stn_sidewalk_lid = 0;
                        str_stn_sidewalk_lid = "";
                    }

                    if (flt_arcs.Count != 0)
                    {
                        flt_volume = Math.Round((flt_arc_area - flt_gutter_area) * 100 * 0.3048 * 0.3048, 1); //浮動型道床仰拱體積:
                        str_flt_volume = (flt_volume / 100).ToString() + "*100";
                        flt_path_volume = Math.Round((solid_arcs[flt_arcs.First().Item2].Volume * Math.Pow(0.3048, 3) * 100) - flt_volume, 1); //浮動式仰拱走道體積 m3
                        str_flt_path_vol = "(" + Math.Round(solid_arcs[flt_arcs.First().Item2].Volume * Math.Pow(0.3048, 3), 3).ToString() + "-" + (flt_volume / 100).ToString() + ")*100";
                        flt_arc_temp = Math.Round(arc_temps[1], 3); //浮動式仰拱模板面積 m2
                        flt_path_temp = Math.Round(path_temp[1].Item1, 1); //浮動式走道模板面積 m2
                        str_flt_path = path_temp[1].Item2;
                    }
                    else
                    {
                        flt_volume = 0;
                        flt_path_volume = 0;
                        flt_arc_temp = 0;
                        flt_path_temp = 0;
                        flt_sidewalk_lid = 0;
                        str_flt_sidewalk_lid = "";
                    }

                    //TaskDialog.Show("Revit", len_add_detail);
                }



                ////PART THREE:槽鋼體積寫入

                FamilyInstance U_shape_steel = new FilteredElementCollector(doc).
                    OfClass(typeof(FamilyInstance)).Where(x => x.Name == "U型槽鋼").Cast<FamilyInstance>().ToList().First();
                Family fa_U_shape_steel = U_shape_steel.Symbol.Family;
                Document fa_doc_U = document.EditFamily(fa_U_shape_steel); //switch to family doc of inverted arcs

                if (null != fa_doc_U && fa_doc_U.IsFamilyDocument == true)
                {
                    FamilyInstance U_shape_steel_fsb = new FilteredElementCollector(fa_doc_U).OfClass(typeof(FamilyInstance)).
                        Cast<FamilyInstance>().ToList().Where(x => x.Name.Contains("U型槽鋼")).First();

                    U_shape_steel_volume = Math.Round(U_shape_steel_fsb.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED)
                        .AsDouble() * Math.Pow(0.3048, 3), 6);
                }


                ////PART FOUR:隧道襯砌環片，隧道預鑄混凝土襯砌環片

                double ada_A_vol_diff = double.NaN;
                double ada_B_vol_diff = double.NaN;
                double ada_K_vol_diff = double.NaN;
                double ada_A_area_diff = double.NaN;
                double ada_B_area_diff = double.NaN;
                double ada_K_area_diff = double.NaN;
                string ada_A_vol_diff_str = "";
                string ada_B_vol_diff_str = "";
                string ada_K_vol_diff_str = "";
                string ada_A_area_diff_str = "";
                string ada_B_area_diff_str = "";
                string ada_K_area_diff_str = "";

                //A.自適應環片
                FamilyInstance adaptive_a = new FilteredElementCollector(doc).
                    OfClass(typeof(FamilyInstance)).Where(x => x.Name == "自適應環形00_A").Cast<FamilyInstance>().ToList().First();

                Family fa_adaptive_a = adaptive_a.Symbol.Family;
                Document fa_doc_adaptive_a = document.EditFamily(fa_adaptive_a); //switch to family doc of adaptive_a

                double ada_A_vol = double.NaN;
                double ada_B_vol = double.NaN;
                double ada_K_vol = double.NaN;
                double ada_A_area = double.NaN;
                double ada_B_area = double.NaN;
                double ada_K_area = double.NaN;

                if (null != fa_doc_adaptive_a && fa_doc_adaptive_a.IsFamilyDocument == true)
                {
                    IList<FamilyInstance> adaptive_a_fsb = new FilteredElementCollector(fa_doc_adaptive_a).
                    OfClass(typeof(FamilyInstance)).Where(x => x.Name.Contains("segment")).Cast<FamilyInstance>().ToList();

                    IList<Tuple<FamilyInstance, double, double>> adaptive_a_vol_area = new List<Tuple<FamilyInstance, double, double>>();

                    foreach (FamilyInstance ada_a in adaptive_a_fsb)
                    {
                        double ada_a_volume = Math.Round(ada_a.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble() * Math.Pow(0.3048, 3), 4);
                        double ada_a_area = Math.Round(ada_a.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble() * Math.Pow(0.3048, 2), 4);
                        adaptive_a_vol_area.Add(Tuple.Create(ada_a, ada_a_volume, ada_a_area));
                    }
                    adaptive_a_vol_area = adaptive_a_vol_area.OrderBy(x => x.Item2).ToList();

                    ada_K_vol = adaptive_a_vol_area.ElementAt(0).Item2;
                    ada_B_vol = adaptive_a_vol_area.ElementAt(2).Item2;
                    ada_A_vol = adaptive_a_vol_area.ElementAt(5).Item2;
                    ada_K_area = adaptive_a_vol_area.ElementAt(0).Item3;
                    ada_B_area = adaptive_a_vol_area.ElementAt(2).Item3;
                    ada_A_area = adaptive_a_vol_area.ElementAt(5).Item3;

                }


                //B.Bolt

                double bolt_vol = 0;
                double KLbolt_vol = 0;
                double KSbolt_vol = 0;
                double bolt_surface = 0;
                string bolt_surface_str = "";
                double KLbolt_surface = 0;
                string KLbolt_surface_str = "";
                double KSbolt_surface = 0;
                string KSbolt_surface_str = "";

                try   //if bolts are exist 
                {
                    FamilyInstance bolt_a = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Where(x => x.Name == "螺栓_A").Cast<FamilyInstance>().ToList().First();

                    Family fa_bolt_a = bolt_a.Symbol.Family;
                    Document fa_doc_bolt_a = document.EditFamily(fa_bolt_a); //switch to family doc of bolt_a

                    if (null != fa_doc_bolt_a && fa_doc_bolt_a.IsFamilyDocument == true)
                    {
                        FamilyInstance bolt_a_fsb = new FilteredElementCollector(fa_doc_bolt_a).
                        OfClass(typeof(FamilyInstance)).Where(x => x.Name.Contains("螺栓雛形")).Cast<FamilyInstance>().ToList().First();

                        Options options = new Options();
                        GeometryElement geoElement = bolt_a_fsb.get_Geometry(options);
                        foreach (GeometryObject geoObject in geoElement)
                        {
                            Solid solid = geoObject as Solid;
                            if (null == solid || solid.Volume == 0)
                                continue;

                            bolt_vol = Math.Round(solid.Volume * Math.Pow(0.3048, 3), 4);

                            double max_area = double.NaN;
                            Face max_face = null;

                            foreach (Face f in solid.Faces)
                            {
                                if (max_face == null)
                                {
                                    max_area = f.Area;
                                    max_face = f;
                                    continue;
                                }

                                if (f.Area > max_area)
                                {
                                    max_area = f.Area;
                                    max_face = f;
                                }
                            }
                            foreach (Face f in solid.Faces)
                            {
                                if (f == max_face)
                                    continue;

                                else
                                {
                                    bolt_surface = bolt_surface + f.Area;
                                    bolt_surface_str = bolt_surface_str + Math.Round(f.Area * Math.Pow(0.3048, 2), 4) + "+";
                                }
                            }
                            bolt_surface = bolt_surface - max_area;
                            bolt_surface = Math.Round(bolt_surface * Math.Pow(0.3048, 2), 4);
                            bolt_surface_str = bolt_surface_str + "(-" + Math.Round(max_area * Math.Pow(0.3048, 2), 4) + ")";
                        }

                        FamilyInstance KLbolt_a_fsb = new FilteredElementCollector(fa_doc_bolt_a).
                            OfClass(typeof(FamilyInstance)).Where(x => x.Name.Contains("K螺栓大凹槽")).Cast<FamilyInstance>().ToList().First();

                        geoElement = KLbolt_a_fsb.get_Geometry(options);
                        foreach (GeometryObject geoObject in geoElement)
                        {
                            Solid solid = geoObject as Solid;
                            if (null == solid || solid.Volume == 0)
                                continue;

                            KLbolt_vol = Math.Round(solid.Volume * Math.Pow(0.3048, 3), 4);

                            double max_area = double.NaN;
                            Face max_face = null;

                            foreach (Face f in solid.Faces)
                            {
                                if (max_face == null)
                                {
                                    max_area = f.Area;
                                    max_face = f;
                                    continue;
                                }

                                if (f.Area > max_area)
                                {
                                    max_area = f.Area;
                                    max_face = f;
                                }
                            }
                            foreach (Face f in solid.Faces)
                            {
                                if (f == max_face)
                                    continue;
                                else
                                {
                                    KLbolt_surface = KLbolt_surface + f.Area;
                                    KLbolt_surface_str = KLbolt_surface_str + Math.Round(f.Area * Math.Pow(0.3048, 2), 4) + "+";
                                }
                            }
                            KLbolt_surface = KLbolt_surface - max_area;
                            KLbolt_surface = Math.Round(KLbolt_surface * Math.Pow(0.3048, 2), 4);
                            KLbolt_surface_str = KLbolt_surface_str + "(-" + Math.Round(max_area * Math.Pow(0.3048, 2), 4) + ")";
                        }

                        FamilyInstance KSbolt_a_fsb = new FilteredElementCollector(fa_doc_bolt_a).
                            OfClass(typeof(FamilyInstance)).Where(x => x.Name.Contains("K螺栓小凹槽")).Cast<FamilyInstance>().ToList().First();

                        geoElement = KSbolt_a_fsb.get_Geometry(options);
                        foreach (GeometryObject geoObject in geoElement)
                        {
                            Solid solid = geoObject as Solid;
                            if (null == solid || solid.Volume == 0)
                                continue;

                            KSbolt_vol = Math.Round(solid.Volume * Math.Pow(0.3048, 3), 4);

                            double max_area = double.NaN;
                            Face max_face = null;

                            foreach (Face f in solid.Faces)
                            {
                                if (max_face == null)
                                {
                                    max_area = f.Area;
                                    max_face = f;
                                    continue;
                                }

                                if (f.Area > max_area)
                                {
                                    max_area = f.Area;
                                    max_face = f;
                                }
                            }
                            foreach (Face f in solid.Faces)
                            {
                                if (f == max_face)
                                    continue;
                                else
                                {
                                    KSbolt_surface = KSbolt_surface + f.Area;
                                    KSbolt_surface_str = KSbolt_surface_str + Math.Round(f.Area * Math.Pow(0.3048, 2), 4) + "+";
                                }
                            }
                            KSbolt_surface = KSbolt_surface - max_area;
                            KSbolt_surface = Math.Round(KSbolt_surface * Math.Pow(0.3048, 2), 4);
                            KSbolt_surface_str = KSbolt_surface_str + "(-" + Math.Round(max_area * Math.Pow(0.3048, 2), 4) + ")";

                        }
                    }

                    ada_A_vol_diff = Math.Round((ada_A_vol + bolt_vol * 8) * 3, 2);
                    ada_B_vol_diff = Math.Round((ada_B_vol + bolt_vol * 8) * 2, 2);
                    ada_K_vol_diff = Math.Round((ada_K_vol + KLbolt_vol + KSbolt_vol) * 1, 2);
                    ada_A_vol_diff_str = "(" + ada_A_vol.ToString() + bolt_vol.ToString() + "*8)*3";
                    ada_B_vol_diff_str = "(" + ada_B_vol.ToString() + bolt_vol.ToString() + "*8)*2";
                    ada_K_vol_diff_str = "(" + ada_K_vol.ToString() + KLbolt_vol.ToString() + KSbolt_vol.ToString() + ")*1";

                    ada_A_area_diff = Math.Round((ada_A_area + bolt_surface * 8) * 3, 1) * 2;
                    ada_B_area_diff = Math.Round((ada_B_area + bolt_surface * 8) * 2, 1) * 2;
                    ada_K_area_diff = Math.Round(ada_K_area + KLbolt_surface + KSbolt_surface, 1) * 2;
                    ada_A_area_diff_str = "(" + ada_A_area.ToString() + "+(" + bolt_surface_str + ")*8)*2*3";
                    ada_B_area_diff_str = "(" + ada_B_area.ToString() + "+(" + bolt_surface_str + ")*8)*2*2";
                    ada_K_area_diff_str = "(" + ada_K_area.ToString() + "+(" + KLbolt_surface_str + "+" + KSbolt_surface_str + "))*2*1";

                }
                catch //bolts are not built
                {
                    //TaskDialog.Show("test", "bolts are not built");

                    ada_A_vol_diff = Math.Round((ada_A_vol) * 3, 2);
                    ada_B_vol_diff = Math.Round((ada_B_vol) * 2, 2);
                    ada_K_vol_diff = Math.Round((ada_K_vol) * 1, 2);
                    ada_A_vol_diff_str = ada_A_vol.ToString() + "*3";
                    ada_B_vol_diff_str = ada_B_vol.ToString() + "*2";
                    ada_K_vol_diff_str = ada_K_vol.ToString() + "*1";

                    ada_A_area_diff = Math.Round((ada_A_area) * 3, 1) * 2;
                    ada_B_area_diff = Math.Round((ada_B_area) * 2, 1) * 2;
                    ada_K_area_diff = Math.Round(ada_K_area, 1) * 2;
                    ada_A_area_diff_str = ada_A_area.ToString() + "*2*3";
                    ada_B_area_diff_str = ada_B_area.ToString() + "*2*2";
                    ada_K_area_diff_str = ada_K_area.ToString() + "*2*1";
                }


                
                ////PART FIVE:環片鋼筋計算
                IList<Rebar> rebars_A = new FilteredElementCollector(doc).
                    OfClass(typeof(Rebar)).Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() == "A").Cast<Rebar>().ToList();

                double A_29M = 0;
                string A_29M_str = string.Empty;
                List<double> A_29M_list = new List<double>();

                double A_25M = 0;
                string A_25M_str = string.Empty;
                List<double> A_25M_list = new List<double>();

                double A_19M = 0;
                string A_19M_str = string.Empty;
                List<double> A_19M_list = new List<double>();

                double A_16M = 0;
                string A_16M_str = string.Empty;
                List<double> A_16M_list = new List<double>();

                double A_13M = 0;
                string A_13M_str = string.Empty;
                List<double> A_13M_list = new List<double>();

                foreach (Rebar rebar_A in rebars_A)
                {
                    if (rebar_A.Name.Contains("29M"))
                    {
                        A_29M += Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        A_29M_list.Add(Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_A.Name.Contains("25M"))
                    {
                        A_25M += Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        A_25M_list.Add(Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_A.Name.Contains("19M"))
                    {
                        A_19M += Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        A_19M_list.Add(Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_A.Name.Contains("16M"))
                    {
                        A_16M += Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        A_16M_list.Add(Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_A.Name.Contains("13M"))
                    {
                        A_13M += Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        A_13M_list.Add(Math.Round(rebar_A.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }
                }

                foreach (var grp in A_29M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    A_29M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in A_25M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    A_25M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in A_19M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    A_19M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in A_16M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    A_16M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in A_13M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    A_13M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                A_29M_str = A_29M_str.Substring(0, A_29M_str.Length - 1);
                A_25M_str = A_25M_str.Substring(0, A_25M_str.Length - 1);
                A_19M_str = A_19M_str.Substring(0, A_19M_str.Length - 1);
                A_16M_str = A_16M_str.Substring(0, A_16M_str.Length - 1);
                A_13M_str = A_13M_str.Substring(0, A_13M_str.Length - 1);


                IList<Rebar> rebars_B = new FilteredElementCollector(doc).
                    OfClass(typeof(Rebar)).Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() == "B").Cast<Rebar>().ToList();

                double B_29M = 0;
                string B_29M_str = string.Empty;
                List<double> B_29M_list = new List<double>();

                double B_25M = 0;
                string B_25M_str = string.Empty;
                List<double> B_25M_list = new List<double>();

                double B_19M = 0;
                string B_19M_str = string.Empty;
                List<double> B_19M_list = new List<double>();

                double B_16M = 0;
                string B_16M_str = string.Empty;
                List<double> B_16M_list = new List<double>();

                double B_13M = 0;
                string B_13M_str = string.Empty;
                List<double> B_13M_list = new List<double>();

                foreach (Rebar rebar_B in rebars_B)
                {
                    if (rebar_B.Name.Contains("29M"))
                    {
                        B_29M += Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        B_29M_list.Add(Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_B.Name.Contains("25M"))
                    {
                        B_25M += Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        B_25M_list.Add(Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_B.Name.Contains("19M"))
                    {
                        B_19M += Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        B_19M_list.Add(Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_B.Name.Contains("16M"))
                    {
                        B_16M += Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        B_16M_list.Add(Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_B.Name.Contains("13M"))
                    {
                        B_13M += Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        B_13M_list.Add(Math.Round(rebar_B.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }
                }

                foreach (var grp in B_29M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    B_29M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in B_25M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    B_25M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in B_19M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    B_19M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in B_16M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    B_16M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in B_13M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    B_13M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                B_29M_str = B_29M_str.Substring(0, B_29M_str.Length - 1);
                B_25M_str = B_25M_str.Substring(0, B_25M_str.Length - 1);
                B_19M_str = B_19M_str.Substring(0, B_19M_str.Length - 1);
                B_16M_str = B_16M_str.Substring(0, B_16M_str.Length - 1);
                B_13M_str = B_13M_str.Substring(0, B_13M_str.Length - 1);


                IList<Rebar> rebars_K = new FilteredElementCollector(doc).
                    OfClass(typeof(Rebar)).Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() == "K").Cast<Rebar>().ToList();

                double K_29M = 0;
                string K_29M_str = string.Empty;
                List<double> K_29M_list = new List<double>();

                double K_25M = 0;
                string K_25M_str = string.Empty;
                List<double> K_25M_list = new List<double>();

                double K_19M = 0;
                string K_19M_str = string.Empty;
                List<double> K_19M_list = new List<double>();

                double K_16M = 0;
                string K_16M_str = string.Empty;
                List<double> K_16M_list = new List<double>();

                double K_13M = 0;
                string K_13M_str = string.Empty;
                List<double> K_13M_list = new List<double>();

                foreach (Rebar rebar_K in rebars_K)
                {
                    if (rebar_K.Name.Contains("29M"))
                    {
                        K_29M += Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        K_29M_list.Add(Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_K.Name.Contains("25M"))
                    {
                        K_25M += Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        K_25M_list.Add(Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_K.Name.Contains("19M"))
                    {
                        K_19M += Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        K_19M_list.Add(Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_K.Name.Contains("16M"))
                    {
                        K_16M += Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        K_16M_list.Add(Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }

                    if (rebar_K.Name.Contains("13M"))
                    {
                        K_13M += Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3);
                        K_13M_list.Add(Math.Round(rebar_K.LookupParameter("鋼筋長度").AsDouble() * 0.3048, 3));
                    }
                }

                foreach (var grp in K_29M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    K_29M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in K_25M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    K_25M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in K_19M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    K_19M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in K_16M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    K_16M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                foreach (var grp in K_13M_list.OrderBy(x => x).ToList().GroupBy(i => i))
                    K_13M_str += grp.Key.ToString() + "*" + grp.Count().ToString() + "+";

                K_29M_str = K_29M_str.Substring(0, K_29M_str.Length - 1);
                K_25M_str = K_25M_str.Substring(0, K_25M_str.Length - 1);
                K_19M_str = K_19M_str.Substring(0, K_19M_str.Length - 1);
                K_16M_str = K_16M_str.Substring(0, K_16M_str.Length - 1);
                K_13M_str = K_13M_str.Substring(0, K_13M_str.Length - 1);
                

                ////PART SIX: 聯絡通道01,02...資料寫入
                IList<FamilyInstance> contact_channels = new FilteredElementCollector(doc).
                    OfClass(typeof(FamilyInstance)).Where(x => x.Name.Contains("聯絡通道")).Cast<FamilyInstance>().ToList();

                List<Tuple<string, double, double, double, int>>
                    contact_tunnel_profile = new List<Tuple<string, double, double, double, int>>();

                List<double> diff_vol_channel = new List<double>(); //體積(扣除內部空間)
                List<int> door_counts = new List<int>(); //防火門數量計算 
                List<double> total_vol_channel = new List<double>(); //體積(加上內部空間)
                List<double> internal_area_channel = new List<double>(); //內部表面積

                Family cnt_channel_fa = null;
                Document cnt_channel_doc = null;

                foreach (FamilyInstance contact_channel in contact_channels)
                {
                    if (contact_channel != null)
                    {
                        //體積(扣除內部空間)
                        double volume = Math.Round(contact_channel.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED)
                        .AsDouble() * Math.Pow(0.3048, 3), 1);
                        diff_vol_channel.Add(volume);

                        cnt_channel_fa = contact_channel.Symbol.Family;
                        cnt_channel_doc = document.EditFamily(cnt_channel_fa);

                        //防火門數量計算 
                        if (null != cnt_channel_doc && cnt_channel_doc.IsFamilyDocument == true)
                        {
                            IList<Sweep> sweeps = new FilteredElementCollector(cnt_channel_doc)
                                .OfClass(typeof(Sweep)).Cast<Sweep>().ToList();
                            int i = 0;
                            foreach (Sweep ele in sweeps)
                            {
                                if (ele.ProfileSymbol.Profile.Name.Contains("防火門框"))
                                    i++;
                            }
                            door_counts.Add(i);
                        }

                        //體積(含內部空間)&內部表面積計算 (a+b)
                        double total_volume = 0;
                        double internal_area = 0;

                        Options options = new Options();
                        List<Solid> solid_sweep = new List<Solid>();

                        GeometryElement geoElement = contact_channel.get_Geometry(options);
                        foreach (GeometryObject geoObject in geoElement)
                        {
                            GeometryInstance instance = geoObject as GeometryInstance;
                            if (null != instance)
                            {
                                foreach (GeometryObject o in instance.SymbolGeometry)
                                {
                                    Solid solid = o as Solid;
                                    if (null == solid || solid.Volume == 0)
                                        continue;
                                    solid_sweep.Add(solid);
                                }
                            }
                        }
                        solid_sweep = solid_sweep.OrderBy(x => x.SurfaceArea).ToList();

                        Solid sweep_body = solid_sweep.Last();
                        Solid sweep_head = solid_sweep.ElementAt(solid_sweep.Count - 2);

                        // a.平台體積(含內部空間)&內部表面積計算
                        if (sweep_body != null)
                        {
                            FaceArray facearray = sweep_body.Faces;
                            if (facearray != null)
                            {
                                foreach (GeometryObject f in facearray)
                                {
                                    PlanarFace fp = f as PlanarFace;
                                    if (null == fp || fp.Area == 0)
                                        continue;

                                    if (fp.EdgeLoops.Size > 1) //有內外雙層loop
                                    {
                                        List<Tuple<EdgeArray, double, double>> edgearray_peri_area = new List<Tuple<EdgeArray, double, double>>();

                                        foreach (EdgeArray edgearray in fp.EdgeLoops)
                                        {
                                            double peri = 0;
                                            List<Line> edgearray_ln = new List<Line>();

                                            foreach (Edge e in edgearray)
                                            {
                                                peri += e.ApproximateLength; //周長(for表面積計算)

                                                Line ln = e.AsCurve() as Line;
                                                if (null == ln)
                                                    continue;
                                                edgearray_ln.Add(ln);
                                            }
                                            edgearray_ln = edgearray_ln.OrderBy(x => x.ApproximateLength).ToList();
                                            edgearray_ln.RemoveAt(0);
                                            double width = edgearray_ln.ElementAt(0).Length;
                                            double length = edgearray_ln.ElementAt(1).Length;
                                            double area = width * length + length * length * 3.14 / 8; //面積(for體積計算)

                                            edgearray_peri_area.Add(Tuple.Create(edgearray, peri, area));
                                        }
                                        edgearray_peri_area = edgearray_peri_area.OrderBy(x => x.Item2).ToList();

                                        double area_diff = edgearray_peri_area[1].Item3 - edgearray_peri_area[0].Item3;
                                        double body_volome = Math.Round(sweep_body.Volume / area_diff * edgearray_peri_area[1].Item3 * Math.Pow(0.3048, 3), 3);
                                        double body_internal_area = Math.Round((sweep_body.SurfaceArea - 2 * area_diff)
                                            / (edgearray_peri_area[1].Item2 + edgearray_peri_area[0].Item2)
                                            * edgearray_peri_area[0].Item2 * Math.Pow(0.3048, 2), 1);

                                        total_volume += body_volome;
                                        internal_area += body_internal_area;
                                    }
                                    break;
                                }
                            }
                        }
                        // b.頭尾體積(含內部空間)&內部表面積計算
                        if (sweep_head != null)
                        {
                            FaceArray facearray = sweep_head.Faces;
                            if (facearray != null)
                            {
                                List<PlanarFace> fp_z = new List<PlanarFace>();
                                List<Tuple<EdgeArray, double, double>> edgearray_w_l = new List<Tuple<EdgeArray, double, double>>();

                                foreach (GeometryObject f in facearray)
                                {
                                    PlanarFace fp = f as PlanarFace;
                                    if (null == fp || fp.Area == 0)
                                        continue;

                                    if (Math.Round(fp.FaceNormal.Z) == 0)
                                        fp_z.Add(fp);

                                    if (fp.EdgeLoops.Size > 1) //有內外雙層loop
                                    {
                                        foreach (EdgeArray edgearray in fp.EdgeLoops)
                                        {
                                            double peri = 0;
                                            List<Curve> edgearray_ln = new List<Curve>();

                                            foreach (Edge e in edgearray)
                                            {
                                                peri += e.ApproximateLength; //周長

                                                Curve c = e.AsCurve() as Curve;
                                                if (null == c)
                                                    continue;
                                                edgearray_ln.Add(c);
                                            }
                                            edgearray_ln = edgearray_ln.OrderBy(x => x.ApproximateLength).ToList();
                                            edgearray_ln.RemoveAt(0);
                                            edgearray_ln.RemoveAt(edgearray_ln.Count() - 1);
                                            double width = edgearray_ln.ElementAt(0).Length;
                                            double length = edgearray_ln.ElementAt(1).Length;

                                            edgearray_w_l.Add(Tuple.Create(edgearray, width, length));
                                        }
                                        edgearray_w_l = edgearray_w_l.OrderBy(x => x.Item2).ToList();
                                    }
                                }
                                fp_z = fp_z.OrderBy(x => x.Area).ToList();
                                int count_fp_z = fp_z.Count;

                                List<Line> ln_up_dn = new List<Line>();
                                foreach (EdgeArray edgearray in fp_z.ElementAt(count_fp_z - 4).EdgeLoops) //-4: 扣除最大的面、以及兩個外圍側面
                                {
                                    foreach (Edge e in edgearray)
                                    {
                                        Line l = e.AsCurve() as Line;
                                        if (null == l)
                                            continue;
                                        ln_up_dn.Add(l);
                                    }
                                }

                                ln_up_dn = ln_up_dn.OrderBy(x => x.ApproximateLength).ToList();
                                ln_up_dn.RemoveAt(ln_up_dn.Count - 1);
                                double dn_bound = ln_up_dn.ElementAt(0).ApproximateLength;
                                double up_bound = ln_up_dn.ElementAt(1).ApproximateLength;
                                double head_internal_area = Math.Round(((dn_bound + up_bound) * edgearray_w_l.ElementAt(0).Item2 + fp_z.ElementAt(count_fp_z - 4).Area * 2)
                                    * Math.Pow(0.3048, 2), 1);
                                internal_area += head_internal_area * 2;

                                double head_volme = Math.Round((fp_z.ElementAt(count_fp_z - 4).Area * edgearray_w_l.ElementAt(0).Item2 + sweep_head.Volume)
                                    * Math.Pow(0.3048, 3), 1);
                                total_volume += head_volme * 2;
                            }
                        }
                        total_vol_channel.Add(total_volume);
                        internal_area_channel.Add(internal_area);
                    }
                }
                for (int r = 0; r < contact_channels.Count; r++)
                {
                    contact_tunnel_profile.Add(Tuple.Create(
                        "聯絡通道" + "0" + (r + 1).ToString(),
                        internal_area_channel.ElementAt(r),
                        total_vol_channel.ElementAt(r),
                        diff_vol_channel.ElementAt(r),
                        door_counts.ElementAt(r)));
                }



                ////PART SEVEN:開始excel資料寫入

                path = Form1.path;
                string saveas_path = Form1.saveas_FilePath;
                string counting_temp_path = Form1.counting_temp_path;

                Excel.Application Eapp = new Excel.Application();
                Excel.Workbook Workbook = Eapp.Workbooks.Open(counting_temp_path); //選取樣板檔
                string new_path = saveas_path; //另存之儲存路徑

                if (File.Exists(new_path))
                    File.Delete(new_path);

                Workbook.SaveAs(new_path, Type.Missing, "", "", Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, 1, false, Type.Missing, Type.Missing, Type.Missing);

                Excel.Workbook n_Workbook = Eapp.Workbooks.Open(new_path);
                Excel.Worksheet n_sheet = (Excel.Worksheet)n_Workbook.Worksheets[1];
                
                Eapp.DisplayAlerts = false;
                Excel.Worksheet n_sheet2 = (Excel.Worksheet)n_Workbook.Worksheets[2];
                n_sheet2.Delete();
                //n_Workbook.Worksheets[2].Delete();
                Eapp.DisplayAlerts = true;

                Excel.Range n_sheet_range = n_sheet.UsedRange;


                //里程數量計算
                n_sheet.Cells[7, 2] = "鑽掘隧道"; n_sheet.Cells[7, 3] = "m"; n_sheet.Cells[7, 4] = data.sum_miles; n_sheet.Cells[7, 5] = data.sum_miles_cal; n_sheet.Cells[7, 7] = "0241400001A";
                n_sheet.Cells[8, 2] = "隧道仰拱(標準型)"; n_sheet.Cells[8, 3] = "m"; n_sheet.Cells[8, 4] = data.inverted_standard; n_sheet.Cells[8, 5] = data.inverted_standard_cal; n_sheet.Cells[8, 7] = "0241400001C1";
                n_sheet.Cells[9, 2] = "隧道仰拱(浮動式道床)"; n_sheet.Cells[9, 3] = "m"; n_sheet.Cells[9, 4] = data.inverted_floating; n_sheet.Cells[9, 5] = data.inverted_floating_cal; n_sheet.Cells[9, 7] = "02414D2001";
                n_sheet.Cells[10, 2] = "隧道仰拱(平板式道床)"; n_sheet.Cells[10, 3] = "m"; n_sheet.Cells[10, 4] = data.inverted_flat; n_sheet.Cells[10, 5] = data.inverted_flat_cal;
                n_sheet.Cells[11, 2] = "隧道步道(標準型)"; n_sheet.Cells[11, 3] = "m"; n_sheet.Cells[11, 4] = data.track_bed_standard; n_sheet.Cells[11, 5] = data.track_bed_standard_cal; n_sheet.Cells[11, 7] = "0241400001D1";
                n_sheet.Cells[12, 2] = "隧道步道(浮動式道床)"; n_sheet.Cells[12, 3] = "m"; n_sheet.Cells[12, 4] = data.track_bed_floating; n_sheet.Cells[12, 5] = data.track_bed_floating_cal;
                n_sheet.Cells[13, 2] = "隧道步道(平板式道床)"; n_sheet.Cells[13, 3] = "m"; n_sheet.Cells[13, 4] = data.track_bed_flat; n_sheet.Cells[13, 5] = data.track_bed_flat_cal;
                n_sheet.Cells[14, 2] = "隧道之軌道排水系統(標準型)"; n_sheet.Cells[14, 3] = "m"; n_sheet.Cells[14, 4] = data.drainage_standard; n_sheet.Cells[14, 5] = data.drainage_standard_cal; n_sheet.Cells[14, 7] = "0262000001R1";
                n_sheet.Cells[15, 2] = "隧道之軌道排水系統(浮動式道床)"; n_sheet.Cells[15, 3] = "m"; n_sheet.Cells[15, 4] = data.drainage_floating; n_sheet.Cells[15, 5] = data.drainage_floating_cal;
                n_sheet.Cells[16, 2] = "槽鋼嵌裝"; n_sheet.Cells[16, 7] = "05500R0001C1";
                n_sheet.Cells[17, 2] = "潛盾工法隧道開挖，皂土漿材料"; n_sheet.Cells[17, 7] = "24140000111";
                n_sheet.Cells[18, 2] = "潛盾工法隧道開挖，隧道鑽掘之運輸設備"; n_sheet.Cells[18, 3] = "m"; n_sheet.Cells[18, 4] = data.sum_miles; n_sheet.Cells[18, 5] = data.sum_miles_cal; n_sheet.Cells[18, 7] = "24140000131";
                n_sheet.Cells[19, 2] = "預鑄鋼筋混凝土環片"; n_sheet.Cells[19, 7] = "24140000121";
                n_sheet.Cells[20, 2] = "預鑄球狀石墨鑄鐵環片"; n_sheet.Cells[20, 3] = "m"; n_sheet.Cells[20, 4] = 0; n_sheet.Cells[20, 7] = "24140000122";
                n_sheet.Cells[21, 2] = "隧道鋼環片"; n_sheet.Cells[21, 3] = "m"; n_sheet.Cells[21, 4] = data.sum_miles; n_sheet.Cells[21, 5] = data.sum_miles_cal; n_sheet.Cells[21, 7] = "24140000123";
                n_sheet.Cells[22, 2] = "潛盾工法隧道開挖，開挖掘進，隧道鑽掘工資"; n_sheet.Cells[22, 3] = "m"; n_sheet.Cells[22, 4] = data.sum_miles; n_sheet.Cells[22, 5] = data.sum_miles_cal; n_sheet.Cells[22, 7] = "24145000111";
                n_sheet.Cells[23, 2] = "潛盾工法隧道開挖，開挖掘進，鑽掘隧道施工中排水"; n_sheet.Cells[23, 3] = "m"; n_sheet.Cells[23, 4] = data.sum_miles; n_sheet.Cells[23, 5] = data.sum_miles_cal; n_sheet.Cells[23, 7] = "24145000112";
                n_sheet.Cells[24, 2] = "潛盾工法隧道開挖，背填灌漿"; n_sheet.Cells[24, 7] = "24140000112";
                n_sheet.Cells[25, 2] = "潛盾工法隧道開挖，清理"; n_sheet.Cells[25, 7] = "24145000113";

                for (int i = 16; i <= 25; i++)
                {
                    if (i == 16 || i == 17 || i == 19 || i == 24 || i == 25)
                    {
                        n_sheet.Cells[i, 3] = "m"; n_sheet.Cells[i, 4] = data.sum_miles; n_sheet.Cells[i, 5] = data.sum_miles_cal;
                    }
                }                
                double.TryParse(GetCell(n_sheet, 20, 4).Value2.ToString(), out double val);
                double.TryParse(GetCell(n_sheet, 19, 4).Value2.ToString(), out double val_tunnel);
                n_sheet.Cells[19, 4] = val_tunnel - val;
                n_sheet.Cells[19, 5] = GetCell(n_sheet, 19, 5).Value2.ToString() + "-" + val.ToString();
                //double.TryParse(n_sheet.Cells[20, 4].Value2.ToString(), out double val);
                //double.TryParse(n_sheet.Cells[19, 4].Value2.ToString(), out double val_tunnel);
                //n_sheet.Cells[19, 4] = val_tunnel - val;
                //n_sheet.Cells[19, 5] = n_sheet.Cells[19, 5].Value2.ToString() + "-" + val.ToString();


                //隧道襯砌環片，隧道預鑄混凝土襯砌環片資料寫入
                n_sheet.Cells[26, 2] = "隧道襯砌環片，隧道預鑄混凝土襯砌環片"; n_sheet.Cells[26, 3] = "m"; n_sheet.Cells[26, 4] = 1;
                n_sheet.Cells[27, 2] = "環片鋼模板(每公尺長度)"; n_sheet.Cells[27, 3] = "m2"; n_sheet.Cells[27, 4] = ada_A_area_diff + ada_B_area_diff + ada_K_area_diff; n_sheet.Cells[27, 5] = ada_A_area_diff + "+" + ada_B_area_diff + "+" + ada_K_area_diff;
                n_sheet.Cells[27, 7] = "M0311090002";
                n_sheet.Cells[28, 2] = "TYPE A"; n_sheet.Cells[28, 3] = "m2"; n_sheet.Cells[28, 4] = ada_A_area_diff; n_sheet.Cells[28, 5] = ada_A_area_diff_str;
                n_sheet.Cells[29, 2] = "TYPE B"; n_sheet.Cells[29, 3] = "m2"; n_sheet.Cells[29, 4] = ada_B_area_diff; n_sheet.Cells[29, 5] = ada_B_area_diff_str;
                n_sheet.Cells[30, 2] = "TYPE K"; n_sheet.Cells[30, 3] = "m2"; n_sheet.Cells[30, 4] = ada_K_area_diff; n_sheet.Cells[30, 5] = ada_K_area_diff_str;
                n_sheet.Cells[31, 2] = "450kg/cm2混凝土(每單位公尺)"; n_sheet.Cells[31, 3] = ""; n_sheet.Cells[31, 4] = ada_A_vol_diff + ada_B_vol_diff + ada_K_vol_diff; n_sheet.Cells[31, 5] = ada_A_vol_diff + "+" + ada_B_vol_diff + "+" + ada_K_vol_diff;
                n_sheet.Cells[31, 7] = "M03050400030B2";
                n_sheet.Cells[32, 2] = "產品，混凝土養護，混凝土澆置及蒸氣養護"; n_sheet.Cells[32, 3] = ""; n_sheet.Cells[32, 4] = ada_A_vol_diff + ada_B_vol_diff + ada_K_vol_diff; n_sheet.Cells[32, 5] = ada_A_vol_diff + "+" + ada_B_vol_diff + "+" + ada_K_vol_diff;
                n_sheet.Cells[32, 7] = "M0339000003A1";
                n_sheet.Cells[33, 2] = "Bolt Pocket & Bolt Hole"; n_sheet.Cells[33, 3] = "m3"; n_sheet.Cells[33, 4] = bolt_vol.ToString().Split(new char[1] { '-' }).Last(); n_sheet.Cells[33, 5] = bolt_vol.ToString().Split(new char[1] { '-' }).Last();
                n_sheet.Cells[34, 2] = "TYPE A"; n_sheet.Cells[34, 3] = "m3"; n_sheet.Cells[34, 4] = ada_A_vol_diff; n_sheet.Cells[34, 5] = ada_A_vol_diff_str;
                n_sheet.Cells[35, 2] = "TYPE B"; n_sheet.Cells[35, 3] = "m3"; n_sheet.Cells[35, 4] = ada_B_vol_diff; n_sheet.Cells[35, 5] = ada_B_vol_diff_str;
                n_sheet.Cells[36, 2] = "TYPE K"; n_sheet.Cells[36, 3] = "m3"; n_sheet.Cells[36, 4] = ada_K_vol_diff; n_sheet.Cells[36, 5] = ada_K_vol_diff_str;

                
                //環片鋼筋數量計算寫入
                n_sheet.Cells[37, 2] = "環片鋼筋";
                n_sheet.Cells[38, 2] = "TYPEA";
                n_sheet.Cells[51, 2] = "∑D29"; n_sheet.Cells[51, 3] = "m"; n_sheet.Cells[51, 4] = A_29M; n_sheet.Cells[51, 5] = A_29M_str;
                n_sheet.Cells[52, 2] = "∑D25"; n_sheet.Cells[52, 3] = "m"; n_sheet.Cells[52, 4] = A_25M; n_sheet.Cells[52, 5] = A_25M_str;
                n_sheet.Cells[53, 2] = "∑D19"; n_sheet.Cells[53, 3] = "m"; n_sheet.Cells[53, 4] = A_19M; n_sheet.Cells[53, 5] = A_19M_str;
                n_sheet.Cells[54, 2] = "∑D16"; n_sheet.Cells[54, 3] = "m"; n_sheet.Cells[54, 4] = A_16M; n_sheet.Cells[54, 5] = A_16M_str;
                n_sheet.Cells[55, 2] = "∑D13"; n_sheet.Cells[55, 3] = "m"; n_sheet.Cells[55, 4] = A_13M; n_sheet.Cells[55, 5] = A_13M_str;

                n_sheet.Cells[57, 2] = "TYPEB";
                n_sheet.Cells[71, 2] = "∑D29"; n_sheet.Cells[71, 3] = "m"; n_sheet.Cells[71, 4] = B_29M; n_sheet.Cells[71, 5] = B_29M_str;
                n_sheet.Cells[72, 2] = "∑D25"; n_sheet.Cells[72, 3] = "m"; n_sheet.Cells[72, 4] = B_25M; n_sheet.Cells[72, 5] = B_25M_str;
                n_sheet.Cells[73, 2] = "∑D19"; n_sheet.Cells[73, 3] = "m"; n_sheet.Cells[73, 4] = B_19M; n_sheet.Cells[73, 5] = B_19M_str;
                n_sheet.Cells[74, 2] = "∑D16"; n_sheet.Cells[74, 3] = "m"; n_sheet.Cells[74, 4] = B_16M; n_sheet.Cells[74, 5] = B_16M_str;
                n_sheet.Cells[75, 2] = "∑D13"; n_sheet.Cells[75, 3] = "m"; n_sheet.Cells[75, 4] = B_13M; n_sheet.Cells[75, 5] = B_13M_str;

                n_sheet.Cells[77, 2] = "TYPEK";
                n_sheet.Cells[89, 2] = "∑D29"; n_sheet.Cells[89, 3] = "m"; n_sheet.Cells[89, 4] = K_29M; n_sheet.Cells[89, 5] = K_29M_str;
                n_sheet.Cells[90, 2] = "∑D25"; n_sheet.Cells[90, 3] = "m"; n_sheet.Cells[90, 4] = K_25M; n_sheet.Cells[90, 5] = K_25M_str;
                n_sheet.Cells[91, 2] = "∑D19"; n_sheet.Cells[91, 3] = "m"; n_sheet.Cells[91, 4] = K_19M; n_sheet.Cells[91, 5] = K_19M_str;
                n_sheet.Cells[92, 2] = "∑D16"; n_sheet.Cells[92, 3] = "m"; n_sheet.Cells[92, 4] = K_16M; n_sheet.Cells[92, 5] = K_16M_str;
                n_sheet.Cells[93, 2] = "∑D13"; n_sheet.Cells[93, 3] = "m"; n_sheet.Cells[93, 4] = K_13M; n_sheet.Cells[93, 5] = K_13M_str;
                n_sheet.Cells[94, 2] = "鋼環片鋼板"; n_sheet.Cells[94, 3] = "m3";
                


                //聯絡通道資料匯入
                n_sheet.Cells[95, 2] = "聯絡通道"; n_sheet.Cells[95, 3] = "處"; n_sheet.Cells[95, 4] = contact_tunnel_profile.Count();
                int s = 0;
                int d = 10;
                Dictionary<int, string> alpha_count = new Dictionary<int, string>() { { 1, "1" }, { 2, "2" }, { 3, "3" }, { 4, "4" }, { 5, "5" }, { 6, "6" } };
                foreach (Tuple<string, double, double, double, int> contact_tunnel in contact_tunnel_profile)

                {
                    n_sheet.Cells[96 + s * d, 2] = contact_tunnel.Item1; n_sheet.Cells[96 + s * d, 7] = "241400004P" + alpha_count[s + 1];
                    n_sheet.Cells[97 + s * d, 2] = "模板"; n_sheet.Cells[97 + s * d, 3] = "m2"; n_sheet.Cells[97 + s * d, 4] = contact_tunnel.Item2; n_sheet.Cells[97 + s * d, 7] = "03110000022V0";
                    n_sheet.Cells[98 + s * d, 2] = "開挖及棄土"; n_sheet.Cells[98 + s * d, 3] = "m3"; n_sheet.Cells[98 + s * d, 4] = contact_tunnel.Item3; n_sheet.Cells[98 + s * d, 7] = "023160T003";
                    n_sheet.Cells[99 + s * d, 2] = "餘方自行處理(含水土保持)，營建剩餘資源處理，土石方處理含運費"; n_sheet.Cells[99 + s * d, 3] = "m3"; n_sheet.Cells[99 + s * d, 4] = contact_tunnel.Item3; n_sheet.Cells[99 + s * d, 7] = "0232340003B0";
                    n_sheet.Cells[100 + s * d, 2] = "預拌混凝土280kg/cm2"; n_sheet.Cells[100 + s * d, 3] = "m3"; n_sheet.Cells[100 + s * d, 4] = contact_tunnel.Item4; n_sheet.Cells[100 + s * d, 7] = "M0305046203";
                    n_sheet.Cells[101 + s * d, 2] = "混凝土澆置，養護，修飾"; n_sheet.Cells[101 + s * d, 3] = "m3"; n_sheet.Cells[101 + s * d, 4] = contact_tunnel.Item4; n_sheet.Cells[101 + s * d, 7] = "03310000Z3A1";
                    n_sheet.Cells[102 + s * d, 2] = "PVC管65mm"; n_sheet.Cells[102 + s * d, 3] = "m";
                    n_sheet.Cells[103 + s * d, 2] = "PVC管200mm"; n_sheet.Cells[103 + s * d, 3] = "m";
                    n_sheet.Cells[104 + s * d, 2] = "不鏽鋼防火門"; n_sheet.Cells[104 + s * d, 3] = "樘"; n_sheet.Cells[104 + s * d, 4] = contact_tunnel.Item5; n_sheet.Cells[104 + s * d, 7] = "08130000000A5W4L21";
                    n_sheet.Cells[105 + s * d, 2] = "臨時安全閘門"; n_sheet.Cells[105 + s * d, 3] = "樘"; n_sheet.Cells[105 + s * d, 4] = contact_tunnel.Item5; n_sheet.Cells[105 + s * d, 7] = "0024140000BD1";
                    s = s + 1;
                }

                //仰拱及走道數量計算
                n_sheet.Cells[116, 2] = "潛盾工法隧道開挖，仰拱(每100m)"; n_sheet.Cells[116, 3] = "100m";
                n_sheet.Cells[117, 2] = "標準型道床仰拱模板"; n_sheet.Cells[117, 3] = "m2"; n_sheet.Cells[117, 4] = stn_arc_temp * 100; n_sheet.Cells[117, 5] = stn_arc_temp.ToString() + "*100"; n_sheet.Cells[117, 7] = "03110000021V0";
                n_sheet.Cells[118, 2] = "標準型道床仰拱混凝土(210kg / cm2)"; n_sheet.Cells[118, 3] = "m3"; n_sheet.Cells[118, 4] = stn_volume; n_sheet.Cells[118, 5] = str_stn_volume; n_sheet.Cells[118, 7] = "M0305044103";
                n_sheet.Cells[119, 2] = "標準型道床混凝土澆置，養護"; n_sheet.Cells[119, 3] = "m3"; n_sheet.Cells[119, 4] = stn_volume; n_sheet.Cells[119, 5] = str_stn_volume; n_sheet.Cells[119, 7] = "03310000Z3A1";
                n_sheet.Cells[120, 2] = "浮動式道床仰拱模板"; n_sheet.Cells[120, 3] = "m2"; n_sheet.Cells[120, 4] = flt_arc_temp * 100; n_sheet.Cells[120, 5] = flt_arc_temp.ToString() + "*100";
                n_sheet.Cells[121, 2] = "浮動式道床仰拱混凝土(210kg / cm2)"; n_sheet.Cells[121, 3] = "m3"; n_sheet.Cells[121, 4] = flt_volume; n_sheet.Cells[121, 5] = str_flt_volume;
                n_sheet.Cells[122, 2] = "浮動式道床混凝土澆置，養護"; n_sheet.Cells[122, 3] = "m3"; n_sheet.Cells[122, 4] = flt_volume; n_sheet.Cells[122, 5] = str_flt_volume;
                n_sheet.Cells[123, 2] = "平板式道床仰拱模板"; n_sheet.Cells[123, 3] = "m2"; n_sheet.Cells[123, 4] = stn_arc_temp * 100; n_sheet.Cells[123, 5] = stn_arc_temp.ToString() + "*100";
                n_sheet.Cells[124, 2] = "平板式道床仰拱混凝土(210kg / cm2)"; n_sheet.Cells[124, 3] = "m3"; n_sheet.Cells[124, 4] = stn_volume; n_sheet.Cells[124, 5] = str_stn_volume;
                n_sheet.Cells[125, 2] = "平板式道床混凝土澆置，養護"; n_sheet.Cells[125, 3] = "m3"; n_sheet.Cells[125, 4] = stn_volume; n_sheet.Cells[125, 5] = str_stn_volume;

                n_sheet.Cells[126, 2] = "潛盾工法隧道開挖，隧道走道(每100m)"; n_sheet.Cells[126, 3] = "100m";
                n_sheet.Cells[127, 2] = "標準型道床走道模板"; n_sheet.Cells[127, 3] = "m2"; n_sheet.Cells[127, 4] = stn_path_temp; n_sheet.Cells[127, 5] = str_stn_path; n_sheet.Cells[127, 7] = "03110000022V0";
                n_sheet.Cells[128, 2] = "標準型道床走道預鑄蓋板"; n_sheet.Cells[128, 3] = "EA"; n_sheet.Cells[128, 4] = stn_sidewalk_lid; n_sheet.Cells[128, 5] = str_stn_sidewalk_lid; n_sheet.Cells[128, 7] = "M034005000A100i";
                n_sheet.Cells[129, 2] = "標準型道床走道混凝土210kg/cm2"; n_sheet.Cells[129, 3] = "m3"; n_sheet.Cells[129, 4] = stn_path_volume; n_sheet.Cells[129, 5] = str_stn_path_vol; n_sheet.Cells[129, 7] = "M0305044103";
                n_sheet.Cells[130, 2] = "標準型道床走道混凝土澆置、養護"; n_sheet.Cells[130, 3] = "m3"; n_sheet.Cells[130, 4] = stn_path_volume; n_sheet.Cells[130, 5] = str_stn_path_vol; n_sheet.Cells[130, 7] = "03310000Z3A1";
                n_sheet.Cells[131, 2] = "浮動式道床走道模板"; n_sheet.Cells[131, 3] = "m2"; n_sheet.Cells[131, 4] = flt_path_temp; n_sheet.Cells[131, 5] = str_flt_path;
                n_sheet.Cells[132, 2] = "浮動式道床走道預鑄蓋板"; n_sheet.Cells[132, 3] = "EA"; n_sheet.Cells[132, 4] = flt_sidewalk_lid; n_sheet.Cells[132, 5] = str_flt_sidewalk_lid;
                n_sheet.Cells[133, 2] = "浮動式道床走道混凝土210kg/cm2"; n_sheet.Cells[133, 3] = "m3"; n_sheet.Cells[133, 4] = flt_path_volume; n_sheet.Cells[133, 5] = str_flt_path_vol;
                n_sheet.Cells[134, 2] = "浮動式道床走道混凝土澆置、養護"; n_sheet.Cells[134, 3] = "m3"; n_sheet.Cells[134, 4] = flt_path_volume; n_sheet.Cells[134, 5] = str_flt_path_vol;
                n_sheet.Cells[135, 2] = "平板式道床走道模板"; n_sheet.Cells[135, 3] = "m2"; n_sheet.Cells[135, 4] = stn_path_temp; n_sheet.Cells[135, 5] = str_stn_path;
                n_sheet.Cells[136, 2] = "平板式道床走道預鑄蓋板"; n_sheet.Cells[136, 3] = "EA"; n_sheet.Cells[136, 4] = stn_sidewalk_lid; n_sheet.Cells[136, 5] = str_stn_sidewalk_lid; ;
                n_sheet.Cells[137, 2] = "平板式道床走道混凝土210kg/cm2"; n_sheet.Cells[137, 3] = "m3"; n_sheet.Cells[137, 4] = stn_path_volume; n_sheet.Cells[137, 5] = str_stn_path_vol;
                n_sheet.Cells[138, 2] = "平板式道床走道混凝土澆置、養護"; n_sheet.Cells[138, 3] = "m3"; n_sheet.Cells[138, 4] = stn_path_volume; n_sheet.Cells[138, 5] = str_stn_path_vol;

                n_sheet.Cells[139, 2] = "潛盾工法隧道開挖，隧道排水系統"; n_sheet.Cells[139, 3] = "100m";
                n_sheet.Cells[140, 2] = "標準型道床排水預鑄蓋板210kg/cm2"; n_sheet.Cells[140, 3] = "EA"; n_sheet.Cells[140, 7] = "M034005000A0X10";
                n_sheet.Cells[141, 2] = "標準型道床排水模板"; n_sheet.Cells[141, 3] = "m2"; n_sheet.Cells[141, 4] = gutter_temp; n_sheet.Cells[141, 5] = str_gutter_temp; n_sheet.Cells[141, 7] = "M0311090002";
                n_sheet.Cells[142, 2] = "標準型道床熱鍍鋅蓋板"; n_sheet.Cells[142, 3] = " "; n_sheet.Cells[142, 4] = 7; n_sheet.Cells[142, 7] = "M055306000A10";

                double.TryParse(GetCell(n_sheet, 142, 4).Value2.ToString(), out double val_cover);
                //double.TryParse(n_sheet.Cells[142, 4].Value2.ToString(), out double val_cover);
                n_sheet.Cells[140, 4] = 1 / 0.5 * 100 - val_cover;
                n_sheet.Cells[140, 5] = stn_sidewalk_lid.ToString() + "-" + val_cover.ToString();

                n_sheet.Cells[143, 2] = "槽鋼";
                n_sheet.Cells[144, 2] = "槽鋼體積"; n_sheet.Cells[144, 3] = "m3"; n_sheet.Cells[144, 4] = U_shape_steel_volume; n_sheet.Cells[144, 7] = "M0506061009C";

                n_Workbook.Save();
                n_sheet = null;
                n_Workbook.Close();
                n_Workbook = null;
                Eapp.Quit();
                Eapp = null;

                TaskDialog.Show("Counting Result", "excel樣板檔寫入完成");
            }
            catch (Exception e)
            {
                TaskDialog.Show("error", e.Message + e.StackTrace);
            }

        }
        /// <summary>
        /// 培文改：.NET Core 8.0 需要使用dynamic來取得Excel的Cell值
        /// </summary>
        /// <param name="xlSheet"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private dynamic GetCell(Excel.Worksheet xlSheet, int row, int col)
        {
            return xlSheet.Cells[row, col];
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
