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

namespace SinoTunnel_2025
{


    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    
    class loft : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            List<string> name_list = new List<string> { "A1", "A2", "A3", "B1", "B2", "K"};
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            string doc_path = doc.PathName;
            Application app = doc.Application;
            UIApplication uiapp = new UIApplication(app);
            
            segment_para segment = new segment_para();
            segment.displacement_angle = 18.0;
            segment.angle_dic = new Dictionary<string, double>();
            segment.angle_dic.Add("K", 7.5);
            segment.angle_dic.Add("B1", 34.0);
            segment.angle_dic.Add("B2", 34.0);
            segment.angle_dic.Add("A1", 34.3);
            segment.angle_dic.Add("A2", 36.0);
            segment.angle_dic.Add("A3", 34.3);

            segment.rotation_dic = new Dictionary<string, double>();
            segment.rotation_dic.Add("K", segment.displacement_angle);
            segment.rotation_dic.Add("B2", segment.rotation_dic["K"] + segment.angle_dic["K"] + segment.angle_dic["B2"]);
            segment.rotation_dic.Add("A3", segment.rotation_dic["B2"] + segment.angle_dic["B2"] + segment.angle_dic["A3"]);
            segment.rotation_dic.Add("A2", segment.rotation_dic["A3"] + segment.angle_dic["A3"] + segment.angle_dic["A2"]);
            segment.rotation_dic.Add("A1", segment.rotation_dic["A2"] + segment.angle_dic["A2"] + segment.angle_dic["A1"]);
            segment.rotation_dic.Add("B1", segment.rotation_dic["A1"] + segment.angle_dic["A1"] + segment.angle_dic["B1"]);

            segment_para segment_b = new segment_para(); ;
            segment_b.displacement_angle = -18.0;
            segment_b.angle_dic = new Dictionary<string, double>();
            segment_b.angle_dic.Add("K", 7.5);
            segment_b.angle_dic.Add("B1", 34.0);
            segment_b.angle_dic.Add("B2", 34.0);
            segment_b.angle_dic.Add("A1", 34.3);
            segment_b.angle_dic.Add("A2", 36.0);
            segment_b.angle_dic.Add("A3", 34.3);
            segment_b.rotation_dic = new Dictionary<string, double>();
            segment_b.rotation_dic.Add("K", segment_b.displacement_angle);
            segment_b.rotation_dic.Add("B2", segment_b.rotation_dic["K"] + segment_b.angle_dic["K"] + segment_b.angle_dic["B2"]);
            segment_b.rotation_dic.Add("A3", segment_b.rotation_dic["B2"] + segment_b.angle_dic["B2"] + segment_b.angle_dic["A3"]);
            segment_b.rotation_dic.Add("A2", segment_b.rotation_dic["A3"] + segment_b.angle_dic["A3"] + segment_b.angle_dic["A2"]);
            segment_b.rotation_dic.Add("A1", segment_b.rotation_dic["A2"] + segment_b.angle_dic["A2"] + segment_b.angle_dic["A1"]);
            segment_b.rotation_dic.Add("B1", segment_b.rotation_dic["A1"] + segment_b.angle_dic["A1"] + segment_b.angle_dic["B1"]);


            segment.displacement_width = 100;
            segment.r2 = 6000;
            segment.r1 = 5000;
            //segment.abangle = 5;
            segment.width = 1000;

            segment_b.displacement_width = 100;
            segment_b.r2 = 6000;
            segment_b.r1 = 5000;
            //segment_b.abangle = 5;
            segment_b.width = 1000;

            List<string> abnormal_segment_parameter_name = new List<string> {"小半弧長", "內徑", "外徑", "寬度", "平移寬度", "旋轉角", "角度"};
            List<string> normal_segment_parameter_name = new List<string> { "小半弧長", "內徑", "外徑", "寬度", "平移寬度", "旋轉角度" };

            foreach(string name in name_list)
            {
                //TaskDialog.Show("message", name);
                loft_segment(doc_path, uiapp, abnormal_segment_parameter_name, segment, name);
                //TaskDialog.Show("message", "pass");
                set_segment(doc_path, uiapp, normal_segment_parameter_name, segment, name, "");
                set_segment(doc_path, uiapp, normal_segment_parameter_name, segment_b, name, "b");
            }
            

           


            
            return Result.Succeeded;
        }

        public void loft_segment(string path, UIApplication uiapp, List<string> abnormal_segment_parameter_name, segment_para segment, string name)
        {
            
            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(@"C:\Users\exist\OneDrive\桌面\0409\異形環片\origin\" + name + ".rfa");
            Document edit_doc = edit_uidoc.Document;

            set_value(edit_doc, abnormal_segment_parameter_name[0], segment.angle_dic[name].ToString());
            set_value(edit_doc, abnormal_segment_parameter_name[1], segment.r1.ToString());
            set_value(edit_doc, abnormal_segment_parameter_name[2], segment.r2.ToString());
            set_value(edit_doc, abnormal_segment_parameter_name[3], segment.width.ToString());
            set_value(edit_doc, abnormal_segment_parameter_name[4], segment.displacement_width.ToString());
            set_value(edit_doc, abnormal_segment_parameter_name[5], segment.rotation_dic[name].ToString());
            //set_value(edit_doc, abnormal_segment_parameter_name[6], segment.abangle.ToString());

            FilteredElementCollector collector = new FilteredElementCollector(edit_doc);
            ElementCategoryFilter model_line_filter = new ElementCategoryFilter(BuiltInCategory.OST_Lines);
            IList<Element> line_list = collector.WherePasses(model_line_filter).WhereElementIsNotElementType().ToElements();

            Tuple<IList<Element>, IList<Element>> sketch_line_list = distinguish_lines(line_list);

            IList<Element> sketch_line_list1 = arrange_lines(sketch_line_list.Item1);
            IList<Element> sketch_line_list2 = arrange_lines(sketch_line_list.Item2);

            ReferenceArrayArray ref_ar_ar = new ReferenceArrayArray();
            ReferenceArray ref_ar = getObjectRef(sketch_line_list1);
            ref_ar_ar.Append(ref_ar);
            ref_ar = getObjectRef(sketch_line_list2);
            ref_ar_ar.Append(ref_ar);

            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;
            Transaction t = new Transaction(edit_doc, "sweep");
            t.Start();
            edit_doc.FamilyCreate.NewLoftForm(true, ref_ar_ar);
            t.Commit();
            edit_doc.SaveAs(@"C:\Users\exist\OneDrive\桌面\0409\異形環片\" + name + ".rfa", save_option);
            uiapp.OpenAndActivateDocument(path);
            edit_doc.Close(false);
        }

        public void set_segment(string path, UIApplication uiapp, List<string> normal_segment_parameter_name, segment_para segment, string name, string tail)
        {
            UIDocument edit_uidoc = uiapp.OpenAndActivateDocument(@"C:\Users\exist\OneDrive\桌面\0409\正常環片\origin\" + name + ".rfa");
            Document edit_doc = edit_uidoc.Document;

            set_value(edit_doc, normal_segment_parameter_name[0], segment.angle_dic[name].ToString());
            set_value(edit_doc, normal_segment_parameter_name[1], segment.r1.ToString());
            set_value(edit_doc, normal_segment_parameter_name[2], segment.r2.ToString());
            set_value(edit_doc, normal_segment_parameter_name[3], segment.width.ToString());
            set_value(edit_doc, normal_segment_parameter_name[4], segment.displacement_width.ToString());
            set_value(edit_doc, normal_segment_parameter_name[5], segment.rotation_dic[name].ToString());

            SaveAsOptions save_option = new SaveAsOptions();
            save_option.OverwriteExistingFile = true;
            edit_doc.SaveAs(@"C:\Users\exist\OneDrive\桌面\0409\正常環片\" + name + tail + ".rfa", save_option);
            uiapp.OpenAndActivateDocument(path);
            edit_doc.Close(false);
        }

        public void set_value(Document edit_doc, string name, string value)
        {
            Transaction t = new Transaction(edit_doc, "set");
            t.Start();
            try
            {
                FamilyManager fm = edit_doc.FamilyManager;
                FamilyParameter p = fm.get_Parameter(name);
                fm.SetValueString(p, value);
                t.Commit();
            }
            catch
            {
                t.Commit();
            }
        }


        public ReferenceArray getObjectRef(IList<Element> modelcurve_list)
        {
            ReferenceArray ra = new ReferenceArray();
            for (int i = 0; i < modelcurve_list.Count; i++)
            {
                ra.Append((modelcurve_list[i] as CurveElement).GeometryCurve.Reference);
            }
            return ra;
        }

        public Curve find_path(IList<Element> line_list1, IList<Element> line_list2)
        {
            XYZ start;
            XYZ end = new XYZ(0, 0, 0);
            if (line_list1.Count > 0)
            {
                start = (line_list1[0].Location as LocationCurve).Curve.GetEndPoint(0);
            }
            else
            {
                return null;
            }

            for (int i = 0; i < line_list2.Count; i++)
            {
                if (start.IsAlmostEqualTo((line_list2[i].Location as LocationCurve).Curve.GetEndPoint(0)))
                {
                    end = (line_list2[i].Location as LocationCurve).Curve.GetEndPoint(0);
                    break;
                }
                else if (start.IsAlmostEqualTo((line_list2[i].Location as LocationCurve).Curve.GetEndPoint(1)))
                {
                    end = (line_list2[i].Location as LocationCurve).Curve.GetEndPoint(1);
                    break;
                }
            }

            Line path = Line.CreateBound(start, end);

            return path;
        }

        public Tuple<IList<Element>, IList<Element>> distinguish_lines(IList<Element> line_list)
        {

            IList<Element> line_list1 = new List<Element>();
            IList<Element> line_list2 = new List<Element>();

            for (int i = 0; i < line_list.Count; i++)
            {
                CurveElement CE = line_list[i] as CurveElement;
                SketchPlane SP = CE.SketchPlane;
                ElementId id = SP.Id;
                if (i == 0)
                {
                    line_list1.Add(line_list[0]);
                }
                else
                {
                    CurveElement tempCE = line_list1[0] as CurveElement;
                    SketchPlane tempSP = tempCE.SketchPlane;
                    ElementId tempid = tempSP.Id;
                    if (tempid.ToString().Equals(id.ToString()))
                    {
                        line_list1.Add(line_list[i]);
                    }
                    else
                    {
                        line_list2.Add(line_list[i]);
                    }
                }
            }

            Tuple<IList<Element>, IList<Element>> sketch_lines = new Tuple<IList<Element>, IList<Element>>(line_list1, line_list2);
            return sketch_lines;
        }

        public IList<Element> arrange_lines(IList<Element> line_list)
        {
            IList<Element> arranged_line_list = new List<Element>();
            IList<Element> line_list_duplicated = line_list;



            if (line_list.Count > 0)
            {
                arranged_line_list.Add(line_list_duplicated[0]);
                line_list_duplicated.RemoveAt(0);
            }

            XYZ temp_start, temp_end, end = new XYZ(0, 0, 0);

            for (int i = 0; i < arranged_line_list.Count; i++)
            {
                if (i == 0)
                {
                    LocationCurve lc1 = arranged_line_list[i].Location as LocationCurve;
                    end = lc1.Curve.GetEndPoint(1);
                }

                for (int j = 0; j < line_list_duplicated.Count; j++)
                {
                    LocationCurve lc2 = line_list_duplicated[j].Location as LocationCurve;
                    temp_start = lc2.Curve.GetEndPoint(0);
                    temp_end = lc2.Curve.GetEndPoint(1);

                    if (end.IsAlmostEqualTo(temp_start))
                    {
                        arranged_line_list.Add(line_list_duplicated[j]);

                        line_list_duplicated.RemoveAt(j);
                        end = temp_end;
                        break;
                    }
                    else if (end.IsAlmostEqualTo(temp_end))
                    {
                        arranged_line_list.Add(line_list_duplicated[j]);
                        line_list_duplicated.RemoveAt(j);
                        end = temp_start;
                        break;
                    }

                }
            }

            return arranged_line_list;
        }

    }
}
