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
    class Attached_pipeline : IExternalEventHandler
    {
        string path;

        public void Execute(UIApplication app)
        {
            try
            {
                // 讀取管線族群檔
                path = Form1.path;
                Document pro_doc = app.ActiveUIDocument.Document;
                UIDocument edit_uidoc = app.OpenAndActivateDocument(path + "附掛設施\\管線模型.rfa");
                Document doc = edit_uidoc.Document;

                // 讀取excel資料
                readfile rf = new readfile();
                rf.read_properties();
                rf.read_target_station();
                rf.read_target_station2();
                rf.read_tunnel_point();
                rf.read_tunnel_point2();
                IList<IList<data_object>> all_data_list_tunnel = new List<IList<data_object>>();
                all_data_list_tunnel.Add(rf.data_list_tunnel);
                all_data_list_tunnel.Add(rf.data_list_tunnel2);
                List<Dictionary<string, string>> facilities_info = rf.Attached_facilities_info();
                IList<IList<string[]>> walk_way_list = new List<IList<string[]>>();

                walk_way_list.Add(rf.setting_Station.walk_way_station);
                walk_way_list.Add(rf.setting_Station2.walk_way_station);

                List<string[]> walk_way = new List<string[]>();
                walk_way.AddRange(rf.setting_Station.walk_way_station);
                walk_way.AddRange(rf.setting_Station2.walk_way_station);

                // 依照走道方向不同分割里程點位
                all_data_list_tunnel = cut_point_list(all_data_list_tunnel, walk_way_list);

                Transaction T = new Transaction(doc);
                FailureHandlingOptions failureHandlingOptions = T.GetFailureHandlingOptions();
                FailureHandler failureHandler = new FailureHandler();
                failureHandlingOptions.SetFailuresPreprocessor(failureHandler);
                failureHandlingOptions.SetClearAfterRollback(true);

                // 建立管線
                int count = 0;
                foreach (IList<data_object> data_list_tunnel in all_data_list_tunnel)
                {
                    count += 1;
                    foreach (Dictionary<string, string> fafacility in facilities_info)
                    {
                        if (fafacility["類型"].Contains("管線") || fafacility["類型"].Contains("電纜"))
                        {
                            for (int pipe_num = 0; pipe_num < int.Parse(fafacility["數量"]); pipe_num++)
                            {
                                // 依照excel資訊建立輪廓
                                T.Start("建置輪廓");

                                FamilySymbol pipe = new FilteredElementCollector(doc)
                                                    .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == fafacility["類型"]).First();
                                FamilySymbol dup_pipe = pipe.Duplicate(fafacility["項目"] + "_" + (pipe_num + 1).ToString() + count.ToString()) as FamilySymbol;
                                dup_pipe.LookupParameter("項目").SetValueString(fafacility["項目"]);
                                dup_pipe.LookupParameter("直徑").SetValueString(fafacility["直徑"]);
                                dup_pipe.LookupParameter("隧道中心點").SetValueString(rf.properties.center_point.ToString());
                                dup_pipe.LookupParameter("隧道內徑").SetValueString(rf.properties.inner_diameter.ToString());

                                if (fafacility["數量"] != "1")
                                {
                                    string Y_pos = (double.Parse(fafacility["Y位置"]) + (double.Parse(fafacility["Y範圍"]) * pipe_num / (double.Parse(fafacility["數量"]) - 1))).ToString();
                                    string X_pos = (double.Parse(fafacility["Y位置"] + rf.properties.center_point) >= rf.properties.inner_diameter / 2)
                                        ? (double.Parse(fafacility["X位置"]) - double.Parse(fafacility["X範圍"]) * pipe_num / (int.Parse(fafacility["數量"]) - 1)).ToString()
                                        : (double.Parse(fafacility["X位置"]) + double.Parse(fafacility["X範圍"]) * pipe_num / (int.Parse(fafacility["數量"]) - 1)).ToString();

                                    
                                    dup_pipe.LookupParameter("Y位置").Set(Y_pos);
                                    dup_pipe.LookupParameter("X位置").SetValueString(X_pos);
                                }
                                else
                                {
                                    string Y_pos = fafacility["Y位置"];
                                    string X_pos = fafacility["X位置"];

                                    dup_pipe.LookupParameter("Y位置").SetValueString(Y_pos);
                                    dup_pipe.LookupParameter("X位置").SetValueString(X_pos);
                                }


                                T.Commit();
                                // 利用以上建置完成之輪廓建立管線
                                T.Start("建置管線");
                                SweepProfile pipe_profile = doc.Application.Create.NewFamilySymbolProfile(dup_pipe);
                                
                                List<ElementId> m_list = new List<ElementId>();
                                ReferenceArray sweep_path = new ReferenceArray();

                                XYZ init_dir = data_list_tunnel[1].start_point - data_list_tunnel[0].start_point;
                                XYZ start = data_list_tunnel[0].start_point;
                                XYZ fake_pt = data_list_tunnel[0].start_point + init_dir / 100;
                                data_list_tunnel[0].start_point = new XYZ(fake_pt.X, fake_pt.Y, start.Z);
                                Line fake_path = Line.CreateBound(start, data_list_tunnel[0].start_point);
                                ModelCurve fake_curve = doc.FamilyCreate.NewModelCurve(fake_path, Sketch_plain(doc, start, data_list_tunnel[0].start_point));
                                m_list.Add(fake_curve.Id);
                                sweep_path.Append(fake_curve.GeometryCurve.Reference);

                                // 依照里程點為建立管線依循線段
                                for (int i = 1; i < data_list_tunnel.Count(); i++)
                                {
                                    Line n_single_path = Line.CreateBound(data_list_tunnel[i - 1].start_point, data_list_tunnel[i].start_point);
                                    ModelCurve m_curve = doc.FamilyCreate.NewModelCurve(n_single_path, Sketch_plain(doc, data_list_tunnel[i - 1].start_point, data_list_tunnel[i].start_point));
                                    m_list.Add(m_curve.Id);
                                    sweep_path.Append(m_curve.GeometryCurve.Reference);
                                }

                                // 依照管線走道側與非走道側改變管線位置
                                Sweep pipeline = doc.FamilyCreate.NewSweep(true, sweep_path, pipe_profile, 0, ProfilePlaneLocation.Start);
                                int turn = 0;
                                if (walk_way[count-1][0] == "右側" && fafacility["側"] == "走道側")
                                    turn = 1;
                                if (walk_way[count-1][0] == "左側" && fafacility["側"] == "非走道側")
                                    turn = 1;
                                T.Commit();

                                T.Start("翻轉輪廓");
                                dup_pipe.LookupParameter("走道側").Set(turn);
                                pipeline.LookupParameter("輪廓翻轉").Set(turn);

                                doc.Delete(m_list);

                                T.Commit();
                            }

                        }
                        else
                        {
                            // 建立附掛設施
                            double station = 0;
                            T.Start("建置設施");
                            // 依照里程點為放置附掛設施元件
                            for (int i = 0; i < data_list_tunnel.Count()-1; i++)
                            {
                                if (i == 0){
                                    station = double.Parse(data_list_tunnel[i].station.Split('+')[0]) * 1000 + double.Parse(data_list_tunnel[i].station.Split('+')[1]);
                                }
                                double tunnel_station = double.Parse(data_list_tunnel[i].station.Split('+')[0])*1000 + double.Parse(data_list_tunnel[i].station.Split('+')[1]);
                                double next_tunnel_station = double.Parse(data_list_tunnel[i+1].station.Split('+')[0]) * 1000 + double.Parse(data_list_tunnel[i+1].station.Split('+')[1]);
                                if (station >= tunnel_station && station < next_tunnel_station)
                                {
                                    double c = station % 1000;
                                    Line toward = Line.CreateBound(data_list_tunnel[i].start_point, data_list_tunnel[i + 1].start_point);
                                    XYZ put_point = data_list_tunnel[i].start_point + (c / 304.8) * toward.Direction + (0.61/0.3048) * toward.Direction;
                                    if(i == 0)
                                    {
                                        put_point = data_list_tunnel[i].start_point + (0.61 / 0.3048) * toward.Direction;
                                    }
                                    FamilySymbol facility = new FilteredElementCollector(doc)
                                                            .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where(x => x.Name == fafacility["項目"]).First();
                                    facility.Activate();
                                    
                                    XYZ x_axis = new XYZ(1, 0, 0);
                                    double x_dis = double.Parse(fafacility["X位置"])/304.8;

                                    // 依照附掛設施走道側與非走道側改變管線位置
                                    bool turn = false;
                                    if (walk_way[count - 1][0] == "右側" && fafacility["側"] == "走道側")
                                        turn = true;
                                    if (walk_way[count - 1][0] == "左側" && fafacility["側"] == "非走道側")
                                        turn = true;
                                    if (turn)
                                    {
                                        x_dis = -x_dis;
                                    }

                                    FamilyInstance FI = doc.FamilyCreate.NewFamilyInstance(put_point + x_axis * x_dis + XYZ.BasisZ * ((double.Parse(fafacility["Y位置"])-rf.properties.center_point) / 304.8), facility, StructuralType.NonStructural);

                                    if(count < walk_way_list[0].Count+1)
                                    {
                                        XYZ loc = put_point + x_axis * x_dis + XYZ.BasisZ * ((double.Parse(fafacility["Y位置"]) - rf.properties.center_point) / 304.8);
                                        Line axis = Line.CreateBound(loc, loc + XYZ.BasisZ); 
                                        ElementTransformUtils.RotateElement(doc, FI.Id, axis, Math.PI);
                                    }


                                    Line rotate_axis = Line.CreateBound(put_point, put_point + XYZ.BasisZ);
                                    double toward_slope = toward.Direction.Y / toward.Direction.X;
                                    double toward_angle = Math.Atan(toward_slope);

                                    if (toward.Direction.X < 0)
                                    {
                                        toward_angle += Math.PI;
                                    }
                                    ElementTransformUtils.RotateElement(doc, FI.Id, rotate_axis, toward_angle + Math.PI / 2);
                                    
                                    // 依照附掛設施間距放置附掛設施
                                    station += double.Parse(fafacility["間距"])/1000;

                                }

                            }
                            T.Commit();
                            
                        }
                    }
                }

                // 儲存附掛設施與管線族群檔
                SaveAsOptions saveAsOptions = new SaveAsOptions { OverwriteExistingFile = true, MaximumBackups = 1 };
                doc.SaveAs(path + "附掛設施\\附掛管線.rfa", saveAsOptions);

                app.OpenAndActivateDocument(pro_doc.PathName);

                doc.Close();

                // 將附掛設施族群檔放入專案中
                Transaction pro_t = new Transaction(pro_doc);

                pro_t.Start("載入族群");

                pro_doc.LoadFamilySymbol(path + "附掛設施\\附掛管線.rfa", "附掛管線");
                FamilySymbol track_bed = new FilteredElementCollector(pro_doc)
                    .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList().Where
                    (x => x.Name == "附掛管線").First();
                track_bed.Activate();
                FamilyInstance object_acr = pro_doc.Create.NewFamilyInstance(XYZ.Zero, track_bed, StructuralType.NonStructural);
                pro_t.Commit();
            }
            catch (Exception e) { TaskDialog.Show("error", e.Message + e.StackTrace); }
        }

        // 依照走道方向不同分割里程點位，給後續建立管線與附掛設施能依照走道側與非走到側建置
        public IList<IList<data_object>> cut_point_list(IList<IList<data_object>> all_data_list_tunnel,IList<IList<string[]>> walk_way_list)
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
        
        // 取得向量之X axis
        public XYZ get_x_axis(XYZ dir)
        {
            dir = new XYZ(dir.X, 0, dir.Z);

            return new XYZ(-dir.Z, 0, dir.X).Normalize();
        }

        // 利用兩個點當作法向量建立平面
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

        // failure handler 與程式目的較不相關
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
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
