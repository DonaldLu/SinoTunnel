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
using System.Windows.Forms;

namespace SinoTunnel
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class section_for_test : IExternalEventHandler
    {
        string path;
        
        public void Execute(UIApplication uiapp)
        {
            // 讀取revit專案檔
            path = Form1.path;
            Autodesk.Revit.DB.Document document = uiapp.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            string doc_path = doc.PathName;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;

            // 選取物件回傳物件id
            IList<Reference> element_list = uiapp.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, "Select element or ESC to reset the view").ToList<Reference>();
            IList<ElementId> ele_id_list = new List<ElementId>();
            foreach (Reference r in element_list)
            {
                ele_id_list.Add(r.ElementId);
            }
            
            // 取得物件剖面
            IList<ViewSection> vs_list = new List<ViewSection>();
            foreach(ElementId id in ele_id_list)
            {
                if(doc.GetElement(id).Category.Name == "一般模型")
                {
                    vs_list.Add(create_section(document, doc.GetElement(id)));
                }
                else if(doc.GetElement(id).Name.Contains("剖"))
                {
                    ViewSection viewPlane = new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Where(y => y.Name == doc.GetElement(id).Name).First();
                    vs_list.Add(viewPlane);
                }
            }
            
            // 將剖面放到圖紙上
            Transaction t = new Transaction(doc);
            //開始圖紙建立
            t.Start("創建圖紙");
            ElementId elementId = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks).ToElements().First().Id;
            //建立圖紙
            ViewSheet viewSheet = ViewSheet.Create(doc, ElementId.InvalidElementId);
            //載入圖紙
            Autodesk.Revit.DB.View view = viewSheet as Autodesk.Revit.DB.View;
            DWGImportOptions dWGImportOptions = new DWGImportOptions();
            dWGImportOptions.ColorMode = ImportColorMode.Preserved;
            dWGImportOptions.Placement = ImportPlacement.Centered;
            dWGImportOptions.Unit = ImportUnit.Centimeter;

            // 將instance匯入圖紙上
            LinkLoadResult linkLoadResult = new LinkLoadResult();
            ImportInstance toz = ImportInstance.Create(doc, view, Form1.frame_path, dWGImportOptions, out linkLoadResult);
            
            // 選出欲匯入圖紙的剖面視圖
            BoundingBoxXYZ xYZ = toz.get_BoundingBox(view);
            

            SetViewport(doc, vs_list, viewSheet, xYZ);
            t.Commit();

            // 完成剖面圖紙出圖

            // 切往所建立圖紙之視圖
            try
            {
                uidoc.ActiveView = view;
                TaskDialog.Show("出圖", "完成出圖!!!");
            }
            catch
            {
                TaskDialog.Show("出圖", "未成功出圖!!!");
            }
            
        }

        // 建立元件剖面
        public ViewSection create_section(Document document, Element e)
        {
            using (Transaction transaction = new Transaction(document, "Modify Section Box"))
            {
                transaction.Start();

                // Find a section view type
                IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(document).OfClass(typeof(ViewFamilyType))
                                                              let type = elem as ViewFamilyType
                                                              where type.ViewFamily == ViewFamily.Section
                                                              select type;

                // Create a BoundingBoxXYZ instance centered on wall
                Curve line = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(1, 0, 0)) as Curve;
                Transform curveTransform = line.ComputeDerivatives(0.5, true);
                // using 0.5 and "true" (to specify that the parameter is normalized) 
                // places the transform's origin at the center of the location curve)


                Transform transform = Transform.CreateRotation(new XYZ(1,0,0), 0*Math.PI);
                transform.Origin = (e as Instance).GetTransform().Origin;

                // can use this simplification because wall's "up" is vertical.
                // For a non-vertical situation (such as section through a sloped floor the surface normal would be needed)
                transform.BasisZ = (e as Instance).GetTransform().BasisX;
                transform.BasisY = (e as Instance).GetTransform().BasisZ;
                transform.BasisX = (e as Instance).GetTransform().BasisY;

                // Min & Max X values (-10 & 10) define the section line length on each side of the wall
                // Max Y (12) is the height of the section box// Max Z (5) is the far clip offset
                //調整剖面框大小
                BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                sectionBox.Transform = transform;
                sectionBox.Max = new XYZ(15, 18, 5);
                sectionBox.Min = new XYZ(-15, -18, -5);



                // Create a new view section.
                ViewSection viewSection = ViewSection.CreateSection(document, viewFamilyTypes.First().Id, sectionBox);

                document.Delete(viewSection.Id);
                viewSection = ViewSection.CreateSection(document, viewFamilyTypes.First().Id, sectionBox);

                transaction.Commit();
                return viewSection;
            }
        }

        // 依照不同的剖面數量排版圖紙上的剖面
        public void SetViewport(Document doc, IList<ViewSection>  vs_list, ViewSheet viewSheet, BoundingBoxXYZ xYZ)
        {
            //剖面圖
            IList<ViewSection> viewSection = vs_list;//new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Where(y => y.Name == sectionName[i]).First();


            IList<XYZ> locations = new List<XYZ>();
            if (vs_list.Count == 1)
            {
                locations.Add(new XYZ((xYZ.Max.X + xYZ.Min.X) / 2, (xYZ.Max.Y + xYZ.Min.Y) / 2, (xYZ.Max.Z + xYZ.Min.Z) / 2));
            }
            else if(vs_list.Count == 2)
            {
                locations.Add(new XYZ((xYZ.Max.X*0.25 + xYZ.Min.X*0.75), (xYZ.Max.Y + xYZ.Min.Y) / 2, (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X*0.65 + xYZ.Min.X*0.35), (xYZ.Max.Y + xYZ.Min.Y) / 2, (xYZ.Max.Z + xYZ.Min.Z) / 2));
            }
            else if(vs_list.Count > 2 && vs_list.Count < 5)
            {
                locations.Add(new XYZ((xYZ.Max.X * 0.25 + xYZ.Min.X * 0.75), (xYZ.Max.Y * 0.75 + xYZ.Min.Y * 0.25), (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X * 0.65 + xYZ.Min.X * 0.35), (xYZ.Max.Y * 0.75 + xYZ.Min.Y * 0.25), (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X * 0.25 + xYZ.Min.X * 0.75), (xYZ.Max.Y * 0.25 + xYZ.Min.Y * 0.75), (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X * 0.65 + xYZ.Min.X * 0.35), (xYZ.Max.Y * 0.25 + xYZ.Min.Y * 0.75), (xYZ.Max.Z + xYZ.Min.Z) / 2));
            }else if(vs_list.Count > 4 && vs_list.Count < 7)
            {
                locations.Add(new XYZ((xYZ.Max.X * 0.225 + xYZ.Min.X * 0.775), (xYZ.Max.Y * 0.75 + xYZ.Min.Y * 0.25), (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X * 0.45 + xYZ.Min.X * 0.55), (xYZ.Max.Y * 0.75 + xYZ.Min.Y * 0.25), (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X * 0.675 + xYZ.Min.X * 0.325), (xYZ.Max.Y * 0.75 + xYZ.Min.Y * 0.25), (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X * 0.225 + xYZ.Min.X * 0.775), (xYZ.Max.Y * 0.25 + xYZ.Min.Y * 0.75), (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X * 0.45 + xYZ.Min.X * 0.55), (xYZ.Max.Y * 0.25 + xYZ.Min.Y * 0.75), (xYZ.Max.Z + xYZ.Min.Z) / 2));
                locations.Add(new XYZ((xYZ.Max.X * 0.675 + xYZ.Min.X * 0.325), (xYZ.Max.Y * 0.25 + xYZ.Min.Y * 0.75), (xYZ.Max.Z + xYZ.Min.Z) / 2));
            }

            try
            {
                for (int i = 0; i < vs_list.Count; i++)
                {
                    set_viewsheet(doc, viewSheet.Id, vs_list[i], locations[i]);
                }
            }
            catch
            {
                TaskDialog.Show("message", "最多選6個");
                doc.Delete(viewSheet.Id);
            }
        }

        // 放置剖面
        public void set_viewsheet(Document doc, ElementId view_sheet_id, ViewSection viewSection, XYZ location) 
        {

            ViewSection depentviewsection = doc.GetElement(viewSection.Duplicate(ViewDuplicateOption.AsDependent)) as ViewSection;

            //剖面圖參數
            depentviewsection.get_Parameter(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC).Set("自訂");
            depentviewsection.Scale = 1 * 50;

            Viewport viewport1 = Viewport.Create(doc, view_sheet_id, viewSection.Id, location);
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
