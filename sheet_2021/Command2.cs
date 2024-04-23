#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.CSharp;
using excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Forms = System.Windows.Forms;
using Autodesk.Revit.DB.Structure;
using System.Text.RegularExpressions;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using Autodesk.Revit.DB.Mechanical;
using System.Windows.Shapes;
using static System.Windows.Forms.LinkLabel;
#endregion

namespace sheet_2021
{
    [Transaction(TransactionMode.Manual)]
    public class Command2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here


            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select Excel File for Views creation";
            dialog.Filter = "Excel files | *.xlsx;*.xls;*.xlsm";
            dialog.Multiselect = false;



            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string file = dialog.FileName;

                excel.Application exc = new excel.Application();
                excel.Workbook wrkbook = exc.Workbooks.Open(file);
                excel.Worksheet wrksheet = wrkbook.Worksheets[1];
                excel.Range rnge = wrksheet.UsedRange;

                int row = rnge.Rows.Count;
                int colmn = rnge.Columns.Count;

                List<List<string>> exceldata = new List<List<string>>();

                for (int i = 1; i <= row; i++)
                {
                    List<string> rowdata = new List<string>();
                    for (int j = 1; j <= colmn; j++)
                    {
                        string cellcontent = wrksheet.Cells[i, j].Value.ToString();
                        rowdata.Add(cellcontent);
                    }
                    exceldata.Add(rowdata);     
                }

                for (int m = 0; m < row; m++)
                {
                    string viewName = exceldata[m][0].ToString();
                    string levelName = exceldata[m][1].ToString();
                    string viewType = exceldata[m][2].ToString();

                    ViewFamilyType viewFamilyType = null;

                    Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>()
                        .FirstOrDefault(l => l.Name.Equals(levelName, StringComparison.OrdinalIgnoreCase));

                    


                    Transaction t = new Transaction(doc);
                    t.Start("Creating Views");

                    {
                        if (level != null)
                        {
                            if (viewType == "FLOOR PLAN")
                            {
                                viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                                                .FirstOrDefault(x => x.ViewFamily == ViewFamily.FloorPlan);

                                ViewPlan planView = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                                planView.Name = viewName;
                                planView.Scale = 150;
                                //planView.ScopeBox = scopeBox.Id;


                            }
                            if (viewType == "REFLECTED CEILING PLAN")
                            {
                                viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                                                .FirstOrDefault(x => x.ViewFamily == ViewFamily.CeilingPlan);
                                ViewPlan planView = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                                planView.Name = viewName;
                                planView.Scale = 150;
                            }
                            if (viewType == "3D VIEWS")
                            {
                                //viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                                //                .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);
                                //View3D view3D = View3D.CreateIsometric(doc, viewFamilyType.Id);
                                View3D originalView3D = Find3DViewByName(doc,"Default");
                                if (originalView3D != null)
                                {
                                    ElementId duplicatedViewId = originalView3D.Duplicate(ViewDuplicateOption.WithDetailing);
                                    BoundingBoxXYZ sectionBox = originalView3D.GetSectionBox();
                                    BoundingBoxXYZ newsectionBox = new BoundingBoxXYZ();
                                    newsectionBox.Min = new XYZ(sectionBox.Min.X, sectionBox.Min.Y, sectionBox.Min.Z + level.Elevation);
                                    newsectionBox.Max = new XYZ(sectionBox.Max.X, sectionBox.Max.Y, sectionBox.Max.Z+level.Elevation);
                                    

                                    View3D duplicatedView3D = doc.GetElement(duplicatedViewId) as View3D;

                                    //duplicatedView3D.IsSectionBoxActive = true;
                                    duplicatedView3D.SetSectionBox(newsectionBox);
                                    duplicatedView3D.Name = viewName;
                                    duplicatedView3D.Scale = 150;
                                }

                                else
                                {
                                    TaskDialog.Show("error", "No Default 3d view found");
                                }
                       

                                //SetSectionBoxBetweenLevels(doc, view3D, level);

                            }
                            //if (viewType == "DETAILED DRAWINGS")
                            //{
                            //    viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                            //                    .FirstOrDefault(x => x.ViewFamily == ViewFamily.Detail);
                            //    ViewPlan detailedView = ViewPlan.CreateDetail(doc, viewFamilyType.Id);
                            //}
                            if (viewType == "SECTION")
                            {
                                viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                                                .FirstOrDefault(x => x.ViewFamily == ViewFamily.Section);
                                BoundingBoxXYZ viewDirection = new BoundingBoxXYZ();
                                ViewSection sectionView = ViewSection.CreateSection(doc, viewFamilyType.Id, viewDirection);
                                sectionView.Name = viewName;
                                sectionView.Scale = 150;
                            }
                            //if (viewType == "ELEVATIONS")
                            //{
                            //    viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                            //    .FirstOrDefault(x => x.ViewFamily == ViewFamily.Section);
                            //    BoundingBoxXYZ viewDirection = new BoundingBoxXYZ();
                            //    ViewSection elevationView = ViewSection.CreateElevation(doc, viewFamilyType.Id, viewDirection);
                            //    elevationView.Name = viewName;
                            //}

                        }
                        else
                        {
                            // Handle the case where the specified level does not exist
                            TaskDialog.Show("Error", "The specified level does not exist.");
                        }
                        t.Commit();
                        t.Dispose();
                    }
                    


                }

                }
                return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }

        public void SetSectionBoxBetweenLevels(Document doc, View3D view3D, Level level1)
        {
            // Activate the section box
            view3D.IsSectionBoxActive = true;

            // Get the elevations of the levels
            double elevation1 = level1.Elevation;
            

            // Calculate the minimum and maximum Z coordinates for the section box
            double minZ = Math.Min(elevation1, elevation1+3600);
            double maxZ = Math.Max(elevation1, elevation1 + 3600);

            // Get the current section box bounds
            BoundingBoxXYZ sectionBox = view3D.GetSectionBox();



            // Update the section box Z coordinates
            sectionBox.Min = new XYZ(sectionBox.Min.X, sectionBox.Min.Y, minZ);
            sectionBox.Max = new XYZ(sectionBox.Max.X, sectionBox.Max.Y, maxZ);

            // Apply the section box to the view
            view3D.SetSectionBox(sectionBox);

            // Refresh the view
            view3D.Document.Regenerate();
            //view3D.Document.RegenerateIfNeeded();
        }

        public View3D Find3DViewByName(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> views = collector.OfClass(typeof(View3D)).ToElements();

            // Find the 3D view with the specified name
            foreach (Element viewElement in views)
            {
                View3D existingview3D = viewElement as View3D;
                if (existingview3D != null && existingview3D.Name.Contains(viewName))
                {
                    return existingview3D; // Return the 3D view with the specified name
                }   

            }

            return null;
        }
    }
}
