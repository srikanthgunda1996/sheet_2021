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
#endregion

namespace sheet_2021
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here

            Transaction t = new Transaction(doc);
            t.Start("Creating Sheets");

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select Excel File";
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
                FilteredElementCollector sheetcollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType();
                FilteredElementCollector viewcollector = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan));
                FilteredElementCollector threedviewcollector = new FilteredElementCollector(doc).OfClass(typeof(View3D));
                FilteredElementCollector sectionviewcollector = new FilteredElementCollector(doc).OfClass(typeof(ViewSection));

                for (int m = 0; m< row; m++)
                {
                    string viewname = exceldata[m][0].ToString();
                    foreach(ViewPlan veiwPlan in viewcollector) 
                    { 
                        if(viewname == veiwPlan.Name)
                        {
                            Element sheetelement = sheetcollector.ElementAt(0);
                            ViewSheet newsheet = ViewSheet.Create(doc, sheetelement.Id);
                            //sheetelement.
                            LocationPoint viewLocation = veiwPlan.Location as LocationPoint;
                            //XYZ viewLocationPoint = viewLocation.Point;
                            //LocationPoint curloc = newsheet.Location as LocationPoint;
                            Viewport.Create(doc, newsheet.Id, veiwPlan.Id, new XYZ(1.802, 3.768, 0));
                            //Viewport.Create(doc, newsheet.Id, veiwPlan.Id, viewLocationPoint);
                            newsheet.Name = exceldata[m][1].ToString();
                            newsheet.SheetNumber = exceldata[m][2].ToString();
                            String subtitleval = exceldata[m][3].ToString();
                            if (subtitleval != "Null") { setparametervalue(newsheet as Element, "Sub Title", subtitleval); }

                        }
                    }

                    foreach (View3D veiwPlan in threedviewcollector)
                    {
                        if (viewname == veiwPlan.Name)
                        {
                            Element sheetelement = sheetcollector.ElementAt(0);
                            ViewSheet newsheet = ViewSheet.Create(doc, sheetelement.Id);
                            //sheetelement.
                            LocationPoint viewLocation = veiwPlan.Location as LocationPoint;
                            //XYZ viewLocationPoint = viewLocation.Point;
                            //LocationPoint curloc = newsheet.Location as LocationPoint;
                            Viewport.Create(doc, newsheet.Id, veiwPlan.Id, new XYZ(0, 0, 0));
                            //Viewport.Create(doc, newsheet.Id, veiwPlan.Id, viewLocationPoint);
                            newsheet.Name = viewname;
                            newsheet.SheetNumber = exceldata[m][2].ToString();
                            String subtitleval = exceldata[m][3].ToString();
                            if (subtitleval != "Null") { setparametervalue(newsheet as Element, "Sub Title", subtitleval); }
                        }
                    }

                    foreach (ViewSection veiwPlan in sectionviewcollector)
                    {
                        if (viewname == veiwPlan.Name)
                        {
                            Element sheetelement = sheetcollector.ElementAt(0);
                            ViewSheet newsheet = ViewSheet.Create(doc, sheetelement.Id);
                            //sheetelement.
                            LocationPoint viewLocation = veiwPlan.Location as LocationPoint;
                            //XYZ viewLocationPoint = viewLocation.Point;
                            //LocationPoint curloc = newsheet.Location as LocationPoint;
                            Viewport.Create(doc, newsheet.Id, veiwPlan.Id, new XYZ(0, 0, 0));
                            //Viewport.Create(doc, newsheet.Id, veiwPlan.Id, viewLocationPoint);
                            newsheet.Name = viewname;
                            newsheet.SheetNumber = exceldata[m][2].ToString();
                            String subtitleval = exceldata[m][3].ToString();
                            if (subtitleval != "Null") { setparametervalue(newsheet, "Sub Title", subtitleval); }

                        }
                    }

                }


            }
            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }

        internal void setparametervalue(Element e, string paramname, string value)
        {
            IList<Parameter> parameters = e.GetParameters(paramname);
            Parameter parameter = parameters.First();
            parameter.Set(value);
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
