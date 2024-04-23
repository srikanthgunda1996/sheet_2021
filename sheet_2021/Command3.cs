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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using Autodesk.Revit.DB.Mechanical;
using System.Windows.Shapes;
using static System.Windows.Forms.LinkLabel;
#endregion

namespace sheet_2021
{
    [Transaction(TransactionMode.Manual)]
    public class Command3 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;
            Transaction t = new Transaction(doc);
            // Your code goes here
            t.Start("Schedules");
            FilteredElementCollector doors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors);
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> sheets = collector.OfCategory(BuiltInCategory.OST_Sheets).ToElements();
            ElementId catgid = new ElementId(BuiltInCategory.OST_Doors);
            foreach (Element sheetElement in sheets)
            {
                ViewSheet sheet = sheetElement as ViewSheet;
                foreach (Element curDoor in doors)
                {
                    if (sheet.Name.Contains("SUPPLY OVERALL PLAN"))
                    {
                        ViewSchedule doorschedule = ViewSchedule.CreateSchedule(doc, catgid);
                        Parameter mark = curDoor.get_Parameter(BuiltInParameter.DOOR_NUMBER);
                        Parameter nam = curDoor.get_Parameter(BuiltInParameter.DOOR_WIDTH);
                        Parameter mark1 = curDoor.LookupParameter("Mark");
                        Parameter nam1 = curDoor.LookupParameter("Name");

                        //ScheduleField schmark = doorschedule.Definition.AddField(ScheduleFieldType.Instance, BuiltInParameter.DOOR_NUMBER.id);
                        //ScheduleField schnam = doorschedule.Definition.AddField(ScheduleFieldType.ElementType, nam.Id);
                        doorschedule.Name = "Door Schedule";

                        //Viewport scheduleViewport = Viewport.Create(doc, sheet.Id, doorschedule.Id, new XYZ(0, 0, 0));
                        break;
                    }
                }
            }

            t.Commit();
            t.Dispose();
            

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
    }
}
