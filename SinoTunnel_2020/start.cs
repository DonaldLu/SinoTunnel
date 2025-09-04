using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace SinoTunnel_2020
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class start : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Form1 form1 = new Form1(commandData.Application.ActiveUIDocument);
            form1.Show();

            return Result.Succeeded;
        }
    }
}
