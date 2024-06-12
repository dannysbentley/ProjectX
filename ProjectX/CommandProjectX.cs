using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ProjectX
{
    [Transaction(TransactionMode.Manual)]
    internal class CommandProjectX : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp?.ActiveUIDocument;
                Document doc = uidoc?.Document;

                return Result.Succeeded;
            }            
            
            catch (Exception ex)
            {
                // Handle exceptions
                message = "An error occurred: " + ex.Message;
                return Result.Failed;
            }
        }
    }
}
