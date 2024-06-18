using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ProjectX
{
    [Transaction(TransactionMode.Manual)]
    public class CommandProjectX : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp?.ActiveUIDocument;
                Document doc = uidoc?.Document;

                if (doc == null)
                {
                    message = "No active document found.";
                    return Result.Failed;
                }

                // Check if "diroot" add-in assembly is loaded
                VariablesData data = CollectInformation(uiapp, uidoc, doc);

                // Display collected information in a dialog
                TaskDialog taskDialog = new TaskDialog("Information");
                taskDialog.MainInstruction = "Revit Project Information";
                taskDialog.MainContent = string.Join("\n", data);
                taskDialog.Show();

                // Optionally, log the information to a file or database
                //LogInformationToFile(infoLines);

                Excel excel = new Excel();
                excel.ExportToExcel(data);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }


        public VariablesData CollectInformation(UIApplication uiapp, UIDocument uidoc, Document doc)
        {
            VariablesData data = new VariablesData();

            try
            {
                data.SessionId = Guid.NewGuid().ToString();
                data.Date = DateTime.Now;
                data.ProjectName = doc.Title;
                data.ProjectNumber = doc.ProjectInformation?.Number ?? "N/A";
                data.ProjectFileName = doc.PathName;
                data.LogUserName = Environment.UserName;
                data.Action = "Opened";
                data.LogStart = DateTime.Now;
                data.Duration = 30;
                data.UserComputer = Environment.MachineName;
                data.RevitServicePack = uiapp.Application.VersionBuild;
                data.Warnings = doc.GetWarnings().Count;
                data.FileSize = GetFileSize(doc.PathName);

                // Worksets count
                try
                {
                    int worksetCount = new FilteredElementCollector(doc)
                        .OfClass(typeof(Workset))
                        .GetElementCount();
                    data.Worksets = worksetCount.ToString();
                }
                catch (Exception ex)
                {
                    data.Worksets = "Error: " + ex.Message;
                }

                // Other counts
                data.LinkedModels = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_RvtLinks)
                    .GetElementCount();

                data.ImportedImages = new FilteredElementCollector(doc)
                    .OfClass(typeof(ImageType))
                    .WhereElementIsNotElementType()
                    .GetElementCount();

                data.Views = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .WhereElementIsNotElementType()
                    .GetElementCount();

                data.Sheets = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .WhereElementIsNotElementType()
                    .GetElementCount();

                // Aggregate multiple built-in categories to represent model elements
                List<BuiltInCategory> modelCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Columns,
                    BuiltInCategory.OST_Doors,
                    BuiltInCategory.OST_Windows,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_Stairs
                };

                int modelElementsCount = modelCategories.Sum(category =>
                    new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .WhereElementIsNotElementType()
                    .GetElementCount());

                data.ModelElements = modelElementsCount;

                data.ModelGroups = new FilteredElementCollector(doc)
                    .OfClass(typeof(Group))
                    .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                    .GetElementCount();

                data.DetailGroups = new FilteredElementCollector(doc)
                    .OfClass(typeof(Group))
                    .OfCategory(BuiltInCategory.OST_IOSDetailGroups)
                    .GetElementCount();

                data.DesignOptions = new FilteredElementCollector(doc)
                    .OfClass(typeof(DesignOption))
                    .GetElementCount();

                string haveDiroots = null;
                try
                {
                    // Check if "diroot" add-in assembly is loaded
                    Assembly dirootAssembly = IsDiRootAssemblyLoaded();
                    haveDiroots = dirootAssembly != null ? "Yes" : "No";
                    data.Diroot = haveDiroots;
                }
                catch (Exception ex)
                {
                    // Handle exception
                }

                return data;
            }
            catch (Exception ex)
            {
                // Handle exception
                return null; // Return null in case of exception
            }
        }

        private string GetFileSize(string filePath)
        {
            if (File.Exists(filePath))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                return $"{fileInfo.Length / 1024 / 1024} MB";
            }
            return "Unknown";
        }

        private Assembly IsDiRootAssemblyLoaded()
        {
            Assembly[] loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();
            string dirootAssemblyName = "DiRootAssemblyName"; // Replace with the actual assembly name of diroot add-in

            foreach (Assembly assembly in loadedAssemblies)
            {
                if (assembly.FullName.Contains(dirootAssemblyName))
                {
                    return assembly;
                }
            }

            return null;
        }
    }
}