using Autodesk.Revit.UI;
using System.Reflection;
using System.Windows.Media.Imaging;
using System;
using System.IO;

namespace ProjectX
{
    internal class appRibbon : IExternalApplication
    {

        public Result OnStartup(UIControlledApplication a)
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string assemblyDir = Path.GetDirectoryName(assemblyPath);

            try
            {
                // Create a new tab named "dwp Tools"
                a.CreateRibbonTab("Project X");

                // Create first panel "dwp | BIM Help"
                RibbonPanel dwpHelpPanel = a.CreateRibbonPanel("Project X", "X | Project");

                // Add buttons to the first panel
                CreateButton(dwpHelpPanel, "Project X", "Project X", assemblyPath, "ProjectX.CommandProjectX", "Image.png", assemblyDir);
            }
            catch (Exception ex)
            {
                // Log or handle exceptions
                Console.WriteLine("An error occurred: " + ex.Message);
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private void CreateButton(RibbonPanel panel, string name, string text, string assemblyPath, string className, string imageName, string imagePath)
        {
            PushButtonData buttonData = new PushButtonData(name, text, assemblyPath, className);
            PushButton button = panel.AddItem(buttonData) as PushButton;
            button.ToolTip = text;
            button.LargeImage = new BitmapImage(new Uri(Path.Combine(imagePath, imageName)));
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
