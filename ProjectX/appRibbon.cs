using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Reflection;
using System.Windows.Media.Imaging;
using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ProjectX
{
    internal class appRibbon : IExternalApplication
    {

        public Result OnStartup(UIControlledApplication a)
        {

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string assemblyDir = Path.GetDirectoryName(assemblyPath);

            a.ControlledApplication.DocumentOpening += new EventHandler<Autodesk.Revit.DB.Events.DocumentOpeningEventArgs>(AppDocOpening);
            a.ControlledApplication.DocumentOpened += new EventHandler<Autodesk.Revit.DB.Events.DocumentOpenedEventArgs>(AppDocOpened);

            a.ControlledApplication.DocumentSynchronizingWithCentral += new EventHandler<Autodesk.Revit.DB.Events.DocumentSynchronizingWithCentralEventArgs>(AppDocSynching);
            a.ControlledApplication.DocumentSynchronizedWithCentral += new EventHandler<Autodesk.Revit.DB.Events.DocumentSynchronizedWithCentralEventArgs>(AppDocSynched);

            a.ControlledApplication.DocumentClosing += new EventHandler<Autodesk.Revit.DB.Events.DocumentClosingEventArgs>(AppDocClosing);
            a.ControlledApplication.DocumentClosed += new EventHandler<Autodesk.Revit.DB.Events.DocumentClosedEventArgs>(AppDocClosed);
            try
            {
                // Create a new tab named "dwp Tools"
                a.CreateRibbonTab("Project X");

                // Create first panel "dwp | BIM Help"
                RibbonPanel projectX = a.CreateRibbonPanel("Project X", "X | Project");

                // Add buttons to the first panel
                CreateButton(projectX, "Project X", "Project X", assemblyPath, "ProjectX.CommandProjectX", "Image.png", assemblyDir);
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
        public Result OnShutdown(UIControlledApplication a)
        {
            a.ControlledApplication.DocumentOpening -= new EventHandler<Autodesk.Revit.DB.Events.DocumentOpeningEventArgs>(AppDocOpening);
            a.ControlledApplication.DocumentOpened -= new EventHandler<Autodesk.Revit.DB.Events.DocumentOpenedEventArgs>(AppDocOpened);

            a.ControlledApplication.DocumentSynchronizingWithCentral -= new EventHandler<Autodesk.Revit.DB.Events.DocumentSynchronizingWithCentralEventArgs>(AppDocSynching);
            a.ControlledApplication.DocumentSynchronizedWithCentral -= new EventHandler<Autodesk.Revit.DB.Events.DocumentSynchronizedWithCentralEventArgs>(AppDocSynched);

            a.ControlledApplication.DocumentClosing -= new EventHandler<Autodesk.Revit.DB.Events.DocumentClosingEventArgs>(AppDocClosing);
            a.ControlledApplication.DocumentClosed -= new EventHandler<Autodesk.Revit.DB.Events.DocumentClosedEventArgs>(AppDocClosed);

            return Result.Succeeded;
        }

        private void AppDocOpening(object sender, Autodesk.Revit.DB.Events.DocumentOpeningEventArgs e)
        {
            TaskDialog taskDialog = new TaskDialog("AppDocOpening");
            taskDialog.MainInstruction = "AppDocOpening";
            taskDialog.Show();
        }

        private void AppDocOpened(object sender, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs e)
        {

            TaskDialog taskDialog = new TaskDialog("AppDocOpened");
            taskDialog.MainInstruction = "AppDocOpened";
            taskDialog.Show();
            

        }
        private void AppDocSynching(object sender, Autodesk.Revit.DB.Events.DocumentSynchronizingWithCentralEventArgs e)
        {
            TaskDialog taskDialog = new TaskDialog("AppDocSynching");
            taskDialog.MainInstruction = "AppDocSynching";
            taskDialog.Show();
        }
        private void AppDocSynched(object sender, Autodesk.Revit.DB.Events.DocumentSynchronizedWithCentralEventArgs e)
        {
            TaskDialog taskDialog = new TaskDialog("AppDocSynched");
            taskDialog.MainInstruction = "AppDocSynched";
            taskDialog.Show();
        }
            private void AppDocClosing(object sender, Autodesk.Revit.DB.Events.DocumentClosingEventArgs e)
        {
            TaskDialog taskDialog = new TaskDialog("AppDocClosing");
            taskDialog.MainInstruction = "AppDocClosing";
            taskDialog.Show();
        }

        private void AppDocClosed(object sender, Autodesk.Revit.DB.Events.DocumentClosedEventArgs e)
        {
            TaskDialog taskDialog = new TaskDialog("AppDocClosed");
            taskDialog.MainInstruction = "AppDocClosed";
            taskDialog.Show();
        }
    }
}