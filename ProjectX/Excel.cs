using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;

namespace ProjectX
{
    internal class Excel
    {
        public void ExportToExcel(VariablesData data)
        {
            List<string> infoLines = FormatInfo(data);
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string assemblyDir = Path.GetDirectoryName(assemblyPath);
            string projectDir = Directory.GetParent(Directory.GetParent(assemblyDir).FullName).FullName;

            string filePath = Path.Combine(projectDir, "InfoLines.xlsx");
            Workbook workbook;
            Worksheet worksheet;

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            if (File.Exists(filePath))
            {
                workbook = excelApp.Workbooks.Open(filePath);
                worksheet = workbook.ActiveSheet;
            }
            else
            {
                workbook = excelApp.Workbooks.Add();
                worksheet = workbook.ActiveSheet;
            }

            // Find the first empty row
            int nextRow = worksheet.Cells[worksheet.Rows.Count, 1].End(XlDirection.xlUp).Row;

            //Set Topic
            if (nextRow == 1)
            {
                for (int i = 1; i < infoLines.Count; i++)
                {
                    worksheet.Cells[nextRow, i] = infoLines[i].Split(new string[] { ": " }, StringSplitOptions.None)[0];
                }
            }

            // Write data
            for (int i = 1; i < infoLines.Count; i++)
            {
                string[] parts = infoLines[i].Split(new string[] { ": " }, StringSplitOptions.None);
                if (parts.Length >= 2)
                {
                    worksheet.Cells[nextRow + 1, i] = parts[1]; // Information in the second column
                }
                else
                {
                    worksheet.Cells[nextRow + 1, i] = infoLines[i];
                }
            }

            // Auto fit columns
            worksheet.Columns.AutoFit();

            // Save the workbook
            workbook.SaveAs(filePath);

            // Clean up
            workbook.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        private List<string> FormatInfo(VariablesData info)
        {
            return new List<string>
    {
        $"Session ID: {info.SessionId}",
        $"Date: {info.Date}",
        $"Project Name: {info.ProjectName}",
        $"Project Number: {info.ProjectNumber}",
        $"Project File Name: {info.ProjectFileName}",
        $"User Name: {info.LogUserName}",
        $"Action: {info.Action}",
        $"Log Start: {info.LogStart}",
        $"Log End: {info.LogEnd}",
        $"Duration: {info.Duration}",
        $"User Computer: {info.UserComputer}",
        $"Revit Service Pack: {info.RevitServicePack}",
        $"Warnings: {info.Warnings}",
        $"File Size: {info.FileSize}",
        $"Worksets: {info.Worksets}",
        $"Linked Models: {info.LinkedModels}",
        $"Imported Images: {info.ImportedImages}",
        $"Views: {info.Views}",
        $"Sheets: {info.Sheets}",
        $"Model Elements: {info.ModelElements}",
        $"Model Groups: {info.ModelGroups}",
        $"Detail Groups: {info.DetailGroups}",
        $"Design Options: {info.DesignOptions}",
        $"Diroot: {info.Diroot}"
    };
        }

    }
}
