using System;

namespace ProjectX
{
    public class VariablesData
    {
        public string SessionId { get; set; }
        public DateTime Date { get; set; }
        public string ProjectName { get; set; }
        public string ProjectNumber { get; set; }
        public string ProjectFileName { get; set; }
        public string LogUserName { get; set; }
        public string Action { get; set; }
        public DateTime LogStart { get; set; }
        public DateTime LogEnd { get; set; }
        public TimeSpan LogDuration { get; set; }
        public int Duration { get; set; }
        public string UserComputer { get; set; }
        public string RevitServicePack { get; set; }
        public int Warnings { get; set; }
        public string FileSize { get; set; }
        public string Worksets { get; set; }
        public int LinkedModels { get; set; }
        public int ImportedImages { get; set; }
        public int Views { get; set; }
        public int Sheets { get; set; }
        public int ModelElements { get; set; }
        public int ModelGroups { get; set; }
        public int DetailGroups { get; set; }
        public int DesignOptions { get; set; }
        public string Diroot { get; set; }
    }
}
