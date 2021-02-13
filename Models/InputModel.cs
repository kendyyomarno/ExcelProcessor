using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelProcessor.Models
{
    public class InputModel
    {
        public string Directory { get; set; }
        public string SearchString { get; set; }
        public string MonthDisplay { get; set; }
        public string Product { get; set; }
        public string ProcessFormula { get; set; }
        public string OutputSuffix { get; set; }
        public AuditColumns auditColumns { get; set; }
        public List<ProcessResult> ProcessResult { get; set; }
    }

    public class ProcessResult 
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }
}
