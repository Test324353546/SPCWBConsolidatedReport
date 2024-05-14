using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPCWBConsolidatedReport
{
   public class clsConsolidatedReportData
    {
        public string ProductCode { get; set; }
        public string Grade { get; set; }
      
        public string FilePath { get; set; }
    
        public List <clsChars> clsChars { get; set; }
    }
    public class clsChars
    {
        public string SubgroupNumber { get; set; }
        public string OrderNumber { get; set; }
        public string TraceCategory { get; set; }
        public string TraceId { get; set; }
        public DateTime SGDate { get; set; }
        public string ParameterName { get; set; }
        public string TOLLower { get; set; }
        public string Target { get; set; }
        public string TOLUpper { get; set; }
        public string ActReading { get; set; }
    }
}
