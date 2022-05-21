using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataStatistics.Main
{
    public record class Config
    {
        public string FileFullPath { get; set; }
        public string DataSheetName { get; set; }
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public string OutputPath { get; set; }
        public string HandingTimeColumnName { get; set; }
        public string SourceColumnName { get; set; }
        public string TypeColumnName { get; set; }
        public string CreaterColumnName { get; set; }
    }
}
