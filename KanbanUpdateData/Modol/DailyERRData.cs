using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KanbanUpdateData.Model
{
    internal class DailyERRData
    {
        public string? DeviceName { get; set; }
        public string? WorkCode { get; set; }
        public string? Line { get; set; }
        public string? Date { get; set; }
        public string? Time { set; get; }
        public string? Type { get; set; }
        public string? Name { get; set; }
        public int? Count { get; set; }
    }
}
