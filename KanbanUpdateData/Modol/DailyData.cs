using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KanbanUpdateData.Model
{
    internal class DailyData
    {
        public string? DeviceName { get; set; }
        public string? Line { get; set; }
        public string? Date { get; set; }
        public string? CAP { get; set; }
        public string? StopRunTime { get; set; }
        public string? WorkCode { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? AllSum { get; set; }
        public string? Availability { get; set; }
        public string? YieId { get; set; }
        public string? Performance { get; set; }
    }
}
