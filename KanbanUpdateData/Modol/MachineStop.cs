using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KanbanUpdateData.Model
{
    internal class MachineStop
    {
        public string? DeviceName { get; set; }
        public string? Date { get; set; }
        public string? WorkCode { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public double SumTime { get; set; }
    }
}
