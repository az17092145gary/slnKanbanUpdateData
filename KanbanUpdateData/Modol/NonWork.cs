using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KanbanUpdateData.Model
{
    internal class NonWork
    {
        public string? DeviceName { get; set; }
        public string? WorkCode { get; set; }
        public string? DeviceOrder { get; set; }
        public bool? Defective { get; set; }
        public bool? Activation { get; set; }
        public bool? Throughput { get; set; }
        public bool? Exception { get; set; }
        public string? Item { get; set; }
        public string? Description { get; set; }
        public string? Factory { get; set; }
        public string? Product { get; set; }
        public string? Line { get; set; }
        public string? Date { get; set; }
        public string? Name { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public double SumTime { get; set; }
    }
}
