using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KanbanUpdateData.Model
{
    internal class LOWDATA
    {
        public string? Factory { get; set; }
        public string? Item { get; set; }
        public string? Product { get; set; }
        public string? Model { get; set; }
        public string? DeviceOrder { get; set; }
        public string? ProductLine { get; set; }
        public string? DeviceName { get; set; }
        public string? Name { get; set; }
        public string? Description { get; set; }
        public string? Quality { get; set; }
        public string? Time { get; set; }
        public string? Value { get; set; }
        public bool? Defective { get; set; }
        public bool? Activation { get;set; }
        public bool? Throughput { get;set; }
        public bool? Exception { get;set; }
    }
}
