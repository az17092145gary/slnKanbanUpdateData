using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KanbanUpdateData.Model
{
    internal class TempData
    {
        public string? DeviceName { get; set; }
        public string? DeviceOrder { get; set; }
        public string? Product { get; set; }
        public string? Factory { get; set; }
        public string? State { get; set; }
        public string? Model { get; set; }
        public bool? Defective { get; set; }
        public bool? Activation { get; set; }
        public bool? Throughput { get; set; }
        public bool? Exception { get; set; }
        public string? Item { get; set; }
        public string? Line { get; set; }
        public string? WorkCode { get; set; }
        public string? MemCode { get; set; }
        public bool? QIMSuperMode { get; set; }
        public bool? DMISuperMode { get; set; }
        public bool? MTCSuperMode { get; set; }
        public string? Date { get; set; }
        public string? ModelStartTime { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public double? SumTime { get; set; }
        public int? FirstSum { get; set; }
        public int? Sum { get; set; }
        public int? NGSum { get; set; }
        public string? RunState { get; set; }
        public string? Quality { get; set; }
        public string? RunStartTime { get; set; }
        public string? RunEndTime { get; set; }
        public string? ERRCountTime { get; set; }

        public double RunSumStopTime { get; set; }
        public double NonSumStopTime { get; set; }
        public List<DailyERRData> ERRAndPATCountList { get; set; }
        public List<MachineStop>? StopTenUP { get; set; }
        public List<MachineStop>? StopTenDown { get; set; }
    }
}
