using System.Collections.Generic;

namespace RanepaSchedule.Models
{
    public class ScheduleModel
    {
        public string DayOfWeek { get; set; }
        public List<string> Times { get; set; } = new List<string>();
        public List<string> Subjects { get; set; } = new List<string>();

    }
}