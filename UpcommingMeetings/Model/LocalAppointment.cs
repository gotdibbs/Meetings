using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpcomingMeetings.Model
{
    public class LocalAppointment
    {
        public string Subject { get; set; }
        public string Location { get; set; }
        public string Uri { get; set; }

        public DateTime StartTime { get; set; }
    }
}
