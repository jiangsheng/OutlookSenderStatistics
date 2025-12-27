using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookSenderStatistics
{
    internal class SenderStatistics
    {
        public string? Sender { get; set; }
        public long MailCount{ get; set; }

        public long TotalMailSize { get; set; }

    }
}
