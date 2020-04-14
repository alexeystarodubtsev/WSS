using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Models
{
    public class OneCall
    {
        public DateTime Date;
        public bool Outgoing;
        public string comment;
        public OneCall(FullCall call)
        {
            Date = call.date;
            this.Outgoing = call.outgoing;
            comment = call.Comment;
        }
    }
}
