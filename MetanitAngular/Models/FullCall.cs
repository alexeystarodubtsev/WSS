using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Models
{
    public class FullCall
    {
        public string phoneNumber;
        public string stage;
        public DateTime date;
        public bool outgoing;
        public string Comment;
        public FullCall(string phoneNumber, string stage, DateTime date, bool outgoing, string Comment)
        {
            this.date = date;
            this.phoneNumber = phoneNumber;
            this.stage = stage;
            this.outgoing = outgoing;
            this.Comment = Comment;
            
        }
    }
}
