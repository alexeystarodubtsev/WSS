using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;

namespace MetanitAngular.Models
{
    public class FullCall
    {
        public string phoneNumber;
        public string stage;
        public DateTime date;
        public bool outgoing;
        public string Comment;
        public string Link;
        public string Manager;
        public string StateDeal;
        public DateTime DateNextContact;
        public FullCall(string phoneNumber, string link, string stage, DateTime date, bool outgoing, string Comment, string Manager,string StateDeal = "",DateTime dateNextContact=new DateTime())
        {
            this.date = date;
            this.phoneNumber = phoneNumber;
            this.stage = stage;
            this.outgoing = outgoing;
            this.Comment = Comment;
            this.Link = link;
            this.Manager = Manager;
            this.StateDeal = StateDeal;
            this.DateNextContact = dateNextContact;
            
        }
    }
}
