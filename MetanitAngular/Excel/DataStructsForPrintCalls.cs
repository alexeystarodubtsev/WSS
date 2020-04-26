using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Excel
{
    public class DataStructsForPrintCalls
    {
        
        public struct CallIncoming 
        {
            public string phoneNumber;
            public string date;
            public string comment;
            public XLHyperlink Link;
            public string Manager;
            public CallIncoming(string phoneNumber, string Link, string date, string comment, string Manager)
            {
                if (Link != "")
                    this.Link = new XLHyperlink(new Uri(Link));
                else
                    this.Link = null;
                this.phoneNumber = phoneNumber;
                this.date = date;
                this.comment = comment;
                this.Manager = Manager;
            }
        }
        public struct ProcessedCall
        {
            public string Client;
            public string Manager;
            public string Comment;
            public string NoticeCRM;
            public string ClientState;
            public DateTime StartDateAnalyze;
            public string Link;
            
        }

        public struct CallPerWeek
        {
            public string phoneNumber;
            public string FirstWeek;
            public string SecondWeek;
            public string ThirdWeek;
            public string comment;
            public XLHyperlink Link;
            public string Manager;
            public CallPerWeek(string phoneNumber, string Link, string FirstWeek, string SecondWeek, string ThirdWeek, string comment, string Manager)
            {
                this.phoneNumber = phoneNumber;
                this.FirstWeek = FirstWeek;
                this.SecondWeek = SecondWeek;
                this.ThirdWeek = ThirdWeek;
                this.comment = comment;
                if (Link != "")
                    this.Link = new XLHyperlink(new Uri(Link));
                else
                    this.Link = null;
                this.Manager = Manager;
            }
        }

        public struct CallOneStage
        {
            public string phoneNumber;
            public string date;
            public string comment;
            public string qty;
            public string stage;
            public XLHyperlink Link;
            public string Manager;
            public CallOneStage(string phoneNumber, string Link, string date, string comment, string stage, string qty, string Manager)
            {
                this.phoneNumber = phoneNumber;
                this.date = date;
                this.comment = comment;
                this.qty = qty;
                this.stage = stage;
                if (Link != "")
                    this.Link = new XLHyperlink(new Uri(Link));
                else
                    this.Link = null;
                this.Manager = Manager;
            }
        }
        public struct CallPreAgreement
        {
            public string phoneNumber;
            public string date;
            public string comment;
            public string stage;
            public XLHyperlink Link;
            public string Manager;
            public CallPreAgreement(string phoneNumber, string Link, string date, string comment, string stage, string Manager)
            {
                this.phoneNumber = phoneNumber;
                this.date = date;
                this.comment = comment;
                this.stage = stage;
                if (Link != "")
                    this.Link = new XLHyperlink(new Uri(Link));
                else
                    this.Link = null;
                this.Manager = Manager;
            }
        }
    }
}
