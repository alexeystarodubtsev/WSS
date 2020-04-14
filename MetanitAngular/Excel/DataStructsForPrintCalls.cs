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

            public CallIncoming(string phoneNumber, string date, string comment)
            {
                this.phoneNumber = phoneNumber;
                this.date = date;
                this.comment = comment;
            }
        }

        public struct CallPerWeek
        {
            public string phoneNumber;
            public string FirstWeek;
            public string SecondWeek;
            public string ThirdWeek;
            public string comment;

            public CallPerWeek(string phoneNumber, string FirstWeek, string SecondWeek, string ThirdWeek, string comment)
            {
                this.phoneNumber = phoneNumber;
                this.FirstWeek = FirstWeek;
                this.SecondWeek = SecondWeek;
                this.ThirdWeek = ThirdWeek;
                this.comment = comment;
            }
        }

        public struct CallOneStage
        {
            public string phoneNumber;
            public string date;
            public string comment;
            public string qty;
            public string stage;

            public CallOneStage(string phoneNumber, string date, string comment, string stage, string qty)
            {
                this.phoneNumber = phoneNumber;
                this.date = date;
                this.comment = comment;
                this.qty = qty;
                this.stage = stage;
            }
        }
        public struct CallPreAgreement
        {
            public string phoneNumber;
            public string date;
            public string comment;
            public string stage;

            public CallPreAgreement(string phoneNumber, string date, string comment, string stage)
            {
                this.phoneNumber = phoneNumber;
                this.date = date;
                this.comment = comment;
                this.stage = stage;
            }
        }
    }
}
