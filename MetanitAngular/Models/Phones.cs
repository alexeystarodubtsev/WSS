using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static MetanitAngular.Excel.DataStructsForPrintCalls;

namespace MetanitAngular.Models
{
    public class Phones
    {
        Dictionary<string, Phone> Calls;
        public Dictionary<string, int> Stages { get; set; }

        public void CleanSuccess(ref List<ProcessedCall> processedCalls)
        {
            List<string> toDel = new List<string>();
            foreach (var call in Calls)
            {
                if (Regex.Match(call.Value.DealState, "Успешн",RegexOptions.IgnoreCase).Success)
                {
                    toDel.Add(call.Key);
                    ProcessedCall procCall = new ProcessedCall();
                    procCall.Client = call.Value.phoneNumber;
                    procCall.Manager = call.Value.GetManager();
                    procCall.StartDateAnalyze = call.Value.DateDeal;
                    procCall.NoticeCRM = call.Value.DealState;
                    procCall.ClientState = "Закрыт";
                    procCall.Link = call.Value.link;
                    procCall.Comment = "";
                    processedCalls.RemoveAll(c => (c.Client == call.Value.phoneNumber && call.Value.link == "") || (c.Link == call.Value.link && call.Value.link != ""));
                    processedCalls.Add(procCall);
                }
                if (call.Value.DealState.ToUpper().Trim() == "В РАБОТЕ" && call.Value.DateDeal > DateTime.Today.AddDays(-1))
                {
                    toDel.Add(call.Key);
                    ProcessedCall procCall = new ProcessedCall();
                    procCall.Client = call.Value.phoneNumber;
                    procCall.Manager = call.Value.GetManager();
                    procCall.StartDateAnalyze = call.Value.DateDeal;
                    procCall.NoticeCRM = call.Value.DealState;
                    procCall.ClientState = "В работе";
                    procCall.Link = call.Value.link;
                    procCall.Comment = "";
                    processedCalls.RemoveAll(c => (c.Client == call.Value.phoneNumber && call.Value.link == "") || (c.Link == call.Value.link && call.Value.link != ""));
                    processedCalls.Add(procCall);
                }
            }
            toDel.ForEach(k => Calls.Remove(k));
        }

        public Phones()
        {
            Calls = new Dictionary<string, Phone>();
            Stages = new Dictionary<string, int>();
        }
        public void AddCall(FullCall call)
        {
            string phoneKey = "";
            if (call.Link != "")
            {
                phoneKey = call.Link;
                
            }
            else
            {
                phoneKey = call.phoneNumber;
            }
            if (!Calls.ContainsKey(phoneKey))
            {
                Calls[phoneKey] = new Phone(call.Link, call.phoneNumber);
            }
            Calls[phoneKey].AddCall(call);

        }
        public Dictionary<string, Phone> getPhones()
        {
            return Calls;
        }
        
    }
}
