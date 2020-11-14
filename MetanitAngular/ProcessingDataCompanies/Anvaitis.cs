using MetanitAngular.Excel;
using MetanitAngular.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static MetanitAngular.Excel.DataStructsForPrintCalls;

namespace MetanitAngular.ProcessingDataCompanies
{
    public class Anvaitis : DefaultCompany
    {
        public Anvaitis(ref List<ProcessedCall> processedCalls, bool DS = false) : base(ref processedCalls, DS)
        {

        }
        public List<CallIncoming> notRecallAnalyze()
        {
            List<CallIncoming> returnCalls = new List<CallIncoming>();
            var Stages = phones.Stages;

            foreach (var call in phones.getPhones())
            {
                FullCall LastCall = new FullCall(call.Value.phoneNumber,
                call.Value.link,
                call.Value.stages.First().Key,
                call.Value.stages.First().Value.First().Date,
                call.Value.stages.First().Value.First().Outgoing,
                call.Value.stages.First().Value.First().comment,
                call.Value.GetManager());
                DateTime datelastcallfirststage = new DateTime();
                DateTime datelastcallsecondstage = new DateTime();
                foreach (var stage in call.Value.stages)
                {
                    foreach (var curCall in stage.Value)
                    {
                        if (curCall.Date > datelastcallfirststage && phones.Stages[stage.Key]==1)
                        {
                            datelastcallfirststage = curCall.Date;
                            LastCall.date = curCall.Date;
                            LastCall.stage = stage.Key;
                            LastCall.outgoing = curCall.Outgoing;
                            LastCall.Comment = curCall.comment;
                        }
                        if (curCall.Date > datelastcallsecondstage && phones.Stages[stage.Key] == 2)
                        {
                            datelastcallfirststage = curCall.Date;
                        }
                    }
                }
                if (datelastcallfirststage != datelastcallsecondstage)
                {
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = call.Value.phoneNumber;
                    AddedCall.Link = call.Value.link;
                    AddedCall.Comment = LastCall.Comment;
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(new CallIncoming(call.Value.phoneNumber, call.Value.link, String.Format("{0:dd.MM.yy}", datelastcallfirststage), LastCall.Comment + "\nДата второго звонка: " + (datelastcallsecondstage.Year >2000 ?  datelastcallsecondstage.ToString("dd.MM.yyyy") : "отсутствует"), call.Value.GetManager(), new ProcessedCall(), call.Value.DealState, call.Value.DateDeal));
                    else
                    {
                        var samecall = InputDoc.getSamePhone(processedCalls, AddedCall);
                        if (samecall.ClientState != null && samecall.ClientState.ToUpper() == "В РАБОТЕ")
                        {
                            returnCalls.Add(new CallIncoming(call.Value.phoneNumber, call.Value.link,
                                String.Format("{0:dd.MM.yy}", datelastcallfirststage),
                                LastCall.Comment + "\nДата второго звонка: " + (datelastcallsecondstage.Year > 2000 ? datelastcallsecondstage.ToString("dd.MM.yyyy") : "отсутствует"), call.Value.GetManager(), samecall, call.Value.DealState, call.Value.DateDeal));
                        }

                    }
                    
                }
            }
            return returnCalls;
        }
    }
}
