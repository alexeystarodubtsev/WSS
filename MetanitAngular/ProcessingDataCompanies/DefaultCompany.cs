using ClosedXML.Excel;
using MetanitAngular.Models;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static MetanitAngular.Excel.DataStructsForPrintCalls;
using  MetanitAngular.Excel;
namespace MetanitAngular.ProcessingDataCompanies
{
    public class DefaultCompany : ICompany
    {
        protected Phones phones = new Phones();
        protected string AgreementStage = "";
        protected string PreAgreementStage = "";
        protected List<ProcessedCall> processedCalls;

        public DefaultCompany(List<ProcessedCall> processedCalls)
        {
            this.processedCalls = processedCalls;
        }
        public void AddCall(FullCall call)
        {
            phones.AddCall(call);
        }
        public List<CallIncoming> getIncomeWithoutOutGoing()
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
                foreach (var stage in call.Value.stages)
                {
                    foreach (var curCall in stage.Value)
                    {

                        if (curCall.Date > LastCall.date)
                        {
                            LastCall.date = curCall.Date;
                            LastCall.stage = stage.Key;
                            LastCall.outgoing = curCall.Outgoing;
                            LastCall.Comment = curCall.comment;

                        }
                    }
                }
                if (!LastCall.outgoing && !call.Value.stages.ContainsKey(AgreementStage))
                {
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = call.Value.phoneNumber;
                    AddedCall.Link = call.Value.link;
                    AddedCall.Comment = LastCall.Comment;
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(new CallIncoming(call.Value.phoneNumber, call.Value.link, String.Format("{0:dd.MM.yy}", LastCall.date), LastCall.Comment, call.Value.GetManager()));
                }
            }


            return returnCalls;
        }

        public List<CallPerWeek> getCallsPerWeek()
        {
            List<CallPerWeek> returnCalls = new List<CallPerWeek>();

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
                foreach (var stage in call.Value.stages)
                {
                    foreach (var curCall in stage.Value)
                    {

                        if (curCall.Date > LastCall.date)
                        {
                            LastCall.date = curCall.Date;
                            LastCall.stage = stage.Key;
                            LastCall.outgoing = curCall.Outgoing;
                            LastCall.Comment = curCall.comment;

                        }
                    }
                }

                TimeSpan t1 = DateTime.Now.Subtract(LastCall.date);

                if (t1.TotalDays >= 23 && !call.Value.stages.ContainsKey(AgreementStage))
                {
                    CallPerWeek curCall = new CallPerWeek();
                    curCall.FirstWeek = "-";
                    curCall.phoneNumber = call.Value.phoneNumber;
                    curCall.Manager = call.Value.GetManager();
                    if (call.Value.link != "")
                        curCall.Link = new XLHyperlink(new Uri(call.Value.link));
                    curCall.comment = LastCall.Comment;
                    if (!LastCall.outgoing)
                        curCall.comment = curCall.comment + " (Входящий)";
                    if (t1.TotalDays >= 30)
                    {
                        curCall.SecondWeek = "-";
                    }
                    else
                    {
                        curCall.SecondWeek = "+";
                    }
                    //if (t1.TotalDays >= 22)
                    //{
                    //    curCall.ThirdWeek = "-";
                    //}
                    //else
                    //{
                    //    curCall.ThirdWeek = "+";
                    //}
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = curCall.phoneNumber;
                    AddedCall.Link = call.Value.link;
                    AddedCall.Comment = curCall.comment;
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(curCall);
                }
                
            }
            
            return returnCalls;
        }

        public List<CallOneStage> getCallsOneStage()
        {
            List<CallOneStage> returnCalls = new List<CallOneStage>();
            var Stages = phones.Stages;

            foreach (var phone in phones.getPhones() )
            {
                string lastStage = getLastStage(phone.Value, phones.Stages);
                OneCall FirstCall = phone.Value.stages[lastStage].First();
                foreach (var call in phone.Value.stages[lastStage])
                {
                    if (call.Date < FirstCall.Date)
                    {
                        FirstCall = call;
                    }
                }
                List<OneCall> dt = new List<OneCall>();
                dt.Add(FirstCall);
                foreach (var stage in phone.Value.stages)
                {
                    foreach (var call in stage.Value)
                    {
                        if (call.Date > FirstCall.Date)
                        {
                            OneCall addcall = (OneCall)call.Clone();
                            if (stage.Key != lastStage)
                            {
                                addcall.comment = call.comment + " (" + stage.Key + ")";
                            }
                            dt.Add(addcall);
                        }
                    }
                }
                
                if (lastStage != AgreementStage && dt.Count > 1)
                {
                    CallOneStage curCall = new CallOneStage();

                    curCall.phoneNumber = phone.Value.phoneNumber;
                    curCall.Manager = phone.Value.GetManager();
                    if (phone.Value.link != "")
                        curCall.Link = new XLHyperlink(new Uri(phone.Value.link));
                    curCall.qty = dt.Count.ToString();
                    curCall.stage = lastStage;
                    curCall.date = "";
                    string comment = "";
                    DateTime lastDate = dt.First().Date;
                    foreach (var call in dt)
                    {
                        curCall.date = curCall.date + String.Format("{0:dd.MM.yy}", call.Date) + ", ";
                        if (lastDate < call.Date)
                        {
                            comment = call.comment;
                            lastDate = call.Date;
                        }
                          
                    }
                    curCall.date = curCall.date.TrimEnd(' ').Trim(',');
                    curCall.comment = comment;
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = phone.Value.phoneNumber;
                    AddedCall.Link = phone.Value.link;
                    AddedCall.Comment = curCall.comment;
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(curCall);
                }
            }
            return returnCalls;
        }

        

        public void FillStageDictionary(XLWorkbook wb)
        {
            int i = 1;
            foreach (var page in wb.Worksheets)
            {
                Regex rx = new Regex("ВХОДЯЩ");
                Match m = rx.Match(page.Name.ToUpper().Trim());
                if (m.Success)
                {
                    phones.Stages[page.Name.ToUpper().Trim()] = -3;
                }
                else
                {
                    rx = new Regex("УТОЧНЯЮЩ");
                    m = rx.Match(page.Name.ToUpper().Trim());
                    if (m.Success)
                    {
                        phones.Stages[page.Name.ToUpper().Trim()] = -2;
                    }
                    else
                    {
                        rx = new Regex("БЫЛО НЕ УДОБНО");
                        m = rx.Match(page.Name.ToUpper().Trim());
                        if (m.Success)
                        {
                            phones.Stages[page.Name.ToUpper().Trim()] = -1;
                        }
                        else
                        {
                            phones.Stages[page.Name.ToUpper().Trim()] = i;
                            i++;
                            rx = new Regex("ДОГОВОР");
                            m = rx.Match(page.Name.ToUpper().Trim());
                            if (m.Success)
                            {
                                AgreementStage = page.Name.ToUpper().Trim();
                            }
                            rx = new Regex("КП ОТПРАВЛЕН");
                            m = rx.Match(page.Name.ToUpper().Trim());
                            if (m.Success)
                            {
                                PreAgreementStage = page.Name.ToUpper().Trim();
                            }
                        }
                    }
                }

                
            }
        }
        
        string getLastStage(Phone ph, Dictionary<string, int> Stages)
        {
            string lastStage = "";
            int numLastStage = -4;
            foreach (var stage in ph.stages)
            {
                if (Stages[stage.Key] > numLastStage)
                {
                    lastStage = stage.Key;
                    numLastStage = Stages[stage.Key];
                }
            }

            return lastStage;
        }


        public List<CallPreAgreement> getCallsPreAgreement()
        {
            List<CallPreAgreement> returnCalls = new List<CallPreAgreement>();
            var Stages = phones.Stages;

            foreach (var phone in phones.getPhones())
            {
                string lastStage = getLastStage(phone.Value, phones.Stages);
                if (lastStage == PreAgreementStage)
                { 
                    OneCall FirstCall = phone.Value.stages[lastStage].First();
                    foreach (var call in phone.Value.stages[lastStage])
                    {
                        if (call.Date < FirstCall.Date)
                        {
                            FirstCall = call;
                        }
                    }
                    OneCall LastCall = FirstCall;
                    foreach (var stage in phone.Value.stages)
                    {
                        foreach (var call in stage.Value)
                        {
                            if ( LastCall.Date < call.Date)
                            {
                                LastCall = (OneCall)call.Clone();
                                if (stage.Key != lastStage)
                                    LastCall.comment = LastCall.comment + " (" + stage.Key + ") ";
                            }
                        }
                    }
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = phone.Value.phoneNumber;
                    AddedCall.Comment = LastCall.comment;
                    if (!InputDoc.hasPhone(processedCalls,AddedCall))
                      returnCalls.Add(new CallPreAgreement(phone.Value.phoneNumber,phone.Value.link, String.Format("{0:dd.MM.yy}", LastCall.Date), LastCall.comment, lastStage, phone.Value.GetManager()));
                }



            }
            return returnCalls;
        }
        public void ParserCheckLists(IEnumerable<IFormFile> files)
        {
            using (var stream = files.First().OpenReadStream())
            {
                XLWorkbook wb = new XLWorkbook(stream);
                FillStageDictionary(wb);
            }

            foreach (var file in files)
            {
                string Manager = Regex.Match(file.FileName, @"(\w+)").Groups[1].Value;
                using (var stream = file.OpenReadStream())
                {
                    XLWorkbook wb = new XLWorkbook(stream);
                    
                    foreach (var page in wb.Worksheets)
                    {
                        var statisticMatch = Regex.Match(page.Name.ToUpper().Trim(), "СТАТИСТИК");
                        var LastTableMatch = Regex.Match(page.Name.ToUpper().Trim(), "СВОДН");
                        if (!statisticMatch.Success && !LastTableMatch.Success)
                        {

                            IXLCell cell = page.Cell(1, 5);
                            DateTime curDate;
                            DateTime.TryParse(cell.GetValue<string>(), new CultureInfo("ru-RU"), DateTimeStyles.None, out curDate);
                            string phoneNumber;
                            int corrRow = 5;
                            Match Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            while (!Mcomment.Success)
                            {
                                corrRow++;
                                Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            }
                            while (!(cell.IsEmpty() && cell.CellRight().IsEmpty() && !cell.IsMerged()))
                            {
                                if (cell.GetValue<string>() != "")
                                {
                                    DateTime.TryParse(cell.GetValue<string>(), new CultureInfo("ru-RU"), DateTimeStyles.None, out curDate);
                                }
                                phoneNumber = cell.CellBelow().GetValue<string>().ToUpper().Trim();
                                var CellPhoneNumber = cell.CellBelow();
                                string link;
                                if (CellPhoneNumber.HasHyperlink)
                                    link = CellPhoneNumber.GetHyperlink().ExternalAddress.AbsoluteUri;
                                else
                                    link = "";
                                
                                if (phoneNumber != "")
                                {
                                    Regex rx = new Regex("ВХОДЯЩ");
                                    Match m = rx.Match(page.Name.ToUpper().Trim());
                                    var exCallSeq = processedCalls.Where(c => (c.Client == phoneNumber && link == "") || (c.Link == link && link != ""));
                                    var exCall = new ProcessedCall();
                                    exCall.StartDateAnalyze = curDate.AddDays(-1);
                                    if (exCallSeq.Count() > 0)
                                        exCall = exCallSeq.First();
                                    else
                                    {

                                    }
                                    if (curDate >= exCall.StartDateAnalyze ||
                                        (
                                          exCall.ClientState.ToUpper() == "В РАБОТЕ") &&
                                          exCall.StartDateAnalyze < DateTime.Now
                                    )
                                        phones.AddCall(new FullCall(phoneNumber, link, page.Name.ToUpper().Trim(), curDate, !m.Success, page.Cell(corrRow, cell.Address.ColumnNumber).GetString(),Manager));

                                }

                                cell = cell.CellRight();
                            }

                        }
                    }
                }
            }
        }

    }
}
