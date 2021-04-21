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
        private bool DS;
        public List<firstCallsToClient> getfirstCallForBelfan()
        {
            return new List<firstCallsToClient>();
        }
        public DefaultCompany(ref List<ProcessedCall> processedCalls,bool DS = false)
        {
            this.processedCalls = processedCalls;
            this.DS = DS;
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

                        if (curCall.Date > LastCall.date || (curCall.Date == LastCall.date && curCall.Outgoing))
                        {
                            LastCall.date = curCall.Date;
                            LastCall.stage = stage.Key;
                            LastCall.outgoing = curCall.Outgoing;
                            LastCall.Comment = curCall.comment;

                        }
                    }
                }
                // !LastCall.outgoing && убрали, так как захотели, чтоб анализировали не только входящие
                if (!call.Value.stages.ContainsKey(AgreementStage))
                {
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = call.Value.phoneNumber;
                    AddedCall.Link = call.Value.link;
                    AddedCall.Comment = LastCall.Comment;
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(new CallIncoming(call.Value.phoneNumber, call.Value.link, String.Format("{0:dd.MM.yy}", LastCall.date), LastCall.Comment, call.Value.GetManager(),new ProcessedCall(), call.Value.DealState, call.Value.DateDeal));
                    else
                    {
                        var samecall = InputDoc.getSamePhone(processedCalls, AddedCall);
                        if (samecall.ClientState != null && samecall.ClientState.ToUpper() == "В РАБОТЕ")
                        {
                            returnCalls.Add(new CallIncoming(call.Value.phoneNumber, call.Value.link, 
                                String.Format("{0:dd.MM.yy}", LastCall.date), 
                                LastCall.Comment, call.Value.GetManager(),samecall, call.Value.DealState, call.Value.DateDeal)); 
                        }

                    }
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

                if (((t1.TotalDays >= 23 && !DS) || (DS && t1.TotalDays > 61))&& !call.Value.stages.ContainsKey(AgreementStage))
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
                    if (t1.TotalDays >= 30 && !DS)
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
                    curCall.DateDeal = "";
                    if (call.Value.DealState.ToUpper() != "В РАБОТЕ" && call.Value.DealState != "")
                    {
                        curCall.DealState = "Закрыт";
                        curCall.NoticeCRM = call.Value.DealState;
                        curCall.DateDeal = call.Value.DateDeal.ToString("dd.MM.yyyy");
                    }
                    if (call.Value.DealState.ToUpper() == "В РАБОТЕ")
                    {
                        curCall.DealState = call.Value.DealState;
                        if (call.Value.DateDeal.Year > 2000)
                        {
                            curCall.DateDeal = call.Value.DateDeal.ToString("dd.MM.yyyy");
                        }
                    }

                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(curCall);
                    else
                    {
                        var samecall = InputDoc.getSamePhone(processedCalls, AddedCall);
                        if (samecall.ClientState != null && samecall.ClientState.ToUpper() == "В РАБОТЕ")
                        {
                            curCall.call = samecall;
                            returnCalls.Add(curCall);
                       
                        }

                    }
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
                
                if (lastStage != AgreementStage && dt.Count > 2)
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
                    List<DateTime> uniqDT = new List<DateTime>();
                    foreach (var call in dt)
                    {
                        if (!uniqDT.Contains(call.Date))
                        {
                            curCall.date = curCall.date + String.Format("{0:dd.MM.yy}", call.Date) + ", ";
                            uniqDT.Add(call.Date);
                        }
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
                    curCall.DateDeal = "";
                    if (phone.Value.DealState.ToUpper() != "В РАБОТЕ" && phone.Value.DealState != "")
                    {
                        curCall.DealState = "Закрыт";
                        curCall.NoticeCRM = phone.Value.DealState;
                        curCall.DateDeal = phone.Value.DateDeal.ToString("dd.MM.yyyy");
                    }
                    if (phone.Value.DealState.ToUpper() == "В РАБОТЕ")
                    {
                        curCall.DealState = phone.Value.DealState;
                        if (phone.Value.DateDeal.Year > 2000)
                        {
                            curCall.DateDeal = phone.Value.DateDeal.ToString("dd.MM.yyyy");
                        }
                    }
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(curCall);
                    else
                    {
                        var samecall = InputDoc.getSamePhone(processedCalls, AddedCall);
                        if (samecall.ClientState != null && samecall.ClientState.ToUpper() == "В РАБОТЕ")
                        {
                            curCall.call = samecall;
                            returnCalls.Add(curCall);
                        }

                    }
                }
            }
            return returnCalls;
        }

        

        public void FillStageDictionary(XLWorkbook wb)
        {
            int i = 1;
            string meet = "Назначена Встреча";
            phones.Stages[meet.ToUpper()] = 50;
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
                    Regex rOpinion =  new Regex("Опросы по качеству", RegexOptions.IgnoreCase);
                    Match mOpinion = rOpinion.Match(page.Name.Trim());
                    m = rx.Match(page.Name.ToUpper().Trim());
                    if (m.Success || mOpinion.Success)
                    {
                        phones.Stages[page.Name.ToUpper().Trim()] = -2;
                        phones.Stages["ОПРОСЫ ПО КАЧЕСТВУ"] = -2;
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
                            rx = new Regex("КП ОТПРАВЛЕН|прайс отправлен", RegexOptions.IgnoreCase);
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
                try
                {
                    if (Stages[stage.Key] > numLastStage)
                    {
                        lastStage = stage.Key;
                        numLastStage = Stages[stage.Key];
                    }
                }
                catch (System.Collections.Generic.KeyNotFoundException)
                {
                    if (stage.Key == "ЗВОНОК ДЛЯ ВЫЯВЛЕНИЯ ЛПР")
                    {
                        if (Stages["ЗВОНОК ДЛЯ ВЫЯВЛЕНИЕ ЛПР"] > numLastStage)
                        {
                            lastStage = stage.Key;
                            numLastStage = Stages["ЗВОНОК ДЛЯ ВЫЯВЛЕНИЕ ЛПР"];
                        }
                    }
                }
            }

            if (lastStage == "")
            {
                return ph.stages.Last().Key;
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
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(new CallPreAgreement(phone.Value.phoneNumber, phone.Value.link, String.Format("{0:dd.MM.yy}", LastCall.Date), LastCall.comment, lastStage, phone.Value.GetManager(), new ProcessedCall(), phone.Value.DealState, phone.Value.DateDeal));
                    else
                    {
                        var samecall = InputDoc.getSamePhone(processedCalls, AddedCall);
                        if (samecall.ClientState != null && samecall.ClientState.ToUpper() == "В РАБОТЕ")
                        {
                            returnCalls.Add(new CallPreAgreement(phone.Value.phoneNumber, phone.Value.link,
                                String.Format("{0:dd.MM.yy}", LastCall.Date),
                                LastCall.comment, lastStage, phone.Value.GetManager(), samecall, phone.Value.DealState, phone.Value.DateDeal));
                        }
                    }
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
                            bool normalDate = false;
                            if (cell.DataType == XLDataType.DateTime)
                            {
                                curDate = cell.GetDateTime();
                                normalDate = true;
                            }
                            else
                            {
                                if (!DateTime.TryParse(cell.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out curDate))
                                {
                                    normalDate = DateTime.TryParse(cell.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out curDate);
                                }
                                else
                                {
                                    normalDate = true;
                                }

                            }
                            string phoneNumber;
                            int corrRow = 5;
                            Match Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            while (!Mcomment.Success)
                            {
                                corrRow++;
                                Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            }
                           while (!(cell.CellBelow().IsEmpty() && cell.CellBelow().CellRight().IsEmpty() && cell.CellBelow().CellBelow().IsEmpty() && cell.CellBelow().CellBelow().CellRight().IsEmpty()))
                                {
                                if (cell.GetValue<string>() != "")
                                {
                                    if (cell.DataType == XLDataType.DateTime)
                                        curDate = cell.GetDateTime();
                                    else
                                    {
                                        if (!DateTime.TryParse(cell.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out curDate))
                                            DateTime.TryParse(cell.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out curDate);

                                    }
                                }
                                phoneNumber = cell.CellBelow().GetValue<string>().ToUpper().Trim();
                                var CellPhoneNumber = cell.CellBelow();
                                string link;
                                if (CellPhoneNumber.HasHyperlink)
                                    link = CellPhoneNumber.GetHyperlink().ExternalAddress.AbsoluteUri;
                                else
                                    link = "";
                                
                                if (link == "")
                                {

                                }

                                if (phoneNumber != "")
                                {
                                    Regex rx = new Regex("ВХОДЯЩ");
                                    Match m = rx.Match(page.Name.ToUpper().Trim());
                                    var exCallSeq = processedCalls.Where(c => (c.Client == phoneNumber && link == "") || (c.Link == link && link != ""));
                                    var exCall = new ProcessedCall();
                                    //exCall.StartDateAnalyze = curDate.AddDays(-1);
                                    if (exCallSeq.Count() > 0)
                                    {
                                        exCall = exCallSeq.First();
                                        //exCall.StartDateAnalyze = curDate.AddDays(-1);
                                    }
                                    else
                                    {
                                        exCall.ClientState = "";
                                        exCall.StartDateAnalyze = DateTime.MinValue;
                                           
                                    }
                                    if ((curDate > exCall.StartDateAnalyze ||
                                        (
                                          exCall.ClientState.ToUpper() == "В РАБОТЕ") &&
                                          exCall.StartDateAnalyze < DateTime.Today.AddDays(1)
                                    ) && normalDate)
                                    {
                                        DateTime DateNext = new DateTime();
                                        var NextContactCell = page.Cell(corrRow + 6, cell.Address.ColumnNumber);
                                        if (NextContactCell.GetString() != "")
                                        {
                                            if (NextContactCell.DataType == XLDataType.DateTime)
                                                DateNext = NextContactCell.GetDateTime();
                                            else
                                            {
                                                if (!DateTime.TryParse(NextContactCell.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out DateNext))
                                                    DateTime.TryParse(NextContactCell.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out DateNext);

                                            }
                                        }
                                        
                                        if (curDate > new DateTime(2020, 5, 5))
                                            phones.AddCall(new FullCall(phoneNumber, link, page.Name.ToUpper().Trim(), curDate, !m.Success, page.Cell(corrRow, cell.Address.ColumnNumber).GetString(), Manager, page.Cell(corrRow + 5, cell.Address.ColumnNumber).GetString(), DateNext));
                                        else
                                        {
                                            phones.AddCall(new FullCall(phoneNumber, link, page.Name.ToUpper().Trim(), curDate, !m.Success, page.Cell(corrRow, cell.Address.ColumnNumber).GetString(), Manager));
                                        }


                                        
                                    }
                                }

                                cell = cell.CellRight();
                            }
                            phones.CleanSuccess(ref processedCalls);

                        }
                    }
                }
            }
        }

    }
}
