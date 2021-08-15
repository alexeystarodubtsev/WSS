using ClosedXML.Excel;
using MetanitAngular.Excel;
using MetanitAngular.Models;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static MetanitAngular.Excel.DataStructsForPrintCalls;

namespace MetanitAngular.ProcessingDataCompanies
{
    public class Belfan : ICompany
    {
        Phones phones = new Phones();
        Phones phonesForFirst = new Phones();
        List<ProcessedCall> processedCalls;
        public Belfan(ref List<ProcessedCall> processedCalls)
        {
            this.processedCalls = processedCalls;
        }
        public void AddCall(FullCall call)
        {
            phones.AddCall(call);
        }
        public List<firstCallsToClient> getfirstCallForBelfan()
        {

            var returnCalls = new List<firstCallsToClient>();
            var Stages = phones.Stages;

            foreach (var call in phonesForFirst.getPhones())
            {
                FullCall LastCall = new FullCall(call.Key,
                    "",
                call.Value.stages.First().Key,
                call.Value.stages.First().Value.First().Date,
                call.Value.stages.First().Value.First().Outgoing,
                call.Value.stages.First().Value.First().comment,
                call.Value.GetManager());
                var rStages = call.Value.stages.Where(s => Regex.Match(s.Key, "Первич|консульт|предварит", RegexOptions.IgnoreCase).Success);
                if (rStages.Any()) {

                    foreach (var stage in call.Value.stages)
                    {
                        foreach (var curCall in stage.Value)
                        {

                            if (curCall.Date > LastCall.date)
                            {
                                if (rStages.Where(s => s.Key == stage.Key).Any())
                                    LastCall.date = curCall.Date;
                                LastCall.stage = stage.Key;
                                LastCall.outgoing = curCall.Outgoing;
                                LastCall.Comment = curCall.comment;

                            }
                        }
                    }
                    var FirstCalltoClient = new firstCallsToClient();
                    FirstCalltoClient.comment = LastCall.Comment;
                    FirstCalltoClient.phoneNumber = call.Value.phoneNumber;
                    if (call.Value.link != "")
                        FirstCalltoClient.Link = new XLHyperlink(new Uri(call.Value.link)); 
                    else
                        FirstCalltoClient.Link = null;
                    FirstCalltoClient.date = String.Format("{0:dd.MM.yy}", LastCall.date);
                    FirstCalltoClient.Manager = call.Value.GetManager();
                    FirstCalltoClient.DealState = call.Value.DealState;
                    FirstCalltoClient.stage = String.Join(", ", rStages.Select(s => s.Key));

                    if (FirstCalltoClient.DealState.ToUpper() != "В РАБОТЕ" && FirstCalltoClient.DealState != "")
                    {
                        FirstCalltoClient.NoticeCRM = FirstCalltoClient.DealState;
                        FirstCalltoClient.DateDeal = call.Value.DateDeal.ToString("dd.MM.yyyy"); 
                    }
                    if (FirstCalltoClient.DealState.ToUpper() == "В РАБОТЕ")
                    {
                        FirstCalltoClient.DateDeal = call.Value.DateDeal.ToString("dd.MM.yyyy");
                        if (call.Value.DateDeal.Year < 2000)
                        {
                            FirstCalltoClient.DateDeal = "";
                        }
                    }
                    returnCalls.Add(FirstCalltoClient);
                }
               
            }
            return returnCalls;
        }
        public List<CallIncoming> getIncomeWithoutOutGoing()
        {
            List<CallIncoming> returnCalls = new List<CallIncoming>();
            var Stages = phones.Stages;

            foreach (var call in phones.getPhones())
            {
                FullCall LastCall = new FullCall(call.Key,
                    "",
                call.Value.stages.First().Key,
                call.Value.stages.First().Value.First().Date,
                call.Value.stages.First().Value.First().Outgoing,
                call.Value.stages.First().Value.First().comment,
                call.Value.GetManager());
                foreach (var stage in call.Value.stages)
                {
                    foreach (var curCall in stage.Value)
                    {
                        
                        if (curCall.Date > LastCall.date && Stages[Regex.Replace(LastCall.stage, @"[\d()]", String.Empty)] <= Stages[Regex.Replace(stage.Key, @"[\d]()", String.Empty)])
                        {
                            LastCall.date = curCall.Date;
                            LastCall.stage = stage.Key;
                            LastCall.outgoing = curCall.Outgoing;
                            LastCall.Comment = curCall.comment;

                        }
                    }
                }
                int LastStage = Stages.Values.Max();
                // !LastCall.outgoing && убрали, так как захотели, чтоб анализировали не только входящие
                if (Stages[Regex.Replace(LastCall.stage.ToUpper(), @"[\d]", String.Empty)] < LastStage - 1)
                {
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = call.Value.phoneNumber;
                    AddedCall.Link = call.Value.link;
                    AddedCall.Comment = LastCall.Comment;
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(new CallIncoming(call.Value.phoneNumber, call.Value.link, String.Format("{0:dd.MM.yy}", LastCall.date), LastCall.Comment, call.Value.GetManager(), new ProcessedCall(), call.Value.DealState, call.Value.DateDeal));
                    else
                    {
                        var samecall = InputDoc.getSamePhone(processedCalls, AddedCall);
                        if (samecall.ClientState != null && samecall.ClientState.ToUpper() == "В РАБОТЕ")
                        {
                            returnCalls.Add(new CallIncoming(call.Value.phoneNumber, call.Value.link, String.Format("{0:dd.MM.yy}", LastCall.date),
                                LastCall.Comment, call.Value.GetManager(), samecall, call.Value.DealState, call.Value.DateDeal));
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
                List<DateTime> dt1 = new List<DateTime>();
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
               
                int NumLastStage = Stages.Values.Max();
                string LastStage = "";
                string prevlastStage = "";
                foreach (var dictStage in Stages)
                {
                    if (dictStage.Value == NumLastStage)
                    {
                        LastStage = dictStage.Key;
                    }
                    if (dictStage.Value == NumLastStage - 1)
                    {
                        prevlastStage = dictStage.Key;
                    }
                }
                if (t1.TotalDays >= 23 && !call.Value.stages.ContainsKey(prevlastStage) && !call.Value.stages.ContainsKey(LastStage))
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

            foreach (var call in phones.getPhones())
            {
                string maxStage = call.Value.stages.First().Key;
                foreach (var stage in call.Value.stages)
                {
                    int bufValueStr;
                    if (Stages.TryGetValue(stage.Key, out bufValueStr))
                    {
                        if (Stages[maxStage] < Stages[stage.Key])
                            maxStage = stage.Key;
                    }
                }
                if (Stages[maxStage] < Stages.Values.Max() && call.Value.stages[maxStage].Count > 2)
                {
                    CallOneStage curCall = new CallOneStage();
                    curCall.phoneNumber = call.Value.phoneNumber;
                    if (call.Value.link != "")
                        curCall.Link = new XLHyperlink(new Uri(call.Value.link));
                    curCall.qty = call.Value.stages[maxStage].Count.ToString();
                    curCall.stage = maxStage;
                    curCall.date = "";
                    curCall.Manager = call.Value.GetManager();
                    string comment = "";
                    foreach (var date in call.Value.stages[maxStage])
                    {
                        curCall.date = curCall.date + String.Format("{0:dd.MM.yy}", date.Date) + ", ";
                        comment = date.comment;
                    }
                    curCall.date = curCall.date.TrimEnd(' ').Trim(',');
                    curCall.comment = comment;
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = curCall.phoneNumber;
                    AddedCall.Link = "";
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
        public List<CallPreAgreement> getCallsPreAgreement()
        {
            List<CallPreAgreement> returnCalls = new List<CallPreAgreement>();
            var Stages = phones.Stages;

            foreach (var call in phones.getPhones())
            {
                int NumMaxStage = Stages.Values.Max();
                string maxStage = call.Value.stages.First().Key;
                foreach (var stage in call.Value.stages)
                {
                    int bufValueStr;
                    if (Stages.TryGetValue(stage.Key, out bufValueStr))
                    {
                        if (Stages[maxStage] < Stages[stage.Key])
                        {
                            maxStage = stage.Key;
                        }
                    }
                }


                if (Stages[maxStage] == NumMaxStage - 3)
                {
                    OneCall compcall = call.Value.stages[maxStage].First();
                    foreach (var date in call.Value.stages[maxStage])
                    {
                        if (compcall.Date < date.Date)
                        {
                            compcall = date;
                        }
                    }
                    var AddedCall = new ProcessedCall();
                    AddedCall.Client = call.Key;
                    AddedCall.Link = "";
                    AddedCall.Comment = compcall.comment;
                    if (!InputDoc.hasPhone(processedCalls, AddedCall))
                        returnCalls.Add(new CallPreAgreement(call.Value.phoneNumber, call.Value.link, String.Format("{0:dd.MM.yy}", compcall.Date), compcall.comment, maxStage, call.Value.GetManager(),new ProcessedCall(), call.Value.DealState, call.Value.DateDeal));
                    else
                    {
                        var samecall = InputDoc.getSamePhone(processedCalls, AddedCall);
                        if (samecall.ClientState != null && samecall.ClientState.ToUpper() == "В РАБОТЕ")
                        {

                            returnCalls.Add(new CallPreAgreement(call.Value.phoneNumber, call.Value.link, String.Format("{0:dd.MM.yy}", compcall.Date), compcall.comment, maxStage, call.Value.GetManager(),
                            samecall, call.Value.DealState, call.Value.DateDeal));
                            
                        }
                    }
                }



            }
            return returnCalls;
        }
        public void FillStageDictionary(XLWorkbook wb)
        {
            int i = 1;
            var page = wb.Worksheets.First();
            var CurCell = page.Cell("A5");
            Regex rx = new Regex("ИТОГ");

            
            Dictionary<string, string> OldNewStage = new Dictionary<string, string>();
            string oldNameStage = "Консультация по коллекции/Макет коллаж 17";
            string newName = "КОНСУЛЬТАЦИЯ ПО КОЛЛЕКЦИИ/МАКЕТ КОЛЛАЖ МАКСИМАЛЬНО БАЛЛОВ БЕЗ УЧЕТА РАСШИРЕНИЯ ЧЕКА (18)";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();
            oldNameStage = "КОНСУЛЬТАЦИЯ ПО КОЛЛЕКЦИИ / МАКЕТ КОЛЛАЖ МАКСИМАЛЬНО БАЛЛОВ БЕЗ УЧЕТА РАСШИРЕНИЯ ЧЕКА(19)";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();

            oldNameStage = "Консультация по коллекции/Макет коллаж Максимально баллов без учета расширения чека (19)";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();

            oldNameStage = "Заказ 11";
            newName = "Заказ 13";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();
            oldNameStage = "Предварительный просчет 17";
            newName = "Предварительный просчет 19";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();
            oldNameStage = "Если нет оплаты 16";
            newName = "Если нет оплаты 18";
            OldNewStage[oldNameStage.ToUpper()] = newName.ToUpper();

            while (!rx.Match(CurCell.GetString().ToUpper()).Success)
            {
                if (CurCell.GetString() != "")
                {
                    string NewStage = CurCell.GetString().Trim().ToUpper();
                    if (OldNewStage.ContainsKey(NewStage))
                        NewStage = OldNewStage[NewStage];
                    phones.Stages[Regex.Replace(NewStage, @"[\d()]", String.Empty).Trim()] = i;
                    if (NewStage == "КОНСУЛЬТАЦИЯ ПО КОЛЛЕКЦИИ/МАКЕТ КОЛЛАЖ МАКСИМАЛЬНО БАЛЛОВ БЕЗ УЧЕТА РАСШИРЕНИЯ ЧЕКА" || OldNewStage.ContainsValue(NewStage))
                        phones.Stages["КОНСУЛЬТАЦИЯ ПО КОЛЛЕКЦИИ/МАКЕТ КОЛЛАЖ"] = i;
                    i++;
                }
                CurCell = CurCell.CellBelow();

            }
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
                    IXLWorksheet page = wb.Worksheets.First();

                    IXLCell cell = page.Cell(1, 5);
                    DateTime curDate;
                    if (cell.DataType == XLDataType.DateTime)
                        curDate = cell.GetDateTime();
                    else
                    {
                        if (!DateTime.TryParse(cell.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out curDate))
                            DateTime.TryParse(cell.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out curDate);

                    }
                    string phoneNumber;
                    IXLCell phoneCell;
                    while (!(cell.IsEmpty() && cell.CellRight().IsEmpty() && !cell.IsMerged()))
                    {
                        if (cell.GetString() != "")
                        {
                            if (cell.DataType == XLDataType.DateTime)
                                curDate = cell.GetDateTime();
                            else
                            {
                                if (!DateTime.TryParse(cell.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out curDate))
                                    DateTime.TryParse(cell.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out curDate);

                            }
                        }

                        phoneCell = cell.CellBelow();
                        if (phoneCell.GetString() == "")
                            phoneCell = phoneCell.CellBelow();
                        if (phoneCell.GetString() != "")
                        {

                            string link;
                            if (phoneCell.HasHyperlink)
                                link = phoneCell.GetHyperlink().ExternalAddress.AbsoluteUri;
                            else
                                link = "";
                            Match outgoing = Regex.Match(phoneCell.GetString().ToUpper(), @"ИСХОДЯЩИЙ");
                            phoneNumber = Regex.Replace(phoneCell.GetString().ToUpper(), @"[^\d]", String.Empty);
                            string oldphonenum = phoneNumber;
                            oldphonenum =  "8 (" + oldphonenum.Substring(1, 3) + ") " + oldphonenum.Substring(4, 3) + "-" + oldphonenum.Substring(7, 2) + "-" + oldphonenum.Substring(9);
                            while (phoneNumber[0] == '0')
                            {
                                phoneNumber = phoneNumber.Substring(1);
                            }
                            if (phoneNumber[0] == '9')
                                phoneNumber = '8' + phoneNumber;
                            if (phoneNumber[0] == '7' || phoneNumber[0] == '8')
                                phoneNumber = "8 (" + phoneNumber.Substring(1, 3) + ") " + phoneNumber.Substring(4, 3) + "-" + phoneNumber.Substring(7, 2) + "-" + phoneNumber.Substring(9);
                            
                            if (processedCalls.Exists(c=> c.Client == oldphonenum) && oldphonenum!= phoneNumber)
                            {
                                var testCall = processedCalls.Where(c => c.Client == oldphonenum).First();
                                processedCalls.Remove(testCall);
                                testCall.Client = phoneNumber;
                                processedCalls.Add(testCall);
                            }
                            if (processedCalls.Exists(c => c.Client == phoneNumber && c.Link == ""))
                            {
                                var testCall = processedCalls.Where(c => c.Client == phoneNumber).First();
                                testCall.Link = link;
                            }

                            var CellStage = page.Cell("A5");
                            Regex rx = new Regex("ИТОГ");
                            int corrRow = 5;
                            Match Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            while (!Mcomment.Success)
                            {
                                corrRow++;
                                Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            }
                            while (!rx.Match(CellStage.GetString().ToUpper()).Success && !rx.Match(CellStage.CellRight().CellRight().CellRight().GetString().ToUpper()).Success)
                            {
                                if (CellStage.GetString() != "" && page.Cell(CellStage.Address.RowNumber, cell.Address.ColumnNumber).GetString() != "")
                                {
                                    var exCallSeq = processedCalls.Where(c => (c.Client == phoneNumber));
                                    var exCall = new ProcessedCall();
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
                                    if (curDate >= exCall.StartDateAnalyze ||
                                        (
                                          exCall.ClientState.ToUpper() == "В РАБОТЕ") &&
                                          exCall.StartDateAnalyze < DateTime.Today.AddDays(1)
                                    )
                                    { DateTime DateNext = new DateTime();
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
                                        if (curDate > DateTime.Now.AddMonths(-1) && Regex.Match(file.Name, "Гакова|Малькова|Лукина|Кожевникова|Рыбачук", RegexOptions.IgnoreCase).Success)
                                            phones.AddCall(new FullCall(phoneNumber, link, Regex.Replace(CellStage.GetString(), @"[\d()]", String.Empty).Trim(), curDate, outgoing.Success, page.Cell(corrRow, cell.Address.ColumnNumber).GetString(), Manager, page.Cell(corrRow + 5, cell.Address.ColumnNumber).GetString(), DateNext));
                                        phonesForFirst.AddCall(new FullCall(phoneNumber, link, Regex.Replace(CellStage.GetString(), @"[\d()]", String.Empty).Trim(), curDate, outgoing.Success, page.Cell(corrRow, cell.Address.ColumnNumber).GetString(), Manager, page.Cell(corrRow + 5, cell.Address.ColumnNumber).GetString(), DateNext));

                                    }
                                }
                                CellStage = CellStage.CellBelow();

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
