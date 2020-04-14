using ClosedXML.Excel;
using MetanitAngular.Models;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static MetanitAngular.Excel.DataStructsForPrintCalls;

namespace MetanitAngular.ProcessingDataCompanies
{
    public class DefaultCompany : ICompany
    {
        Phones phones = new Phones();
        string AgreementStage = "";
        string PreAgreementStage = "";
        
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
                FullCall LastCall = new FullCall(call.Key,
                call.Value.stages.First().Key,
                call.Value.stages.First().Value.First().Date,
                call.Value.stages.First().Value.First().Outgoing,
                call.Value.stages.First().Value.First().comment);
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
                    returnCalls.Add(new CallIncoming(call.Key, String.Format("{0:dd.MM.yy}", LastCall.date), LastCall.Comment));
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

                FullCall LastCall = new FullCall(call.Key,
                call.Value.stages.First().Key,
                call.Value.stages.First().Value.First().Date,
                call.Value.stages.First().Value.First().Outgoing,
                call.Value.stages.First().Value.First().comment);
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

                if (t1.TotalDays >= 7 && !call.Value.stages.ContainsKey(AgreementStage))
                {
                    CallPerWeek curCall = new CallPerWeek();
                    curCall.FirstWeek = "-";
                    curCall.phoneNumber = call.Key;
                    curCall.comment = LastCall.Comment;
                    if (!LastCall.outgoing)
                        curCall.comment = curCall.comment + " (Входящий)";
                    if (t1.TotalDays >= 14)
                    {
                        curCall.SecondWeek = "-";
                    }
                    else
                    {
                        curCall.SecondWeek = "+";
                    }
                    if (t1.TotalDays >= 21)
                    {
                        curCall.ThirdWeek = "-";
                    }
                    else
                    {
                        curCall.ThirdWeek = "+";
                    }
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
                            if (stage.Key != lastStage)
                            {
                                call.comment = call.comment + " (" + stage.Key + ")";
                            }
                            dt.Add(call);
                        }
                    }
                }
                
                if (lastStage != AgreementStage && dt.Count > 1)
                {
                    CallOneStage curCall = new CallOneStage();
                    curCall.phoneNumber = phone.Key;
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
                                LastCall = call;
                                if (stage.Key != lastStage)
                                    LastCall.comment = LastCall.comment + " (" + stage.Key + ") ";
                            }
                        }
                    }
                    returnCalls.Add(new CallPreAgreement(phone.Key, String.Format("{0:dd.MM.yy}", LastCall.Date), LastCall.comment, lastStage));
                }



            }
            return returnCalls;
        }
        public void ParserCheckLists(IFormFileCollection files)
        {
            using (var stream = files[0].OpenReadStream())
            {
                XLWorkbook wb = new XLWorkbook(stream);
                FillStageDictionary(wb);
            }

            foreach (var file in files)
            {
                using (var stream = file.OpenReadStream())
                {
                    XLWorkbook wb = new XLWorkbook(stream);
                    foreach (var page in wb.Worksheets)
                    {
                        if (page.Name.ToUpper().Trim() != "СТАТИСТИКА" && page.Name.ToUpper().Trim() != "СВОДНАЯ")
                        {

                            IXLCell cell = page.Cell(1, 5);
                            DateTime curDate;
                            DateTime.TryParse(cell.GetValue<string>(), out curDate);
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
                                    DateTime.TryParse(cell.GetValue<string>(), out curDate);
                                }
                                phoneNumber = cell.CellBelow().GetValue<string>().ToUpper().Trim();
                                if (phoneNumber != "")
                                {
                                    Regex rx = new Regex("ВХОДЯЩ");
                                    Match m = rx.Match(page.Name.ToUpper().Trim());
                                    phones.AddCall(new FullCall(phoneNumber, page.Name.ToUpper().Trim(), curDate, !m.Success, page.Cell(corrRow, cell.Address.ColumnNumber).GetString()));

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
