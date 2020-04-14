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
    public class Belfan : ICompany
    {
        Phones phones = new Phones();
        
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
                        
                        if (curCall.Date > LastCall.date && Stages[LastCall.stage] <= Stages[stage.Key])
                        {
                            LastCall.date = curCall.Date;
                            LastCall.stage = stage.Key;
                            LastCall.outgoing = curCall.Outgoing;
                            LastCall.Comment = curCall.comment;

                        }
                    }
                }
                int LastStage = Stages.Values.Max();
                if (!LastCall.outgoing && Stages[LastCall.stage.ToUpper()] < LastStage - 1)
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
                List<DateTime> dt1 = new List<DateTime>();
                foreach (var stage in call.Value.stages)
                {
                    foreach (var onecall in stage.Value)
                    {
                        if (onecall.Outgoing)
                            dt1.Add(onecall.Date);
                    }
                }
                if (dt1.Count > 0)
                {
                    dt1.Sort();
                    dt1.Reverse();
                    DateTime prevDate = dt1.First();
                    foreach (var date in dt1)
                    {
                        TimeSpan t1 = prevDate.Subtract(date);
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
                        if (t1.TotalDays >= 7 && !call.Value.stages.ContainsKey(prevlastStage) && !call.Value.stages.ContainsKey(LastStage))
                        {
                            CallPerWeek curCall = new CallPerWeek();
                            curCall.FirstWeek = "-";
                            curCall.phoneNumber = call.Key;
                            foreach (var stage in call.Value.stages)
                            {
                                foreach (var onecall in stage.Value)
                                {
                                    if (onecall.Outgoing && onecall.Date == date)
                                        curCall.comment = onecall.comment;
                                }
                            }
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
                            break;
                        }
                        prevDate = date;
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
                    if (Stages[maxStage] < Stages[stage.Key])
                        maxStage = stage.Key;
                }
                if (Stages[maxStage] < Stages.Values.Max() && call.Value.stages[maxStage].Count > 1)
                {
                    CallOneStage curCall = new CallOneStage();
                    curCall.phoneNumber = call.Key;
                    curCall.qty = call.Value.stages[maxStage].Count.ToString();
                    curCall.stage = maxStage;
                    curCall.date = "";
                    string comment = "";
                    foreach (var date in call.Value.stages[maxStage])
                    {
                        curCall.date = curCall.date + String.Format("{0:dd.MM.yy}", date.Date) + ", ";
                        comment = date.comment;
                    }
                    curCall.date = curCall.date.TrimEnd(' ').Trim(',');
                    curCall.comment = comment;
                    returnCalls.Add(curCall);
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
                    if (Stages[maxStage] < Stages[stage.Key])
                    {
                        maxStage = stage.Key;
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
                    returnCalls.Add(new CallPreAgreement(call.Key, String.Format("{0:dd.MM.yy}", compcall.Date), compcall.comment, maxStage));
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
            while (!rx.Match(CurCell.GetString().ToUpper()).Success)
            {
                if (CurCell.GetString() != "")
                {
                    phones.Stages[CurCell.GetString().Trim().ToUpper()] = i;
                    i++;
                }
                CurCell = CurCell.CellBelow();

            }
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
                    IXLWorksheet page = wb.Worksheets.First();

                    IXLCell cell = page.Cell(1, 5);
                    DateTime curDate;
                    DateTime.TryParse(cell.GetValue<string>(), out curDate);
                    string phoneNumber;
                    string phoneCell;
                    while (!(cell.IsEmpty() && cell.CellRight().IsEmpty() && !cell.IsMerged()))
                    {
                        if (cell.GetString() != "")
                        {
                            DateTime.TryParse(cell.GetValue<string>(), out curDate);
                        }
                        phoneCell = cell.CellBelow().CellBelow().GetValue<string>().ToUpper().Trim();
                        if (phoneCell != "")
                        {
                            Match outgoing = Regex.Match(phoneCell.ToUpper(), @"ИСХОДЯЩИЙ");
                            phoneNumber = Regex.Replace(phoneCell.ToUpper(), @"[^\d]", String.Empty);
                            phoneNumber = "8 (" + phoneNumber.Substring(1, 3) + ") " + phoneNumber.Substring(4, 3) + "-" + phoneNumber.Substring(7, 2) + "-" + phoneNumber.Substring(9);
                            var CellStage = page.Cell("A5");
                            Regex rx = new Regex("ИТОГ");
                            int corrRow = 5;
                            Match Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            while (!Mcomment.Success)
                            {
                                corrRow++;
                                Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            }
                            while (!rx.Match(CellStage.GetString().ToUpper()).Success)
                            {
                                if (CellStage.GetString() != "" && page.Cell(CellStage.Address.RowNumber, cell.Address.ColumnNumber).GetString() != "")
                                {
                                    phones.AddCall(new FullCall(phoneNumber, CellStage.GetString(), curDate, outgoing.Success, page.Cell(corrRow, cell.Address.ColumnNumber).GetString()));
                                }
                                CellStage = CellStage.CellBelow();

                            }

                        }

                        cell = cell.CellRight();
                    }

                }
            }
        }
    }
}
