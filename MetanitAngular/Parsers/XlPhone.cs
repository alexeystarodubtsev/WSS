using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;
using MetanitAngular.Excel;
using MetanitAngular.Models;
using MetanitAngular.ProcessingDataCompanies;
using Microsoft.AspNetCore.Http;
using static MetanitAngular.Excel.DataStructsForPrintCalls;

namespace MetanitAngular.Parsers
{
    public class XlPhone
    {


        //public static Tuple<List<Phone>, List<Phone>> getPhone(IFormFileCollection files)
        //{
        //    Dictionary<string, Dictionary<string, Tuple<int, string>>> dict = new Dictionary<string, Dictionary<string, Tuple<int, string>>>();
        //    Dictionary<string, Tuple<int,string,string>> dictPhone = new Dictionary<string, Tuple<int, string, string>>();
        //    List<Phone> phones = new List<Phone>();
        //    List<Phone> OnlyOneCall = new List<Phone>();

        //    foreach (var file in files)
        //    {
        //        using (var stream = file.OpenReadStream())
        //        {
        //            XLWorkbook wb = new XLWorkbook(stream);
        //            foreach (var page in wb.Worksheets)
        //            {

        //                if (page.Name.ToUpper().Trim() != "СТАТИСТИКА" && page.Name.ToUpper().Trim() != "СВОДНАЯ" )
        //                {
        //                    if (!dict.ContainsKey(page.Name.ToUpper().Trim()))
        //                    {
        //                        dict[page.Name.ToUpper().Trim()] = new Dictionary<string, Tuple<int, string>>();
        //                    }
        //                    IXLCell cell = page.Cell(1, 5);
        //                    DateTime curDate;
        //                    DateTime.TryParse(cell.GetValue<string>(), out curDate);
        //                    string phoneNumber;
        //                    while (!(cell.IsEmpty() && cell.CellRight().IsEmpty() && !cell.IsMerged()))
        //                    {
        //                        if (cell.GetValue<string>() != "")
        //                        {
        //                            DateTime.TryParse(cell.GetValue<string>(), out curDate);
        //                        }

        //                        phoneNumber = cell.CellBelow().GetValue<string>().ToUpper();
        //                        if (phoneNumber != "")
        //                        {

        //                            if (!dict[page.Name.ToUpper().Trim()].ContainsKey(phoneNumber))
        //                            {
        //                                dict[page.Name.ToUpper().Trim()][phoneNumber] = new Tuple<int, string>(1, String.Format("{0:dd.MM.yy}", curDate));
        //                            }
        //                            else
        //                            {
        //                                int curQty = dict[page.Name.ToUpper().Trim()][phoneNumber].Item1;
        //                                string curDates = dict[page.Name.ToUpper().Trim()][phoneNumber].Item2;
        //                                curQty++;
        //                                curDates = curDates + ", " + String.Format("{0:dd.MM.yy}", curDate);
        //                                dict[page.Name.ToUpper().Trim()][phoneNumber] = new Tuple<int, string>(curQty, curDates);
        //                            }
        //                            if (!dictPhone.ContainsKey(phoneNumber))
        //                            {
        //                                var data = new Tuple<int, string, string>(1, page.Name.ToUpper().Trim(), String.Format("{0:dd.MM.yy}", curDate));
        //                                dictPhone[phoneNumber] = data;
        //                            }
        //                            else
        //                            {
        //                                dictPhone[phoneNumber] = new Tuple<int, string, string>(3, "", "");
        //                            }
        //                        }

        //                        cell = cell.CellRight();
        //                    }

        //                }
        //            }
        //        }

        //    }

        //    foreach (var dictStage in dict)
        //    {
        //        foreach (var phone in dictStage.Value)
        //        {
        //            if (phone.Value.Item1 > 1)
        //            {
        //                phones.Add(new Phone(phone.Key, phone.Value.Item1, dictStage.Key, phone.Value.Item2));
        //            }
        //        }
        //    }
        //    foreach (var phone in dictPhone)
        //    {
        //        if (phone.Value.Item1 == 1)
        //        {
        //            OnlyOneCall.Add(new Phone(phone.Key, phone.Value.Item1, phone.Value.Item2, phone.Value.Item3));
        //        }

        //    }
        //    Tuple<List<Phone>, List<Phone>> returnPhones;
        //    returnPhones = new Tuple<List<Phone>, List<Phone>>(phones, OnlyOneCall);

        //    string fileoutput = "C:\\Users\\xiaomi\\source\\repos\\MetanitAngular\\MetanitAngular\\OutputAnalitics";
        //    XLWorkbook wbout = new XLWorkbook();
        //    var worksheet =  wbout.Worksheets.Add("Вх, на которые не перезв");
        //    //создадим заголовки у столбцов
        //    worksheet.Cell("A" + 1).Value = "Имя";
        //    worksheet.Cell("B" + 1).Value = "Фамиля";
        //    worksheet.Cell("C" + 1).Value = "Возраст";

        //    // 

        //    worksheet.Cell("A" + 2).Value = "Иван";
        //    worksheet.Cell("B" + 2).Value = "Иванов";
        //    worksheet.Cell("C" + 2).Value = 18;
        //    //пример изменения стиля ячейки
        //    worksheet.Cell("B" + 2).Style.Fill.BackgroundColor = XLColor.Red;

        //    // пример создания сетки в диапазоне


        //    worksheet = wbout.Worksheets.Add("Застрявшие на 1 этапе");

        //    worksheet.Cell(1, 1).Value = "Клиент";
        //    worksheet.Cell(1, 1).Style.Font.Bold = true;
        //    worksheet.Cell(1, 2).Value = "Этап";
        //    worksheet.Cell(1, 2).Style.Font.Bold = true;
        //    worksheet.Cell(1, 3).Value = "Количество повторных звонков";
        //    worksheet.Cell(1, 3).Style.Font.Bold = true;
        //    worksheet.Cell(1, 4).Value = "Даты звонков";
        //    worksheet.Cell(1, 4).Style.Font.Bold = true;


        //    int i = 1;
        //    foreach (var phone in phones)
        //    {
        //        i++;
        //        worksheet.Cell(i, 1).Value = phone.PhoneNumber;
        //        worksheet.Cell(i, 2).Value = phone.Stage;
        //        worksheet.Cell(i, 3).Value = phone.Qty;
        //        worksheet.Cell(i, 4).Value = phone.Date;

        //    }

        //    var rngTable = worksheet.Range("A1:D" + i);
        //    rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        //    rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //    rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        //    worksheet.Columns().AdjustToContents(); //ширина столбца
        //    wbout.SaveAs(fileoutput + "\\Деловой союз.xlsx");
        //    return returnPhones;
        //}

        public static Tuple<List<Phone>, List<Phone>> getPhoneNew(IFormFileCollection files, string nameoutput)
        {
            List<ProcessedCall> processedCalls = new List<ProcessedCall>();
            try
            {
                var lastCommentsFile = files.Where(c => Regex.Match(c.FileName.ToUpper(),"ПРЕДЫДУЩАЯ АНАЛИТИКА").Success).First();
                processedCalls = InputDoc.GetProcessedCalls(lastCommentsFile);
            }
            catch (InvalidOperationException)
            {

            }
            var FilesManagers = files.Where(c => !Regex.Match(c.FileName.ToUpper(), "ПРЕДЫДУЩАЯ АНАЛИТИКА").Success);
            //Dictionary<string, List<DataCall>> phones = new Dictionary<string, List<DataCall>>();
            Regex rNameOut = new Regex("БЕЛФАН");
            Regex rRNR = new Regex("РНР");
            Regex rAvers = new Regex("АВЕРС");
            Regex rDS = new Regex("Деловой союз", RegexOptions.IgnoreCase);
            ICompany company;
            if (!rNameOut.Match(nameoutput.ToUpper()).Success)
            {
                if (!rRNR.Match(nameoutput.ToUpper()).Success)
                {
                    if (!rAvers.Match(nameoutput.ToUpper()).Success)
                        company = new DefaultCompany(processedCalls, rDS.Match(nameoutput).Success);
                    else
                        company = new Avers(processedCalls);
                }
                else
                    company = new RNRHouse(processedCalls);
            }
            else
            {
                company = new Belfan(ref processedCalls);
            }
            company.ParserCheckLists(FilesManagers);

            string fileoutput = "C:\\Users\\xiaomi\\source\\repos\\MetanitAngular\\MetanitAngular\\OutputAnalitics";
            OutputDoc wbout = new OutputDoc();
            wbout.setProcessedCalls(processedCalls);
            wbout.FillIncoming(company.getIncomeWithoutOutGoing());
            wbout.FillOutGoingPerWeeks(company.getCallsPerWeek(), rDS.Match(nameoutput).Success);
            wbout.FillCallsOnSameStage(company.getCallsOneStage());
            wbout.FillCallsWithoutAgreement(company.getCallsPreAgreement());
            wbout.FillArchive();
            var wboutFile = wbout.getFile();

            string dirpath = fileoutput;
            if (nameoutput == "")
                fileoutput = fileoutput + "\\" + "default" + ".xlsx";
            else
                fileoutput = fileoutput + "\\" + nameoutput + ".xlsx";
            try
            {
                wboutFile.SaveAs(fileoutput);
            }
            catch (System.IO.IOException)
            {
                fileoutput = dirpath + "\\Копия " + nameoutput + ".xlsx";
            }
            Tuple<List<Phone>, List<Phone>> returnPhones = new Tuple<List<Phone>, List<Phone>>(new List<Phone>(), new List<Phone>());
            return returnPhones;
        }



        //public static Tuple<List<Phone>, List<Phone>> getPhoneNew(IFormFileCollection files, string nameoutput)
        //{
        //    //Dictionary<string, List<DataCall>> phones = new Dictionary<string, List<DataCall>>();
        //    Regex rNameOut = new Regex("БЕЛФАН");
        //    if (!rNameOut.Match(nameoutput.ToUpper()).Success)
        //    {
        //        Tuple<Dictionary<string, Dictionary<string, List<DateTime>>>, Dictionary<string, int>> tup1 = TypicalParser(files);
        //        Dictionary<string, Dictionary<string, List<DateTime>>> phones = tup1.Item1;
        //        Dictionary<string, int> DictPages = tup1.Item2;
        //        string AgreementStage = "";
        //        string preAgreementStage = "";
        //        foreach (var page in DictPages)
        //        {
        //            Regex rx = new Regex("ДОГОВОР");
        //            Match m = rx.Match(page.Key);
        //            if (m.Success)
        //            {
        //                AgreementStage = page.Key;
        //            }
        //            rx = new Regex("КП ОТПРАВЛЕН");
        //            m = rx.Match(page.Key);
        //            if (m.Success)
        //            {
        //                preAgreementStage = page.Key;
        //            }
        //        }


        //        string fileoutput = "C:\\Users\\xiaomi\\source\\repos\\MetanitAngular\\MetanitAngular\\OutputAnalitics";
        //        XLWorkbook wbout = new XLWorkbook();
        //        var worksheet = wbout.Worksheets.Add("Вх, на которые не перезвон"); //Нужно видеть входящие звонки на которые не перезвонили
        //        worksheet.Cell(1, 1).Value = "Клиент";
        //        worksheet.Cell(1, 1).Style.Font.Bold = true;
        //        worksheet.Cell(1, 2).Value = "Дата";
        //        worksheet.Cell(1, 2).Style.Font.Bold = true;
        //        int curRow = 1;
        //        foreach (var phone in phones)
        //        {
        //            Dictionary<string, List<DateTime>> stages = phone.Value;
        //            DataCall LastCall = new DataCall(stages.First().Key, stages.First().Value[0]);

        //            foreach (var stage in stages)
        //            {
        //                foreach (var date in stage.Value)
        //                {
        //                    if (date > LastCall.Date)
        //                    {
        //                        LastCall = new DataCall(stage.Key, date);
        //                    }
        //                }
        //            }

        //            Regex rx = new Regex("ВХОДЯЩ");
        //            Match m = rx.Match(LastCall.Stage);
        //            if (m.Success)
        //            {

        //                curRow++;
        //                worksheet.Cell(curRow, 1).Value = phone.Key;
        //                worksheet.Cell(curRow, 2).Value = String.Format("{0:dd.MM.yy}", LastCall.Date);

        //            }
        //        }
        //        var rngTable = worksheet.Range("A1:B" + curRow);
        //        rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        //        rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //        rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        //        //worksheet.Style.Alignment.WrapText = true;
        //        worksheet.Columns().AdjustToContents(); //ширина столбца


        //        worksheet = wbout.Worksheets.Add("Сделанные раз за 1,2,3 недели"); //2. нужно вдеть исходящие звонки которе сделаны всего один раз за 2 недели или за неделю( в зависимости от специфики)
        //                                                                           //создадим заголовки у столбцов
        //        worksheet.Cell(1, 1).Value = "Клиент";
        //        worksheet.Cell(1, 1).Style.Font.Bold = true;
        //        worksheet.Cell(1, 2).Value = "1 неделя";
        //        worksheet.Cell(1, 2).Style.Font.Bold = true;
        //        worksheet.Cell(1, 3).Value = "2 неделя";
        //        worksheet.Cell(1, 3).Style.Font.Bold = true;
        //        worksheet.Cell(1, 4).Value = "3 неделя";
        //        worksheet.Cell(1, 4).Style.Font.Bold = true;

        //        curRow = 1;
        //        foreach (var phone in phones)
        //        {
        //            Dictionary<string, List<DateTime>> stages = phone.Value;
        //            List<DateTime> dt1 = new List<DateTime>();
        //            foreach (var stage in stages)
        //            {
        //                Regex rx = new Regex("ВХОДЯЩ");
        //                Match m = rx.Match(stage.Key);
        //                if (!m.Success)
        //                {
        //                    foreach (var date in stage.Value)
        //                    {
        //                        dt1.Add(date);
        //                    }


        //                }

        //            }
        //            if (dt1.Count > 0)
        //            {
        //                dt1.Sort();

        //                DateTime prevDate = dt1.First();
        //                foreach (var date in dt1)
        //                {
        //                    TimeSpan t1 = date.Subtract(prevDate);
        //                    if (t1.TotalDays >= 7 && !phone.Value.ContainsKey(AgreementStage))
        //                    {
        //                        curRow++;
        //                        worksheet.Cell(curRow, 1).Value = phone.Key;
        //                        worksheet.Cell(curRow, 2).Value = "-";
        //                        if (t1.TotalDays >= 14)
        //                        {
        //                            worksheet.Cell(curRow, 3).Value = "-";
        //                        }
        //                        else
        //                        {
        //                            worksheet.Cell(curRow, 3).Value = "+";
        //                        }
        //                        if (t1.TotalDays >= 21)
        //                        {
        //                            worksheet.Cell(curRow, 4).Value = "-";
        //                        }
        //                        else
        //                        {
        //                            worksheet.Cell(curRow, 4).Value = "+";
        //                        }
        //                        break;
        //                    }
        //                    prevDate = date;
        //                }
        //            }

        //        }

        //        rngTable = worksheet.Range("A1:D" + curRow);
        //        rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        //        rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //        rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        //        //worksheet.Style.Alignment.WrapText = true;
        //        worksheet.Columns().AdjustToContents(); //ширина столбца

        //        worksheet = wbout.Worksheets.Add("Застрявшие на 1 этапе"); //4. Видеть звонки которые задержались на одном и том же этапе

        //        worksheet.Cell(1, 1).Value = "Клиент";
        //        worksheet.Cell(1, 1).Style.Font.Bold = true;
        //        worksheet.Cell(1, 2).Value = "Этап";
        //        worksheet.Cell(1, 2).Style.Font.Bold = true;
        //        worksheet.Cell(1, 3).Value = "Количество повторных звонков";
        //        worksheet.Cell(1, 3).Style.Font.Bold = true;
        //        worksheet.Cell(1, 4).Value = "Даты звонков";
        //        worksheet.Cell(1, 4).Style.Font.Bold = true;

        //        curRow = 1;

        //        foreach (var phone in phones)
        //        {
        //            Dictionary<string, List<DateTime>> stages = phone.Value;
        //            DataCall LastCall = new DataCall(stages.First().Key, stages.First().Value[0]);
        //            DataCall PreLastCall = new DataCall(stages.First().Key, stages.First().Value[0]);
        //            foreach (var stage in stages)
        //            {

        //                foreach (var date in stage.Value)
        //                {
        //                    Regex rx = new Regex("ВХОДЯЩ");
        //                    Match m = rx.Match(stage.Key);

        //                    if (date > LastCall.Date && stage.Key != new string("Уточняющее касание").ToUpper() && !m.Success)
        //                    {
        //                        PreLastCall = LastCall;
        //                        LastCall = new DataCall(stage.Key, date);
        //                    }
        //                }


        //            }
        //            if (PreLastCall.Date != LastCall.Date && PreLastCall.Stage == LastCall.Stage)
        //            {
        //                curRow++;
        //                worksheet.Cell(curRow, 1).Value = phone.Key;
        //                worksheet.Cell(curRow, 2).Value = LastCall.Stage;
        //                worksheet.Cell(curRow, 3).Value = stages[LastCall.Stage].Count;
        //                string dates = "";
        //                foreach (var date in stages[LastCall.Stage])
        //                {
        //                    dates = dates + String.Format("{0:dd.MM.yy}", date) + ", ";
        //                }

        //                worksheet.Cell(curRow, 4).Value = dates.TrimEnd(' ').TrimEnd(',');

        //            }



        //        }

        //        rngTable = worksheet.Range("A1:D" + curRow);
        //        rngTable.Style.Alignment.WrapText = true;
        //        rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        //        rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //        rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        //        //worksheet.Style.Alignment.WrapText = true;
        //        worksheet.Columns().AdjustToContents(); //ширина столбца

        //        worksheet = wbout.Worksheets.Add("Не закрытые в договор"); //5. не закрытые в договор

        //        worksheet.Cell(1, 1).Value = "Клиент";
        //        worksheet.Cell(1, 1).Style.Font.Bold = true;
        //        worksheet.Cell(1, 2).Value = "Этап";
        //        worksheet.Cell(1, 2).Style.Font.Bold = true;

        //        worksheet.Cell(1, 3).Value = "Дата последнего звонка";
        //        worksheet.Cell(1, 3).Style.Font.Bold = true;
        //        curRow = 1;
        //        foreach (var phone in phones)
        //        {
        //            Dictionary<string, List<DateTime>> stages = phone.Value;
        //            //string maxStage = stages.First().Key;
        //            //foreach (var stage in stages)
        //            //{
        //            //    if (DictPages[stage.Key] > DictPages[maxStage] && DictPages[stage.Key] <= AgreementStageNum)
        //            //    {
        //            //        maxStage = stage.Key;
        //            //    }

        //            //}
        //            //if (DictPages[maxStage] < AgreementStageNum && DictPages[maxStage] >= AgreementStageNum - 2)
        //            //{
        //            //    curRow++;
        //            //    worksheet.Cell(curRow, 1).Value = phone.Key;
        //            //    worksheet.Cell(curRow, 2).Value = maxStage;
        //            //    DateTime maxDate = stages[maxStage][0];
        //            //    foreach (var date in stages[maxStage])
        //            //    {
        //            //        if (date > maxDate)
        //            //        {
        //            //            maxDate = date;
        //            //        }
        //            //    }
        //            //    worksheet.Cell(curRow, 3).Value = String.Format("{0:dd.MM.yy}", maxDate);
        //            //}
        //            if (stages.ContainsKey(preAgreementStage) && !stages.ContainsKey(AgreementStage))
        //            {
        //                curRow++;
        //                worksheet.Cell(curRow, 1).Value = phone.Key;
        //                worksheet.Cell(curRow, 2).Value = preAgreementStage;
        //                DateTime maxDate = stages[preAgreementStage][0];
        //                foreach (var date in stages[preAgreementStage])
        //                {
        //                    if (date > maxDate)
        //                    {
        //                        maxDate = date;
        //                    }
        //                }
        //                worksheet.Cell(curRow, 3).Value = String.Format("{0:dd.MM.yy}", maxDate);
        //            }


        //        }

        //        rngTable = worksheet.Range("A1:C" + curRow);

        //        rngTable.Style.Alignment.WrapText = true;
        //        rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        //        rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //        rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        //        //worksheet.Style.Alignment.WrapText = true;
        //        worksheet.Columns().AdjustToContents(); //ширина столбца

        //        if (nameoutput == "")
        //            fileoutput = fileoutput + "\\" + "default" + ".xlsx";
        //        else
        //            fileoutput = fileoutput + "\\" + nameoutput + ".xlsx";
        //        wbout.SaveAs(fileoutput);
        //    }
        //    else
        //    {
        //        getCallsForBelfan(files, nameoutput);
        //    }
        //    Tuple<List<Phone>, List<Phone>> returnPhones = new Tuple<List<Phone>, List<Phone>>(new List<Phone>(), new List<Phone>());
        //    return returnPhones;
        //}


        //private static Tuple<Dictionary<string, Dictionary<string, List<DateTime>>>, Dictionary<string, int>> TypicalParser(IFormFileCollection files)
        //{
        //    Dictionary<string, Dictionary<string, List<DateTime>>> phones = new Dictionary<string, Dictionary<string, List<DateTime>>>();
        //    Dictionary<string, int> DictPages = new Dictionary<string, int>();
        //    //CallsBelfan phones = new CallsBelfan();

        //    string AgreementStage = "";
        //    string preAgreementStage = "";
        //    using (var stream = files[0].OpenReadStream())
        //    {
        //        XLWorkbook wb = new XLWorkbook(stream);
        //        int i = 1;
        //        foreach (var page in wb.Worksheets)
        //        {
        //            DictPages[page.Name.ToUpper().Trim()] = i;
        //            i++;
        //            Regex rx = new Regex("ДОГОВОР");
        //            Match m = rx.Match(page.Name.ToUpper().Trim());
        //            if (m.Success)
        //            {
        //                AgreementStage = page.Name.ToUpper().Trim();
        //            }
        //            rx = new Regex("КП ОТПРАВЛЕН");
        //            m = rx.Match(page.Name.ToUpper().Trim());
        //            if (m.Success)
        //            {
        //                preAgreementStage = page.Name.ToUpper().Trim();
        //            }
        //        }
        //    }

        //    foreach (var file in files)
        //    {
        //        using (var stream = file.OpenReadStream())
        //        {
        //            XLWorkbook wb = new XLWorkbook(stream);
        //            foreach (var page in wb.Worksheets)
        //            {
        //                if (page.Name.ToUpper().Trim() != "СТАТИСТИКА" && page.Name.ToUpper().Trim() != "СВОДНАЯ")
        //                {

        //                    IXLCell cell = page.Cell(1, 5);
        //                    DateTime curDate;
        //                    DateTime.TryParse(cell.GetValue<string>(), out curDate);
        //                    string phoneNumber;
        //                    while (!(cell.IsEmpty() && cell.CellRight().IsEmpty() && !cell.IsMerged()))
        //                    {
        //                        if (cell.GetValue<string>() != "")
        //                        {
        //                            DateTime.TryParse(cell.GetValue<string>(), out curDate);
        //                        }
        //                        phoneNumber = cell.CellBelow().GetValue<string>().ToUpper().Trim();
        //                        if (phoneNumber != "")
        //                        {
        //                            if (!phones.ContainsKey(phoneNumber))
        //                            {
        //                                phones[phoneNumber] = new Dictionary<string, List<DateTime>>();
        //                            }

        //                            if (!phones[phoneNumber].ContainsKey(page.Name.ToUpper().Trim()))
        //                            {
        //                                phones[phoneNumber][page.Name.ToUpper().Trim()] = new List<DateTime>();
        //                            }
        //                            phones[phoneNumber][page.Name.ToUpper().Trim()].Add(curDate);
        //                        }

        //                        cell = cell.CellRight();
        //                    }

        //                }
        //            }
        //        }
        //    }
        //    return new Tuple<Dictionary<string, Dictionary<string, List<DateTime>>>, Dictionary<string, int>>(phones, DictPages);
        //}


    }      
}
