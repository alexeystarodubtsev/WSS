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

namespace MetanitAngular.ProcessingDataCompanies
{
    public class Avers : DefaultCompany, ICompany
    {
        public Avers(List<ProcessedCall> processedCalls) : base(processedCalls)
        {

        }
        public new void ParserCheckLists(IEnumerable<IFormFile> files)
        {
            var DefaultFiles = files.Where(f => Regex.Match(f.FileName, "newf").Success);
            var filesAvers = files.Where(f => !Regex.Match(f.FileName, "newf").Success);
            base.ParserCheckLists(DefaultFiles);
            //using (var stream = files.First().OpenReadStream())
            //{
            //    XLWorkbook wb = new XLWorkbook(stream);
            //    FillStageDictionary(wb);
            //}

            foreach (var file in filesAvers)
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

                            IXLCell cell = page.Cell(2, 5);
                            DateTime curDate;
                            if (cell.DataType == XLDataType.DateTime)
                                curDate = cell.GetDateTime();
                            else
                            {
                                if (!DateTime.TryParse(cell.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out curDate))
                                    DateTime.TryParse(cell.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out curDate);

                            }
                            string phoneNumber;
                            int corrRow = 5;
                            Match Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            while (!Mcomment.Success)
                            {
                                corrRow++;
                                Mcomment = Regex.Match(page.Cell(corrRow, 1).GetString().ToUpper(), @"КОРРЕКЦИИ");
                            }
                            while (!(cell.CellBelow().IsEmpty() && cell.CellBelow().CellRight().IsEmpty() && !cell.CellBelow().IsMerged()))
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


                                if (phoneNumber != "")
                                {
                                    var CellPhoneNumber = cell.CellBelow();
                                    string link;
                                    if (CellPhoneNumber.HasHyperlink)
                                        link = CellPhoneNumber.GetHyperlink().ExternalAddress.AbsoluteUri;
                                    else
                                        link = "";
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
                                          exCall.StartDateAnalyze < DateTime.Today
                                    )
                                        phones.AddCall(new FullCall(phoneNumber, link, page.Name.ToUpper().Trim(), curDate, !m.Success, page.Cell(corrRow, cell.Address.ColumnNumber).GetString(), Manager));

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
