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

namespace MetanitAngular.Excel
{
    public static class InputDoc
    {
        static List<ProcessedCall> calls = new List<ProcessedCall>();
        public static List<ProcessedCall> GetProcessedCalls(IFormFile fIle, bool Belfan = false)
        {
            using (var stream = fIle.OpenReadStream())
            {
                XLWorkbook wb = new XLWorkbook(stream);
                foreach (var page in wb.Worksheets)
                {
                    ProcessWorkSheet(page, Belfan);
                }
            }
            return calls;
        }
        static void ProcessWorkSheet(IXLWorksheet sheet, bool Belfan)
        {
            var data = sheet.RangeUsed();
            int lastcol = data.LastColumn().ColumnNumber();
            for (int numRow = data.LastRow().RowNumber(); numRow >=2; numRow--)
            {
                var row = sheet.Row(numRow);
                ProcessedCall call = new ProcessedCall();
                var curCell = row.Cell(1);
                call.Client = curCell.GetString();
                if (curCell.HasHyperlink && curCell.GetHyperlink().IsExternal )
                {
                    call.Link = curCell.Hyperlink.ExternalAddress.AbsoluteUri;
                }
                curCell = row.Cell(lastcol - 5);
                call.Manager = curCell.GetString();
                curCell = curCell.CellRight(); // -4
                call.Comment = curCell.GetString();
                curCell = curCell.CellRight(); // -3
                call.NoticeCRM = curCell.GetString();  
                curCell = curCell.CellRight(); // -2
                call.ClientState = curCell.GetString();
                curCell = curCell.CellRight(); // -1
                if (curCell.GetString() != "" && curCell.DataType == XLDataType.DateTime)
                    call.StartDateAnalyze = curCell.GetDateTime();
                else
                {
                    if (!DateTime.TryParse(curCell.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out call.StartDateAnalyze))
                        DateTime.TryParse(curCell.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out call.StartDateAnalyze);

                }
                curCell = curCell.CellRight(); // last col
                call.CommentOfCustomer = curCell.GetString();

                if ((!Belfan && !calls.Exists(c => (c.Client == call.Client && call.Link == "") || (c.Link == call.Link && call.Link !=null) )) || (Belfan && !calls.Exists(c => (c.Client == call.Client))))
                    calls.Add(call);
                else
                {
                    ProcessedCall exCall;
                    if (!Belfan)
                    {
                        exCall = calls.Where(c => (c.Client == call.Client && call.Link == "") || (c.Link == call.Link && call.Link != null)).First();
                    }
                    else
                    {
                        exCall = calls.Where(c => (c.Client == call.Client)).First();
                    }
                    if (exCall.StartDateAnalyze < call.StartDateAnalyze)
                    {
                        calls.Remove(exCall);
                        calls.Add(call);
                    }
                    else
                    {
                        if (exCall.CommentOfCustomer == "")
                        {
                            exCall.CommentOfCustomer = call.CommentOfCustomer;
                        }
                    }
                }
                
            }
        }
        public static bool hasPhone(List<ProcessedCall> calls, ProcessedCall phone)
        {
           
            if (calls.Exists(c => 
                ((c.Client == phone.Client && phone.Link == "" && c.Link == "") || (c.Link == phone.Link && phone.Link != ""))
               // && 
                //(
                //Regex.Match(phone.Comment.Substring(0,20),c.Comment.Substring(0, 20), RegexOptions.IgnoreCase).Success
                //   || Regex.Match(c.Comment.Substring(0, 20), phone.Comment.Substring(0, 20), RegexOptions.IgnoreCase).Success
                //   || 
                  // c.Comment == phone.Comment
                //) 
               )
           )
                return true;
            else
            {
               
                return false;   
            }
                
        }
        public static ProcessedCall getSamePhone(List<ProcessedCall> calls, ProcessedCall phone)
        {
            return calls.Where(c =>
                ((c.Client == phone.Client && phone.Link == "" && c.Link == "") || (c.Link == phone.Link && phone.Link != ""))
                //&&
                //(
                ////Regex.Match(phone.Comment.Substring(0, 20), c.Comment.Substring(0, 20), RegexOptions.IgnoreCase).Success
                ////   || Regex.Match(c.Comment.Substring(0, 20), phone.Comment.Substring(0, 20), RegexOptions.IgnoreCase).Success
                ////   ||
                //   c.Comment == phone.Comment
                //  )
                //&& !Regex.Match(c.ClientState, "Закрыт", RegexOptions.IgnoreCase).Success
                //&& (c.StartDateAnalyze < DateTime.Today)
               ).FirstOrDefault();

        }

    }
}
