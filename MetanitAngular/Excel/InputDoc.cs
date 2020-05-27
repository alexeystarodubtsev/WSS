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
        public static List<ProcessedCall> GetProcessedCalls(IFormFile fIle)
        {
            using (var stream = fIle.OpenReadStream())
            {
                XLWorkbook wb = new XLWorkbook(stream);
                foreach (var page in wb.Worksheets)
                {
                    ProcessWorkSheet(page);
                }
            }
            return calls;
        }
        static void ProcessWorkSheet(IXLWorksheet sheet)
        {
            var data = sheet.RangeUsed();
            int lastcol = data.LastColumn().ColumnNumber();
            for (int numRow = 2; numRow <= data.LastRow().RowNumber(); numRow++)
            {
                var row = sheet.Row(numRow);
                ProcessedCall call = new ProcessedCall();
                var curCell = row.Cell(1);
                call.Client = curCell.GetString();
                if (curCell.HasHyperlink && curCell.GetHyperlink().IsExternal )
                {
                    call.Link = curCell.Hyperlink.ExternalAddress.AbsoluteUri;
                }
                curCell = row.Cell(lastcol - 4);
                call.Manager = curCell.GetString();
                curCell = curCell.CellRight();
                call.Comment = curCell.GetString();
                curCell = curCell.CellRight();
                call.NoticeCRM = curCell.GetString();
                curCell = curCell.CellRight();
                call.ClientState = curCell.GetString();
                curCell = curCell.CellRight();
                if (curCell.GetString() != "" && curCell.DataType == XLDataType.DateTime)
                    call.StartDateAnalyze = curCell.GetDateTime();
                else
                {
                    if (!DateTime.TryParse(curCell.GetString(), new CultureInfo("ru-RU"), DateTimeStyles.None, out call.StartDateAnalyze))
                        DateTime.TryParse(curCell.GetString(), new CultureInfo("en-US"), DateTimeStyles.None, out call.StartDateAnalyze);

                }

                if (!calls.Exists(c => (c.Client == call.Client && call.Link == "") || (c.Link == call.Link && call.Link !=null) ))
                    calls.Add(call);
                
            }
        }
        public static bool hasPhone(List<ProcessedCall> calls, ProcessedCall phone)
        {
            if (calls.Exists(c => 
                ((c.Client == phone.Client && phone.Link == "") || (c.Link == phone.Link && phone.Link != ""))
                && 
                (Regex.Match(phone.Comment,c.Comment,RegexOptions.IgnoreCase).Success
                   || Regex.Match(c.Comment, phone.Comment, RegexOptions.IgnoreCase).Success
                   || c.Comment == phone.Comment
                  ) 
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
                ((c.Client == phone.Client && phone.Link == "") || (c.Link == phone.Link && phone.Link != ""))
                &&
                (Regex.Match(phone.Comment, c.Comment, RegexOptions.IgnoreCase).Success
                   || Regex.Match(c.Comment, phone.Comment, RegexOptions.IgnoreCase).Success
                   || c.Comment == phone.Comment
                  )
                && !Regex.Match(c.ClientState, "Закрыт", RegexOptions.IgnoreCase).Success
                && (c.StartDateAnalyze < DateTime.Today)
               ).FirstOrDefault();

        }

    }
}
