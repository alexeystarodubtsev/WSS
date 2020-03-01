using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using MetanitAngular.Models;

namespace MetanitAngular.Parsers
{
    public class XlPhone
    {
        public static IEnumerable<Phone> getPhone(string filePath)
        {
            List<Phone> phones = new List<Phone>();

            XLWorkbook wb = new XLWorkbook(filePath);
            
            
            IXLWorksheet sheet = wb.Worksheets.First();
            int curRow = 2;
            int startCol = 2;
            int numColStage = 1;
            IXLCell cell = sheet.Cell(curRow, startCol);
            while (!(cell.GetValue<string>() == ""))
            {
                Dictionary<string, int> dictPhones = new Dictionary<string, int>();
                while (!(cell.GetValue<string>() == ""))
                {
                    string cphone = cell.GetValue<string>();
                    if (!dictPhones.ContainsKey(cphone))
                    {
                        dictPhones[cphone] = 0;
                    }
                    dictPhones[cphone]++;
                    cell = cell.CellRight();

                }
                foreach (var keyValue in dictPhones)
                {
                    if (keyValue.Value > 1)
                    {
                        Phone phone = new Phone();
                        phone.phone = keyValue.Key;
                        phone.qty = keyValue.Value;
                        phone.stage = sheet.Cell(curRow, numColStage).GetValue<string>();
                        phones.Add(phone);
                    }
                }

                curRow++;
                cell = sheet.Cell(curRow, startCol);
            }

            return phones;
        }
    }
}
