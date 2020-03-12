using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using MetanitAngular.Models;
using Microsoft.AspNetCore.Http;

namespace MetanitAngular.Parsers
{
    public class XlPhone
    {
        

        public static Tuple<List<Phone>, List<Phone>> getPhone(IFormFileCollection files)
        {
            Dictionary<string, Dictionary<string, Tuple<int, string>>> dict = new Dictionary<string, Dictionary<string, Tuple<int, string>>>();
            Dictionary<string, Tuple<int,string,string>> dictPhone = new Dictionary<string, Tuple<int, string, string>>();
            List<Phone> phones = new List<Phone>();
            List<Phone> OnlyOneCall = new List<Phone>();
             
            foreach (var file in files)
            {
                using (var stream = file.OpenReadStream())
                {
                    XLWorkbook wb = new XLWorkbook(stream);
                    foreach (var page in wb.Worksheets)
                    {
                        if (page.Name.ToUpper() != "СТАТИСТИКА" && page.Name.ToUpper() != "СВОДНАЯ")
                        {
                            if (!dict.ContainsKey(page.Name.ToUpper()))
                            {
                                dict[page.Name.ToUpper()] = new Dictionary<string, Tuple<int, string>>();
                            }
                            IXLCell cell = page.Cell(1, 5);
                            DateTime curDate;
                            DateTime.TryParse(cell.GetValue<string>(), out curDate);
                            string phoneNumber;
                            while (!(cell.IsEmpty() && cell.CellRight().IsEmpty() && !cell.IsMerged()))
                            {
                                if (cell.GetValue<string>() != "")
                                {
                                    DateTime.TryParse(cell.GetValue<string>(), out curDate);
                                }

                                phoneNumber = cell.CellBelow().GetValue<string>().ToUpper();
                                if (phoneNumber != "")
                                {

                                    if (!dict[page.Name.ToUpper()].ContainsKey(phoneNumber))
                                    {
                                        dict[page.Name.ToUpper()][phoneNumber] = new Tuple<int, string>(1, String.Format("{0:dd.MM.yy}", curDate));
                                    }
                                    else
                                    {
                                        int curQty = dict[page.Name.ToUpper()][phoneNumber].Item1;
                                        string curDates = dict[page.Name.ToUpper()][phoneNumber].Item2;
                                        curQty++;
                                        curDates = curDates + ", " + String.Format("{0:dd.MM.yy}", curDate);
                                        dict[page.Name.ToUpper()][phoneNumber] = new Tuple<int, string>(curQty, curDates);
                                    }
                                    if (!dictPhone.ContainsKey(phoneNumber))
                                    {
                                        var data = new Tuple<int, string, string>(1, page.Name.ToUpper(), String.Format("{0:dd.MM.yy}", curDate));
                                        dictPhone[phoneNumber] = data;
                                    }
                                    else
                                    {
                                        dictPhone[phoneNumber] = new Tuple<int, string, string>(3, "", "");
                                    }
                                }

                                cell = cell.CellRight();
                            }

                        }
                    }
                }
                
            }
            
            foreach (var dictStage in dict)
            {
                foreach (var phone in dictStage.Value)
                {
                    if (phone.Value.Item1 > 1)
                    {
                        phones.Add(new Phone(phone.Key, phone.Value.Item1, dictStage.Key, phone.Value.Item2));
                    }
                }
            }
            foreach (var phone in dictPhone)
            {
                if (phone.Value.Item1 == 1)
                {
                    OnlyOneCall.Add(new Phone(phone.Key, phone.Value.Item1, phone.Value.Item2, phone.Value.Item3));
                }

            }
            Tuple<List<Phone>, List<Phone>> returnPhones;
            returnPhones = new Tuple<List<Phone>, List<Phone>>(phones, OnlyOneCall);
                return returnPhones;
        }
    }
}
