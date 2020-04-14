using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetanitAngular.Excel
{
    public class OutputDoc
    {
        XLWorkbook wbout = new XLWorkbook();
        IXLWorksheet worksheet;
        public void FillIncoming(List<DataStructsForPrintCalls.CallIncoming> Calls) //Нужно видеть входящие звонки на которые не перезвонили
        {
            worksheet = wbout.Worksheets.Add("Вх, на которые не перезвон"); //Нужно видеть входящие звонки на которые не перезвонили
            worksheet.Cell(1, 1).Value = "Клиент";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 2).Value = "Дата";
            worksheet.Cell(1, 2).Style.Font.Bold = true;
            worksheet.Cell(1, 3).Value = "Примечание";
            worksheet.Cell(1, 3).Style.Font.Bold = true;
            int curRow = 1;
            foreach (var phone in Calls)
            {
                curRow++;
                worksheet.Cell(curRow, 1).Value = phone.phoneNumber;
                worksheet.Cell(curRow, 2).Value = phone.date;
                worksheet.Cell(curRow, 3).Value = phone.comment;
            }
            RangeSheets(curRow, 3);


        }
        public void FillOutGoingPerWeeks(List<DataStructsForPrintCalls.CallPerWeek> CallsPerWeek) //2. нужно вдеть исходящие звонки которе сделаны всего один раз за 2 недели или за неделю( в зависимости от специфики)
        {
            worksheet = wbout.Worksheets.Add("Сделанные раз за 1,2,3 недели"); //2. нужно вдеть исходящие звонки которе сделаны всего один раз за 2 недели или за неделю( в зависимости от специфики)
                                                                               //создадим заголовки у столбцов
            worksheet.Cell(1, 1).Value = "Клиент";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 2).Value = "1 неделя";
            worksheet.Cell(1, 2).Style.Font.Bold = true;
            worksheet.Cell(1, 3).Value = "2 неделя";
            worksheet.Cell(1, 3).Style.Font.Bold = true;
            worksheet.Cell(1, 4).Value = "3 неделя";
            worksheet.Cell(1, 4).Style.Font.Bold = true;
            worksheet.Cell(1, 5).Value = "Примечание";
            worksheet.Cell(1, 5).Style.Font.Bold = true;

            int curRow = 1;
            
            foreach (DataStructsForPrintCalls.CallPerWeek phone in CallsPerWeek)
            {
                curRow++;
                worksheet.Cell(curRow, 1).Value = phone.phoneNumber;
                worksheet.Cell(curRow, 2).Value = phone.FirstWeek;
                worksheet.Cell(curRow, 3).Value = phone.SecondWeek;
                worksheet.Cell(curRow, 4).Value = phone.ThirdWeek;
                worksheet.Cell(curRow, 5).Value = phone.comment;
                if (phone.ThirdWeek == "-")
                {
                    worksheet.Cell(curRow, 5).Style.Fill.BackgroundColor = XLColor.Red;
                }
            }

            RangeSheets(curRow, 5);
        }
        public void FillCallsOnSameStage(List<DataStructsForPrintCalls.CallOneStage> CallsOneStage)   //4. Видеть звонки которые задержались на одном и том же этапе
        {
            worksheet = wbout.Worksheets.Add("Застрявшие на 1 этапе"); //4. Видеть звонки которые задержались на одном и том же этапе

            worksheet.Cell(1, 1).Value = "Клиент";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 2).Value = "Этап";
            worksheet.Cell(1, 2).Style.Font.Bold = true;
            worksheet.Cell(1, 3).Value = "Количество повторных звонков";
            worksheet.Cell(1, 3).Style.Font.Bold = true;
            worksheet.Cell(1, 4).Value = "Даты звонков";
            worksheet.Cell(1, 4).Style.Font.Bold = true;
            worksheet.Cell(1, 5).Value = "Примечание последнего звонка";
            worksheet.Cell(1, 5).Style.Font.Bold = true;

            int curRow = 1;
            foreach (DataStructsForPrintCalls.CallOneStage phone in CallsOneStage)
            {
                curRow++;
                worksheet.Cell(curRow, 1).Value = phone.phoneNumber;
                worksheet.Cell(curRow, 2).Value = phone.stage;
                worksheet.Cell(curRow, 3).Value = phone.qty;
                worksheet.Cell(curRow, 4).Value = phone.date;
                worksheet.Cell(curRow, 5).Value = phone.comment;
            }
            RangeSheets(curRow, 5);
        }
        public void FillCallsWithoutAgreement(List<DataStructsForPrintCalls.CallPreAgreement> CallsPreAgreement) //Не закрытые в договор
        {
            worksheet = wbout.Worksheets.Add("Не закрытые в договор"); //5. не закрытые в договор

            worksheet.Cell(1, 1).Value = "Клиент";
            worksheet.Cell(1, 1).Style.Font.Bold = true;
            worksheet.Cell(1, 2).Value = "Этап";
            worksheet.Cell(1, 2).Style.Font.Bold = true;

            worksheet.Cell(1, 3).Value = "Дата последнего звонка";
            worksheet.Cell(1, 3).Style.Font.Bold = true;
            worksheet.Cell(1, 4).Value = "Примечание";
            worksheet.Cell(1, 4).Style.Font.Bold = true;
            int curRow = 1;
            foreach (var phone in CallsPreAgreement)
            {
                curRow++;
                worksheet.Cell(curRow, 1).Value = phone.phoneNumber;
                worksheet.Cell(curRow, 2).Value = phone.stage;
                worksheet.Cell(curRow, 3).Value = phone.date;
                worksheet.Cell(curRow, 4).Value = phone.comment;
            }

            RangeSheets(curRow, 4);
        }
        public void RangeSheets(int row, int col) //Установка размеров
        {
            var rngTable = worksheet.Range(1, 1, row, col);
            rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Column(col).Width = 60;
            worksheet.Column(col).Style.Alignment.WrapText = true;
            worksheet.Columns(1, col - 1).AdjustToContents(); //ширина столбца
        }
        public XLWorkbook getFile()
        {
            return wbout;
        }

    }
}
