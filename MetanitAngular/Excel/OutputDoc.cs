using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static MetanitAngular.Excel.DataStructsForPrintCalls;

namespace MetanitAngular.Excel
{
    public class OutputDoc
    {
        XLWorkbook wbout = new XLWorkbook();
        IXLWorksheet worksheet;
        List<ProcessedCall> ProccessedCalls;
        public void FillIncoming(List<DataStructsForPrintCalls.CallIncoming> Calls) //Нужно видеть входящие звонки на которые не перезвонили
        {
            worksheet = wbout.Worksheets.Add("Вх, на которые не перезвон"); //Нужно видеть входящие звонки на которые не перезвонили
            worksheet.Cell(1, 1).Value = "Клиент";
            worksheet.Cell(1, 2).Value = "Дата";
            worksheet.Cell(1, 3).Value = "Ответственный";
            worksheet.Cell(1, 4).Value = "Примечание";
            worksheet.Cell(1, 5).Value = "Примечание по CRM";
            worksheet.Cell(1, 6).Value = "В работе или закрыт";
            worksheet.Cell(1, 7).Value = "Дата назначенного контакта или дата закрытия сделки";
            
            int curRow = 1;
            foreach (var phone in Calls)
            {
                curRow++;
                worksheet.Cell(curRow, 1).Value = phone.phoneNumber;
                worksheet.Cell(curRow, 1).Hyperlink = phone.Link;
                worksheet.Cell(curRow, 2).SetValue<string>(phone.date);
                //worksheet.Cell(curRow, 2).Value = phone.date;
                worksheet.Cell(curRow, 4).Value = phone.comment;
                worksheet.Cell(curRow, 3).Value = phone.Manager;
                
                if (phone.DealState != "" && phone.DealState != null)
                {
                    worksheet.Cell(curRow, 6).Value = phone.DealState;
                    worksheet.Cell(curRow, 6).Style.Font.Italic = true;
                    worksheet.Cell(curRow, 6).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(curRow, 7).SetValue<string>(phone.DateDeal);
                    worksheet.Cell(curRow, 5).Value = phone.NoticeCRM;
                    worksheet.Cell(curRow, 5).Style.Font.Italic = true;
                }
                else
                {
                    worksheet.Cell(curRow, 5).Value = phone.call.NoticeCRM;
                    if (phone.call.ClientState != "")
                    {
                        worksheet.Cell(curRow, 6).Value = phone.call.ClientState;
                        worksheet.Cell(curRow, 6).Style.Font.FontColor = XLColor.Red;
                    }
                    if (phone.call.StartDateAnalyze.Year > 2000)
                    {
                        worksheet.Cell(curRow, 7).SetValue<string>(String.Format("{0:dd.MM.yyyy}", phone.call.StartDateAnalyze));
                    }
                }
                //worksheet.Cell(curRow, 7).Style.NumberFormat.NumberFormatId = 14;
            }
            RangeSheets(curRow, 7);


        }
        public void FillOutGoingPerWeeks(List<DataStructsForPrintCalls.CallPerWeek> CallsPerWeek,bool DS = false) //2. нужно вдеть исходящие звонки которе сделаны всего один раз за 2 недели или за неделю( в зависимости от специфики)
        {
            worksheet = wbout.Worksheets.Add("Сделанные раз за 3,4 недели"); //2. нужно вдеть исходящие звонки которе сделаны всего один раз за 2 недели или за неделю( в зависимости от специфики)
                                                                             //создадим заголовки у столбцов
            int curCol = 1;
            worksheet.Cell(1, curCol++).Value = "Клиент";
            
            if (!DS)
            {
                worksheet.Cell(1, curCol++).Value = "3 неделя";

                worksheet.Cell(1, curCol++).Value = "4 неделя";
            }
            else
            {
                worksheet.Cell(1, curCol++).Value = "60 дней";
            }

            worksheet.Cell(1, curCol++).Value = "Ответственный";
            worksheet.Cell(1, curCol++).Value = "Примечание";

            worksheet.Cell(1, curCol++).Value = "Примечание по CRM";
            worksheet.Cell(1, curCol++).Value = "В работе или закрыт";
            worksheet.Cell(1, curCol++).Value = "Дата назначенного контакта или дата закрытия сделки";

            int curRow = 1;
            
            foreach (DataStructsForPrintCalls.CallPerWeek phone in CallsPerWeek)
            {
                curRow++;
                curCol = 1;
                worksheet.Cell(curRow, curCol).Value = phone.phoneNumber;
                worksheet.Cell(curRow, curCol++).Hyperlink = phone.Link;
                worksheet.Cell(curRow, curCol++).Value = phone.FirstWeek;
                if (!DS)
                {
                    worksheet.Cell(curRow, curCol++).Value = phone.SecondWeek;
                }
                //worksheet.Cell(curRow, 4).Value = phone.ThirdWeek;
                worksheet.Cell(curRow, curCol++).Value = phone.Manager;
                worksheet.Cell(curRow, curCol).Value = phone.comment;
                if (phone.SecondWeek == "-" && !DS)
                {
                    worksheet.Cell(curRow, curCol).Style.Fill.BackgroundColor = XLColor.Red;
                }
                curCol++;
                
                if (phone.DealState != "" && phone.DealState != null)
                {
                    worksheet.Cell(curRow, curCol + 1).Value = phone.DealState;
                    worksheet.Cell(curRow, curCol + 1).Style.Font.Italic = true;
                    worksheet.Cell(curRow, curCol + 2).SetValue<string>(phone.DateDeal);

                    worksheet.Cell(curRow, curCol).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(curRow, curCol++).Value = phone.NoticeCRM;
                    worksheet.Cell(curRow, curCol).Style.Font.Italic = true;
                }
                else
                {
                    worksheet.Cell(curRow, curCol++).Value = phone.call.NoticeCRM;

                    if (phone.call.ClientState != "")
                    {
                        worksheet.Cell(curRow, curCol).Value = phone.call.ClientState;
                        worksheet.Cell(curRow, curCol).Style.Font.FontColor = XLColor.Red;
                    }
                    curCol++;
                    if (phone.call.StartDateAnalyze.Year > 2000)
                    {
                        worksheet.Cell(curRow, curCol).SetValue<string>(String.Format("{0:dd.MM.yyyy}", phone.call.StartDateAnalyze));
                    }
                }
                
                
                
                //worksheet.Cell(curRow, curCol).Style.NumberFormat.NumberFormatId = 14;


            }

            RangeSheets(curRow, curCol);
        }
        public void FillCallsOnSameStage(List<DataStructsForPrintCalls.CallOneStage> CallsOneStage)   //4. Видеть звонки которые задержались на одном и том же этапе
        {
            worksheet = wbout.Worksheets.Add("Застрявшие на 1 этапе"); //4. Видеть звонки которые задержались на одном и том же этапе

            worksheet.Cell(1, 1).Value = "Клиент";
            worksheet.Cell(1, 2).Value = "Этап";
            worksheet.Cell(1, 3).Value = "Количество повторных звонков";
            worksheet.Cell(1, 4).Value = "Даты звонков";
            worksheet.Cell(1, 5).Value = "Ответственный";
            worksheet.Cell(1, 6).Value = "Примечание последнего звонка";
            worksheet.Cell(1, 7).Value = "Примечание по CRM";
            worksheet.Cell(1, 8).Value = "В работе или закрыт";
            worksheet.Cell(1, 9).Value = "Дата назначенного контакта или дата закрытия сделки";

            int curRow = 1;
            foreach (DataStructsForPrintCalls.CallOneStage phone in CallsOneStage)
            {
                curRow++;
                worksheet.Cell(curRow, 1).Value = phone.phoneNumber;
                worksheet.Cell(curRow, 1).Hyperlink = phone.Link;
                worksheet.Cell(curRow, 2).Value = phone.stage;
                worksheet.Cell(curRow, 3).Value = phone.qty;
                worksheet.Cell(curRow, 4).SetValue<string>(phone.date);
                //worksheet.Cell(curRow, 4).Value = phone.date;
                worksheet.Cell(curRow, 5).Value = phone.Manager;
                worksheet.Cell(curRow, 6).Value = phone.comment;
                
                if (phone.DealState != "" && phone.DealState != null)
                {
                    worksheet.Cell(curRow, 8).Value = phone.DealState;
                    worksheet.Cell(curRow, 8).Style.Font.Italic = true;
                    worksheet.Cell(curRow, 8 + 1).SetValue<string>(phone.DateDeal);

                    worksheet.Cell(curRow, 8).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(curRow, 7).Value = phone.NoticeCRM;
                    worksheet.Cell(curRow, 7).Style.Font.Italic = true;
                }
                else
                {
                    worksheet.Cell(curRow, 7).Value = phone.call.NoticeCRM;
                    if (phone.call.ClientState != "")
                    {
                        worksheet.Cell(curRow, 8).Value = phone.call.ClientState;
                        worksheet.Cell(curRow, 8).Style.Font.FontColor = XLColor.Red;
                    }
                    if (phone.call.StartDateAnalyze.Year > 2000)
                    {
                        worksheet.Cell(curRow, 9).SetValue<string>(String.Format("{0:dd.MM.yyyy}", phone.call.StartDateAnalyze));
                    }
                }
                //worksheet.Cell(curRow, 9).Style.NumberFormat.NumberFormatId = 14;
            }
            RangeSheets(curRow, 9);
        }
        public void FillCallsWithoutAgreement(List<DataStructsForPrintCalls.CallPreAgreement> CallsPreAgreement) //Не закрытые в договор
        {
            worksheet = wbout.Worksheets.Add("Не закрытые в договор"); //5. не закрытые в договор

            worksheet.Cell(1, 1).Value = "Клиент";
            worksheet.Cell(1, 2).Value = "Этап";

            worksheet.Cell(1, 3).Value = "Дата последнего звонка";
            worksheet.Cell(1, 4).Value = "Ответственный";
            worksheet.Cell(1, 5).Value = "Примечание";
            worksheet.Cell(1, 6).Value = "Примечание по CRM";
            worksheet.Cell(1, 7).Value = "В работе или закрыт";
            worksheet.Cell(1, 8).Value = "Дата назначенного контакта или дата закрытия сделки";
            int curRow = 1;
            foreach (var phone in CallsPreAgreement)
            {
                curRow++;
                worksheet.Cell(curRow, 1).Value = phone.phoneNumber;
                worksheet.Cell(curRow, 1).Hyperlink = phone.Link;
                worksheet.Cell(curRow, 2).Value = phone.stage;
                //worksheet.Cell(curRow, 3).Value = phone.date;
                worksheet.Cell(curRow, 3).SetValue<string>(phone.date);
                worksheet.Cell(curRow, 4).Value = phone.Manager;
                worksheet.Cell(curRow, 5).Value = phone.comment;
                
                if (phone.DealState != "" && phone.DealState != null)
                {
                    worksheet.Cell(curRow, 7).Value = phone.DealState;
                    worksheet.Cell(curRow, 7).Style.Font.Italic = true;
                    worksheet.Cell(curRow, 7 + 1).SetValue<string>(phone.DateDeal);

                    worksheet.Cell(curRow, 7).Style.Font.FontColor = XLColor.Black;
                    worksheet.Cell(curRow, 6).Value = phone.NoticeCRM;
                    worksheet.Cell(curRow, 6).Style.Font.Italic = true;
                }
                else
                {
                    worksheet.Cell(curRow, 6).Value = phone.call.NoticeCRM;
                    if (phone.call.ClientState != "")
                    {
                        worksheet.Cell(curRow, 7).Value = phone.call.ClientState;
                        worksheet.Cell(curRow, 7).Style.Font.FontColor = XLColor.Red;
                    }
                    if (phone.call.StartDateAnalyze.Year > 2000)
                    {
                        worksheet.Cell(curRow, 8).SetValue<string>(String.Format("{0:dd.MM.yyyy}", phone.call.StartDateAnalyze));
                    }
                }
                //worksheet.Cell(curRow, 8).Style.NumberFormat.NumberFormatId = 14;
            }

            RangeSheets(curRow, 8);
        }
        public void FillArchive() //Архив
        {
            worksheet = wbout.Worksheets.Add("Архив"); //5. Архив

            worksheet.Cell(1, 1).Value = "Клиент";
            
            worksheet.Cell(1, 2).Value = "Ответственный";
            worksheet.Cell(1, 3).Value = "Примечание";
            worksheet.Cell(1, 4).Value = "Примечание по CRM";
            worksheet.Cell(1, 5).Value = "В работе или закрыт";
            worksheet.Cell(1, 6).Value = "Дата назначенного контакта или дата закрытия сделки";
            int curRow = 1;
            foreach (var call in ProccessedCalls)
            {
                curRow++;
                worksheet.Cell(curRow, 1).Value = call.Client;
                if (call.Link != "" && call.Link != null)
                  worksheet.Cell(curRow, 1).Hyperlink = new XLHyperlink(call.Link);
                if (call.Client == "8 (950) 461-54-94")
                {

                }
                worksheet.Cell(curRow, 2).Value = call.Manager;
                worksheet.Cell(curRow, 3).Value = call.Comment;
                worksheet.Cell(curRow, 4).Value = call.NoticeCRM;
                worksheet.Cell(curRow, 5).Value = call.ClientState;
                //worksheet.Cell(curRow, 6).Style.NumberFormat.NumberFormatId = 14;
                if (call.StartDateAnalyze.Year > 2000)
                  worksheet.Cell(curRow, 6).SetValue<string>(String.Format("{0:dd.MM.yyyy}", call.StartDateAnalyze));
                
            }

            RangeSheets(curRow, 6);
        }
        public void RangeSheets(int row, int col) //Установка размеров
        {
            var rngTable = worksheet.Range(1, 1, row, col);
            rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Range(1, 1, 1, col).Style.Font.Bold = true;
            
            //worksheet.Columns(1, col - 5).AdjustToContents(); //ширина столбца
            worksheet.Columns(1, col).Width = 15;
            worksheet.Column(col - 3).Width = 60;
            worksheet.Column(col - 2).Width = 40;
            worksheet.Column(1).Width = 28;
            worksheet.Columns(1, col).Style.Alignment.WrapText = true;

        }
        public void setProcessedCalls(List<ProcessedCall> proccessedCalls)
        {
            ProccessedCalls = proccessedCalls;
        }
        public XLWorkbook getFile()
        {
            return wbout;
        }

    }
}
