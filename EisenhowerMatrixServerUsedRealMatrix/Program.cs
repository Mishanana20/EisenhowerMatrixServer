///считыванияе текста из документа MS Excel (по ячейкам):
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using SqlConn;
using ASquare.WindowsTaskScheduler;
using ASquare.WindowsTaskScheduler.Models;

namespace EisenhowerMatrixServer
{
    internal class Program
    {
        //определить какие аргументы следует передавать. Вроде это было Частота обращения к Excel таблице.
        //и возможно задание критериев? или лучше было бы их сделать в конфиге
        static void Main(string[] args)
        {
            //путь к файлу 
            //string FileName = @"C:\Users\user\Desktop\первая задача\массив.xlsx";
            string FileName = @"C:\Users\Home\Desktop\первая задача\массив.xlsx";
            object readOnly = true;
            object SaveChanges = false;
            //для предоставления отсутствующих значений, параметры которых будут вызваны по умолчанию
            object MissingObj = System.Reflection.Missing.Value;

            //создаем новый объект Excel формата
            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            try
            {
                //Если метод UpdateLink вызван без параметров, Excel по умолчанию обновит все ссылки на таблицу.
                workbooks = app.Workbooks;
                workbook = workbooks.Open(FileName, MissingObj, readOnly, //задали путь к файлу и только для чтения
                           MissingObj, MissingObj, MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                            MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);

                // Получение всех страниц документа. У нас всего одна такая страница
                sheets = workbook.Sheets;

                //удаляем все записи с таблицы, чтобы не было дубликатов (смотреть сразу чтобы это были уникальные значениея?)
                SQLScripts.DeleteAllFromTable();

                foreach (Excel.Worksheet worksheet in sheets)
                {
                    // Получаем диапазон используемых на странице ячеек
                    Excel.Range UsedRange = worksheet.UsedRange;
                    // Получаем строки в используемом диапазоне
                    Excel.Range urRows = UsedRange.Rows;
                    // Получаем столбцы в используемом диапазоне
                    Excel.Range urColums = UsedRange.Columns;

                    // Количества строк и столбцов
                    int RowsCount = urRows.Count;
                    int ColumnsCount = urColums.Count;


                    string id = null; string name = null; string value = null; string valueTime = null;
                    for (int i = 5; i <= RowsCount; i++) //пропускаем строку с названием
                    {
                        name = null; value = null; id = null; valueTime = null;
                        for (int j = 1; j <= ColumnsCount; j++)
                        {
                            Excel.Range CellRange = UsedRange.Cells[i, j];
                            if (j == 1) //если это не столбец с датой и временем
                            {
                                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                (CellRange as Excel.Range).Value2.ToString();
                                if (CellText != null)
                                {
                                    id = CellText;
                                    Console.Write($"{CellText} \t");
                                }
                            }
                            if (j == 2) //если это не столбец с датой и временем
                            {
                                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                (CellRange as Excel.Range).Value2.ToString();
                                if (CellText != null)
                                {
                                    name = CellText;
                                    Console.Write($"{CellText} \t \t \t");
                                }
                            }
                            if (j == 3)
                            {
                                //да, дублирую, а что?
                                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                (CellRange as Excel.Range).Value2.ToString();
                                if (CellText != null)
                                {
                                    value = CellText;
                                    Console.Write($"{CellText} \t");
                                }
                            }
                            if (j == 4)
                            {
                                //да, дублирую, а что?
                                string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                (CellRange as Excel.Range).Value2.ToString();
                                if (CellText != null)
                                {
                                    valueTime = CellText;
                                    Console.Write($"{CellText} \t \n");
                                }
                            }
                            
                        }
                        SQLScripts.InsertToDataBase(id, name, value, valueTime);
                    }
                        // Очистка неуправляемых ресурсов на каждой итерации
                        if (urRows != null) Marshal.ReleaseComObject(urRows);
                        if (urColums != null) Marshal.ReleaseComObject(urColums);
                        if (UsedRange != null) Marshal.ReleaseComObject(UsedRange);
                        if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                    }
                
            }
            catch (Exception ex)
            {
                /* Обработка исключений */
                Console.WriteLine("Файл не найден" + ex);
            }
            finally
            {
                /* Очистка оставшихся неуправляемых ресурсов */
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (workbook != null)
                {
                    workbook.Close(SaveChanges);
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }

                if (workbooks != null)
                {
                    workbooks.Close();
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
                Console.ReadKey(true);
            }
        }
    }
}

