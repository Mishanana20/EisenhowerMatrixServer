///считыванияе текста из документа MS Excel (по ячейкам):
using System;
using System.Runtime.InteropServices;
//using Excel = Microsoft.Office.Interop.Excel;
using SqlConn;
using ASquare.WindowsTaskScheduler;
using ASquare.WindowsTaskScheduler.Models;
using Microsoft.Win32.TaskScheduler;
using System.Linq;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using MySql.Data.MySqlClient;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace EisenhowerMatrixServer
{
    internal class Program
    {
        [DllImport("kernel32.dll")]
        public static extern bool FreeConsole();

        static void Main(string[] args)
        {
            try
            {
                string ReportPath = ConfigurationManager.AppSettings.Get("pathToExcel");

                MySqlConnection conn = DBUtils.GetDBConnection();
                SQLScripts.Conn(conn);
                //удаляем все записи с бд, чтобы не было дубликатов (ведь могут поменять исходный файл Excel)
                SQLScripts.DeleteAllFromTable(conn);

                string fileName = ReportPath;

                var workbook = new XLWorkbook(new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));//new XLWorkbook(ReportPath);
                var worksheet = workbook.Worksheet(1);
                // получим все строки в файле
                var rows = worksheet.RangeUsed().RowsUsed(); // Skip header row

                int i = 0; int j = 0;
                foreach (var row in rows)
                {
                    //пропускаем 2 заголовочные строки
                    if (i < 2) { i++; }
                    // Вместо строки можно заносить в базу согласно модели.
                    else
                    {
                        SQLScripts.InsertToDataBase( row.Cell(1).Value.ToString(), row.Cell(2).Value.ToString(), row.Cell(3).Value.ToString(), row.Cell(4).Value.ToString(), conn);
                    }
                }
                SQLScripts.ConnClose(conn); 
            }
            catch (Exception ex)
            {
                Console.WriteLine("Файл не найден" + ex);
            }
            finally
            {
            }
        }
    }
}

