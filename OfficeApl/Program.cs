using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Reflection;
using WorkBook = Microsoft.Office.Interop.Excel.Workbook;
using Word = Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Excel.Application;


using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Npgsql;
namespace OfficeApl
{
    public class DataToBase
    {
        string connectionString = "Server=127.0.0.1;Port=5432;User Id=postgres;Password=supervisor;Database=gisdb;";
        private NpgsqlConnection getConnection(String conString){  
            NpgsqlConnection conn = new NpgsqlConnection(conString);
            return conn;
        }

        public static void saveDataToBase(NpgsqlConnection conn, NpgsqlCommand comm)
        {
            // string sql = "insert into rawdata (id, polygon, clss) values ("+arr[1,1]+", "+arr[1,3]+", 'fff');";
            try
            {
                if (conn.State.ToString().ToLower() == "open")
                {


                    Console.WriteLine(conn.State);
                    Console.WriteLine("закрываем соединение");
                    Console.ReadLine();
                    conn.Close(); //Закрываем соединение.
                    Console.WriteLine(conn.State);
                    Console.ReadLine();
                    comm.Connection = conn;
                    //comm.CommandText = sql;
                    //comm.ExecuteNonQuery();//.ExecuteScalar().ToString(); //Выполняем нашу команду.
                    //comm.Dispose();

                }
                else //(getCon.State.ToString().ToLower() == "closed")
                {
                    Console.WriteLine(conn.State);
                    Console.WriteLine("открываем соединение");
                    Console.ReadLine();
                    conn.Open(); //Открываем соединение.
                    Console.WriteLine(conn.State);
                    Console.ReadLine();

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Ошибка подключения к БД "+ex.Message);
                Console.WriteLine(ex.Message);
                Console.ReadLine();
                conn.Close(); //Закрываем соединение.
            }


        }

    }
    class Program
    {



        public static void loadPostgres(object[,] arr, NpgsqlConnection conn)
        {

        }
        private static object[,] loadCellByCell(int row, int maxColNum, _Worksheet osheet)
        {
            var list = new object[2, maxColNum + 1];
            for (int i = 1; i <= maxColNum; i++)
            {
                var RealExcelRangeLoc = osheet.Range[(object)osheet.Cells[row, i], (object)osheet.Cells[row, i]];
                object valarrCheck;
                try
                {
                    valarrCheck = RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];
                }
                catch
                {
                    valarrCheck = (object)RealExcelRangeLoc.Value2;
                }
                list[1, i] = valarrCheck;
            }
            return list;
        }


        static void Main(string[] args)
        {
        string connectionString = "Server=127.0.0.1;Port=5432;User Id=postgres;Password=supervisor;Database=gisdb;";
        NpgsqlConnection conn = new NpgsqlConnection(connectionString);
        conn.Open();
        NpgsqlCommand comm = new NpgsqlCommand();
        comm.Connection = conn;



             Application ExcelObj = null;
            WorkBook excelbook = null;
        try{

        
            //Word.Application application = new Word.Application();
            //Word.Document document;

            ExcelObj = new Application();
            ExcelObj.DisplayAlerts = false;
            const string f = @"C:\book.xlsx";
            excelbook = ExcelObj.Workbooks.Open(f, 0, true, 5, "", "", false, XlPlatform.xlWindows);

            var sheets = excelbook.Sheets;
            var maxNumSheet = sheets.Count;


            for (int i = 1; i <= maxNumSheet; i++)
                {

                    var osheet = (_Worksheet) excelbook.Sheets[i];
                    Range excelRange = osheet.UsedRange;

                    int maxColNum;
                    int lastRow;
                    try
                    {
                        maxColNum = excelRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
                        lastRow = excelRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                    }
                    catch
                    {
                        maxColNum = excelRange.Columns.Count;
                        lastRow = excelRange.Rows.Count;

                    }

                    for (int l = 1; l <= lastRow; l++)
                    {
                        Range RealExcelRangeLoc = osheet.Range[(object) osheet.Cells[l, 1], (object) osheet.Cells[l, maxColNum]];
                        object[,] valarr = null;
                        try
                        {
                            var valarrCheck = RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];
                            if (valarrCheck is object[,] || valarrCheck == null)
                                valarr = (object[,]) RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];

                            Console.WriteLine(valarr[1, 1] + " " + valarr[1, 3]);

                            string sql = "insert into rawdata (id, polygon, clss) values ('" + valarr[1, 1] + "', '" + valarr[1, 3] + "', '" + valarr[1, 5] + "');";
                            comm.CommandText = sql;
                            comm.ExecuteNonQuery();//.ExecuteScalar().ToString(); //Выполняем нашу команду.
                            comm.Dispose();

                            
                        }
                        catch
                        {
                            valarr = loadCellByCell(l, maxColNum, osheet);
                        }
                        //SaveDataToBase(valarr);

                        }
                }
        }
        finally
        {
            conn.Close();
            if (excelbook != null)
            {
                excelbook.Close();
                Marshal.ReleaseComObject(excelbook);
            }
            if (ExcelObj != null) ExcelObj.Quit();
        }
        }
    }
}
