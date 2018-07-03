using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;


namespace ServicioSiigo
{
    public class Operaciones
    {

        public DataTable getExcelFile(string input)
        {
            string url = "C:\\inetpub\\wwwroot\\ServicioSiigo\\SiigoServicio\\" + input + ".xls";
            DataTable dataTable = new DataTable("DATA");
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(url);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                DataRow dat = dataTable.NewRow();
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (i == 1)
                    {
                        dataTable.Columns.Add(xlRange.Cells[i, j].Value2.ToString());
                    }
                    else
                    {
                        dat[j - 1] = xlRange.Cells[i, j].Value2.ToString();
                    }
                }
                dataTable.Rows.Add(dat);
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            if (dataTable != null)
            {
                return dataTable;
            }
            else
            {
                return new DataTable();
            }
        }

        public bool CrearExcel(DataTable dat, string ouput)
        {
            Excel.Application xlApp = new Excel.Application();

            if (xlApp == null)
            {
                //MessageBox.Show("Excel is not properly installed!!");
                return false;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dat.Rows.Count; i++)
            {
                for (int j = 0; j < dat.Rows[i].ItemArray.Length; j++)
                {
                    xlWorkSheet.Cells[i, j] = dat.Rows[i][j];
                }
            }

            //xlWorkSheet.Cells[1, 1] = "ID";
            //xlWorkSheet.Cells[1, 2] = "Name";
            //xlWorkSheet.Cells[2, 1] = "1";
            //xlWorkSheet.Cells[2, 2] = "One";
            //xlWorkSheet.Cells[3, 1] = "2";
            //xlWorkSheet.Cells[3, 2] = "Two";



            xlWorkBook.SaveAs(@"C:\SiigoServicio\" + ouput, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return true;
        }
        public string crearnombre(string param)
        {
            DateTime dateTime = DateTime.Now;
            return param + "" + dateTime.Year + "" + dateTime.Month + "" + dateTime.Day + "" + dateTime.Hour + "" + dateTime.Minute + ".xls";
        }
        public string CrearExcelModificado(DataTable dat, string input, string ouput)
        {
            string url = "C:\\inetpub\\wwwroot\\ServicioSiigo\\SiigoServicio" + input + ".xls";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(url);
            Excel._Worksheet xlWorkSheet = xlWorkBook.Sheets[1];

            string direccion = "C:\\inetpub\\wwwroot\\ServicioSiigo\\SiigoServicio" + crearnombre(ouput);
            if (xlApp == null)
            {
                //MessageBox.Show("Excel is not properly installed!!");
                return "Never";
            }

            for (int i = 0; i < dat.Rows.Count; i++)
            {
                for (int j = 0; j < dat.Rows[i].ItemArray.Length; j++)
                {
                    xlWorkSheet.Cells[i + 5, j + 1] = dat.Rows[i][j];
                }
            }
            xlWorkBook.SaveAs(direccion);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlWorkSheet);

            //close and release
            xlWorkBook.Close();
            Marshal.ReleaseComObject(xlWorkBook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return direccion;
        }
    }
}