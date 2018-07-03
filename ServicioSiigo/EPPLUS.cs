using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
namespace ServicioSiigo
{
    public class EPPLUS
    {


        private DataTable WorksheetToDataTable(ExcelWorksheet oSheet)
        {
            int totalRows = oSheet.Dimension.End.Row;
            int totalCols = oSheet.Dimension.End.Column;
            DataTable dt = new DataTable(oSheet.Name);
            DataRow dr = null;
            for (int i = 1; i <= totalRows; i++)
            {
                if (i > 1) dr = dt.Rows.Add();
                for (int j = 1; j <= totalCols; j++)
                {
                    if (i == 1)
                        dt.Columns.Add(oSheet.Cells[i, j].Value.ToString());
                    else
                        dr[j - 1] = oSheet.Cells[i, j].Value.ToString();
                }
            }
            return dt;
        }
        private ExcelWorksheet DataTableToWorksheet(ExcelWorksheet oSheet,DataTable dat)
        {
            int totalRows = oSheet.Dimension.End.Row;
            int totalCols = oSheet.Dimension.End.Column;
            for (int i = 0; i < dat.Rows.Count; i++)
            {
                for (int j = 0; j < dat.Rows[i].ItemArray.Length; j++)
                {
                    oSheet.Cells[i + 5, j + 1].Value = dat.Rows[i][j];
                }
            }
            return oSheet;
        }

        public ExcelPackage crearExel(string nombre)
        {
            string name = "C:\\inetpub\\wwwroot\\" + nombre + ".xls";
            ExcelPackage excel = new ExcelPackage(new FileInfo(name));
            excel.Workbook.Worksheets.Add("Hoja prueba");
            excel.Save();
            return excel;
        }
        public string crearnombre(string param)
        {
            DateTime dateTime = DateTime.Now;
            return param + "" + dateTime.Year + "" + dateTime.Month + "" + dateTime.Day + "" + dateTime.Hour + "" + dateTime.Minute + ".xls";
        }
        public DataTable LeerExcel(string path)
        {
            try
            {
                ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path));
                return WorksheetToDataTable(excelPackage.Workbook.Worksheets[1]);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string CrearExcel(string plantilla,DataTable entrada, string path)
        {
            try
            {
                string pat = crearnombre(path+"PUSHMOV");
                ExcelPackage excelPackage = new ExcelPackage(new FileInfo(plantilla));
                var worksheet = DataTableToWorksheet(excelPackage.Workbook.Worksheets[1], entrada);
                excelPackage.SaveAs(new FileInfo(pat));
                return pat;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}