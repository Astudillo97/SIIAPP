using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;

namespace ServicioSiigo
{
    public class Conexion
    {

        public DataTable ConsultarExel(string SlnoAbbreviation,string sql)
        {
            OleDbConnection oledbConn = null;
            try
            {
                oledbConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SlnoAbbreviation + ";Extended Properties = 'Excel 12.0;HDR=YES;IMEX=1;'; ");

                oledbConn.Open();
                OleDbCommand cmd = new OleDbCommand(); ;
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();

                // passing list to drop-down list

                // selecting distinct list of Slno 
                cmd.Connection = oledbConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [Hoja1$]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds);

                // binding form data with grid view
                return ds.Tables[0];

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oledbConn.Close();
            }
        }// close of method GemerateExceLData

        public bool OperarExel(string SlnoAbbreviation, string sql)
        {
            try
            {
                OleDbConnection MyConnection;
                OleDbCommand myCommand = new OleDbCommand();
                MyConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SlnoAbbreviation + ";Extended Properties = 'Excel 12.0;HDR=YES;IMEX=1;'; ");
                MyConnection.Open();
                myCommand.Connection = MyConnection;
                myCommand.CommandText = sql;//"Insert into [Hoja1$] (id,name) values('5','e')"
                myCommand.ExecuteNonQuery();
                MyConnection.Close();
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}