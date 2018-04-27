using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Services;

namespace ServicioSiigo
{
    /// <summary>
    /// Descripción breve de MovCont
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente. 
    // [System.Web.Script.Services.ScriptService]
    public class MovCont : WebService,Imovcon
    {
        [WebMethod]
        public DataTable ConsultaMovimiento(string nombre)
        {
            Operaciones op = new Operaciones();
            DataTable DAT= op.getExcelFile(nombre);
            return DAT;
        }
        [WebMethod]
        public bool RegistrarMovimiento(DataTable dat,string nombre)
        {
            Operaciones op = new Operaciones();
            return op.CrearExcel(dat, nombre);
        }

        [WebMethod]
        public string CrearExcelModificado(DataTable dat, string input, string ouput)
        {
            Operaciones op = new Operaciones();
            return op.CrearExcelModificado(dat, input, ouput);
        }
    }
}
