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
    public class MovCont : WebService, Imovcon
    {

        [WebMethod]
        public DataTable ConsultaMovimiento(string nombre)
        {
            EPPLUS EP = new EPPLUS();
            return EP.LeerExcel(nombre);
        }
        [WebMethod]
        public string RegistrarMovimiento(string plnt, DataTable dat, string path)
        {
            EPPLUS EP = new EPPLUS();
            return EP.CrearExcel(plnt, dat, path);
        }


    }
}
