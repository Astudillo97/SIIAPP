using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace ServicioSiigo
{
    // NOTA: puede usar el comando "Rename" del menú "Refactorizar" para cambiar el nombre de interfaz "Imovcon" en el código y en el archivo de configuración a la vez.
    [ServiceContract]
    public interface Imovcon
    {
        [OperationContract]
        DataTable ConsultaMovimiento(string nombre);

        [OperationContract]
        bool RegistrarMovimiento(DataTable dat, string nombre);

        [OperationContract]
        string CrearExcelModificado(DataTable dat, string input, string ouput);
    }
}
