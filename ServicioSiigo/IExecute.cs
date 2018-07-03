using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace ServicioSiigo
{
    // NOTA: puede usar el comando "Rename" del menú "Refactorizar" para cambiar el nombre de interfaz "IExecute" en el código y en el archivo de configuración a la vez.
    [ServiceContract]
    public interface IExecute
    {
        [OperationContract]
        string ExecuteCommand(string ExcelSIIGO, string RutaEmpresa, string Año, string PUSHMOV, string Norma, string Usuario, string Clave, string NombreLog, string NombreArchivoExcelEntrada, string encrypt);
        [OperationContract]
        string OperarCmd(string secuencia);
    }
}
