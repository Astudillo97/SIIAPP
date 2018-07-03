using System;
using System.Web.Services;
using System.Configuration;
using System.Threading;
using System.Security.Cryptography;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace ServicioSiigo
{
    /// <summary>
    /// Descripción breve de ExecuteCmd
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente. 
    // [System.Web.Script.Services.ScriptService]
    public class ExecuteCmd : WebService
    {

        [WebMethod]
        public string PUSHMOVIM(string ExcelSIIGO, string RutaEmpresa, string Año, string PUSHMOV, string Norma, string Usuario, string Clave, string NombreLog, string NombreArchivoExcelEntrada, string encrypt)
        {
            try
            {
                string cmdexecute = ExcelSIIGO + " " + RutaEmpresa + " " + Año + " " + PUSHMOV + " " + Norma + " " + Usuario + " " + Clave + " " + NombreLog + " " + NombreArchivoExcelEntrada;

                Process cmd = new Process();
                cmd.StartInfo.FileName = "cmd.exe";
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = false;
                cmd.StartInfo.UseShellExecute = false;
                cmd.StartInfo.UserName = "Administrador";
                System.Security.SecureString theSecureString = new System.Security.SecureString();
                theSecureString.AppendChar('A');
                theSecureString.AppendChar('d');
                theSecureString.AppendChar('m');
                theSecureString.AppendChar('i');
                theSecureString.AppendChar('n');
                theSecureString.AppendChar('.');
                theSecureString.AppendChar('3');
                theSecureString.AppendChar('2');
                theSecureString.AppendChar('1');
                cmd.StartInfo.Password = theSecureString;
                cmd.Start();

                cmd.StandardInput.WriteLine(cmdexecute);
                cmd.StandardInput.Flush();
                cmd.StandardInput.Close();
                cmd.WaitForExit();
                while (VerificarMovimiento(NombreLog) == null)
                {
                    Console.Write("Esperado");
                }
                return VerificarMovimiento(NombreLog);

            }
            catch (Exception ex)
            {

                return ex.Message;
            }
            
        }

        [WebMethod]
        public string GETMOVIM(string ExcelSIIGO, string RutaEmpresa, string Año, string PUSHMOV, string Norma, string Usuario, string Clave, string NombreLog, string parametrizacion, string NombreArchivoExcelEntrada, string encrypt)
        {
            try
            {
                string cmdexecute = ExcelSIIGO + RutaEmpresa + Año + PUSHMOV + Norma + Usuario + Clave + NombreLog  +parametrizacion+ NombreArchivoExcelEntrada;

                Process cmd = new Process();
                cmd.StartInfo.FileName = "cmd.exe";
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = false;
                cmd.StartInfo.UseShellExecute = false;
                cmd.StartInfo.UserName = "Administrador";
                System.Security.SecureString theSecureString = new System.Security.SecureString();
                theSecureString.AppendChar('A');
                theSecureString.AppendChar('d');
                theSecureString.AppendChar('m');
                theSecureString.AppendChar('i');
                theSecureString.AppendChar('n');
                theSecureString.AppendChar('.');
                theSecureString.AppendChar('3');
                theSecureString.AppendChar('2');
                theSecureString.AppendChar('1');
                cmd.StartInfo.Password = theSecureString;
                cmd.Start();

                cmd.StandardInput.WriteLine(cmdexecute);
                cmd.StandardInput.Flush();
                cmd.StandardInput.Close();
                cmd.WaitForExit();
                while (VerificarMovimiento(NombreLog)==null)
                {
                    Console.Write("Esperado");
                }
                return VerificarMovimiento(NombreLog);

            }
            catch (Exception ex)
            {

                return ex.Message;
            }

        }
        public string VerificarMovimiento(string urllog)
        {
            string line;
            string lineas = null;
            try
            {
                System.IO.StreamReader file = new System.IO.StreamReader(urllog);
                //while ((line = file.ReadLine()) != null)
                //{
                //    lineas += line;
                //}
                lineas = file.ReadLine();
                file.Close();
                return lineas;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string Decrypt(string cipherString, bool useHashing)
        {
            byte[] keyArray;
            //get the byte code of the string

            byte[] toEncryptArray = Convert.FromBase64String(cipherString);

            System.Configuration.AppSettingsReader settingsReader =
                                                new AppSettingsReader();
            //Get your key from config file to open the lock!
            string key = (string)settingsReader.GetValue("SecurityKey", typeof(String));

            if (useHashing)
            {
                //if hashing was used get the hash code with regards to your key
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
                //release any resource held by the MD5CryptoServiceProvider

                hashmd5.Clear();
            }
            else
            {
                //if hashing was not implemented get the byte code of the key
                keyArray = UTF8Encoding.UTF8.GetBytes(key);
            }

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            //set the secret key for the tripleDES algorithm
            tdes.Key = keyArray;
            //mode of operation. there are other 4 modes. 
            //We choose ECB(Electronic code Book)

            tdes.Mode = CipherMode.ECB;
            //padding mode(if any extra byte added)
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(
                                 toEncryptArray, 0, toEncryptArray.Length);
            //Release resources held by TripleDes Encryptor                
            tdes.Clear();
            //return the Clear decrypted TEXT
            return UTF8Encoding.UTF8.GetString(resultArray);
        }

    }
}
