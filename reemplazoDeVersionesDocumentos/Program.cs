using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Xml;
using System.Net;
using System.Net.Http;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;

namespace reemplazoDeVersionesDocumentos
{
    class Program
    {

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static string[] Scopes = { SheetsService.Scope.Spreadsheets }; // static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "ReporteOTC";




        static void Main(string[] args)
        {
            Console.WriteLine("\nIniciando ...");

            /**
             * Inicio de la configuración con la 
             * Hoja de Google
             */


            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read)) // Cliente Json Descargado de Google
            {
                string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal); // de Donde jala el Json

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "evazquez@logsys.com.mx",  // Cuenta de google con la que se accedera
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            String spreadsheetId = "1CAeC3cNHTLoevsqyrN5ryurz7m_iuTeMYCpLoYJAnQI"; //ID de Sheets, ultima parte de la URL
            String range = "NuevasVersiones!A2"; //+celda;  // Que Celda se Actualizara
            ValueRange valueRange = new ValueRange();

            var oblist=new List<object>();

            /**
            * Fin de la configuración con la 
            * Hoja de Google
            */

            String p = ConfigurationManager.AppSettings["rutaVersionNueva"];
            String ruta =ConfigurationManager.AppSettings["ruta"];
            String[]nuevos=Directory.GetFiles(p,"*");
            /*
             * Obtiene los documentos que corresponden a las nuevas versiones.
             */
            String path;
            /*
             * Esta variable se utilizara para
             * buscar el path del archivo con la version
             * vieja.
             */
            String nameFile;
            String[] s;
            String rename;

            foreach (String archivo in nuevos)
            {
                oblist = new List<object>();
                /*
                 * Se creara un nuevo objeto de tipo
                 * oblist por cada archivo que se tenga.
                 */
                path = "";
                path=ruta;
                nameFile= Path.GetFileName(archivo);
                s=Directory.GetFiles(path, nameFile , SearchOption.AllDirectories);
                if (s.Length != 0)
                {
                    /**
                     * Si entra al if significa que si encontro al 
                     * archivo con la version vieja.
                     */
                    if (s.Length > 1)
                    {
                        log.Info("\nEl archivo: " + archivo + " se duplica en : ");
                        oblist.Add((object)Path.GetFileName(archivo));
                        foreach (String i in s)
                        {
                            oblist.Add((object)i);
                            log.Info(i);

                        }
                        
                        valueRange.Values = new List<IList<object>> { oblist };
                        
                        SpreadsheetsResource.ValuesResource.AppendRequest update = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
                        update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                        AppendValuesResponse result = update.Execute();

                    }
                    else
                    {
                        rename = Path.GetDirectoryName(s[0]) + "\\" + Path.GetFileNameWithoutExtension(s[0]) + "__" + Path.GetExtension(s[0]);
                        File.Move(archivo, rename);
                        Console.WriteLine(archivo);
                        Console.WriteLine(rename);
                    }
                }
                else
                {
                    /*
                     * Significa que el nombre del archivo con la nueva
                     * versión no coincide con el nombre del archivo de la version
                     * vieja o que el archivo no existe.
                     */
                    log.Info("El archivo: " + nameFile + " no se encuentra ");
                    oblist.Add((object)nameFile);

                    valueRange.Values = new List<IList<object>> { oblist };
                    SpreadsheetsResource.ValuesResource.AppendRequest update = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
                    
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    AppendValuesResponse result = update.Execute();

                }

            }
            Console.ReadKey();

        }
    }
}
