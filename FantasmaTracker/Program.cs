using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text;
using System.Globalization;

namespace FantasmaTracker
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static void Main(string[] args)
        {
            //Toma de argumentos e inicialización de variables
            int argumentoMes = 0;
            string idArchivo = "";
            int columnas = 4;
            string logTexto = "";
            string nombreMes = "";

            DateTimeFormatInfo dtinfo = new CultureInfo("es-ES", false).DateTimeFormat;
            nombreMes = dtinfo.GetMonthName(1);

            if (args != null || args.Length > 0)
            {
                try
                {
                    argumentoMes = int.Parse(args[0].ToString());
                    idArchivo = args[1].ToString();
                }
                catch (Exception ex)
                {
                    argumentoMes = 3;
                    //idArchivo = "1WW216OeWLhhHkCo4_4o3NxBJF-ErZcYzfdnqOPZgO6k"; //*ID Tracker original del área de Transformación Digital Grupo Leonera*/
                    idArchivo = "1H5l5OuecIsyZr0Gx9nQ70POM4jqJCgbr3OCZ9zmbN90"; //*ID presupuesto 1*/
                }
            }
            else
            {
                //Si no hay argumentos, el default es de tres días antes del vencimiento del proyecto y el Tracker será el del área Transformación Digital 
                argumentoMes = 1;
                idArchivo = "1H5l5OuecIsyZr0Gx9nQ70POM4jqJCgbr3OCZ9zmbN90";
            }

            //Introducción / Presentación

            
            Console.WriteLine("SERVICIO FANTASMA DE ALERTA PRESUPUESTO");
            Console.WriteLine("====================================================================================");
            Console.WriteLine("Área Transformación Digital, Grupo Leonera.");
            Console.WriteLine("2022 Saúl Candia - Holding Leonera");
            //String spreadsheetId = "1ekSXRUep9T0U7ZuVCDz0fJ7X_trGZp8JmrzL7b8OJb4"; /*ID copia Tracker para pruebas*/
            String spreadsheetId = idArchivo;

            try
            {
                logTexto = DateTime.Now.ToString() + " Servicio Inicializado.";
                WriteLog(logTexto);

                logTexto = DateTime.Now.ToString()+" Contectando a API Google Sheets...";
                WriteLog(logTexto);

                //Conecta a API

                Console.WriteLine("Conectando con API Google Sheets");
                string[] scopes = new string[] { SheetsService.Scope.Drive };
                String serviceAccountEmail = "presupuestoapi@presupuestoapi.iam.gserviceaccount.com";

                //Crea instancia de conexión
                var initializer = new ServiceAccountCredential.Initializer(serviceAccountEmail)
                {
                    Scopes = scopes,
                    User = "scandia@leonera.cl"
                };

                //Consigue clave privada desde servicios JSON
                var credential = new ServiceAccountCredential(initializer.FromPrivateKey("-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDSnp1pfVi3kWiC\nTTqliK71qibzr7Dbrs/8F+iBNTHCkPvWXfxKqtRhGjPh8p/885kMc5lqYKhRUis0\nn1vfmVznqsJMHIn8u1cjxobEV9+W06SQpQKrua5+wGqKZw/bOXK9lV1zY/j3anQ+\nGRL6oetTf43Gbb8BCTkgRRq0sMBS5JNF+d7hBp3eBQ6IyaO8RC+sNg5GqVRLAcb6\nCFOU9QAoRPHZH6mp/N4OvEvUvUJevcPiFniHc2onm480Yc1DIGbgunxw4v9STeQK\nuDcL2jRWKZ+88shOXOj72smRJzKQCJwiTzgcUr71U1zIsbiAsll2el1XlnYWy9vE\n9uzPCaC/AgMBAAECggEABT0IsTz63eXx8Xu2P7O8lkObIGh4P56DccOudrg5+prc\niKJhygGhsqCSNcZxEDuGzPZ7FFg/F3axuGdWQ6Nu2hw3JOl4zR5jtnITnAKLfxbY\nevh/roG5w1FJ1RNnI460Od7jKiGMaaruJTU+cZlhXvxHLG5CV+ZA03qkhWX4Ape9\nO4QgZFQzan9aPzu165ydG1fVsQqjFRmsF06rP104hdEwG6ekKAopHdCHk4TJBHaJ\nqaEz/Sa1FxxgTa/O++eHgHi6I6GDnsyGI3x8ydNgBL1NyZZTt6JmZdo6fpKqoZ+/\nOuYQ04qoaUlToYwq0ukAZYNmk2irbrG+ei0Y727OEQKBgQDvsjY7IYT9tbUlt1TO\nmn/ov29Yt7QCqNHQHLcsWE92BXmXnzOfYPOZAoIruNjxIFCS3R3Zk0GG293t/CM/\nvEfvn/jFjsrKSl4iFZZTqw3g52Pa4QDNK5e2I01VgFqgmSmjgrXgmeN+FECPdoJu\nvAOUT8VWfqOy8+bwmb0hsr0omwKBgQDg8hkXlVKMkqzE6ufuDZRrXlYYXZdhYVyx\n66WP1Jo/18R0KkGCAvCWaAC9KBXc07LnfcSJ2lQWhj/Dcqa7bpQVQBevSc4x1zoN\n+H1ATj1uHTXS2qu6v86d6vyYD/0RcpIxsKvq45NiLwOUpNPS3YZb1+WzKrxJ/VSO\n1Tnu1fqQrQKBgFvhsowkIzimCNR2XFn+O33atDIL6UMDt7nQ6B5lk8AoBR4r9rvn\njDlhDsj3yKFVw80oWaLnobyyV3Y8qr5pzCF87v276Nx2eXMTV1anQWCvEkX67jW3\nuiYljiVyWEsrqxx0pId+NghEdyMHSKRuCek2Uuz/Cn00pZghNrDONVh1AoGAYed9\nHFKVdzFvmNVU1Lt8Wa7ZcglqFaw2mAmkKZGzAQ58JsMtd9Snug7SI4IK4e4R88c9\nf3JTHuqXXg3Mm89pDEa1CEnrQK4YSnRYr2BeREraXkdmbwWEfB8GiXiMAMgI8S+f\n47/hKd6khFGpECHylI7HHs/+24UzBGexq03enJECgYEAmjLcyHatnGORwUszDxnp\nNLYjxXX/5cBgAl46zg1o1QJfaHDWDQEb5h8c3Onjx5eOpk0aGtzlsw6ntg805qrK\ng1/tQC6j6TtrlEtRZ8b852VwZ7iy70Qg0r3Ne94z1xUw9mM3yQZsyzY0DmRQHAEo\nBtL4IAolVS5C5wwEq8mGPh8=\n-----END PRIVATE KEY-----\n"));

                //Crea una instancia con el servicio API
                Google.Apis.Sheets.v4.SheetsService service = new Google.Apis.Sheets.v4.SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    //ApplicationName = "GoogleDriveRestAPI-v3",
                    ApplicationName = "DriveAPI",
                });
                service.HttpClient.Timeout = TimeSpan.FromMinutes(20);

                String range = "Reporte!D6:AM6";

                SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

                var response = request.Execute();

                IList<IList<Object>> values = response.Values;



                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nConectado exitosamente a la API\n");
                Console.WriteLine("ID Documento: " + spreadsheetId);
                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nLeyendo presupuestos... espere por favor...");

                logTexto = DateTime.Now.ToString() + " Contectado exitosamente a API Google Sheets";
                WriteLog(logTexto);

                logTexto = DateTime.Now.ToString() + " ID Documento: " + spreadsheetId;
                WriteLog(logTexto);

                //Declaración de variables
                int numFila = 1;
                int limite = values.Count;
                var hoy = DateTime.Today;
                int contadorElementos = 0;
                Boolean result = false;

                //Declaración de matriz bidimensional [cantidad elementos encontrados en Hoja, 7]
                //      0                   1                   2                       3                   4                   5                   6               7
                //[nombreProyecto], [fechaCompromiso], [nombreClienteInterno], [nombreResponsable], [correoResponsable], [estadoProyecto], [rowPositionIndex], [diasDiferencia]
                string[,] datosTracker = new string[limite, columnas];

                //PARTE 1: REVISA EL ESTADO DE TAREAS Y DETERMINA POSICIÓN DE FILA X EN LAS MISMAS

                logTexto = DateTime.Now.ToString() + " PARTE 1: Revisa el estado de tareas y determina ubicación en las mismas.";
                WriteLog(logTexto);

                if (values != null && values.Count > 0)
                {
                    int rowPosition = 0;
                    foreach (var rowEstado in values)
                    {
                        datosTracker[rowPosition, 5] = rowEstado[0].ToString();
                        datosTracker[rowPosition, 6] = numFila.ToString();
                        rowPosition = rowPosition + 1;
                        numFila = numFila + 1;
                    }
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío y se cerrará la aplicación.");
                    logTexto = DateTime.Now.ToString() + " El tracker se encuentra vacío y se cerrará la aplicación.";
                    WriteLog(logTexto);
                    Console.WriteLine("Esta ventana se cerrará automáticamente en cuatro (4) segundos");
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("© 2022 Saúl Candia");

                    //Contador para cierre de app consola:
                    Thread.Sleep(4000);
                    Environment.Exit(0);
                }

                logTexto = DateTime.Now.ToString() + " PARTE 1: Resultado Exitoso.";
                WriteLog(logTexto);

                //PARTE 2: REVISA EL NOMBRE DE LAS TAREAS, LA FECHA DE COMPROMISO INICIAL Y LOS DIAS DE DIFERENCIA ENTRE HOY Y ESA FECHA, SEGÚN POSICIÓN DE FILA

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nBuscando nombre de proyectos, espere por favor... \n");

                logTexto = DateTime.Now.ToString() + " PARTE 2: Buscando nombre de proyectos.";
                WriteLog(logTexto);


                range = "Tracker!G:H";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                response = request.Execute();
                values = response.Values;

                if (values != null && values.Count > 0)
                {
                    int rowPosition = 0;
                    DateTime fechaCompromiso = DateTime.Today;
                    foreach (var rowNombreTareas in values)
                    {
                        try
                        {
                            datosTracker[rowPosition, 0] = rowNombreTareas[0].ToString();
                            datosTracker[rowPosition, 1] = rowNombreTareas[1].ToString();

                            if (rowNombreTareas[1].ToString() != "Fecha Comp Inicial")
                            {
                                //Cálculo de fechas
                                fechaCompromiso = Convert.ToDateTime(rowNombreTareas[1].ToString());
                                TimeSpan resta = fechaCompromiso.Date - hoy;
                                int difDias = int.Parse(resta.Days.ToString());
                                datosTracker[rowPosition, 7] = difDias.ToString();
                            }
                            else
                            {
                                datosTracker[rowPosition, 7] = "0";
                            }
                        }
                        catch (Exception)
                        {
                            datosTracker[rowPosition, 1] = "";
                        }

                        rowPosition = rowPosition + 1;
                    }
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío");
                    logTexto = DateTime.Now.ToString() + " El tracker se encuentra vacío";
                    WriteLog(logTexto);
                    Console.WriteLine("Esta ventana se cerrará automáticamente en cuatro (4) segundos");
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("© 2022 Saúl Candia");

                    //Contador para cierre de app consola:
                    Thread.Sleep(4000);
                    Environment.Exit(0);
                }

                logTexto = DateTime.Now.ToString() + " PARTE 2: Resultado Exitoso.";
                WriteLog(logTexto);


                //PARTE 3: REVISA AL CLIENTE INTERNO Y AL RESPONSABLE

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nRevisando cliente interno y responsable de la tarea, espere por favor... \n");

                logTexto = DateTime.Now.ToString() + " PARTE 3: Revisando cliente interno y responsable de cada tarea.";
                WriteLog(logTexto);

                range = "Tracker!K:L";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                response = request.Execute();
                values = response.Values;

                if (values != null && values.Count > 0)
                {
                    int rowPosition = 0;
                    foreach (var rowClienteInterno in values)
                    {
                        datosTracker[rowPosition, 2] = rowClienteInterno[0].ToString();
                        datosTracker[rowPosition, 3] = rowClienteInterno[1].ToString();
                        rowPosition = rowPosition + 1;
                    }
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío");
                }

                logTexto = DateTime.Now.ToString() + " PARTE 3: Resultado Exitoso.";
                WriteLog(logTexto);


                //PARTE 4: REVISA CORREOS DE LOS RESPONSABLES

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nRevisando listado de correos, espere por favor... \n");


                logTexto = DateTime.Now.ToString() + " PARTE 4: Revisando listado de correos de cada responsable.";
                WriteLog(logTexto);


                range = "DataCorreo!A:B";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                response = request.Execute();
                values = response.Values;

                if (values != null && values.Count > 0)
                {
                    int w = 0;
                    foreach (var rowCorreoResponsable in values)
                    {
                        for (w = 1; w < limite; w++)
                        {
                            if (rowCorreoResponsable[0].ToString() == datosTracker[w, 3].ToString())
                            {
                                datosTracker[w, 4] = rowCorreoResponsable[1].ToString();
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío");
                }

                logTexto = DateTime.Now.ToString() + " PARTE 4: Resultado Exitoso.";
                WriteLog(logTexto);

                //PARTE 5: SELECCIÓN DE TAREAS VENCIDAS Y POR VENCER SEGÚN CANTIDAD DE DIAS

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nEvaluando tareas vencidas y por vencer, espere por favor... \n");

                logTexto = DateTime.Now.ToString() + " PARTE 5: Evaluando tareas vencidas y por vencer.";
                WriteLog(logTexto);

                int i;
                int diferenciaDias;
                for (i = 1; i < limite; i++)
                {
                    diferenciaDias = 0;
                    if (datosTracker[i, 5].ToString() == "En Curso" || datosTracker[i, 5].ToString() == "Atrasado")
                    {
                        diferenciaDias = int.Parse(datosTracker[i, 7].ToString());
                        if (diferenciaDias <= argumentoMes)
                        {
                            contadorElementos = contadorElementos + 1;
                        }
                    }
                }


                //Declaración de matriz bidimensional con datos en limpio [cantidad elementos encontrados en Hoja, 7]
                //      0                   1                   2                       3                   4                   5                   6               7
                //[nombreProyecto], [fechaCompromiso], [nombreClienteInterno], [nombreResponsable], [correoResponsable], [estadoProyecto], [rowPositionIndex], [diasDiferencia]
                string[,] datosTrackerFinal = new string[contadorElementos, columnas];


                int a = 0;
                int b = 0;
                int ultimaposicion = 0;

                for (i = 1; i < limite; i++)
                {
                    diferenciaDias = int.Parse(datosTracker[i, 7].ToString());
                    if (diferenciaDias <= argumentoMes)
                    {
                        if (datosTracker[i, 5].ToString() == "En Curso" || datosTracker[i, 5].ToString() == "Atrasado")
                        {
                            for (a = ultimaposicion; a < contadorElementos; a++)
                            {
                                for (b = 0; b < columnas; b++)
                                {
                                    datosTrackerFinal[a, b] = datosTracker[i, b].ToString();
                                }
                                ultimaposicion++; ;
                                break;
                            }
                        }
                    }
                }

                logTexto = DateTime.Now.ToString() + " PARTE 5: Resultado Exitoso.";
                WriteLog(logTexto);

                //PARTE 6: Enviar el correo:

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nEnviando correos a responsables...\nEsto podría tomar unos minutos dependiendo de la cantidad de correspondencia a enviar. Espere por favor...");
                Console.WriteLine("\nNO CIERRE EL SERVICIO EN ESTE MOMENTO...");

                logTexto = DateTime.Now.ToString() + " PARTE 6: Enviado correos a cada destinatario responsable.";
                WriteLog(logTexto);

                //Envia correo

                result = SendMail(datosTrackerFinal, contadorElementos);

                if (result == true)
                {
                    logTexto = DateTime.Now.ToString() + " PARTE 6: Resultado Exitoso.";
                    WriteLog(logTexto);

                    Console.WriteLine("\nEl correo fue enviado exitosamente. :-)");
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("Esta ventana se cerrará automáticamente en cuatro (4) segundos");
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("© 2022 Saúl Candia");

                    logTexto = DateTime.Now.ToString() + " El servicio finalizará.";
                    WriteLog(logTexto);

                    //Contador para cierre de app consola:
                    Thread.Sleep(4000);
                    Environment.Exit(0);
                }
                else
                {
                    Console.WriteLine("\nEl correo no fue enviado.\nProbablemente se deba a que no hay correo definido para el responsable o porque el mismo no existe.");
                    
                    logTexto = DateTime.Now.ToString() + " El correo no fue enviado. Probablemente se deba a que no hay correo definido para el responsable o porque el mismo no existe.";
                    WriteLog(logTexto);

                    //Contador para cierre de app consola:
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("Esta ventana se cerrará automáticamente en cuatro (4) segundos");
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("© 2022 Saúl Candia");
                    Thread.Sleep(4000);
                    Environment.Exit(0);
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("===========================================================================================");
                Console.WriteLine("\nHUBO UN ERROR AL REVISAR EL DOCUMENTO. Revise el error a continuación para más detalles\n");
                Console.WriteLine("\n\n" + ex.Message.ToString() + "\n\n");
            }
        }

        static Boolean SendMail(string[,] listaDefinitiva, int contadorElementos)
        {
            int a = 0;
            int b = 0;
            int i = 0;
            string nombreResponsable = "";
            string correoResponsable;
            string logTexto;

            string[] nombreProyecto = new string[contadorElementos];
            string[] fechaCompromiso = new string[contadorElementos];
            string[] nombreClienteInterno = new string[contadorElementos];
            string[] estadoTarea = new string[contadorElementos];
            string[] diferenciaDias = new string[contadorElementos];

            for (a = 0; a < contadorElementos; a++)
            {
                if (nombreResponsable != listaDefinitiva[a, 3].ToString() || listaDefinitiva[a, 3].ToString() != "ENVIADO" || nombreResponsable != "ENVIADO")
                {
                    nombreResponsable = listaDefinitiva[a, 3].ToString();
                    correoResponsable = listaDefinitiva[a, 4].ToString();

                    for (b = 0; b < contadorElementos; b++)
                    {
                        if (listaDefinitiva[b, 3].ToString() == nombreResponsable)
                        {
                            nombreProyecto[b] = listaDefinitiva[b, 0].ToString();
                            fechaCompromiso[b] = listaDefinitiva[b, 1].ToString();
                            nombreClienteInterno[b] = listaDefinitiva[b, 2].ToString();
                            estadoTarea[b] = listaDefinitiva[b, 5].ToString();
                            diferenciaDias[b] = listaDefinitiva[b, 7].ToString();
                        }
                    }

                    // Aqui se genera y envia el correo acumulado:
                    try
                    {
                        StringBuilder builder = new StringBuilder();

                        builder.Append("<!DOCTYPE html><html><body><center><p><img src='cid:leoneraLogo'  width='100' height='75' /></p><h2>CORREO INFORMATIVO</h2><p>Hola, <b>" + nombreResponsable + "</b><br><br>Por medio de este correo, se le informa que, de acuerdo a nuestros registros usted tiene proyectos que deben ser revisados porque están por vencer o ya vencieron.");

                        for (i = 0; i < contadorElementos; i++)
                        {
                            if (nombreProyecto[i] != null)
                            {
                                builder.Append("<hr style='width: 600px; color:gray;'>" +
                                    "<p><table style='table-layout: fixed;'><tr>" +
                                          "<td style='width: 200px;'><b>Nombre del proyecto</b></td>" +
                                          "<td style='width: 400px;'>" + nombreProyecto[i].ToString() + "</td>" +
                                      "</tr>" +
                                       "<tr>" +
                                          "<td style='width: 200px;'><b>Cliente Interno</b></td>" +
                                          "<td style='width: 400px;'>" + nombreClienteInterno[i].ToString() + "</td>" +
                                      "</tr>" +
                                      "<tr>" +
                                          "<td style='width: 200px;'><b>Responsable</b></td>" +
                                          "<td style='width: 400px;'>" + nombreResponsable + "</td>" +
                                      "</tr>" +
                                      "<tr>" +
                                          "<td style='width: 200px;'><b>Fecha de compromiso</b></td>" +
                                          "<td style='width: 400px; color:blue;'>" + fechaCompromiso[i].ToString() + "</td>" +
                                      "</tr>" +
                                      "<tr>" +
                                          "<td style='width: 200px;'><b>Estado</b></td>" +
                                          "<td style='width: 400px; color:red;'><b>" + estadoTarea[i] + "</b></td>" +
                                      "</tr>" +
                                  "</table></p>");
                            }

                        }
                        builder.Append("<hr style='width: 600px; color:gray;'>");
                        builder.Append("<p style='color:green'><b>Se solicita regularizar la situación a la brevedad o en su defecto, conversar con solicitante para acordar nueva fecha de compromiso.</b></p>");
                        builder.Append("<p><b>Por favor NO RESPONDA A ESTE CORREO</b> ya que fue generado de manera automática.</p>");
                        builder.Append("<p>Deseándole mucho éxito en sus proyectos, se despide atentamente, <b>GRUPO LEONERA.</b></p></center></body></html> <br/>");

                        //GENERA EL CORREO Y LO ENVÍA


                        MailMessage mm = new MailMessage();
                        mm.To.Add(correoResponsable);
                        //mm.CC.Add(search.emailRepresentante.ToString());
                        mm.From = new MailAddress("mensajero@leonera.cl", "Mensajes Leonera (NO RESPONDER)");
                        //mm.CC.Add(copiaGuardias);
                        mm.Subject = "Alerta de Tracker";
                        mm.Body = builder.ToString();

                        //AGREGA IMÁGENES
                        AlternateView aw = AlternateView.CreateAlternateViewFromString(mm.Body, null, MediaTypeNames.Text.Html);
                        string currentDir = System.IO.Directory.GetCurrentDirectory();
                        LinkedResource LOGO = new LinkedResource(currentDir+"/GrupoLeoneraLogo.jpg", "image/jpg");
                        //LinkedResource LOGO = new LinkedResource(Path.Combine(Directory.GetCurrentDirectory(), "D:/GrupoLeoneraLogo.jpg"), "image/jpg");
                        LOGO.ContentId = "leoneraLogo";
                        aw.LinkedResources.Add(LOGO);
                        mm.AlternateViews.Add(aw);
                        mm.Body = LOGO.ContentId;
                        SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                        smtp.EnableSsl = true;
                        NetworkCredential nc = new NetworkCredential("swdocumentalcg@leonera.cl", "L3oner4F0rest@l!");
                        smtp.Credentials = nc;

                        logTexto = DateTime.Now.ToString() + " Enviando correo a "+ correoResponsable;
                        WriteLog(logTexto);

                        smtp.Send(mm);

                        //BORRA DATOS YA ENVIADOS

                        logTexto = DateTime.Now.ToString() + " Correo enviado exitosamente a " + correoResponsable;
                        WriteLog(logTexto);

                        for (b = 0; b < contadorElementos; b++)
                        {
                            if (listaDefinitiva[b, 3].ToString() == nombreResponsable)
                            {
                                nombreProyecto[b] = null;
                                fechaCompromiso[b] = null;
                                nombreClienteInterno[b] = null;
                                estadoTarea[b] = null;
                                diferenciaDias[b] = null;
                                listaDefinitiva[b, 3] = "ENVIADO";
                            }
                        }
                        nombreResponsable = "ENVIADO";
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: \n" + ex.Message.ToString());
                        logTexto = DateTime.Now.ToString() + " ERROR: "+ex.Message.ToString();
                        WriteLog(logTexto);
                        return false;
                    }

                }
                else
                {
                    a = a + 0;
                    //nombreResponsable = "";
                }
            }
            return true;
        }

        //Escribe en log de archivo en caso de errores
        public static void WriteLog(string strLog)
        {
            StreamWriter log;
            FileStream fileStream = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;

            string logFilePath = "C:\\Logs\\";
            logFilePath = logFilePath + "FantasmaTracker-Log" + System.DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
            logFileInfo = new FileInfo(logFilePath);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists) logDirInfo.Create();
            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);
            }
            log = new StreamWriter(fileStream);
            log.WriteLine(strLog);
            log.Close();
        }
    }
}
