using EntregasRendir.Context.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Dynamic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Text.RegularExpressions;

namespace EntregasRendir
{
    public class EntregasRendir
    {
        public void EntregasRendirMigracionHelmAOfisis()
        {
            #region Inicializacion de Cliente HTTP - Helm
            HttpClient client = new HttpClient();
            string URI = ConfigurationManager.AppSettings["uriBase"].ToString();
            string mediaType = ConfigurationManager.AppSettings["mediaType"].ToString();
            string apikey = ConfigurationManager.AppSettings["apiKey"].ToString();
            client.BaseAddress = new System.Uri(URI);
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(mediaType));
            client.DefaultRequestHeaders.Add("API-Key", apikey);
            #endregion

            #region Obtener todos las entregas de los reportes de Helm
            string ReporteEntregaRendirConfigID = ConfigurationManager.AppSettings["ReporteEntregaRendirConfigID"].ToString();
            string ReporteEntregaRendirMasivoConfigID = ConfigurationManager.AppSettings["ReporteEntregaRendirMasivoConfigID"].ToString();
            string TimeZone = ConfigurationManager.AppSettings["TimeZone"].ToString();
            string EntregaRendirUriRequest = $"api/v1/jobs/reports/tables/csvlink?configId={ReporteEntregaRendirConfigID}&timezone={TimeZone}";
            string EntregaRendirMasivoUriRequest = $"api/v1/jobs/reports/tables/csvlink?configId={ReporteEntregaRendirMasivoConfigID}&timezone={TimeZone}";

            bool UtilizarFecha = Convert.ToBoolean(ConfigurationManager.AppSettings["UtilizarFechaMigracion"].ToString());
            DateTime FechaMigracion = Convert.ToDateTime(ConfigurationManager.AppSettings["FechaMigracion"].ToString());

            //Obtener de Helm Reporte Individuales
            var EntregaRendirResponse = GetSync($"{EntregaRendirUriRequest}", client);
            //Obtener de Helm Reporte Masivo
            var EntregaRendirMasivoResponse = GetSync($"{EntregaRendirMasivoUriRequest}", client);

            //Parsear a JSON desde la respuesta que esta en CSV
            var JSONEntregaRendir = JsonConvert.SerializeObject(CsvToDynamicData(EntregaRendirResponse));
            var JSONEntregaRendirMasivo = JsonConvert.SerializeObject(CsvToDynamicData(EntregaRendirMasivoResponse));



            List<EntregaRendirReporte> EntregasRendir = new List<EntregaRendirReporte>();
            var registrosMasivo = new List<EntregaRendirReporte>();
            try
            {
                //Desearlizar JSON's de Reportes a Listas de Clase.
                var entregaRendirUnico = JsonConvert.DeserializeObject<List<EntregaRendirReporte>>(JSONEntregaRendir.ToString());
                var entregaRendirMasivo = JsonConvert.DeserializeObject<List<EntregaRendirReporte>>(JSONEntregaRendirMasivo.ToString());

                EntregasRendir.AddRange(entregaRendirUnico);
                registrosMasivo.AddRange(entregaRendirMasivo);

                //EntregasRendir.AddRange(entregaRendirMasivo);

            }
            catch (Exception ex)
            {
                Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
            }

            
            //Individual
            var entregasRendirDTO = EntregasRendir.GroupBy(er => er.External_Number.Trim() )
                            .Select(ner => new EntregaRendirDTO
                            {
                                ExternalNumber = ner.Key,
                                
                                Trabajador = ner.Where(x => x.Item_Values_Item_Description == "Trabajador").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                DNI = ner.Where(x => x.Item_Values_Item_Description == "DNI").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                Categoría = ner.Where(x => x.Item_Values_Item_Description == "Categoría").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                Motivo = ner.Where(x => x.Item_Values_Item_Description == "Motivo").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                Detalle = (ner.Where(x => x.Item_Values_Item_Description == "Detalle").Select(n => n.Item_Values_Value).FirstOrDefault() == null) ? string.Empty : ner.Where(x => x.Item_Values_Item_Description == "Detalle").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                FechaCreacion = Convert.ToDateTime(ner.Select(x => x.Created_Date).FirstOrDefault()),
                                FechaAprobacion = Convert.ToDateTime(ner.Select(x => x.Approval_Date).FirstOrDefault()),
                                Glosa = ner.Where(x => x.Item_Values_Item_Description.StartsWith("**")).Select(n => n.Item_Values_Item_Description).FirstOrDefault(),
                                CantidadDías = ner.Where(x => x.Item_Values_Item_Description == "Cantidad de Días").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                Monto = ner.Where(x => x.Item_Values_Item_Description  == "Monto").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                Moneda = ner.Where(x => x.Item_Values_Item_Description == "Moneda").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                CentroCosto = ner.Where(x => x.Item_Values_Item_Description == "Centro de Costo").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                AutorizadoPor = ner.Where(x => x.Item_Values_Item_Description == "Autorizado por:").Select(n => n.Item_Values_Value).FirstOrDefault(),
                                RegistradoPor = ner.Select(x => x.Filled_By).FirstOrDefault(),
                                Correlativo =  1,
                                Empresa = ner.Where(x => x.Item_Values_Item_Description == "Empresa").Select(n => n.Item_Values_Value).FirstOrDefault()
                            })
                            .ToList();

            //Generar Registros Masivos
            var GrupoMasivos = registrosMasivo.GroupBy(er => er.External_Number.Trim()).ToList();
            List<EntregaRendirDTO> totalEntregaMasivo = new List<EntregaRendirDTO>();
            foreach (var masivo in GrupoMasivos)
            {
                string ExternalNumber = masivo.Key;
                var registros = masivo.ToList();
                int correlativo = 0;
                string Trabajador = string.Empty;
                string DNI = string.Empty;
                string Categoría = string.Empty;
                string Motivo = string.Empty;
                string Glosa = string.Empty;
                string Detalle = string.Empty;
                string CantidadDías = string.Empty;
                string Monto = string.Empty;
                string Moneda = string.Empty;
                string CentroCosto = string.Empty;
                string AutorizadoPor = string.Empty;
                string RegistradoPor = string.Empty;
                string Empresa = string.Empty;
                bool trabajador = false;
                bool dni = false;
                List<EntregaRendirDTO> entregaMasivo = new List<EntregaRendirDTO>();
                foreach (var item in registros)
                {
                    switch (item.Item_Values_Item_Description.Trim())
                    {
                        case "Trabajador":
                            correlativo += 1;
                            trabajador = true;
                            Trabajador = item.Item_Values_Value.Trim();
                            break;
                        case "DNI":
                            if (trabajador)
                            {
                                dni = true;
                                DNI = item.Item_Values_Value.Trim();
                            }
                            break;
                        case "Monto":
                            if (trabajador && dni)
                            {
                                dni = false;
                                trabajador = false;
                                Monto = item.Item_Values_Value.Trim();
                                if (Trabajador != string.Empty && DNI != string.Empty && Monto != string.Empty)
                                {
                                    EntregaRendirDTO entregaRendirDTO = new EntregaRendirDTO();
                                    entregaRendirDTO.ExternalNumber = ExternalNumber;
                                    entregaRendirDTO.Trabajador = Trabajador;
                                    entregaRendirDTO.DNI = DNI;
                                    entregaRendirDTO.Correlativo = correlativo;
                                    entregaRendirDTO.Monto = Monto;
                                    entregaRendirDTO.Moneda = Moneda;
                                    entregaRendirDTO.Glosa = Glosa;
                                    entregaRendirDTO.Categoría = Categoría;
                                    entregaRendirDTO.CantidadDías = CantidadDías;
                                    entregaRendirDTO.CentroCosto = CentroCosto;
                                    entregaRendirDTO.Detalle = Detalle;
                                    entregaRendirDTO.FechaCreacion = Convert.ToDateTime(item.Created_Date);
                                    entregaRendirDTO.FechaAprobacion = Convert.ToDateTime(item.Approval_Date);
                                    entregaRendirDTO.AutorizadoPor = AutorizadoPor;
                                    entregaRendirDTO.RegistradoPor = item.Filled_By;
                                    entregaRendirDTO.Empresa = Empresa;
                                    entregaMasivo.Add(entregaRendirDTO);
                                    Trabajador = string.Empty;
                                    DNI = string.Empty;
                                    Monto = string.Empty;
                                }
                            }
                            break;
                        case "Categoría":
                            Categoría = item.Item_Values_Value.Trim();
                            break;
                        case "Motivo":
                            Motivo = item.Item_Values_Value.Trim();
                            break;
                        case "Detalle":
                            Detalle = item.Item_Values_Value.Trim();
                            break;
                        case "Cantidad de Días":
                            CantidadDías = item.Item_Values_Value.Trim();
                            break;
                        case "Moneda":
                            Moneda = item.Item_Values_Value.Trim();
                            break;
                        case "Centro de Costo":
                            CentroCosto = item.Item_Values_Value.Trim();
                            break;
                        case "Autorizado por:":
                            AutorizadoPor = item.Item_Values_Value.Trim();
                            break;
                        case "Empresa":
                            Empresa = item.Item_Values_Value.Trim();
                            break;
                        default:
                            break;
                    }
                }

                
                entregaMasivo.ForEach(x =>
                {
                    x.Motivo = Motivo;
                    x.Glosa = Glosa;
                    x.Detalle = Detalle;
                    x.CantidadDías = CantidadDías;
                    x.Moneda = Moneda;
                    x.CentroCosto = CentroCosto;
                    x.AutorizadoPor = AutorizadoPor;
                    x.Empresa = Empresa;
                   
                });

                totalEntregaMasivo.AddRange(entregaMasivo);
            }

            entregasRendirDTO.AddRange(totalEntregaMasivo);



            #endregion
            #region Obtener Fecha Inicial del periodo de tesoreria de OFISIS
            DateTime FechaPeriodo = DateTime.Now;
            //List<ListadoPeriodosDTO> listadoPeriodosDTO  = new List<ListadoPeriodosDTO>();
            TRAMARSAContext OfisisDB = new TRAMARSAContext();
            //ListadoPeriodosDTO PeriodoDesde = new ListadoPeriodosDTO();
            //ListadoPeriodosDTO PeriodoHasta = new ListadoPeriodosDTO();
            //try
            //{
            //    listadoPeriodosDTO = OfisisDB.Set<ListadoPeriodosDTO>().FromSqlRaw($"exec USP_LISTAR_PERIODOS_OFISIS_ER").ToList();
            //    PeriodoDesde = listadoPeriodosDTO.Where(x => x.GRTPAR_MODULO.Equals("CJ") && x.DESPAR.Equals("Desde")).FirstOrDefault();
            //    PeriodoHasta = listadoPeriodosDTO.Where(x => x.GRTPAR_MODULO.Equals("CJ") && x.DESPAR.Equals("Hasta")).FirstOrDefault();
            //}
            //catch (Exception ex)
            //{

            //    throw;
            //}
            #endregion

            //if (UtilizarFecha)
            //{
            //    entregasRendirDTO = entregasRendirDTO.Where(x => x.FechaAprobacion.Date == FechaMigracion.Date).ToList();
            //}else
            //{
            //    DateTime FechaHoy = DateTime.Now;
            //    entregasRendirDTO = entregasRendirDTO.Where(x => x.FechaAprobacion.Date == FechaHoy.Date).ToList();
            //}
            bool envioCorreo = Convert.ToBoolean(ConfigurationManager.AppSettings["envioCorreo"].ToString());
            List<User> users = GetUsers();
            //TRAMARSAContext OfisisDB = new TRAMARSAContext();
            
            for (int i = 0; i < entregasRendirDTO.Count; i++)
            {
                var entrega = entregasRendirDTO[i];
                try
                {
                    string CodigoEmpresa = entrega.Empresa.Substring(0, entrega.Empresa.IndexOf("-")).Trim();
                    string CodigoComprobante = string.Empty;
                    string CodigoConcepto = string.Empty;
                    string CodigoMoneda = string.Empty;
                    if (entrega.Moneda.Equals("USD"))
                    {
                        CodigoMoneda = "D";
                    }
                    else if (entrega.Moneda.Equals("PEN"))
                    {
                        CodigoMoneda = "S";
                    }

                    switch (CodigoEmpresa)
                    {
                        case "01":
                            CodigoComprobante = $"IENTTF";
                            CodigoConcepto = $"TFC{CodigoMoneda}T1";
                            break;
                        case "02":
                            CodigoComprobante = $"IENTNT";
                            CodigoConcepto = $"NTC{CodigoMoneda}T1";
                            break;
                        case "03":
                            CodigoComprobante = $"IENTDP";
                            CodigoConcepto = $"DPC{CodigoMoneda}T1";
                            break;
                        default:
                            break;
                    }

                    var registro = OfisisDB.USR_CJRMVX
                            .Where(x => 
                                x.USR_CJRMVX_CODEMP.Equals(CodigoEmpresa) &&
                                x.USR_CJRMVX_CODFOR.Equals(CodigoComprobante) &&
                                x.USR_CJRMVX_IDHELM.Equals(entrega.ExternalNumber) &&
                                x.USR_CJRMVX_NROITM.Equals(entrega.Correlativo)).FirstOrDefault();

                    var registropendiente = OfisisDB.USR_CJRMVX
                            .Where(x =>
                                x.USR_CJRMVX_CODEMP.Equals(CodigoEmpresa) &&
                                x.USR_CJRMVX_CODFOR.Equals(CodigoComprobante) &&
                                x.USR_CJRMVX_IDHELM.Equals(entrega.ExternalNumber) &&
                                x.USR_CJRMVX_NROITM.Equals(entrega.Correlativo) &&
                                x.USR_CJRMVX_STVALI.Equals("N")).FirstOrDefault();

            var UsuarioTrabajador = users.Where(x => x.FullName.Trim() == entrega.Trabajador.Trim()).FirstOrDefault();
                    if (UsuarioTrabajador == null)
                    {
                       if (envioCorreo)
                        {
                            Correo correo = new Correo();
                            try
                            {
                                string cuerpo = $"<table border='0' align='left' cellpadding='0' cellspacing='0' style='width: 100%'>" +
                                        "<tr>" +
                                   $"<td align='left' valign='top'> Estimados,</td>" +
                               "</tr>" +
                               "<tr>" +
                                   $"<td align='left' valign='top'> El Trabajador {entrega.Trabajador}, no se logro encontrar en Helm, por tanto no se registro la entrega a rendir {entrega.ExternalNumber}</td>" +
                               "</tr>" +
                                    "<tr style='' align='center'>" +
                                        "<td align='left' valign='top'>" +
                                            "<span> Saludos, </span> <br>" +
                                            "<b> " + "Area de TI - PSAM <br>" +
                                "<b> " + " <br> </b></td></tr><br></table>";
                                correo.Asunto = $"Entrega a rendir no registrada por trabajador no encontrado en HELM";
                                ConstruirCorreoError(correo, entrega.Trabajador, null, cuerpo);
                                var envio = EnviarCorreoElectronico(correo, true);
                                continue;
                            }
                            catch (Exception ex)
                            {
                                Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
                                continue;
                                //new ErrorHandler(new LoggerTXT()).Handle(ex);//Guardamos logErrorTXT
                            }
                        }
                    }

                    var param = new SqlParameter("@DocumentoIdentidad", UsuarioTrabajador.EmployeeNumber);
                    var result = new SqlParameter("@result", SqlDbType.Bit) { Direction = ParameterDirection.Output };
                    OfisisDB.Database.ExecuteSqlRaw($"exec USP_EXISTE_EMPLEADO_OFISIS @DocumentoIdentidad, @Result output", param, result);
                    var UsuarioTrabajadorOFISIS = (bool)result.Value;
                    if (!UsuarioTrabajadorOFISIS)
                    {
                        if (envioCorreo)
                        {
                            Correo correo = new Correo();
                            try
                            {
                                string cuerpo = $"<table border='0' align='left' cellpadding='0' cellspacing='0' style='width: 100%'>" +
                                        "<tr>" +
                                   $"<td align='left' valign='top'> Estimados,</td>" +
                               "</tr>" +
                               "<tr>" +
                                   $"<td align='left' valign='top'> El Trabajador {entrega.Trabajador}, esta registrando la ER {entrega.ExternalNumber}, por favor registrar al empleado en OFISIS ERP</td>" +
                               "</tr>" +
                                    "<tr style='' align='center'>" +
                                        "<td align='left' valign='top'>" +
                                            "<span> Saludos, </span> <br>" +
                                            "<b> " + "Area de TI - PSAM<br>" +
                                "<b> " + " <br> </b></td></tr><br></table>";
                                correo.Asunto = $"Entrega a rendir no registrada por trabajador no encontrado en OFISIS ERP";
                                ConstruirCorreoError(correo, entrega.Trabajador, null, cuerpo);
                                var envio = EnviarCorreoElectronico(correo, true);
                                continue;
                            }
                            catch (Exception ex)
                            {
                                Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
                                continue;
                                //new ErrorHandler(new LoggerTXT()).Handle(ex);//Guardamos logErrorTXT
                            }
                        }
                        continue;
                    }

                    // Fecha de Entrega
                    //DateTime FechaEntrega = DateTime.Now;
                    //if (entrega.FechaAprobacion.Date < Convert.ToDateTime(PeriodoDesde.VALPAR).Date)
                    //{
                    //    FechaEntrega = Convert.ToDateTime(PeriodoDesde.VALPAR).Date;
                    //}
                    //else
                    //{
                    //    FechaEntrega = entrega.FechaAprobacion.Date;
                    //}
                    //

                    // Solicitado por Joe, las ER deben tener fecha de hoy
                    DateTime FechaEntrega = DateTime.Now; 
                    //if ((entrega.FechaAprobacion.Date >= Convert.ToDateTime(PeriodoDesde.VALPAR).Date) && (entrega.FechaAprobacion.Date <= Convert.ToDateTime(PeriodoHasta.VALPAR).Date))
                    //{
                    //    FechaEntrega = entrega.FechaAprobacion.Date;
                    //}
                    //else
                    //{
                    //    FechaEntrega = Convert.ToDateTime(PeriodoDesde.VALPAR).Date;
                    //}

                    if (registro == null)
                    {
                        #region Registro de nueva Entrega a rendir

                        USR_CJRMVX entregaRendir = new USR_CJRMVX();

                        
                        string[] CentrosCosto = entrega.CentroCosto.Split(",");
                        string CC = string.Empty;
                        foreach (var CentroCosto in CentrosCosto)
                        {
                            CC += $" {CentroCosto.Substring(0, CentroCosto.IndexOf("-")).Trim()} ";
                            if (CentrosCosto.Length > 1)
                            {
                                CC += ",";
                            }
                        }

                        //DateTime FechaEntrega = DateTime.Now;
                        //if (entrega.FechaAprobacion.Date < Convert.ToDateTime(PeriodoDesde.VALPAR).Date)
                        //{
                        //    FechaEntrega = Convert.ToDateTime(PeriodoDesde.VALPAR).Date;
                        //}else
                        //{
                        //    FechaEntrega = entrega.FechaAprobacion.Date;
                        //}

                        string Glosa = $"{entrega.Motivo} - {entrega.Detalle} - {CC}";
                        entregaRendir.USR_CJRMVX_CODEMP = CodigoEmpresa;
                        entregaRendir.USR_CJRMVX_CODFOR = CodigoComprobante;
                        entregaRendir.USR_CJRMVX_IDHELM = entrega.ExternalNumber;
                        entregaRendir.USR_CJRMVX_NROITM = entrega.Correlativo;
                        entregaRendir.USR_CJRMVX_FCHMOV = FechaEntrega;
                        entregaRendir.USR_CJRMVX_CAMSEC = 1;
                        entregaRendir.USR_CJRMVX_TEXTOS = Glosa;
                        entregaRendir.USR_CJRMVX_SUBCUE = UsuarioTrabajador.EmployeeNumber;
                        entregaRendir.USR_CJRMVX_TIPCPT = "B";
                        entregaRendir.USR_CJRMVX_CODCPT = CodigoConcepto;
                        entregaRendir.USR_CJRMVX_MONEDA = entrega.Moneda;
                        entregaRendir.USR_CJRMVX_IMPORT = decimal.Parse(entrega.Monto);
                        entregaRendir.USR_CJRMVX_STVALI = "N";
                        entregaRendir.USR_CJRMVX_MSVALI = string.Empty;
                        entregaRendir.USR_CJRMVX_REGISTRADOR = entrega.RegistradoPor.Trim();
                        entregaRendir.USR_CJRMVX_FECHAENVIO = DateTime.Now;
                        entregaRendir.USR_CJ_FECALT = DateTime.Now;
                        entregaRendir.USR_CJ_USERID = "Interface";
                        entregaRendir.USR_CJ_ULTOPR = "A";
                        entregaRendir.USR_CJ_DEBAJA = "N";

                        OfisisDB.USR_CJRMVX.Add(entregaRendir);
                        OfisisDB.SaveChanges();
                        #endregion
                    }
                    //else
                    if (registropendiente != null)
                    {
                        string[] CentrosCosto = entrega.CentroCosto.Split(",");
                        string CC = string.Empty;
                        foreach (var CentroCosto in CentrosCosto)
                        {
                            CC += $" {CentroCosto.Substring(0, CentroCosto.IndexOf("-")).Trim()} ";
                            if (CentrosCosto.Length > 1)
                            {
                                CC += ",";
                            }
                        }

                        string Glosa = $"{entrega.Motivo} - {entrega.Detalle} - {CC}";
                        registro.USR_CJRMVX_CODEMP = CodigoEmpresa;
                        registro.USR_CJRMVX_CODFOR = CodigoComprobante;
                        registro.USR_CJRMVX_IDHELM = entrega.ExternalNumber;
                        registro.USR_CJRMVX_NROITM = entrega.Correlativo;
                        registro.USR_CJRMVX_FCHMOV = FechaEntrega;
                        registro.USR_CJRMVX_CAMSEC = 1;
                        registro.USR_CJRMVX_TEXTOS = Glosa;
                        registro.USR_CJRMVX_SUBCUE = UsuarioTrabajador.EmployeeNumber;
                        registro.USR_CJRMVX_TIPCPT = "B";
                        registro.USR_CJRMVX_CODCPT = CodigoConcepto;
                        registro.USR_CJRMVX_MONEDA = entrega.Moneda;
                        registro.USR_CJRMVX_IMPORT = decimal.Parse(entrega.Monto);
                        registro.USR_CJRMVX_REGISTRADOR = entrega.RegistradoPor.Trim();
                        registro.USR_CJ_FECMOD = DateTime.Now;
                        registro.USR_CJ_USERID = "Interface";
                        registro.USR_CJ_ULTOPR = "A";
                        registro.USR_CJ_DEBAJA = "N";
                        OfisisDB.SaveChanges();
                    }
                }
                catch (Exception ex)
                {
                    Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
                    continue;
                }                
            }

            #region Envio de Correo de Observaciones
           
            if (envioCorreo)
            {
                List<USR_CJRMVX> entregasRendirObservaciones = OfisisDB.USR_CJRMVX
                                      .Where(x => x.USR_CJRMVX_STVALI == "N" && x.USR_CJRMVX_MSVALI.Length > 0)
                                      .ToList();

                if (entregasRendirObservaciones.Count > 0)
                {
                    //Agrupar observaciones por DNI
                    List<IGrouping<string, USR_CJRMVX>> obsercionesAgrupadasPorDNI = entregasRendirObservaciones.GroupBy(x => x.USR_CJRMVX_SUBCUE).ToList();
                    //Obtener usuarios de Helm                    

                    foreach (var observacionesPorDNI in obsercionesAgrupadasPorDNI)
                    {
                        var Usuario = users.Where(x => x.EmployeeNumber.Trim() == observacionesPorDNI.Key.ToString().Trim()).FirstOrDefault();
                        var observacionesUsuario = observacionesPorDNI.ToList();

                        List<EntregaRendirObservacionDTO> observaciones = new List<EntregaRendirObservacionDTO>();
                        string NombreRegistrador = observacionesUsuario.FirstOrDefault().USR_CJRMVX_REGISTRADOR;

                        observacionesUsuario = observacionesUsuario.Where(x => ((DateTime.Now).Date - x.USR_CJRMVX_FECHAENVIO.Date).TotalDays < 4).ToList();
                        if (observacionesUsuario.Count > 0)
                        {
                            foreach (var entrega in observacionesUsuario)
                            {
                                EntregaRendirObservacionDTO observacion = new EntregaRendirObservacionDTO();

                                observacion.NroExternoHelm = entrega.USR_CJRMVX_IDHELM;
                                observacion.FechaAprobacion = entrega.USR_CJRMVX_FCHMOV.ToString();
                                observacion.Moneda = entrega.USR_CJRMVX_MONEDA;
                                observacion.Monto = entrega.USR_CJRMVX_IMPORT;
                                observacion.Observacion = entrega.USR_CJRMVX_MSVALI;

                                observaciones.Add(observacion);
                            }
                            //Generar Excel con observaciones de entregas a rendir
                            byte[] excelObservaciones = CreateExcelLog(observaciones);
                            string excelName = string.Format("ObservacionesEntregas-{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                            Correo correo = new Correo();
                            string NombreTrabajador = string.Empty;
                            String EmailPattern = "^[\\w!#$%&'*+\\-/=?\\^_`{|}~]+(\\.[\\w!#$%&'*+\\-/=?\\^_`{|}~]+)*@((([\\-\\w]+\\.)+[a-zA-Z]{2,4})|(([0-9]{1,3}\\.){3}[0-9]{1,3}))\\z";
                            if (Usuario != null)
                            {
                                correo.Asunto = $"Entregas a rendir observadas en Ofisis - DNI {observacionesPorDNI.Key} - {Usuario.FullName} ";
                                NombreTrabajador = Usuario.FullName;
                                if (Usuario.Email != null)
                                {
                                    if (Regex.IsMatch(Usuario.Email.Trim(), EmailPattern))
                                    {
                                        correo.CorreosPara.Add(new StructMail
                                        {
                                            Mail = Usuario.Email,
                                            NameMail = Usuario.FullName
                                        });
                                    }
                                    else
                                    {
                                        var UsuarioRegistrador = users.Where(x => x.FullName.Trim() == NombreRegistrador.Trim()).FirstOrDefault();
                                        correo.Asunto = $"Entregas a rendir observadas en Ofisis - DNI {observacionesPorDNI.Key}";
                                        if (UsuarioRegistrador != null)
                                        {
                                            NombreTrabajador = UsuarioRegistrador.FullName;
                                            if (Regex.IsMatch(UsuarioRegistrador.Email, EmailPattern))
                                            {
                                                correo.CorreosPara.Add(new StructMail
                                                {
                                                    Mail = UsuarioRegistrador.Email,
                                                    NameMail = UsuarioRegistrador.FullName
                                                });
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    var UsuarioRegistrador = users.Where(x => x.FullName.Trim() == NombreRegistrador.Trim()).FirstOrDefault();
                                    correo.Asunto = $"Entregas a rendir observadas en Ofisis - DNI {observacionesPorDNI.Key}";
                                    if (UsuarioRegistrador != null)
                                    {
                                        NombreTrabajador = UsuarioRegistrador.FullName;
                                        if (Regex.IsMatch(UsuarioRegistrador.Email, EmailPattern))
                                        {
                                            correo.CorreosPara.Add(new StructMail
                                            {
                                                Mail = UsuarioRegistrador.Email,
                                                NameMail = UsuarioRegistrador.FullName
                                            });
                                        }
                                    }
                                }

                            }
                            else
                            {
                                var UsuarioRegistrador = users.Where(x => x.FullName.Trim() == NombreRegistrador.Trim()).FirstOrDefault();
                                correo.Asunto = $"Entregas a rendir observadas en Ofisis - DNI {observacionesPorDNI.Key}";
                                if (UsuarioRegistrador != null)
                                {
                                    NombreTrabajador = UsuarioRegistrador.FullName;
                                    if (Regex.IsMatch(UsuarioRegistrador.Email, EmailPattern))
                                    {
                                        correo.CorreosPara.Add(new StructMail
                                        {
                                            Mail = UsuarioRegistrador.Email,
                                            NameMail = UsuarioRegistrador.FullName
                                        });
                                    }
                                }

                            }

                            correo.Adjuntos.Add(new Adjunto
                            {
                                archivo = excelObservaciones,
                                nombreArchivo = excelName
                            });

                            if (correo.CorreosPara.Count == 0)
                            {
                                var correosPara = (ConfigurationManager.AppSettings["CorreoErrorPara"].ToString().IndexOf(',') > -1) ? ConfigurationManager.AppSettings["CorreoErrorPara"].ToString().Split(',') : new string[] { ConfigurationManager.AppSettings["CorreoErrorPara"].ToString() };
                                correosPara = (string.IsNullOrEmpty(correosPara[0])) ? null : correosPara;
                                if (correo.CorreosPara.Count == 0)
                                {
                                    if (correosPara != null)
                                    {
                                        foreach (var correoPara in correosPara)
                                        {
                                            correo.CorreosPara.Add(new StructMail() { Mail = correoPara, NameMail = "", Password = string.Empty });
                                        }
                                    }
                                }


                            }

                            try
                            {
                                string entregasError = string.Empty;
                                ConstruirCorreoError(correo, NombreTrabajador, entregasError);
                                var envio = EnviarCorreoElectronico(correo, true);
                            }
                            catch (Exception ex)
                            {
                                Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
                                //new ErrorHandler(new LoggerTXT()).Handle(ex);//Guardamos logErrorTXT
                            }


                        }

                    }
                }
                

            }
            #endregion
        }


        //VALIDACION DE DNI, IMPORTE > 0, PERIODO DE MODULO ABIERTO (FECHA)
        private List<User> GetUsers ()
        {
            List<User> Users = new List<User>();
            string URIUsers = "api/v1/Jobs/Users/FindUsers?page=";
            int page = 1;
            int qtyByPage = 100;
            #region Inicializacion de Cliente HTTP - Helm
            HttpClient clientHttp = new HttpClient();
            string URI = ConfigurationManager.AppSettings["uriBase"].ToString();
            string mediaType = ConfigurationManager.AppSettings["mediaType"].ToString();
            string apikey = ConfigurationManager.AppSettings["apiKeyUsuarios"].ToString(); 
            clientHttp.BaseAddress = new System.Uri(URI);
            clientHttp.DefaultRequestHeaders.Accept.Clear();
            clientHttp.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(mediaType));
            clientHttp.DefaultRequestHeaders.Add("API-Key", apikey);
            #endregion
            
            HelmResponse helmResponse = JObject.Parse(GetSync($"{URIUsers}{page}", clientHttp)).ToObject<HelmResponse>();

            int TotalpageQty = QuantityPages(qtyByPage, helmResponse.Data.TotalCount);

            for (int i = page; i <= TotalpageQty; i++)
            {
                var usersData = JObject.Parse(GetSync($"{URIUsers}{i}", clientHttp)).ToObject<HelmResponse>().Data.Page;
                string json = JsonConvert.SerializeObject(usersData);
                List<User> usersPage = JsonConvert.DeserializeObject<List<User>>(json);
                Users.AddRange(usersPage);
            }
            return Users;
        }


        #region Utils
        private string GetSync(string apiMetodo, HttpClient http)
        {
            string retorno = "";

            try
            {
                var response = http.GetAsync(apiMetodo).Result;

                if (response.IsSuccessStatusCode)
                {
                    retorno = response.Content.ReadAsStringAsync().Result;
                }
                else
                {
                    retorno = "";
                }
            }
            catch (Exception ex)
            {
                return retorno = "Error crítico: " + ex.Message;

            }

            return retorno;
        }

        internal static List<dynamic> CsvToDynamicData(string csv)
        {
            var headers = new List<string>();
            var dataRows = new List<dynamic>();
            using (TextReader reader = new StringReader(csv))
            {
                using (var parser = new TextFieldParser(reader))
                {
                    parser.Delimiters = new[] { "," };
                    parser.HasFieldsEnclosedInQuotes = true;
                    parser.TrimWhiteSpace = true;

                    var rowIdx = 0;

                    while (!parser.EndOfData)
                    {
                        var colIdx = 0;
                        dynamic rowData = new ExpandoObject();
                        var rowDataAsDictionary = (IDictionary<string, object>)rowData;

                        foreach (var field in parser.ReadFields().AsEnumerable())
                        {
                            if (rowIdx == 0)
                            {
                                // header
                                var newfield = RemoveDiacritics(field).ToString();
                                headers.Add(newfield.Replace("\\", "_").Replace("/", "_").Replace(",", "_").Replace(" ", "_"));
                            }
                            else
                            {
                                if (field == "null" || field == "NULL")
                                {
                                    rowDataAsDictionary.Add(headers[colIdx], null);
                                }
                                else
                                {
                                    rowDataAsDictionary.Add(headers[colIdx], field);


                                }

                            }
                            colIdx++;
                        }

                        if (rowDataAsDictionary.Keys.Any())
                        {
                            dataRows.Add(rowData);
                        }

                        rowIdx++;
                    }
                }
            }

            return dataRows;
        }

        internal static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }

        private byte[] CreateExcelLog(List<EntregaRendirObservacionDTO> Errores)
        {

            var stream = new MemoryStream();
            try
            {
                using (var package = new ExcelPackage(stream))
                {
                    var workSheet = package.Workbook.Worksheets.Add("Observaciones");
                    workSheet.Cells.LoadFromCollection(Errores, true);
                    package.Save();
                }
                stream.Position = 0;
                return stream.ToArray();
                //string archivoBase64 = Convert.ToBase64String(bytes);
            }
            catch (Exception ex)
            {
                Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
                throw;
            }
            

        }

        public static int QuantityPages(int quantityByPage, int totalRows)
        {
            int pages = (Math.Floor((decimal)(totalRows / quantityByPage)) > 0) ? (totalRows / quantityByPage) + 1 : (totalRows / quantityByPage);
            return (pages > 0) ? pages : 0;
        }

        public static void ConstruirCorreoError(Correo correo, string NombreUsuario, string InformacionAlternativa = null,string CuerpoAlternativo = null)
        {
            try
            {
                if(CuerpoAlternativo == null)
                {
                    correo.Cuerpo = $"<table border='0' align='left' cellpadding='0' cellspacing='0' style='width: 100%'>" +
                        "<tr>" +
                   $"<td align='left' valign='top'> Estimado(a) {NombreUsuario},</td>" +
               "</tr>" +
               "<tr>" +
                   $"<td align='left' valign='top'> Ud. tiene entregas a rendir pendiente de trasladar a Ofisis, se adjunta Excel para mayor detalle</td>" +
               "</tr>" +
                    "<tr style='' align='center'>" +
                        "<td align='left' valign='top'>" +
                            "<span> Saludos, </span> <br>" +
                            "<b> " + "Area de TI <br>" +
                            "<b> " + " <br> </b></td></tr><br></table>";

                }else
                {
                    correo.Cuerpo = CuerpoAlternativo;
                }

                var correoEnvio = ConfigurationManager.AppSettings["CorreoErrorEnvio"].ToString().Split('|');
                var correosPara = (ConfigurationManager.AppSettings["CorreoErrorPara"].ToString().IndexOf(',') > -1) ? ConfigurationManager.AppSettings["CorreoErrorPara"].ToString().Split(',') : new string[] { ConfigurationManager.AppSettings["CorreoErrorPara"].ToString() };
                var correosCC = (ConfigurationManager.AppSettings["CorreoErrorCC"].ToString().IndexOf(',') > -1) ? ConfigurationManager.AppSettings["CorreoErrorCC"].ToString().Split(',') : new string[] { ConfigurationManager.AppSettings["CorreoErrorCC"].ToString() };
                var correosCO = (ConfigurationManager.AppSettings["CorreoErrorCO"].ToString().IndexOf(',') > -1) ? ConfigurationManager.AppSettings["CorreoErrorCO"].ToString().Split(',') : new string[] { ConfigurationManager.AppSettings["CorreoErrorCO"].ToString() };
                List<StructMail> correosErrorPara = new List<StructMail>();
                List<StructMail> correosErrorCC = new List<StructMail>();
                List<StructMail> correosErrorCO = new List<StructMail>();

                correo.CorreoEmisor.Mail  = correoEnvio[0];
                correo.CorreoEmisor.Password = correoEnvio[1];
                correo.CorreoEmisor.NameMail = correoEnvio[2];
                correosPara = (string.IsNullOrEmpty(correosPara[0])) ? null : correosPara;
                correosCC = (string.IsNullOrEmpty(correosCC[0])) ? null : correosCC;
                correosCO = (string.IsNullOrEmpty(correosCO[0])) ? null : correosCO;
                if (correo.CorreosPara.Count == 0) {
                    if (correosPara != null)
                    {
                        foreach (var correoPara in correosPara)
                        {
                            correo.CorreosPara.Add(new StructMail() { Mail = correoPara, NameMail = "", Password = string.Empty });
                        }
                    }
                }
                
                if (correosCC != null)
                {
                    foreach (var correoCC in correosCC)
                    {
                       
                        correo.CorreosCC.Add(new StructMail() { Mail = correoCC, NameMail = "", Password = string.Empty });
                    }
                }
                if (correosCO != null)
                {
                    foreach (var correoCO in correosCO)
                    {
                       
                        correo.CorreosCCO.Add(new StructMail() { Mail = correoCO, NameMail = "", Password = string.Empty });
                    }
                }


                

                 
            }
            catch (Exception ex)
            {
                Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
                //new ErrorHandler(new LoggerTXT()).Handle(ex);//Guardamos logErrorTXT
            }


        }

        public static bool EnviarCorreoElectronico(Correo correo, bool esHTML,
           bool AcuseRecibo = true)
        {
            var mailMsg = new MailMessage();
            mailMsg.From = new MailAddress(correo.CorreoEmisor.Mail, correo.CorreoEmisor.NameMail);
            foreach (StructMail correopara in correo.CorreosPara)
                mailMsg.To.Add(new MailAddress(correopara.Mail, correopara.NameMail));
            foreach (StructMail correocc in correo.CorreosCC)
                mailMsg.CC.Add(new MailAddress(correocc.Mail, correocc.NameMail));
            foreach (StructMail correocco in correo.CorreosCCO)
                mailMsg.Bcc.Add(new MailAddress(correocco.Mail, correocco.NameMail));
            mailMsg.Subject = correo.Asunto;
            mailMsg.Body = correo.Cuerpo;
            if (esHTML)
            {
                AlternateView htmlView;
                htmlView = AlternateView.CreateAlternateViewFromString(correo.Cuerpo, Encoding.UTF8, "text/html");
                mailMsg.AlternateViews.Add(htmlView);
            }

            if (correo.Adjuntos.Count > 0)
            {
                foreach (var adjunto in correo.Adjuntos)
                {
                    Attachment att = new Attachment(new MemoryStream(adjunto.archivo), adjunto.nombreArchivo);
                    mailMsg.Attachments.Add(att);
                }
            }

            var SmtpClient = new SmtpClient();
            try
            {
                SmtpClient = new SmtpClient("smtp.office365.com", 587);//Cambiar por proveedor de appsettings.json
                SmtpClient.Credentials = new NetworkCredential(correo.CorreoEmisor.Mail, correo.CorreoEmisor.Password);
                SmtpClient.EnableSsl = true; //correoEmisor.SMTPSSL;
            }
            catch (Exception ex)
            {
                //new ErrorHandler(new LoggerTXT()).Handle(new Exception("No se pudo configurar la cuenta de correo electrónico. (Puerto)"));//Guardamos logErrorTXT

            }

            if (AcuseRecibo)
            {
                //SOLICITAR ACUSE DE RECIBO Y LECTURA
                mailMsg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure | DeliveryNotificationOptions.OnSuccess | DeliveryNotificationOptions.Delay;
                mailMsg.Headers.Add("Disposition-Notification-To", correo.CorreoEmisor.Mail); //solicitar acuse de recibo al abrir mensaje
            }
            try
            {
                SmtpClient.Send(mailMsg);
            }
            catch (Exception ex)
            {
                try
                {
                    //reenviando en caso de error
                    mailMsg.DeliveryNotificationOptions = DeliveryNotificationOptions.None;
                    mailMsg.Headers.Remove("Disposition-Notification-To");
                    SmtpClient.Send(mailMsg);
                }
                catch (Exception exc)
                {
                    Logger.WriteLine($"{exc.Message}\n{exc.InnerException?.Message}");
                    //new ErrorHandler(new LoggerTXT()).Handle(ex);//Guardamos logErrorTXT
                }

            }

            return true;
        }


        #endregion

    }

    public class EntregaRendirReporte {
        public string External_Number {get; set;}
        public string Created_Date {get; set;}
        public string Approval_Date { get; set; }
        public string Item_Values_Form_Category_Approval_Date {get; set;}
        public string Item_Values_Item_Description {get; set;}
        public string Item_Values_Value { get; set; } 
        public string Form_Name { get; set; }
        public string Filled_By { get; set; } 
    }

    public class EntregaRendirDTO
    {
        public string ExternalNumber { get; set; }
        public string CodigoComprobante { get; set; }
        public string Trabajador { get; set; }
        public string DNI { get; set; }
        public string Empresa { get; set; }
        public string Categoría { get; set; }
        public DateTime FechaAprobacion { get; set; }
        public DateTime FechaCreacion { get; set; }
        public string Motivo { get; set; }
        public string Glosa { get; set; }
        public string Detalle { get; set; }
        public string CantidadDías { get; set; }
        public string Monto { get; set; }
        public string Moneda { get; set; }
        public string CentroCosto { get; set; }
        public string AutorizadoPor { get; set; }
        public string RegistradoPor { get; set; }
        
        public int Correlativo { get; set; }

        public EntregaRendirDTO()
        {
            Categoría = string.Empty;
            Motivo = string.Empty;
            Glosa = string.Empty;
            Detalle = string.Empty;
            CantidadDías = string.Empty;
            Monto = string.Empty;
            Moneda = string.Empty;
            CentroCosto = string.Empty;
            AutorizadoPor = string.Empty;
            RegistradoPor = string.Empty;
        }
    }

    public class EntregaRendirObservacionDTO
    {
        public string NroExternoHelm { get; set; }
        public string FechaAprobacion { get; set; }
        public string Moneda { get; set; }
        public decimal Monto { get; set; }
        public string Observacion { get; set; }
    }

    public class Correo
    {

        public StructMail CorreoEmisor { get; set; }
        public List<StructMail> CorreosPara { get; set; }
        public List<StructMail> CorreosCC { get; set; }
        public List<StructMail> CorreosCCO { get; set; }
        public string Asunto { get; set; }
        public string Cuerpo { get; set; }
        public List<Adjunto> Adjuntos { get; set; }
        public Correo()
        {
            CorreoEmisor = new StructMail() { Mail = string.Empty, NameMail = string.Empty, Password = string.Empty };
            CorreosPara = new List<StructMail>();
            CorreosCC = new List<StructMail>();
            CorreosCCO = new List<StructMail>();
            Asunto = string.Empty;
            Cuerpo = string.Empty;
            Adjuntos = new List<Adjunto>();
        }
    }

    public class StructMail
    {
        public string Mail { get; set; }
        public string Password { get; set; }
        public string NameMail { get; set; }
    }

    public struct Adjunto
    {
        public byte[] archivo { get; set; }
        public string nombreArchivo { get; set; }
    }

    public class HelmResponse
    {
        public HelmDataReponse Data { get; set; }
    }
    public class HelmDataReponse
    {
        public int TotalCount { get; set; }
        public List<dynamic> Page { get; set; }
    }

    public class ListadoPeriodosDTO
    {
        public string GRTPAR_CODPAR { get; set; }
        public string GRTPAR_MODULO { get; set; }
        public string MODULO_DESC { get; set; }
        public string DESPAR { get; set; }
        public string VALPAR { get; set; }
        public int CONTAD { get; set; }
    }
}
