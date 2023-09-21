using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;

using System.Runtime.InteropServices; //DLLImport
using System.Security.Principal; //WindowsImpersonationContext
using System.Security.Permissions; //PermissionSetAttribute 

//using System.Messaging;
using System.Text;
using System.Net.Security;
using System.Net;
using System.Net.Mail;

using System.IO;

namespace CFDIEnvioAutomatico
{
    public partial class Form1 : Form
    {
        /*20140929 Se desea que se envie a un destinatario en específico de cada agencia el log cuando hay errores en el envio
         * a fin de que corrijan los mismos via pantalla Personas en BPro. el correo que recibirán deberá ser una plantilla.
         
         * 20200611 se agregó ToUpper siempre estaba mandando a false
         * 20220208 se pide nunca enviarle las facturas a José Abascal...
         * 20230412 En las carpetas donde se alojan físicamente los archivos xml y pdf se hará una subclasificación por Año Mes1, Mes2, MesN a partir de una determinada fechaSubClasificacion
        */

        #region Impersonacion en el servidor remoto
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        //[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        //private unsafe static extern int FormatMessage(int dwFlags, ref IntPtr lpSource, int dwMessageId, int dwLanguageId, ref String lpBuffer, int nSize, IntPtr* arguments);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public extern static bool DuplicateToken(IntPtr existingTokenHandle, int SECURITY_IMPERSONATION_LEVEL, ref IntPtr duplicateTokenHandle);

        // logon types
        const int LOGON32_LOGON_INTERACTIVE = 2;
        const int LOGON32_LOGON_NETWORK = 3;
        const int LOGON32_LOGON_NEW_CREDENTIALS = 9;

        // logon providers
        const int LOGON32_PROVIDER_DEFAULT = 0;
        const int LOGON32_PROVIDER_WINNT50 = 3;
        const int LOGON32_PROVIDER_WINNT40 = 2;
        const int LOGON32_PROVIDER_WINNT35 = 1;

        #region manejo de errores
        // GetErrorMessage formats and returns an error message
        // corresponding to the input errorCode.
        public static string GetErrorMessage(int errorCode)
        {
            int FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100;
            int FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;
            int FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

            int messageSize = 255;
            string lpMsgBuf = "";
            int dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;

            IntPtr ptrlpSource = IntPtr.Zero;
            IntPtr ptrArguments = IntPtr.Zero;

            int retVal = 1; //FormatMessage(dwFlags, ref ptrlpSource, errorCode, 0, ref lpMsgBuf, messageSize, &ptrArguments);
            if (retVal == 0)
            {
                throw new ApplicationException(string.Format("Failed to format message for error code '{0}'.", errorCode));
            }

            return lpMsgBuf;
        }

        private static void RaiseLastError()
        {
            int errorCode = Marshal.GetLastWin32Error();
            string errorMessage = "Error LB"; //GetErrorMessage(errorCode);

            throw new ApplicationException(errorMessage);
        }

        #endregion
        #endregion


        #region variables de configuracion
                

        string smtpserverhost1 = System.Configuration.ConfigurationSettings.AppSettings["smtpserverhost"];
        string smtpport1 = System.Configuration.ConfigurationSettings.AppSettings["smtpport"];
        string usrcredential1 = System.Configuration.ConfigurationSettings.AppSettings["usrcredential"];
        string usrpassword1 = System.Configuration.ConfigurationSettings.AppSettings["usrpassword"];
        string EnableSsl1 = System.Configuration.ConfigurationSettings.AppSettings["EnableSsl"];
        string emailsavisar = System.Configuration.ConfigurationSettings.AppSettings["emailsavisar"];
        string PreSubject = System.Configuration.ConfigurationSettings.AppSettings["PreSubject"];
        string EnviaCorreoPrincipal = System.Configuration.ConfigurationSettings.AppSettings["EnviaCorreoPrincipal"];
        string modo = System.Configuration.ConfigurationSettings.AppSettings["modo"];
        string InfEmailLogNoHayArchivo = System.Configuration.ConfigurationSettings.AppSettings["InfEmailLogNoHayArchivo"];
        string DiasAtras = System.Configuration.ConfigurationSettings.AppSettings["DiasAtras"];
        string correodepurador = System.Configuration.ConfigurationSettings.AppSettings["correodepurador"];

        string ConnectionString = System.Configuration.ConfigurationSettings.AppSettings["ConnectionString"]; //es la cadena de conexión a la base de datos principal
        
        string id_agencia_depuracion = System.Configuration.ConfigurationSettings.AppSettings["id_agencia_depuracion"];

        string cadenaLog = "";

        string cad_series_consulta = ""; //En realidad son las series que no se deben enviar.

        ConexionBD objDB = null;
        #endregion


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            FileInfo archivoExecutalbe = new FileInfo(Application.ExecutablePath.Trim());
            string NombreProceso = archivoExecutalbe.Name; 
            NombreProceso = NombreProceso.Replace(".exe", "");
            NombreProceso = NombreProceso.Replace(".EXE", "");
            if (CuentaInstancias(NombreProceso) == 1)
            {

                this.objDB = new ConexionBD(this.ConnectionString);

                Utilerias.WriteToLog(" ", " ", Application.StartupPath + "\\Log.txt");
                Utilerias.WriteToLog("Inicio de Operaciones", "Load", Application.StartupPath + "\\Log.txt");
                ConsultaSeriesEnvio();
                ChecaYEnvia();
                if (this.EnviaCorreoPrincipal.ToUpper().Trim() == "SI")
                    EnviaLog();

                Utilerias.WriteToLog("Termino de Operaciones el sistema se cerrará", "Load", Application.StartupPath + "\\Log.txt");                            
            }
            else {
                Utilerias.WriteToLog("Ya existe una instancia de: " + NombreProceso + " se conserva la instancia actual", "Sincronizador_Load", Application.StartupPath + "\\Log.txt");                               
            }
            Application.Exit();
        }


        public string EnviaErroresEnvioResponsablesxAgencia(string id_agencia,string NombreAgencia,string ErroresxAgencia,string CorreosEnviar)
        {
            string res = "";
            if (ErroresxAgencia.Trim() != "")
            {
                try
                {

                    string rutaplantilla = Application.StartupPath;
                    rutaplantilla += "\\PlantillaAvisoErrorEnvio.txt";
                    clsEmail correoLog = new clsEmail(this.smtpserverhost1.Trim(), Convert.ToInt16(this.smtpport1), this.usrcredential1.Trim(), this.usrpassword1.Trim(), this.EnableSsl1.Trim());
                    MailMessage mensaje = new MailMessage();
                    mensaje.Priority = System.Net.Mail.MailPriority.Normal;
                    mensaje.IsBodyHtml = false;
                    mensaje.Subject = this.PreSubject.Trim() + " " + "CFDI´s log del envio automático Errores en envio: " + DateTime.Now.ToString();
                    string Remitente = "Sistemas de Grupo Andrade";

                    if (this.correodepurador.Trim() != "")
                        CorreosEnviar += "," + this.correodepurador.Trim();
                    string[] EmailsEspeciales = CorreosEnviar.Split(',');

                    foreach (string Email in EmailsEspeciales)
                    {
                        mensaje.To.Add(new MailAddress(Email.Trim()));
                    }

                    mensaje.From = new MailAddress(usrcredential1.Trim(), Remitente.Trim());
                    //mnsj.Attachments.Add(new Attachment(ArchZip));

                    Dictionary<string, string> TextoIncluir = new Dictionary<string, string>();

                    TextoIncluir.Add("fecha", DateTime.Now.ToString("dd-MM-yyyy"));
                    TextoIncluir.Add("hora", DateTime.Now.ToString("HH:mm:ss"));
                    TextoIncluir.Add("NombreSucursal", id_agencia + "-" + NombreAgencia);

                    TextoIncluir.Add("LogEjecucion", ErroresxAgencia);

                    //AlternateView vistaplana = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoPlano(TextoIncluir), null, "text/plain");
                    AlternateView vistahtml = AlternateView.CreateAlternateViewFromString(correoLog.CreaCuerpoHTML(rutaplantilla, TextoIncluir).ToString(), null, "text/plain");

                    //LinkedResource logo = new LinkedResource(rutalogo);
                    //logo.ContentId = "companylogo";
                    //vistahtml.LinkedResources.Add(logo);

                    //mensaje.AlternateViews.Add(vistaplana);
                    mensaje.AlternateViews.Add(vistahtml);
                    correoLog.MandarCorreo(mensaje);
                    res = "Envio exitoso del Log";
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    Utilerias.WriteToLog(ex.Message, "EnviaLog", Application.StartupPath + "\\Log.txt");
                    res = ex.Message;
                }
            }

            return res;
        }


        public string EnviaLog()
        {
            string res = "";
            if (this.cadenaLog.Trim() != "")
            {
                try
                {

                    string rutaplantilla = Application.StartupPath;
                    rutaplantilla += "\\PlantillaAvisoLog.txt";
                    clsEmail correoLog = new clsEmail(this.smtpserverhost1.Trim(), Convert.ToInt16(this.smtpport1), this.usrcredential1.Trim(), this.usrpassword1.Trim(), this.EnableSsl1.Trim());
                    MailMessage mensaje = new MailMessage();
                    mensaje.Priority = System.Net.Mail.MailPriority.Normal;
                    mensaje.IsBodyHtml = false;
                    mensaje.Subject = this.PreSubject.Trim() + " " + "CFDI´s log del envio automático : " + DateTime.Now.ToString();
                    string Remitente = "Sistemas de Grupo Andrade";

                    string[] EmailsEspeciales = this.emailsavisar.Split(',');

                    foreach (string Email in EmailsEspeciales)
                    {
                        mensaje.To.Add(new MailAddress(Email.Trim()));
                    }

                    mensaje.From = new MailAddress(usrcredential1.Trim(), Remitente.Trim());
                    //mnsj.Attachments.Add(new Attachment(ArchZip));

                    Dictionary<string, string> TextoIncluir = new Dictionary<string, string>();

                    TextoIncluir.Add("fecha", DateTime.Now.ToString("dd-MM-yyyy"));
                    TextoIncluir.Add("hora", DateTime.Now.ToString("HH:mm:ss"));
                    TextoIncluir.Add("LogEjecucion", this.cadenaLog);

                    //AlternateView vistaplana = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoPlano(TextoIncluir), null, "text/plain");
                    AlternateView vistahtml = AlternateView.CreateAlternateViewFromString(correoLog.CreaCuerpoHTML(rutaplantilla, TextoIncluir).ToString(), null, "text/plain");

                    //LinkedResource logo = new LinkedResource(rutalogo);
                    //logo.ContentId = "companylogo";
                    //vistahtml.LinkedResources.Add(logo);

                    //mensaje.AlternateViews.Add(vistaplana);
                    mensaje.AlternateViews.Add(vistahtml);
                    correoLog.MandarCorreo(mensaje);
                    res = "Envio exitoso del Log";
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    Utilerias.WriteToLog(ex.Message, "EnviaLog", Application.StartupPath + "\\Log.txt");
                    res = ex.Message;
                }
            }

            return res;
        }

        public void ConsultaSeriesEnvio()
        {
            this.cad_series_consulta = "";
            string Q = "Select id_serie from BP_SERIES_ENVIOAUTOMATICO";
            DataSet ds = this.objDB.Consulta(Q);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow reg in ds.Tables[0].Rows)
                {
                    this.cad_series_consulta += "'" + reg["id_serie"].ToString().Trim() + "',";
                }
            }
            if (this.cad_series_consulta.Trim() != "")
                this.cad_series_consulta = this.cad_series_consulta.Substring(0, this.cad_series_consulta.Length - 1);
        }

        public string ConsultaSeriesXSucursal(string id_agencia,string strconexionABussinesProBaseOperativa)
        {
            string res = "";

            SqlConnection conBPBaseOperativa = new SqlConnection();
            //SqlCommand bp_comandBaseOperativa = new SqlCommand();
            
            if (conBPBaseOperativa.State.ToString().ToUpper().Trim() == "CLOSED")
                {
                    try
                    {
                        conBPBaseOperativa.ConnectionString = strconexionABussinesProBaseOperativa;
                        conBPBaseOperativa.Open();
                        //bp_comandBaseOperativa.Connection = conBPBaseOperativa;
                    }
                    catch (Exception ex10)
                    {
                        Utilerias.WriteToLog("Imposible conexion con BD de BP:" + ex10.Message + " " + strconexionABussinesProBaseOperativa + " id_agencia= " + id_agencia, "ConsultaSeriesXSucursal", Application.StartupPath + "\\Log.txt");
                        return res;
                    }
                }

                string Q = "select FCF_SERIE from ADE_CFDFOLIOS ";

                            SqlCommand command = new SqlCommand(Q, conBPBaseOperativa);
                            command.CommandTimeout = 300; //5 * 60 segundos = 5 minutos (300segundos)  
                            DataSet dsBP = new DataSet();
                            System.Data.SqlClient.SqlDataAdapter objAdaptador = new System.Data.SqlClient.SqlDataAdapter(command);
                        try
                        {    
                        objAdaptador.Fill(dsBP, "Resultados");

                        if (dsBP != null && dsBP.Tables.Count > 0 && dsBP.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow registro in dsBP.Tables[0].Rows)
                            {
                                res += "'" + registro["FCF_SERIE"].ToString().Trim() + "',";
                            }
                        }

                        if (res.Trim() != "")
                            res = res.Substring(0, res.Length - 1);
                        
                        }
                        catch (Exception exSQL)
                        {
                            Utilerias.WriteToLog("Error Al hacer la consulta a la tabla de ADE_CFDFOLIOS " + exSQL.Message + " id_agencia: " + id_agencia, "ConsultaSeriesXSucursal", Application.StartupPath + "\\Log.txt");
                        }
            
            return res;
        }


        public string ChecaYEnvia()
        {
            string res = "";
            string strconexionABussinesPro = "";
            string strconexionABussinesProBaseOperativa = "";
            string strUsrRemoto = "";
            string strPassRemoto = "";
            string strDirectorioRemotoXML = "";
            string strDirectorioRemotoPDF = "";
            string strIPFileStorage = "";
            string RFCEmpresa = "";
            string NombreEmpresa = "";
            string vde_idemisor = "";
            string rutalogo = "";
            string rutaplantillaHTML = "";
            string textocuerpoplano = "Aviso de Recepción de Comprobante Fiscal Digital \\n";
            string Q = "";
            string id_agencia="";
            string base_operativa = ""; //Es decir la base de datos de BPro que no es la concentradora.
            string seriesxbaseoperativa = ""; //Son las series que maneja cada sucursal nos sirve para poder diferenciar en las bases Concentradora de dónde proviene cada factura.
            string fechasubclasif = "";

            this.cadenaLog = "";


            textocuerpoplano += "Por éste conducto le enviamos adjuntos los archivos de su Comprobante Fiscal Digital: [serie] / [folio]";
            textocuerpoplano += "Fecha de envio: [fecha]";
            textocuerpoplano += "El archivo adjunto contine la información del Comprobante Fiscal Digital emitido por: [rfcemisor] , [nombreemisor]";

            rutaplantillaHTML = ""; //+ "\\bp_PlantillaCFDAvisoCliente.html"; 
            rutalogo = ""; //+"\\Imagenes\\"; //image2993.png";
            ConexionBD objDB = new ConexionBD(this.ConnectionString);
                       

            #region Consulta de los datos para el Logueo en el Servidor Remoto para traer los CFD´s
            
            SqlConnection conBP = new SqlConnection();
            SqlCommand bp_comand = new SqlCommand();


            //conociendo el id_agencia procedemos a consultar los datos de conexion en la tabla transferencia
            Q = "Select ip,usr_bd,pass_bd,nombre_bd, bd_alterna, dir_remoto_xml, dir_remoto_pdf,usr_remoto,pass_remoto, ip_almacen_archivos, smtpserverhost, smtpport, usrcredential, usrpassword, plantillaHTML, logo, cuenta_from, enable_ssl, id_agencia, isnull(Convert(char(8),fechasubclasif,112),'19000101') as fechasubclasif ";
            Q += " From TRANSMISION ";  //where id_agencia='" + id_agencia + "'";
            if (this.id_agencia_depuracion.Trim() != "") 
                Q += " where id_agencia='" + this.id_agencia_depuracion.Trim() + "'";            
            Q += " order by id_agencia";

            DataSet ds = objDB.Consulta(Q);

            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
               foreach (DataRow regConexion in ds.Tables[0].Rows)
               {
                string ErroresxAgencia = "";

                //DataRow regConexion = ds.Tables[0].Rows[0];
                strconexionABussinesPro = string.Format("Data Source={0};Initial Catalog={1}; Persist Security Info=True; User ID={2};Password={3}", regConexion["ip"].ToString(), regConexion["nombre_bd"].ToString(), regConexion["usr_bd"].ToString(), regConexion["pass_bd"].ToString());
                strUsrRemoto = regConexion["usr_remoto"].ToString().Trim();
                strPassRemoto = regConexion["pass_remoto"].ToString().Trim();
                strDirectorioRemotoXML = regConexion["dir_remoto_xml"].ToString().Trim();
                strDirectorioRemotoPDF = regConexion["dir_remoto_pdf"].ToString().Trim();
                strIPFileStorage = regConexion["ip_almacen_archivos"].ToString().Trim();
                if (regConexion["nombre_bd"].ToString().IndexOf("Concen") > 0)
                    base_operativa = regConexion["bd_alterna"].ToString();
                else
                    base_operativa = regConexion["nombre_bd"].ToString();

                strconexionABussinesProBaseOperativa = string.Format("Data Source={0};Initial Catalog={1}; Persist Security Info=True; User ID={2};Password={3}", regConexion["ip"].ToString(),base_operativa.Trim(), regConexion["usr_bd"].ToString(), regConexion["pass_bd"].ToString());
                seriesxbaseoperativa = ConsultaSeriesXSucursal(id_agencia, strconexionABussinesProBaseOperativa);

                id_agencia = regConexion["id_agencia"].ToString().Trim();
                RFCEmpresa = objDB.ConsultaUnSoloCampo("Select rfc from AGENCIAS where id_agencia='" + id_agencia + "'");
                NombreEmpresa = objDB.ConsultaUnSoloCampo("Select nombre from AGENCIAS where id_agencia='" + id_agencia + "'");
                vde_idemisor = objDB.ConsultaUnSoloCampo("Select vde_idemisor from AGENCIAS where id_agencia='" + id_agencia + "'");
                vde_idemisor = vde_idemisor == "" ? "1" : vde_idemisor;
                string smtpserverhost = regConexion["smtpserverhost"].ToString().Trim();
                string smtpport = regConexion["smtpport"].ToString().Trim();
                string usrcredential = regConexion["usrcredential"].ToString().Trim();
                string usrpassword = regConexion["usrpassword"].ToString().Trim();
                rutaplantillaHTML = Application.StartupPath + "\\" + regConexion["plantillaHTML"].ToString().Trim();
                rutalogo = Application.StartupPath + "\\" + regConexion["logo"].ToString().Trim();
                string cuenta_from = regConexion["cuenta_from"].ToString().Trim();
                string enable_ssl = regConexion["enable_ssl"].ToString().ToUpper().Trim();
                string cuentaslogerrores = "";
                string correosconcopia = "";
                fechasubclasif = regConexion["fechasubclasif"].ToString().Trim();

                #region Consulta cuentas de correo de errores y de copia
                DataSet dserr = this.objDB.Consulta("Select e_mail from CUENTASENVIARERROR where id_agencia='" + id_agencia + "'");
                foreach(DataRow emailerr in dserr.Tables[0].Rows)
                {
                   cuentaslogerrores += emailerr["e_mail"].ToString().Trim() + ","; 
                }
                   cuentaslogerrores = cuentaslogerrores.Trim()==""? cuentaslogerrores.Trim() : cuentaslogerrores.Substring(0,cuentaslogerrores.Length-1);

                   DataSet dsccopia = this.objDB.Consulta("Select e_mail from CUENTAS_CONCOPIA where id_agencia='" + id_agencia + "' and activo='TRUE'");
                   foreach (DataRow emailconcopia in dsccopia.Tables[0].Rows)
                   {
                       correosconcopia += emailconcopia["e_mail"].ToString().Trim() + ",";
                   }
                   correosconcopia = correosconcopia.Trim() == "" ? correosconcopia.Trim() : correosconcopia.Substring(0, correosconcopia.Length - 1);
                #endregion

                   #region Conecciones a bases de BPRO  
                   if (conBP.State.ToString().ToUpper().Trim() == "CLOSED")
                {
                    try
                    {
                        conBP.ConnectionString = strconexionABussinesPro;
                        conBP.Open();
                        bp_comand.Connection = conBP;
                    }
                    catch (Exception ex1)
                    {
                        Utilerias.WriteToLog("Imposible conexion con BD de BP:" + ex1.Message + " " + strconexionABussinesPro, "ChecaYEnvia", Application.StartupPath + "\\Log.txt");
                        res = "Imposible conexion con BD de BP:" + ex1.Message;                        
                    }
                }

                   #endregion



                DateTime DiaBase = DateTime.Now.AddDays(-1 * Convert.ToInt32(this.DiasAtras));

                    //if (id_agencia == "85" || id_agencia == "86" || id_agencia == "87" || id_agencia == "88")
                    //{//Prueba de donde se extrae la cuenta de correo en un principio solo para las CRAs.
                        Q = " SELECT distinct cfds.VDE_SERIE, cfds.VDE_FOLIO, cfds.VDE_DOCTO, per.PER_EMAIL as VDE_RMAIL, Convert(char(8),Convert(datetime,cfds.VDE_FECHOPE),112) as fecha "; //20230412 se agrega la fecha de la factura para poder comparar y decidir si ya está subclasificado en año mes
                        Q += " FROM ADE_VTACFD cfds, ADE_CANCFD cfdscancelados, PER_PERSONAS per ";
                        //Q += " FROM ADE_VTACFD cfds, ADE_CANCFD cfdscancelados, PER_PERSONAS per, AGENCIAS ag";
                        Q += " WHERE per.PER_IDPERSONA = cfds.VDE_IDRECEPTOR ";
                        Q += " and Len(rtrim(ltrim(per.PER_EMAIL))) > 0 "; //--Que tengan correo electrónico
                        Q += " and patindex('%@%',per.PER_EMAIL)>0 "; //--Que tenga el arroba en el correo electrónico
                        Q += " and  Convert(char(8),Convert(datetime,cfds.VDE_FECHAORACAR),112) > '" + DiaBase.ToString("yyyyMMdd") + "'";
                        Q += " and cfds.VDE_SERIE not in (" + this.cad_series_consulta.Trim() + ")";
                        Q += " and cfds.VDE_DOCTO <> cfdscancelados.CDE_DOCTO "; //y que no esten cancelados.
                        //Q += " and ag.id_agencia = " + id_agencia.Trim(); //20180213 para identificar en las bd Concentradoras de donde viene la factura. otra manera es consultar la serie en la tabla ADE_CFDFOLIOS de la base de datos Operativa i.e. no de la concentradora.
                        //Q += " and ag.vde_idemisor = cfds.VDE_IDEMISOR"; //20180213
                        //Q += " and cfds.VDE_IDEMISOR=" + vde_idemisor.Trim();
                        Q += " and cfds.VDE_SERIE in (" + seriesxbaseoperativa + ")"; 
                        //Q +=  "and  cfds.VDE_SERIE = 'FA'";
                        //Q += " and  cfds.VDE_FOLIO = '74570'";
                    /*}
                    else
                    {
                        //Consultamos en la bd de esta agencia los CFDI generados desde el 2 de Agosto y que tienen EMAIL en BPRo
                        Q = " SELECT distinct cfds.VDE_SERIE, cfds.VDE_FOLIO, cfds.VDE_DOCTO, cfds.VDE_RMAIL";
                        Q += " FROM ADE_VTACFD cfds, ADE_CANCFD cfdscancelados";
                        Q += " WHERE Len(rtrim(ltrim(cfds.VDE_RMAIL))) > 0"; //Que tengan correo electrónico
                        Q += " and patindex('%@%',cfds.VDE_RMAIL)>0 "; //Que tenga el arroba en el correo electrónico
                        //Q += " and  Convert(char(8),Convert(datetime,cfds.VDE_FECHAORACAR),112) = '" + DateTime.Now.ToString("yyyyMMdd") + "'";
                        Q += " and  Convert(char(8),Convert(datetime,cfds.VDE_FECHAORACAR),112) > '" + DiaBase.ToString("yyyyMMdd") + "'";
                        Q += " and cfds.VDE_SERIE not in (" + this.cad_series_consulta.Trim() + ")";
                        Q += " and cfds.VDE_DOCTO <> cfdscancelados.CDE_DOCTO"; //y que no esten cancelados.
                    }*/
                        
                            SqlCommand command = new SqlCommand(Q, conBP);
                            command.CommandTimeout = 300; //5 * 60 segundos = 5 minutos (300segundos)  
                            DataSet dsBP = new DataSet();
                            System.Data.SqlClient.SqlDataAdapter objAdaptador = new System.Data.SqlClient.SqlDataAdapter(command);
                        try
                        {    
                        objAdaptador.Fill(dsBP, "Resultados");
                        }
                        catch (Exception exSQL)
                        {
                            Utilerias.WriteToLog("Error Al hacer la consulta a la tabla de Personas " + exSQL.Message + "Query:" + Q + " agencia: " + id_agencia + " " + NombreEmpresa, "ChecaYEnvia", Application.StartupPath + "\\Log.txt");
                        }
                   
                   if (dsBP != null && dsBP.Tables.Count > 0 && dsBP.Tables[0].Rows.Count > 0)
                   {                       
                       foreach (DataRow registro in dsBP.Tables[0].Rows)
                       { 
                          //buscamos que no este en la bitacora i.e. que nunca antes hayan sido enviados.
                           Q = "select * from BP_BITACORA";
                           Q += " where que = 'Envio de CFD con serie/folio: " + registro["VDE_SERIE"].ToString().Trim() + "/" + registro["VDE_FOLIO"].ToString().Trim() + "'";
                           Q += " and id_agencia=" + id_agencia;

                           DataSet dsr = objDB.Consulta(Q);
                           if (dsr != null && dsr.Tables.Count > 0 && dsr.Tables[0].Rows.Count > 0)
                           {
                               //si el registro ya fue enviado por lo menos una vez entonces avisamos en el log el registro de la Bitacora.
                               if (this.InfEmailLogNoHayArchivo.ToUpper().Trim() == "SI")
                                 this.cadenaLog += "Previamente el usuario: " + dsr.Tables[0].Rows[0]["quien"].ToString().Trim() + "\t" + dsr.Tables[0].Rows[0]["que"].ToString().Trim() + " a " + dsr.Tables[0].Rows[0]["aquien"].ToString().Trim() + " el " + dsr.Tables[0].Rows[0]["fecha"].ToString().Trim() + " agencia: " + dsr.Tables[0].Rows[0]["id_agencia"].ToString().Trim() + " " + NombreEmpresa + "\n" + "\r";
                               
                               Utilerias.WriteToLog("Previamente el usuario: " + dsr.Tables[0].Rows[0]["quien"].ToString().Trim() + "\t" + dsr.Tables[0].Rows[0]["que"].ToString().Trim() + " a " + dsr.Tables[0].Rows[0]["aquien"].ToString().Trim() + " el " + dsr.Tables[0].Rows[0]["fecha"].ToString().Trim() + " agencia: " + dsr.Tables[0].Rows[0]["id_agencia"].ToString().Trim() + " " + NombreEmpresa, "ChecaYenvia", Application.StartupPath + "\\Log.txt");
                           }
                           else {
                               //si el registro no ha sido enviado nunca, lo enviamos, registramos en la bitacora y avisamo en el log.                           
                               #region recuperacion de los archivos e intento de envio al destinatario 
                               string serie = registro["VDE_SERIE"].ToString().Trim();
                               string folio = registro["VDE_FOLIO"].ToString().Trim();
                               //string dircorreo = this.modo.ToUpper().Trim()=="DEBUGGING"?"luis.bonnet@grupoandrade.com.mx" : registro["VDE_RMAIL"].ToString().Trim();
                               string dircorreo = registro["VDE_RMAIL"].ToString().Trim();
                               //20220208 
                               dircorreo = dircorreo.ToUpper().Trim() == "JOSE.ABASCAL@INTEGRASOFOM.COM" ? "veronica.barreto@integrasofom.com" : dircorreo.Trim();
                               //20230412
                               string fechaope = registro["fecha"].ToString().Trim();

                               string documento = registro["VDE_DOCTO"].ToString().Trim();
                               
                               strUsrRemoto = regConexion["usr_remoto"].ToString().Trim();
                               string strDominio = "";

                               if (strUsrRemoto.IndexOf("\\") > -1)
                               {   // DANDRADE\sistemas     DANDRADE = dominio sistemas=usuario
                                   strDominio = strUsrRemoto.Substring(0, strUsrRemoto.IndexOf("\\"));
                                   strUsrRemoto = strUsrRemoto.Substring(strUsrRemoto.IndexOf("\\") + 1);
                               }

                               #region funciones de logueo
                               IntPtr token = IntPtr.Zero;
                               IntPtr dupToken = IntPtr.Zero;
                               //primero intentamos el logueo en el servidor remoto
                               bool isSuccess = false;
                               if (strDominio.Trim() != "") //cuando la impersonizacion es en un servidor que pertenece a un dominio entonces es necesario autenticarse haciendo uso del dominio y no de la ip.
                                   isSuccess = LogonUser(strUsrRemoto, strDominio, strPassRemoto, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);
                               else
                                   isSuccess = LogonUser(strUsrRemoto, strIPFileStorage, strPassRemoto, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT, ref token);

                               if (!isSuccess)
                               {
                                   RaiseLastError();
                               }

                               isSuccess = DuplicateToken(token, 2, ref dupToken);
                               if (!isSuccess)
                               {
                                   RaiseLastError();
                               }

                               WindowsIdentity newIdentity = new WindowsIdentity(dupToken);
                               #endregion

                               //En este punto ya debemos tener acceso al servidor remoto para poder traer los archivos;
                               using (newIdentity.Impersonate())
                               {
                                   try
                                   {
                                               string rutaArchivoXML = string.Format("\\\\{0}\\{1}\\{2}.xml", strIPFileStorage, strDirectorioRemotoXML, documento);
                                               string rutaArchivoPDF = string.Format("\\\\{0}\\{1}\\{2}.pdf", strIPFileStorage, strDirectorioRemotoPDF, documento);

                                               if (Convert.ToDouble(fechaope) >= Convert.ToDouble(fechasubclasif) && Convert.ToDouble(fechasubclasif) != 19000101) //yyyyMMdd
                                               { //20230412 se registra en la subcarpeta de cada Año / Mes 
                                                   string AnioFactura = fechaope.Substring(0, 4);
                                                   string MesFactura = fechaope.Substring(4, 2);
                                                   rutaArchivoXML = string.Format("\\\\{0}\\{1}\\{3}{4}\\{2}.xml", strIPFileStorage, strDirectorioRemotoXML, documento, AnioFactura, MesFactura);
                                                   rutaArchivoPDF = string.Format("\\\\{0}\\{1}\\{3}{4}\\{2}.pdf", strIPFileStorage, strDirectorioRemotoPDF, documento, AnioFactura, MesFactura);
                                               }

                                               bool solouna = false;
                                               if (!File.Exists(rutaArchivoXML))
                                               {
                                                   if (this.InfEmailLogNoHayArchivo.ToUpper().Trim()=="SI")
                                                       this.cadenaLog += "No existe el archivo XML: " + rutaArchivoXML + " NO SE ENVIA agencia: " + id_agencia + " " + NombreEmpresa + "\n" + "\r";
                                                   Utilerias.WriteToLog("No existe el archivo XML: " + rutaArchivoXML + " NO SE ENVIA agencia: " + id_agencia  + " " + NombreEmpresa, "ChecaYEnvia", Application.StartupPath + "\\Log.txt"); 
                                                   solouna = true;
                                               }
                                               if (!File.Exists(rutaArchivoPDF) && solouna==false)
                                               {
                                                   if (this.InfEmailLogNoHayArchivo.ToUpper().Trim() == "SI")
                                                        this.cadenaLog += "No existe el archivo PDF: " + rutaArchivoPDF + " NO SE ENVIA agencia: " + id_agencia  + " " + NombreEmpresa  + "\n" + "\r";

                                                   Utilerias.WriteToLog("No existe el archivo PDF: " + rutaArchivoPDF + " NO SE ENVIA agencia: " + id_agencia  + " " + NombreEmpresa, "ChecaYEnvia", Application.StartupPath + "\\Log.txt");
                                               }

                                               if (File.Exists(rutaArchivoXML) && File.Exists(rutaArchivoPDF))
                                               {
                                                   #region Envio del correo al cliente final
                                                   
                                                        if (correosconcopia.Trim() != "")
                                                        {
                                                            dircorreo += "," + correosconcopia.Trim();                                                                                                                         
                                                        }

                                                        string[] EmailsConCopia = dircorreo.Split(',');
                                                   try
                                                   {
                                                       clsEmail correo = new clsEmail(smtpserverhost.Trim(), Convert.ToInt16(smtpport), usrcredential.Trim(), usrpassword.Trim(), enable_ssl.Trim());
                                                       MailMessage mnsj = new MailMessage();
                                                       mnsj.BodyEncoding = System.Text.Encoding.UTF8;
                                                       mnsj.Priority = System.Net.Mail.MailPriority.Normal;
                                                       mnsj.IsBodyHtml = true;
                                                       mnsj.Subject = "Has recibido un nuevo CFD(I)";                                                       
                                                       //mnsj.To.Add(new MailAddress(dircorreo.Trim()));                                                       
                                                       foreach (string Email in EmailsConCopia)
                                                       {
                                                           mnsj.To.Add(new MailAddress(Email.Trim()));
                                                       }

                                                       mnsj.From = new MailAddress(cuenta_from, "Factura Electronica");
                                                       if (File.Exists(rutaArchivoXML))
                                                           mnsj.Attachments.Add(new Attachment(rutaArchivoXML));

                                                       if (File.Exists(rutaArchivoPDF))
                                                           mnsj.Attachments.Add(new Attachment(rutaArchivoPDF));
                                                       
                                                       Dictionary<string, string> TextoIncluir = new Dictionary<string, string>();
                                                       TextoIncluir.Add("fecha", DateTime.Now.ToString());
                                                       TextoIncluir.Add("serie", serie);
                                                       TextoIncluir.Add("folio", folio);
                                                       TextoIncluir.Add("rfcemisor", RFCEmpresa);
                                                       TextoIncluir.Add("nombreemisor", NombreEmpresa);

                                                       AlternateView vistaplana = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoPlano(textocuerpoplano, TextoIncluir), null, "text/plain");
                                                       AlternateView vistahtml = AlternateView.CreateAlternateViewFromString(correo.CreaCuerpoHTML(rutaplantillaHTML, TextoIncluir).ToString(), null, "text/html");
                                                       

                                                       LinkedResource logo = new LinkedResource(rutalogo);
                                                       logo.ContentId = "companylogo";
                                                       vistahtml.LinkedResources.Add(logo);

                                                       mnsj.AlternateViews.Add(vistaplana);
                                                       mnsj.AlternateViews.Add(vistahtml);
                                                       
                                                       if (this.modo.ToUpper().Trim()!="DEBUGGER")
                                                            correo.MandarCorreo(mnsj);
                                                       
                                                       Q = " Insert into BP_BITACORA(fecha,quien,que,aquien,id_agencia)";
                                                       Q += " values (getdate(),'Envio Automatico Programado','Envio de CFD con serie/folio: " + serie + "/" + folio + "','" + dircorreo + "','" + id_agencia + "')";
                                                       if (this.modo.ToUpper().Trim() != "DEBUGGER")
                                                            objDB.EjecUnaInstruccion(Q);

                                                       this.cadenaLog += "Envio Automatico Programado envió el CFDI " + serie + "/" + folio + " a " + dircorreo + " el " + DateTime.Now.ToString() + "\t" + " de la agencia: " + id_agencia + " " + NombreEmpresa + "\n" + "\r";                                                   
                                                   }
                                                   catch (Exception ex)
                                                   {
                                                       res = ex.Message;
                                                       Debug.WriteLine(ex.Message);                                                       
                                                       this.cadenaLog += "Error al intentar enviar el CFDI:  " + serie + "/" + folio + " a " + dircorreo + " de la agencia: " + id_agencia + " " + NombreEmpresa + " " + ex.Message + "\n" + "\r";
                                                       ErroresxAgencia += "Error al intentar enviar el CFDI:  " + serie + "/" + folio + " a " + dircorreo + " de la agencia: " + id_agencia + " " + NombreEmpresa + " " + ex.Message + "\n" + "\r";
                                                   }
                                                   #endregion
                                               }//de que existen los archivos xml y pdf en el servidor remoto. 
                                               
                                               dircorreo = "";
                                               
                                   }//del try
                                   catch (Exception exguia)
                                   {
                                       Debug.WriteLine(exguia.Message);
                                       res = exguia.Message;
                                   }

                                   isSuccess = CloseHandle(token);
                                   if (!isSuccess)
                                   {
                                       RaiseLastError();
                                   }
                               }//del using del usuario autenticado                               
                               #endregion
                           }
                       } //ciclo por cada registro en el día de hoy
                   } //si hubo registros en BP para el dia de hoy
                   if (ErroresxAgencia.Trim() != "" && cuentaslogerrores.Trim() != "")
                           EnviaErroresEnvioResponsablesxAgencia(id_agencia, NombreEmpresa, ErroresxAgencia,cuentaslogerrores.Trim());  
                   conBP.Close(); //para que en la siguiente agencia intente abrir la conexion
               }//foreach ciclo por cada agencia declarada en la BD en la tabla transmision
            }
        else
            {
                Utilerias.WriteToLog("No fue posible conectarse a la BD local tabla Transmision " + this.ConnectionString, "ChecaYEnvia", Application.StartupPath + "\\Log.txt");
                return res = "No fue posible conectarse a la BD local tabla Transmision";
            }
            #endregion  
      
            return res;    
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            this.Hide();
            this.Visible = false;
        }

        //le debe llegar sin la extension .exe
        private int CuentaInstancias(string NombreProceso)
        {
            int res = 0;
            try
            {
                if (NombreProceso.Trim() != "")
                {
                    Process[] localByName = Process.GetProcessesByName(NombreProceso);
                    res = localByName.Length;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return res;
        }
    }
}
