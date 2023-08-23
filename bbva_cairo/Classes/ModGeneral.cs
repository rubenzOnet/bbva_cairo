using ADODB;
using static ADODB.LockTypeEnum;
using static ADODB.CursorTypeEnum;


using static System.Collections.Specialized.BitVector32;
using bbva_cairo.Formularios;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Security.Cryptography;
using System.Collections.Immutable;

namespace bbva_cairo.Classes
{



    public static class ModGeneral
    {


        //_______________________________________________________________
        //Título: Rutinas Generales para el Proyecto CAIRO
        //Modulo: ModGeneral.bas
        //Versión: 1.0
        //Fecha:    20/06/2000
        //Autor:    Roberto Reyes I - Gerardo Acosta.
        //Modificación:
        //Fecha de Modificacion:
        //_______________________________________________________________
        //Descripción:
        //Este Módulo muestra las Funciones Generales para el Proyecto Cairo.
        //Así como la declaración de Variables Globales para el Proyecto.
        //_______________________________________________________________DB.

        //Variables Globales
        //20/06/2000 RReyes
        public static string gsUsrRed;               //Variable de Usuario de RED
        public static string gsUsrBD;                //Variable Usuario de Base de Datos
        public static string gsUsrBD_Net;            //Variable Usuario de Base de Datos .Net
        public static string gsUsrSistema;           //Usuario que se firma en el Sistema
        public static string gsHostName;            //HostName de la Máquina Local
        public static string gsDSN;                   //Data Source Name al cual nos Conectamos
        public static string gsDSN_Net;               //Data Source Name al cual nos Conectamos para .Net
        public static string gsDSN_NetCrystal;              //Data Source Name al cual nos Conectamos para .Net

        public static string gsServidor;              //Servidor que utilizamos para la conexion
        public static string gsServidor_Net;          //Servidor que utilizamos para la conexion a .Net
        // *** Héctor García 10/dic/2010 Se agregar variable goblal para determinar el nombre del servidor de Componentes (MTS)
        public static string gsServidorMTS;          //Servidor de MTS
        public static string gsConexion = @"Provider=sqloledb;Server=DESKTOP-123TQML\SQLEXPRESS; Database=BBVASysView;User Id=Onet_Sa;Password=Onet_Sa;";            //Es el String para ADO al hacer la conexion a la Base de Datos
        public static string gsConexion_Net;         //Es el String para ADO al hacer la conexion a la Base de Datos para .Net
        public static string gsConexionCrystal;      //Es una Conexion ADO para hacer la conexio a .net y CRYSTAL REPORT
        public static string gsPathRpts;               //Path de los Reportes
        public static string gsVerSistema;           //Variable para la Version del Sietema
        public static string gsFechaSistema;          //Variable para la Fecha del Sistema, esta es Modificable
        public static int giIDUsuario = 10115213;      //ID del Usuario Conectado
        public static string gsModulo;              //Indica la etiqueta a Desplegar en los mensajes
        public static string gsBaseDatos;            //Indica la Base de Datos a la cual nos conectamos
        public static string gsPwdBD;                //Indica el Password de la Base de Datos a la cual nos conectamos
        public static string gsPwdBD_Net;           //Indica el Password de la Base de Datos a la cual nos conectamos para .Net
        public static string gsProvider;             //Indica el Proveedor de la Base de Datos
        public static string gsReportes;             //Indica el Proveedor de la Base de Datos
        public static string vgNombre_ini;           //Variable para localizacion del Archivo C:\WINDOWS\Cairo.ini (Temporal)
        //Variable de la Tabla Parametros
        //18/09/2000 RReyes
        public static double gdPjePago;              //Porcentaje de Pago
        public static double gdComision;             //Comision
        public static double gcMttoAdic;      //Monto Adicional
        public static string gsPathBase;              //Path de la Base de Archivos Generados y Guardados
        public static string gsDirCNSF;               //Direccion para la conexión a la Página de la CNSF
        public static int gsDParamSUC;           //Días para Mensaje de Parametros del SUC
        public static string gsPagAyuda;              //Variable para direccionar la Página de Ayuda en Linea
        public static int MTS;                   //Variable para controlar la bandera para Transacciones MTS
        //SDC 2007-03-01 Nombre del usuario
        public static string gsNombreUsrSistema;           //Usuario que se firma en el Sistema

        //Front Dispersion Pagos(Semaforo) JCMN 04/09/2014
        public static string gsPathSemaforo;
        public static bool ParamGenSemaf;
        public static bool FileDispersionBoolean;

        //AGonzalez 13 de agosto de 2005
        public static bool gbDoctosSobrevivencia; //Bandera para saber si están impriendo carátulas de Sobrevivencia

        //Variable de la Clase General
        //20/06/2000 RReyes
        public static ClsGeneral gClsGeneral;

        //Variables de los Componentes Generales
        //20/06/2000 RReyes
        public static object gObjErrorSuc;          //Variable para manejar el Componente MTS de Errores de SUC
        public static object gObjParametros;         //Variable para manejar el Componente MTS de Parametros

        //20/08/2000 AGonzalez
        public static object gObjCotizacion;

        //15/05/2001 AGonzalez
        public static object gObjCSeguridad;

        //Variables de laEmpresa Nombre y ID de la Empresa
        //20/06/2000 GAcosta
        public static string gsEmpresa;
        public static int giIDEmpresa;
        public static Recordset grsEmpresas;

        //Tipo para el resultado de EjecutaSql
        //20/06/2000 GAcosta
        //mafm 05052022

        // hace falta implementar esta variable
        // 19 mayo 2023
        // RGB
        //public Itemx As MSComctlLib.ListItem

        //mafm 05052022

        public enum TipoResultado
        {
            DatosOK = 1,
            NoHayDatos = 2,
            ExisteError = 3
        }


        public static int DatosOK = 1;
        public static int NoHayDatos = 2;
        public static int ExisteError = 3;




        //Tipo para el resultado de la Carga de la Forma FrmCarga.frm
        //20/06/2000 AGonzalez
        public enum TipoCarga
        {
            Int_Oferta = 1,
            Carga_Oferta = 2,
            Carga_Beneficios = 3,
            Int_Resol = 4,
            Base_Resol = 5,
            Suc_Aseg = 6,
            Suc_Benef = 7,
            Int_Ingresos = 8,
            Base_Ingresos = 9,
            Int_Titular = 10,
            Base_Titular = 11,
            Concilia_IngPend = 12,
            Cuentas_Titular = 13,
            Polizas_Tramite = 14,
            Polizas_Emision = 15,
            Imp_Caratulas = 16,
            Imp_ProgPagos = 17,
            Imp_ReciboDoctos = 18
        }

        //Arreglo para manejar las rutas de los archivos que se cargan en la B.D.
        //31/08/2000 Agonzalez
        public static string[] sArrArchivo = new string[3];

        public static frmMensaje gfMsgbox = new frmMensaje();            //Es la forma de los mensajes
        public static string gsErrVB;                                //Contiene los errores generados por VB en el cliente
        public static Recordset grsErrADO;      //Contiene los errores enviados por la clase

        public const string gsINGLESA = "mm/dd/yyyy";
        public const string gsFRANCESA = "dd/mmm/yyyy";

        //Variables para tipos de InstituciónSS
        public const int gnIMSS_97 = 1;
        public const int gnIMSS_08 = 2;
        public const int gnISSSTE = 3;
        //Variables para tipos de Archivos de Ingresos
        public const string gsARCH_BMX = "B";
        public const string gsARCH_PRC = "P";

        //Front Dispersion Pagos(Semaforo) JCMN 04/09/2014
        public static int contadorTime;
        public static object objSemaforoScan;
        //JCMN 17/09/2014 adecuacion eliminacion de archivos de dispersión de prestamos(SEMAFORO)
        public static string[] arrayTxtDispersion = new string[4];
        //Tipo para el Manejo de los datos Erroneos de la Revision de Errores del SUC
        //31/08/2000 RReyes
        public enum ErrorSuc                                           //Tipo para el resultado del Chequeo de Errores de SUC
        {
            Datos_OK = 1,                                                   //Datos Correctos, Error Cero para las Cargas del SUC
            Datos_Erroneos = 2,                                           //Se encontraron datos Erroneos, pero se prosigue a la Carga del SUC
            Suspension_Carga = 3                                        //Se suspende la carga del SUC por encontrar Errores Criticos
        }

        //Tipo para el Manejo de los Reportes
        //02/04/2001 RReyes
        public enum Reportes                                           //Tipo para el resultado de los Reportes
        {
            TXCondPago = 1,                                              //Reporte de Titulares por Conductos de Pago
            TAperturaCuentas = 2,                                        //Reporte de Titulares por Apertura de Cuenta
            TMovimientos = 3,                                            //Reporte de Titulares por Movimientos
            RDoctosC = 4,                                                //Reporte de Resoluciones Con Documentación Completa
            RDoctosI = 5,                                                //Reporte de Resoluciones Con Documentación Incompleta
            RDoctosSD = 6,                                               //Reporte de Resoluciones Sin Documentación
            RPbancomer = 7,                                              //Reporte de Resoluciones de Pensiones Bancomer
            PPagadas = 8,                                                //Reporte de Polizas Pagadas
            PPendientes = 9,                                             //Reporte de Polizas Pendientes
            PPEnviar = 10,                                               //Reporte de Polizas Por Enviar
            PRPolizaMes = 11,                                            //Reporte de Resumen de Polizas en el Mes
            PPolEmi = 12,                                                //Reporte de Polizas Emitidas para prospect
            PGastosFunerarios = 13,                                      //Reporte de Gastos Funerarios
            PPolizasEmitidas = 14,                                       //Reporte de Polizas Emitidas
            PSiniestro = 15,                                             //Reporte de Siniestros
            PSusFall = 16,                                               //Reporte de Suspendidos-Fallecidos
            SUC = 17,                                                    //Generar archivo de póliza para SUC
            EEndosos = 18,                                               //SDC 2007-03-22 Reportes de endosos
            NFechaAplicPagel = 19,                                       //SDC 2007-06-13 Fecha de Aplicacion para Layout de Pagel
            //Incio Reporte archivo pagos Erick Bejarano 26/07/2017
            PPagos = 21,                                                 //Reporte de archivo de pagos
            //Fin Reporte archivo pagos Erick Bejarano 26/07/2017
            //Inicio Reporte RFc Beneficiarios -- Julio Cesar Martínez Nava 30/01/2019
            RptRFCBenefs = 22,
            RptPEmis = 23,
            //Inicio Reporte Datos Fiscales -- Alexander Hdez 2022-02-09
            //Inicio Reporte RFc Beneficiarios -- Julio Cesar Martínez Nava 30/01/2019
            RptDatosFiscales = 24,
            //Fin Reporte Datos Fiscales -- Alexander Hdez 2022-02-09
        }

        //Tipo para el Manejo de las Carátulas
        //02/04/2001 RReyes
        public enum Caratulas                                          //Tipo para el resultado de las Caratulas
        { 
            CPoliza = 0,                                                 //Carátulas de Pólizas
            CReciboDoctos = 1,                                           //Recibo de Documentos
            CProgPagos = 2,                                              //Programas de Pagos
            CEndosos = 3,                                                //Endosos
            CCartaCompromiso = 4                                        //Carta Compromiso
        }

        //Tipo para el Manejo de los Endosos Internos
        //12/09/2002 RReyes
        public enum Endosos                                            //Tipo para selección del Endoso
        { 
            Alta_Componente = 0,                                         //Forma de Endoso de Alta de Componente
            Baja_Fallecimiento = 1,                                         //Forma de Endoso de Baja de Componente
            Baja_Nupcias = 2,                                         //Forma de Endoso de Baja de Componente
            Baja_Improcedencia = 3,                                         //Forma de Endoso de Baja de Componente
            Cambio_Ramo = 4,                                             //Forma de Cambio de Ramo
            DivGrupos_Fam = 5                                           //Forma de División de Grupos Familiares
        }

        public static object vErrorSuc;                                 //Variable  para el Regreso de los Num_Oferta erroneos

        //SDC 20/02/2004
        public static string WinDir;


        [DllImport("kernel32.dll", EntryPoint = "GetWindowsDirectory")]
        public static extern long GetWindowsDirectoryA(string lpBuffer, long nSize);

        [DllImport("kernel32.dll", EntryPoint = "GetSystemDirectory")]
        public static extern long GetSystemDirectoryA(string lpBuffer, long nSize);

        //SDC 2007-06-26 Para utilizar en GeneraLayouts para la referencia de pago
        //Dar de alta aquí en un futuro si hay otro proceso de pago
        public enum ProcesoPago
        { 
            PNomina = 0,
            PPagosVencidos = 1,
            PEndosos = 2,
            PSobrevivencia = 3,
            PArticulo14 = 4,
            PSuspensionEncuesta = 5 //Este no se utiliza con GeneraLayouts, pero esta reservado para llamar Pagel directamente
        }

        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileStringA")]
        public static extern long GetPrivateProfileString(string lpApplicationName,object lpKeyName, string lpDefault, string lpReturnedString, long nSize, string lpFileName);

        [DllImport("kernel32.dll", EntryPoint = "WritePrivateProfileStringA")]
        public static extern long WritePrivateProfileString(string lpApplicationName, object lpKeyName, object lpString, string lpFileName );

        // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
        // Juan Martínez Díaz
        // 13/05/2022

        [DllImport("kernel32.dll", EntryPoint = "WritePrivateProfileSectionA")]
        public static extern long WritePrivateProfileSection(string lpAppName, string lpString, string lpFileName);



        //Función Api SetWindowLong
        [DllImport("user32.dll", EntryPoint = "WritePrivateProfileSectionA")]
        public static extern long SetWindowLong(long hwnd, long nIndex, long dwNewLong );


        // Función Api CallWindowProc
        [DllImport("user32.dll", EntryPoint = "CallWindowProcA")]
        public static extern long CallWindowProc(long lpPrevWndFunc, long hwnd, long Msg, long wParam, long lParam );

        //constantes
        ////////////////////////////////////////////////

        // mensaje para el menú contextual
        public const long WM_CONTEXTMENU = 0x007B;

        public static long lpPrevWndProc;


        //JCMN modificación a Polizas en ROPC 02102014
        public static double IdPolizaROPC;
        //Ajuste con beneficios Duales en front(Aguinaldo) JCMN 14-05-2015 --Inicio
        public static double IdPolizaMasUna;
        //Ajuste con beneficios Duales en front(Aguinaldo) JCMN 14-05-2015 --Fin
        public static int Id_GrupoROPC;
        public static double IdPagoPolizaROPC;
        public static double IdPolBandera;
        public static int columnaROPC; // String
        public static object MontoROPCActual;
        public static object MontoROPCNuevo;
        public static int RowPolBenef;
        public static string dFecha_SAOR;  //< Almacena la fecha del control txtResult(11) del formulario frmEndosoCET EBS 17/02/2016
        public static string sCboEndoso; //< Almacena la opcion elegida por el usuario CboEndoso en formulario frmEndosoCET EBS 17/02/2016
        public static DateTime dFecha_Emision ; //< Almacena la fecha de endoso de la poliza EBS 17/02/2016

        // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
        // Juan Martínez Díaz
        // 2022-11-23
        public static string gsOdbcActivo;
        public static string gsConexion2019Dualidad;    //Es el String para ADO al hacer la conexion a la Base de Datos de 2019
        public static string gsConexionDualidadSel;

        public static void Def_VarsPublic()
        {
            //On Error GoTo msgerror

            //SE OBTIENE EL NOMBRE DEL DSN QUE SE VA A UTILIZAR
            WinDir = GetWinDir();
            //FJMH 04/01/2021 Inicio Ruta archivo INI
            //vgNombre_ini = WinDir & "\Cairo2000.ini"
            vgNombre_ini = @"c:\Cairo\Cairo2000.ini";
            //FJMH 04/01/2021 Fin Ruta archivo INI
            if (File.Exists(vgNombre_ini))
            {
                ReadINI(vgNombre_ini, "SERVIDOR", "Servidor", gsServidor);
                ReadINI(vgNombre_ini, "ODBC", "DSN", gsDSN);
                ReadINI(vgNombre_ini, "USUARIO", "UID", gsUsrBD);
                ReadINI(vgNombre_ini, "PASSWORD", "PWD", gsPwdBD);

                //Variables de .Net

                ReadINI(vgNombre_ini, "ODBC_NETCRYSTAL", "DSN", gsDSN_NetCrystal);
                ReadINI(vgNombre_ini, "SERVIDOR_NET", "Servidor", gsServidor_Net);
                ReadINI(vgNombre_ini, "ODBC_NET", "DSN", gsDSN_Net);
                ReadINI(vgNombre_ini, "USUARIO_NET", "UID", gsUsrBD_Net);
                ReadINI(vgNombre_ini, "PASSWORD_NET", "PWD", gsPwdBD_Net);

                //I *** Héctor García 10-dic-2010 Recupera el nombre del Servidor de Componentes MTS
                ReadINI(vgNombre_ini, "SERVIDORMTS", "Servidormts", gsServidorMTS);
                if (gsServidorMTS.Trim() == string.Empty)
                {
                    gsServidorMTS = gsServidor;
                }
                //F *** Héctor García 10-dic-2010 Recupera el nombre del Servidor de Componentes MTS

                //ReadINI vgNombre_ini, "USUARIO", "UID", gsUsrBD
                //ReadINI vgNombre_ini, "PASSWORD", "PWD", gsPwdBD
                //        ReadINI vgNombre_ini, "VERSION", "VerProspect", vgVersion
                //if vgUID = "Desarrollo" Then
                //   vgPWD = "site"
                //else
                //   vgPWD = "bancomer"
                //End if

                // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
                // Juan Martínez Díaz
                // 2022-11-23
                // validamos si no existe la conexión a la base 2019 y agregamos sus valores
                //ReadINI vgNombre_ini, "ODBC_ACTIVO", "OdbcActivo", gsOdbcActivo
                //if Trim(gsOdbcActivo) = "" Then
                //   WriteIni vgNombre_ini, "ODBC_ACTIVO", "OdbcActivo", "000020755259605A5857421277" & vbCr
                //  WriteIni vgNombre_ini, "SERVIDOR_2", "Servidor", "000040280A1531031F0A0154193D123E3C282472272A73" & vbCr
                // WriteIni vgNombre_ini, "ODBC_2", "DSN", "00001807243F022045554319" & vbCr
                //WriteIni vgNombre_ini, "USUARIO_2", "UID", "0000043704" & vbCr
                // WriteIni vgNombre_ini, "PASSWORD_2", "PWD", "000018141703630D37415C0E" & vbCr

                // variables de .Net
                // WriteIni vgNombre_ini, "ODBC_NETCRYSTAL_2", "DSN", "00003207243F0220282B3774063C1A212D3138" & vbCr
                // WriteIni vgNombre_ini, "SERVIDOR_NET_2", "Servidor", "0000307550467E5E47555C117C586D404942" & vbCr
                // WriteIni vgNombre_ini, "ODBC_NET_2", "DSN", "00001807243F0220282B3774" & vbCr
                // WriteIni vgNombre_ini, "USUARIO_NET_2", "UID", "000014141703350D1616" & vbCr
                // WriteIni vgNombre_ini, "PASSWORD_NET_2", "PWD", "000016141703630D37415C" & vbCr

                // WriteIni vgNombre_ini, "SERVIDORMTS_2", "Servidormts", "0000307550467E5E47555C117C586D404942" & vbCr

                // ReadINI vgNombre_ini, "ODBC_ACTIVO", "OdbcActivo", gsOdbcActivo
                // End if
            }

            //Definiendo la conección para los reportes al Crystal
            //vgConecta = "SERVER=" & vgServidor & ";DSN=" & vgDSN & ";DBQ=;UID=" & vgUID & ";PWD=" & vgPWD & ";"
            //Definiendo Variables Globales para la Seguridad
            //vgUsuario = ""
            //vgPassword = ""
        }

    //    public static int ExistFile(string FileName)
    //    {
    //        Dim Filenum

    //        On Error GoTo Exist

    //        Filenum = FreeFile()
    //        FileOpen(Filenum, FileName, OpenMode.Input)
    //        ExistFile = True
    //        FileClose(Filenum)
    //        Exit Function

    //Exist:
    //        Select Case Err().Number
    //            Case 53       //file not found
    //                ExistFile = False
    //            Case 55       //file already open
    //                ExistFile = True
    //            Case 58       //file already exists
    //                ExistFile = True
    //            Case 76       //path not found
    //                ExistFile = False
    //        End Select

    //        FileClose(Filenum)
    //        Exit Function

    //    }

        private static void ReadINI(string INifile, string Seccion, string VARIABLE, object Resultado)
        {
            string res = new string('*', 255);
            long l = 0;

            l = GetPrivateProfileString(Seccion, VARIABLE, "", res, 255, INifile);
            string ressult = res.Substring(0, Convert.ToInt32(l));  //Left$(res, l)
        }

        // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
        // Juan Martínez Díaz
        // 13/05/2022
        public static void WriteIni(string FileName, string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, FileName);
        }

        // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
        // Juan Martínez Díaz
        // 13/05/2022
        public static void WriteIniSection(string FileName, string Section, string Value)
        {
            WritePrivateProfileSection(Section, Value, FileName);
        }


        //Funcion General para el comodín % - * en las consultas
        //20/06/2000 RReyes
        public static string sComodin(string sCriterio, char vCarater = '%')
        {
            int iPosicion = 0;

            if (sCriterio == string.Empty)
                sCriterio = vCarater.ToString();
            else if (vCarater != '*')
                iPosicion = 1;
            while (iPosicion > 0)
            { 
                iPosicion = sCriterio.IndexOf("*", 1);  //InStr(1, sCriterio, "*")
                if (iPosicion > 0)
                {
                    // Mid(sCriterio, iPosicion, 1) = vCarater;   
                }
            }

            return sCriterio;
        }

        //Funcion General para Numeros 0 al 9 y punto
        //20/06/2000 RReyes
        public static bool iNumeros(KeyPressEventArgs e)
        {
            //Dim char As String
            //Permitiendo unicamente la captura de Numeros
            //if (!IsNumeric(char(KeyAscii)) && KeyAscii != 8 && KeyAscii != 42 && KeyAscii != 46) KeyAscii = 0;
            //return KeyAscii;
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Ignorar el carácter ingresado
                return false;
            }
            return true;
        }

        //Funcion General para Telefono
        //21/10/2004 SDC
        //public static int iTelefono(KeyAscii As Integer) As Integer
        //    //Permitiendo unicamente la captura de Numeros
        //    if(KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = Asc(vbBack) Then
        //        iTelefono = KeyAscii
        //    else
        //        iTelefono = 0
        //    End if
        //End Function

        //Funcion General para únicamente validación de Numeros Enteros 0 al 9
        //20/06/2000 RReyes
        //public Function iEnteros(KeyAscii As Integer) As Integer
        //    //Dim char As String
        //    //Permitiendo unicamente la captura de Numeros
        //    if Not IsNumeric(Chr(KeyAscii)) And KeyAscii!= 8 Then KeyAscii = 0
        //    iEnteros = KeyAscii
        //End Function

        //Funcion General para Validacion de Mayusculas y otro caracter
        //20/06/2000 RReyes
        //public Function iMayusculas(KeyAscii As Integer)
        //    Dim _char As String
        //    //Permitiendo unicamente la captura de mayusculas
        //    if Not((KeyAscii >= 65 And KeyAscii <= 90) Or(KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = Asc("ñ") Or KeyAscii = Asc("Ñ") Or KeyAscii = Asc(" ") Or KeyAscii = Asc(vbBack)) Then
        //        Exit Function
        //    End if
        //    _char = Chr(KeyAscii)
        //    iMayusculas = Asc(UCase(_char))
        //End Function


        //Funcion General para Validacion de Mayusculas Sin Restricciones de caracter
        //16/06/2005 SDC
        //public Function iUCase(KeyAscii As Integer)
        //    iUCase = Asc(UCase(Chr(KeyAscii)))
        //End Function

        //Funcion General para Validacion Formato de Salida al SUC.
        //Asigna el formato Espacios a un tamaño especifico.
        //21/07/2000 RReyes
        //public Function sPadR(sCad As String, byTam As Byte) As String
        //    if Len(sCad) > byTam Then
        //        sPadR = Mid$(sCad, 1, byTam)
        //    else
        //        sPadR = sCad & Space(byTam - Len(sCad))
        //    End if
        //End Function

        //Funcion General para Validacion Formato de Salida al SUC.
        //Asigna el formato Money al campo para el Archivo de Salida del SUC.
        //21/07/2000 RReyes
        //public Function sFormaSal(cNum As Double) As String
        //    Dim sCNum As String
        //    Dim byTam As Byte
        //    Dim byPos As Byte

        //    byTam = 13
        //    if cNum!= 0 Then
        //        sCNum = Strings.Format$(cNum, "##########0.00")
        //        sCNum = LTrim(sCNum)
        //        byPos = InStr(sCNum, ".")
        //        sCNum = Mid(sCNum, 1, byPos - 1) & Mid(sCNum, byPos + 1)
        //        While Len(sCNum) < byTam
        //            sCNum = "0" & sCNum
        //        End While
        //    else
        //        sCNum = Space(byTam)
        //    End if
        //    sFormaSal = sCNum
        //End Function

        //Funcion General para Validacion Formato de Salida al SUC.
        //Quita los caractéres raros y limpia la cadena dejando solo los carcteres.
        //21/07/2000 RReyes
        //public Function sFrmtEsp(iNum As Integer, Optional iTam As Integer = 5) As String
        //    Dim SNum5 As String //* 5
        //    Dim SNum4 As String //* 4
        //    Dim sCX As String

        //    if iTam = 5 Then
        //        if iNum!= 0 Then
        //            sCX = LTrim(Str(iNum * 100))
        //            SNum5 = Iif(Len(sCX) < 5, Str(Val(sCX)), sCX)
        //        else
        //            SNum5 = Space(5)
        //        End if
        //        sFrmtEsp = SNum5
        //    elseif iTam = 4 Then
        //        sCX = LTrim(RTrim(iNum))
        //        if iNum!= 0 Then
        //            SNum4 = Iif(Len(sCX) < 4, Str(Val(sCX)), sCX)
        //        else
        //            SNum4 = Space(4)
        //        End if
        //        sFrmtEsp = SNum4
        //    End if
        //End Function

        //Funcion General para El Recordset Desconectado
        //21/07/2000 GAcosta
        public static Recordset rsRecordset(Recordset objRecordset, LockTypeEnum LockType = adLockBatchOptimistic, CursorTypeEnum CursorType = adOpenDynamic)
        {
            try
            {
                // On Error GoTo msgerror
                Recordset objNewRS;
                object objField;
                int lngCnt;

                string errVB;

                objNewRS = new Recordset();
                objNewRS.CursorLocation = CursorLocationEnum.adUseClient;
                objNewRS.LockType = LockType;
                objNewRS.CursorType = CursorType;

                //foreach (var objField in objRecordset.Fields)
                //{
                //    objNewRS.Fields.Append(objField.Name, objField.Type, objField.DefinedSize, objField.Attributes);
                //}

                objNewRS.Open();

                objNewRS = objRecordset;



                //if Not objRecordset.RecordCount = 0 Then
                //    objRecordset.MoveFirst()

                //    While Not objRecordset.EOF
                //        objNewRS.AddNew()

                //        For lngCnt = 0 To objRecordset.Fields.Count - 1
                //            objNewRS.Fields(lngCnt).Value = objRecordset.Fields(lngCnt).Value
                //        Next lngCnt
                //        objRecordset.MoveNext()
                //    End While
                //End if
                
                }
            catch (Exception Err)
            {
                //Screen.MousePointer = vbDefault
                //errVB = Strings.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description;
                //gfMsgbox.Mensaje(" Error del Sistema Prospect ", frmMensaje.ETipos.vbOKDetails + vbCritical, "P R O S P E C T", , errVB)
            }

            return objNewRS;
        }


        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //public Sub IniciaProceso(Mensaje As String)
        //    Dim i As Long

        //    frmRentas.Animation1.Open(app.path + "\DOWNLOAD.AVI")
        //    frmRentas.Animation1.Play(-1)
        //    frmRentas.Animation1.Height = 240
        //    frmRentas.abdRentas.Bands("stbPrincipal").Tools("txtStatus").Caption = Mensaje

        //    With frmRentas.abdRentas.Bands("stbPrincipal").Tools("ctlAnimation")
        //        .Custom = frmRentas.Animation1
        //        .Visible = True
        //        // Tool.Height/width should be in twips, its in pixels.
        //        .Height = 225 //* Screen.TwipsPerPixelY
        //        .Width = 225 //* Screen.TwipsPerPixelX

        //    End With

        //    frmRentas.Animation1.Visible = True
        //    frmRentas.abdRentas.RecalcLayout
        //    frmRentas.abdRentas.Refresh

        //End Sub


        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //public Sub FinProceso()

        //    frmRentas.abdRentas.Bands("stbPrincipal").Tools("ctlAnimation").Visible = False
        //    frmRentas.Animation1.Visible = False
        //    frmRentas.Animation1.Stop
        //    frmRentas.abdRentas.Bands("stbPrincipal").Tools("txtStatus").Caption = "Listo....."
        //    frmRentas.abdRentas.RecalcLayout

        //End Sub



        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //    //* Llena un combo con el objeto RecorSet que se le pase
        //    public Sub LlenaCombo(rsData As ADODB.Recordset, cboCombo As ComboBox, Optional bCampoFijo As Boolean = False, Optional svalor As String = " ")
        //        Dim Columna As Field
        //        Dim sDato As String
        //        Dim ErroresVB As String

        //        On Error GoTo LlenaCombo

        //        cboCombo.Clear
        //        //Checamos si la comobo va a llevar un campo fijo
        //        if bCampoFijo Then
        //            cboCombo.AddItem(svalor)
        //            cboCombo.ItemData(cboCombo.NewIndex) = -1
        //        End if

        //        rsData.MoveFirst()

        //        While Not rsData.EOF
        //            if IsNothing(rsData.Fields.Item(1).Value) Then
        //                cboCombo.AddItem("")
        //            else
        //                cboCombo.AddItem(rsData.Fields.Item(1).Value)
        //            End if
        //            cboCombo.ItemData(cboCombo.NewIndex) = rsData.Fields.Item(0).Value
        //            rsData.MoveNext()
        //        End While
        //        rsData.Close()

        //        Exit Sub

        //LlenaCombo:
        //        ErroresVB = Strings.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        gfMsgbox.Mensaje(" Error al Llenar catalogos  ", vbOKDetails + vbCritical, "C A I R O", , ErroresVB)

        //    End Sub



        //Divide un String con ApP ApM Nom concatenados, los separa.
        //2/10/2000 AGonzalez
        public static bool bDivideNombre(ref string sApP, ref string sApM, ref string sNom, ref string Nombre)
        {

            int iLong = 0;
            string sOriginal = "";
            int iR = 0;
            string sPaterno = "";
            string sMaterno = "";
            string sNombre = "";
            int[] iPosicion = new int[5];

            // On Error GoTo msgerror
            try
            {


                sOriginal = Nombre;
                iPosicion[0] = 1;

                iLong = sOriginal.Length;

                //OBTENEMOS EL APELLIDO PATERNO

                for (int i = 0; i < iLong; i++)
                {
                    if (sOriginal.Substring(iR, 1) == " ") {
                        sPaterno = sOriginal.Substring(iPosicion[0], iR - 1);
                        iPosicion[1] = iR + 1;
                        break;
                    }
                }

                for (int i = 0; i < iLong; i++)
                {
                    // if Mid$(sOriginal, iR, 1) = " " Then
                    if (sOriginal.Substring(iR, 1) == " ")
                    {
                        sPaterno = sOriginal.Substring(iPosicion[0], iR - 1);  //Mid$(sOriginal, iPosicion(0), iR - 1)
                        iPosicion[1] = iR + 1;
                        break;
                    }
                }


                if ( (sPaterno.Trim() == "DE") || (sPaterno.Trim() == "Y") || (sPaterno.Trim() == "DEL") || (sPaterno.Trim() == "LA"))
                {
                
                    for (int i = iR + 1; i <= iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sPaterno = sPaterno + " " + sOriginal.Substring(iPosicion[1], iR - iPosicion[1]);
                            iPosicion[1] = iR + 1;
                            break;
                        }
                    }
                }
                //////////////Falta validacion para la Y y L
                if (sOriginal.Substring(iR + 1, 2) == "Y")
                {
                    for (int i = iR + 3; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ") {
                            sPaterno = sPaterno + " " + sOriginal.Substring(iPosicion[1], iR - iPosicion[1]);
                            iPosicion[1] = iR + 1;
                            break;
                        }
                    }
                }

                if (sOriginal.Substring(iR - 3, 3) == " LA") 
                {
                    for (int i = iR + 1; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sPaterno = sPaterno + " " + sOriginal.Substring(iPosicion[1], iR - iPosicion[1]);
                            iPosicion[1] = iR + 1;
                            break;
                        }
                    }
                }

                //////////////////////////////////////////////////////////////////////////
                if (sOriginal.Substring(iR + 1, 2) == "Y ") 
                {
                    for (int i = iR + 3; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sPaterno = sPaterno + " " + sOriginal.Substring(iPosicion[1], iR - iPosicion[1]);
                            iPosicion[1] = iR + 1;
                            break;
                        }
                    }
                }

                if (sOriginal.Substring(iR - 3, 3) == " LA")
                {
                    for (int i = iR + 1; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sPaterno = sPaterno + " " + sOriginal.Substring(iPosicion[1], iR - iPosicion[1]);
                            iPosicion[1] = iR + 1;
                            break;
                        }
                    }
                }

                //////////////////////////////////////////////////////////////////////////

                //AHORA EL MATERNO
                for (int i = iR + 1; i < iLong; i++)
                {
                    if (sOriginal.Substring(iR, 1) == " ")
                    {
                        sMaterno = sOriginal.Substring(iPosicion[1], iR - iPosicion[1]);
                        iPosicion[2] = iR + 1;
                        break;
                    }
                }

                if (sMaterno.Trim() == "DE" || sMaterno.Trim() == "Y" || sMaterno.Trim() == "DEL")
                {
                    for (int i = iR + 1; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sMaterno = sMaterno + " " + sOriginal.Substring(iPosicion[2], iR - iPosicion[2]);
                            iPosicion[2] = iR + 1;
                            break;
                        }
                    }
                }

                if (sOriginal.Substring(iR + 1, 2) == "Y ")
                {
                    for (int i = iR + 3; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sMaterno = sMaterno + " " + sOriginal.Substring(iPosicion[2], iR - iPosicion[2]);
                            iPosicion[2] = iR + 1;
                            break;
                        }
                    }
                }

                if (sOriginal.Substring(iR - 3, 3) == " LA")
                {
                    for (int i = iR + 1; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sMaterno = sMaterno + " " + sOriginal.Substring(iPosicion[2], iR - iPosicion[2]);
                            iPosicion[2] = iR + 1;
                            break;
                        }
                    }
                }

                //////////////////////////////////////////////////////////////////////////
                if (sOriginal.Substring(iR + 1, 2) == "Y ")
                {
                    for (int i = iR + 3; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sMaterno = sMaterno + " " + sOriginal.Substring(iPosicion[2], iR - iPosicion[2]);
                            iPosicion[2] = iR + 1;
                            break;
                        }
                    }
                }

                if ( sOriginal.Substring(iR - 3, 3) == " LA")
                {
                    for (int i = iR + 1; i < iLong; i++)
                    {
                        if ( sOriginal.Substring(iR, 1) == " ")
                        {
                            sMaterno = sMaterno + " " + sOriginal.Substring(iPosicion[2], iR - iPosicion[2]);
                            iPosicion[2] = iR + 1;
                            break;
                        }
                    }
                }

                if ( (sOriginal.Substring(iR - 3, 3) == "VDA") || (sOriginal.Substring(iR - 4, 4) == "VDA.") )
                {
                    for (int i = iR + 1; i < iLong; i++)
                    {
                        if (sOriginal.Substring(iR, 1) == " ")
                        {
                            sMaterno = sMaterno + " " + sOriginal.Substring(iPosicion[2], iR - iPosicion[2]);
                            iPosicion[2] = iR + 1;
                            break;
                        }
                    }
                }

                //////////////////////////////////////////////////////////////////////////

                //AHORA EL(LOS) NOMBRE(S)
                if (iPosicion[2] == 0)
                {
                    sNombre = sOriginal.Substring(iPosicion[1], iLong + 1 - iPosicion[1]);
                    sApP = sPaterno;
                    sApM = new string(' ', 20);
                    sNom = sNombre;
                }
                else
                {
                    sNombre = sOriginal.Substring(iPosicion[2], iLong + 1 - iPosicion[2]);
                    sApP = sPaterno;
                    sApM = sMaterno;
                    sNom = sNombre;
                }

                return true;

                // Exit Function
            }
            catch (Exception)
            {
                return false;
            }
        }

        // hace falta implementar esta funcion
        // 22 agosto 2023
        // hacer falta la funcion 
        // RGB
        //public static bool bVal_NSS(string NSS)
        //{
        //    int DigitoVerificador;
        //    int DigitoVerificadorReal;
        //    int DigitProd;
        //    int SumaTotal;
        //    //byte i;

        //    for (int i = 1; i < 10; i++)
        //    {
        //        DigitProd(i, 1) = NSS.Substring(i, 1);
        //    }

        //    DigitoVerificadorReal = Convert.ToInt32(NSS.Substring(11, 1));

        //    //Empiezo las multiplicaciones
        //    for (int i = 1; i < 10; i++)
        //    {
        //        if ((i % 2) == 0)
        //        {
        //            DigitProd(i, 2) = DigitProd(i, 1) * 2;
        //        }
        //        else
        //        {
        //            DigitProd(i, 2) = DigitProd(i, 1);
        //        }
        //    }

        //    //Veo si el producto tiene dos dígitos
        //    for (int i = 1; i < 10; i++)
        //    {
        //        if ( Convert.ToString(DigitProd(i, 2)).Trim().Length() == 2 )
        //        { 
        //            DigitProd(i, 2) = Convert.ToString(DigitProd(i, 2)).Substring(1, 1) + Convert.ToString(DigitProd(i, 2)).Substring(2, 1);
        //        }
        //    }

        //    SumaTotal = 0;
        //    for (int i = 1; i < 10; i++)
        //    {
        //        SumaTotal = SumaTotal + DigitProd(i, 2);
        //    }

        //    if (SumaTotal > 0 && SumaTotal< 10)
        //        DigitoVerificador = 10 - SumaTotal;

        //    if (SumaTotal > 10 && SumaTotal < 20)
        //        DigitoVerificador = 20 - SumaTotal;

        //    if (SumaTotal > 20 && SumaTotal < 30)
        //        DigitoVerificador = 30 - SumaTotal;

        //    if (SumaTotal > 30 && SumaTotal < 40)
        //        DigitoVerificador = 40 - SumaTotal;

        //    if (SumaTotal > 40 && SumaTotal < 50)
        //        DigitoVerificador = 50 - SumaTotal;

        //    if (SumaTotal > 50 && SumaTotal < 60)
        //        DigitoVerificador = 60 - SumaTotal;

        //    if (SumaTotal > 60 && SumaTotal < 70)
        //        DigitoVerificador = 70 - SumaTotal;

        //    if (SumaTotal > 70 && SumaTotal < 80)
        //        DigitoVerificador = 80 - SumaTotal;

        //    if (SumaTotal > 80 && SumaTotal < 90)
        //        DigitoVerificador = 90 - SumaTotal;

        //    if (SumaTotal > 90 && SumaTotal < 100)
        //        DigitoVerificador = 100 - SumaTotal;

        //    if (DigitoVerificador != DigitoVerificadorReal)
        //        return false;

        //    return true;
        //}



        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //public Function ScanItem2(ByVal Item As String, CboSrc As ComboBox) As Integer
        //    Dim i As Integer

        //    ScanItem2 = -1
        //    For i = 0 To CboSrc.ListCount - 1
        //        if CboSrc.List(i) = Item Then
        //            ScanItem2 = i
        //            Exit For
        //        End if
        //    Next i

        //End Function


        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //public Function ScanItem(ByVal Item As Long, CboSrc As ComboBox) As Integer
        //    Dim i As Integer

        //    ScanItem = -1
        //    For i = 0 To CboSrc.ListCount - 1
        //        if CboSrc.ItemData(i) = Item Then
        //            ScanItem = i
        //            Exit For
        //        End if
        //    Next i

        //End Function

        //Valida que la entrada en una se un caracter numerico o el comodin asterico
        //31/10/2000 G@cost@
        public static int iValida09(int ikeyascii)
        {
            string sChar;
            //Permitiendo unicamente la captura de mayusculas
            if ((int.TryParse(ikeyascii.ToString(), out _)) && ikeyascii != 8 && ikeyascii != 42)
                ikeyascii = 0;
            return ikeyascii;
        }



        // Esta funcion ya se encuentra en otra parte
        // 19 mayo 2023
        // RGB
        //    Sub CopiaDatosGrid(TDBGrid1 As TDBGrid, rsdatos As ADODB.Recordset)
        //        Dim iCol As Integer
        //        Dim iCol2 As Integer
        //        Dim iRow As Integer
        //        Dim i As Integer
        //        Dim sRow As String
        //        Dim sTemp As String

        //        On Error GoTo GridCopyError

        //        Screen.MousePointer = vbHourglass

        //        sTemp = ""
        //        With TDBGrid1
        //            //Si existen columnas seleccionadas, copiamos todos los renglones de esas columnas
        //            if .SelStartCol != -1 Then
        //                For iRow = 1 To rsdatos.RecordCount
        //                    sRow = ""
        //                    For iCol = .SelStartCol To .SelEndCol
        //                        //Obtenemos el indice real de la columna, esto se da si la columna ha sido movida
        //                        For i = 0 To .Columns.Count
        //                            if iCol = .Columns(i).Order Then
        //                                iCol2 = i
        //                                Exit For
        //                            End if
        //                        Next i
        //                        if .Columns(iCol2).Visible Then
        //                            if sRow != "" Then
        //                                sRow = sRow & vbTab
        //                            End if
        //                            if IsNull(.Columns(iCol2).CellValue(iRow)) Then
        //                                sRow = sRow & " "
        //                            else
        //                                sRow = sRow & .Columns(iCol2).CellValue(iRow)
        //                            End if
        //                        End if
        //                    Next
        //                    if sTemp != "" Then sTemp = sTemp & vbCrLf
        //                    sTemp = sTemp & sRow
        //                Next
        //            else
        //                //Si no copiamos todo el grid
        //                For iRow = 1 To rsdatos.RecordCount
        //                    if .IsSelected(iRow) = -1 Then
        //                        sRow = ""
        //                        For iCol = 0 To .Columns.Count - 1
        //                            For i = 0 To .Columns.Count
        //                                if iCol = .Columns(i).Order Then
        //                                    iCol2 = i
        //                                    Exit For
        //                                End if
        //                            Next i
        //                            if .Columns(iCol2).Visible Then
        //                                if sRow != "" Then sRow = sRow & vbTab
        //                                if IsNull(.Columns(iCol2).CellValue(iRow)) Then
        //                                    sRow = sRow & " "
        //                                else
        //                                    sRow = sRow & .Columns(iCol2).CellValue(iRow)
        //                                End if
        //                            End if
        //                        Next
        //                        if sTemp != "" Then sTemp = sTemp & vbCrLf
        //                        sTemp = sTemp & sRow
        //                    End if
        //                Next
        //            End if
        //        End With

        //        Clipboard.Clear()
        //        Clipboard.SetText sTemp

        //Screen.MousePointer = vbDefault
        //        Exit Sub

        //GridCopyError:
        //        Screen.MousePointer = vbDefault
        //        gfMsgbox.Mensaje "Error  al copiar datos del grid " & vbLf & Error$, vbOKOnly

        //End Sub



        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //    public Sub ExpGridXLS(grdCont As TDBGrid)
        //        Dim sarchivo As String
        //        Dim gobjComisionEj As Object


        //        On Error GoTo msgerror

        //        sarchivo = ObtenNombreArchivo("xls", "Exportar grid de " & grdCont.Caption & " a Excel")
        //        if sarchivo != Empty Then
        //            grdCont.ExportToFile sarchivo, False
        //        gfMsgbox.Mensaje "Archivo Generado : " & sarchivo
        //    End if
        //        Exit Sub
        //msgerror:
        //        Screen.MousePointer = vbDefault
        //        SetStatusBar("Listo")
        //        gsErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        gfMsgbox.Mensaje "Error al exportar el grid a Excel.", vbCritical + vbOKDetails, grdCont.Caption, grsErrADO, gsErrVB
        //End Sub



        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //public Sub ImprimeTDBGrid(grdPrint As TDBGrid, bPreview As Boolean, sTitulo As String, iTop As Double, iBottom As Double, iLeft As Double, iRight As Double)
        //    grdPrint.PrintInfo.SettingsMarginBottom = iBottom * grdPrint.PrintInfo.UnitsPerInch
        //    grdPrint.PrintInfo.SettingsMarginLeft = iLeft * grdPrint.PrintInfo.UnitsPerInch
        //    grdPrint.PrintInfo.SettingsMarginRight = iRight * grdPrint.PrintInfo.UnitsPerInch
        //    grdPrint.PrintInfo.SettingsMarginTop = iTop * grdPrint.PrintInfo.UnitsPerInch
        //    grdPrint.PrintInfo.RepeatColumnHeaders = True
        //    grdPrint.PrintInfo.Draft = True

        //    if grdPrint.Width < grdPrint.PrintInfo.UnitsPerInch * 8 Then
        //        grdPrint.PrintInfo.SettingsOrientation = 1
        //    else
        //        grdPrint.PrintInfo.SettingsOrientation = 2
        //    End if

        //    if bPreview = True Then
        //        grdPrint.PrintInfo.PrintPreview
        //    else
        //        grdPrint.PrintInfo.PrintData
        //    End if
        //End Sub

        //Sub ImprimeGrid(grdCopia As TDBGrid, Titulo As String)

        //    if Not (grdCopia.DataSource Is Nothing) Then

        //        With grdCopia.PrintInfo
        //            // Set the page header
        //            .PageHeaderFont.Italic = True
        //            .PageHeader = Titulo

        //            // Column headers will be on every page
        //            .RepeatColumnHeaders = True

        //            //Colocamos los margenes
        //            .SettingsMarginLeft = 700
        //            .SettingsMarginRight = 700
        //            .SettingsMarginTop = 700
        //            .SettingsMarginBottom = 700

        //            // Display page numbers (centered)
        //            .PageFooter = "\tPage: \p"
        //            // Invoke Print Preview
        //            //.PrintPreview
        //            .PageSetup
        //            .PrintData
        //        End With

        //    End if


        //End Sub


        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        // ////Sub LlenaComboEmpresas(sUsuario As String, cboDatos As ComboBox)
        // Sub LlenaComboEmpresas(cboEmpresa As ComboBox)

        //     //Agregamos los valores del rs
        //     grsEmpresas.MoveFirst()

        //     While Not grsEmpresas.EOF
        //         if Not IsNull(grsEmpresas.Fields.Item(1).Value) Then
        //             cboEmpresa.AddItem grsEmpresas.Fields.Item(1).Value
        //     cboEmpresa.ItemData(cboEmpresa.NewIndex) = grsEmpresas.Fields.Item(0).Value
        //         End if
        //         grsEmpresas.MoveNext()
        //Wend

        ////Agregamos el valor fijo de todas las empresas
        //cboEmpresa.AddItem "Grupo BBVA-Bancomer"
        //cboEmpresa.ItemData(cboEmpresa.NewIndex) = 99


        // End Sub



        //Verifica si existe el folder si no lo crea
        public static void VerificaFolder(string folder)
        {
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
        }

        //SDC 20/02/2004
        //Regresa el directorio de sistema operativo
        public static string GetSysDir()
        {
            string Temp = new string(' ', 256);
            long  X;

            X = GetSystemDirectoryA(Temp, Temp.Length); // Make API Call (Temp will hold return value)
            return Temp.Substring(0, Convert.ToInt32(X));              // Trim Buffer and return string
        }

        //SDC 20/02/2004
        //Regresa el directorio de Windows
        private static string GetWinDir() 
        {
            string Temp = new string('*', 255);
            long X;
            X = GetWindowsDirectoryA(Temp, Convert.ToInt64(Temp.Length));    // Make API Call (Temp will hold return value)
            return  Temp.Substring(0, Convert.ToInt32(X));  //Left$(Temp, X)               // Trim Buffer and return string
        }

        //SDC 08/06/2004
        //Devuelve la fecha del último día del mes y año de Fecha
        public static DateTime tUltimoDiaDeMes(ref DateTime Fecha)
        {
            int year = Fecha.Year;
            int month = Fecha.Month;

            // Obtener el último día del mes
            int lastDayOfMonth = DateTime.DaysInMonth(year, month);
            DateTime lastDateOfMonth = new DateTime(year, month, lastDayOfMonth);

            return lastDateOfMonth;
        }



        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //SDC 08/06/2004
        //Escribe un texto en la barra de status
        //public Sub SetStatusBar(ByRef sText As String)
        //    frmRentas.abdRentas.Bands("stbPrincipal").Tools("txtStatus").Caption = sText
        //    frmRentas.abdRentas.Refresh
        //    DoEvents
        //End Sub



        //SDC 2004-12-06
        //Genera los dos archivos de entrada para el SUC
        public Sub GeneraArchivosSUC(ByRef polizas As String, ByRef tFP As Date)
            Dim rsRecord As ADODB.Recordset
            Dim vParametros(4) As Object
            Dim iRes As Integer
            Dim objPolizas As Object
            On Error GoTo msgerror

            //SetStatusBar("Generando archivos de entrada al SUC")
            //Screen.MousePointer = vbHourglass

            Call VerificaFolder("C:\SUC2000\DATOS\")
            objPolizas = CreateObject("MTSEndososCET.clsEndososCET")

            vParametros(0) = 0
            vParametros(1) = polizas
            vParametros(2) = Strings.Format(tFP, "yyyy-mm-dd")
            vParametros(3) = giIDEmpresa
            rsRecord = Nothing
            iRes = objPolizas.iDatosSUC(grsErrADO, gsConexion, rsRecord, vParametros)
            if iRes = DatosOK Then
                rsRecord.MoveFirst()
                FileOpen(1, "C:\SUC2000\DATOS\asegs.txt", OpenMode.Output)
                Do While Not rsRecord.EOF
                    Print(1, rsRecord.Fields("Texto").Value)
                    rsRecord.MoveNext()
                Loop
                FileClose(1)
                vParametros(0) = 1
                rsRecord = Nothing
                iRes = objPolizas.iDatosSUC(grsErrADO, gsConexion, rsRecord, vParametros)
                if iRes = DatosOK Then
                    rsRecord.MoveFirst()
                    FileOpen(2, "C:\SUC2000\DATOS\benefs.txt", OpenMode.Output)
                    Do While Not rsRecord.EOF
                        Print(2, rsRecord.Fields("Texto").Value)
                        rsRecord.MoveNext()
                    Loop
                    FileClose(2)
                    gfMsgbox.Mensaje("Archivos asegs.txt y benefs.txt generados en C:\SUC2000\DATOS", vbInformation, "CAIRO")
                elseif iRes = NoHayDatos Then
                    gfMsgbox.Mensaje("Archivo asegs.txt generado en C:\SUC2000\DATOS", vbInformation, "CAIRO")
                else
                    gfMsgbox.Mensaje("Error del Sistema Cairo", frmMensaje.ETipos.vbOKDetails + vbCritical, "C A I R O", grsErrADO)
                End if
            elseif iRes = NoHayDatos Then
                gfMsgbox.Mensaje("No existen datos.", vbInformation + vbOKOnly, "C A I R O")
            else
                gfMsgbox.Mensaje("Error del Sistema Cairo", frmMensaje.ETipos.vbOKDetails + vbCritical, "C A I R O", grsErrADO)
            End if

            //SetStatusBar("Listo")
            //Screen.MousePointer = vbDefault
            objPolizas = Nothing
            Exit Sub
    msgerror:
            objPolizas = Nothing
            //Screen.MousePointer = vbDefault
            gsErrVB = Strings.Format$((Err.Number) & vbTab & Err.Source & vbTab & Err.Description)
            gfMsgbox.Mensaje(" Error del Sistema Cairo ", frmMensaje.ETipos.vbOKDetails + vbCritical, "C A I R O", , gsErrVB)
        End Sub


        //_______________________________________________________________
        //Título: Verificacion de radiobutton
        //Subrutina: iRadioOption
        //Versión: 1.0
        //Fecha:    30/03/2005
        //Autor:    Samuel Dueñas
        //Modificación:
        //Fecha de Modificacion:
        //_______________________________________________________________
        //Descripción:
        //       Regresa el Index del radiobutton seleccionado.
        //_______________________________________________________________


        public Function iRadioOption(ByRef RadioOption As Object, max_count As Integer) As Integer
            Dim iCont As Integer
            For iCont = 0 To max_count
                On Error Resume Next
                if RadioOption(iCont).Value = True Then
                    if Err.Number!= 340 Then
                        iRadioOption = iCont
                    End if
                End if
            Next
        End Function

        //_______________________________________________________________
        //Título: Obtener nombre de archivo para escritura
        //Subrutina: ObtenNombreArchivo
        //Versión: 1.0
        //Fecha:    07/04/2005
        //Autor:    Samuel Dueñas
        //Modificación:
        //Fecha de Modificacion:
        //_______________________________________________________________
        //Descripción:
        //       Regresa el Nombre de un archivo escrito por el usuario.
        //_______________________________________________________________
        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //    public Function ObtenNombreArchivo(Extension As String, Titulo As String, Optional InitDir As String, Optional OverwritePrompt As Boolean = True, Optional Name As String) As String
        //        On Error GoTo fin

        //        ObtenNombreArchivo = Empty
        //        Extension = Replace(Extension, ".", "")
        //        With frmRentas.Dialog
        //            .Filter = "Archivos de Tipo " & Extension & "|*." & Extension & "|Todos los Archivos|*.*"
        //            .CancelError = True
        //            .DialogTitle = Titulo
        //            .DefaultExt = Extension
        //            .InitDir = InitDir
        //            .Flags = cdlOFNPathMustExist
        //            //I Héctor García  18ene2012 Optimizar Proceso de Comisiones
        //            .FileName = Name
        //            //F Héctor García  18ene2012 Optimizar Proceso de Comisiones
        //            if OverwritePrompt Then
        //                .Flags = cdlOFNOverwritePrompt
        //                .ShowSave
        //            else
        //                .ShowOpen
        //            End if

        //            if .FileName = "" Then
        //                ObtenNombreArchivo = Empty
        //            else
        //                ObtenNombreArchivo = .FileName
        //            End if
        //        End With
        //fin:
        //    End Function

        //_______________________________________________________________
        //Título: Generar archivo separado por comas de un recordset
        //Subrutina: GeneraReporte
        //Versión: 1.0
        //Fecha:    02/08/2005
        //Autor:    Samuel Dueñas
        //Modificación:
        //Fecha de Modificacion:
        //_______________________________________________________________
        //Descripción:
        //       Crea un archivo separado por comas para ser abierto en excel.
        //   Se puede usar también para generar TXT del estilo del IMSS.
        //_______________________________________________________________

        public Sub GeneraCSV(ByRef rsData As Recordset, ByRef sNombreArchivo As String, bTitulos As Boolean, Optional bMensaje As Boolean = True)
            Dim nCanal As Integer
            Dim sCadena As String
            Dim X As Integer

            nCanal = FreeFile()
            On Error Resume Next
            FileOpen(1, sNombreArchivo, OpenMode.Output)
            if Err().Number!= 0 Then
                gfMsgbox.Mensaje("Imposible abrir el Archivo para escritura.", vbInformation, gsModulo)
                Exit Sub
            End if

            rsData.MoveFirst()

            if Not rsData.EOF Then
                sCadena = ""

                if bTitulos Then
                    For X = 0 To rsData.Fields.Count - 1
                        sCadena = sCadena & Chr(34) & rsData(X).Name & Chr(34) & ","
                    Next
                    Print(1, sCadena)
                End if

                While Not rsData.EOF
                    sCadena = ""
                    For X = 0 To rsData.Fields.Count - 1
                        sCadena = sCadena + rsData(X).Value + Iif(rsData.Fields.Count = 1, "", ",")
                    Next
                    Print(1, sCadena)
                    rsData.MoveNext()
                End While
                if bMensaje Then
                    gfMsgbox.Mensaje("Archivo generado exitosamente en " & sNombreArchivo, vbInformation, gsModulo)
                End if
            End if

            FileClose(1)
            On Error GoTo 0

        End Sub

        public Function dMaximo(ByRef dUno As Double, ByRef dDos As Double) As Double
            if dUno > dDos Then
                dMaximo = dUno
            else
                dMaximo = dDos
            End if
        End Function

        public Function iMaximo(ByRef iUno As Integer, ByRef iDos As Integer) As Integer
            if iUno > iDos Then
                iMaximo = iUno
            else
                iMaximo = iDos
            End if
        End Function

        //_______________________________________________________________
        //Título: Generar archivo de Excel a partir de un recordset
        //Subrutina: GeneraXLS
        //Versión: 1.0
        //Fecha:    2007-06-15
        //Autor:    Samuel Dueñas
        //Modificación:
        //Fecha de Modificacion:
        //_______________________________________________________________
        //Descripción:
        //   Crea un archivo de Excel con los mismos parametros de entrada
        //   que la funcion GeneraCSV. Pueden ser intercambiables, sin embargo,
        //   esta función es muy lenta.
        //_______________________________________________________________
        public Sub GeneraXLS(ByRef rsData As Recordset, ByRef sNombreArchivo As String, bTitulos As Boolean, Optional bMensaje As Boolean = True)
            Dim objXcel As Excel.Worksheet
            Dim AppExcel As Excel.Application
            Dim lRow As Long
            Dim iCol As Integer
            Dim lTotal As Long
            On Error GoTo msgerror

            //Screen.MousePointer = vbHourglass

            lRow = 1
            iCol = 1

            rsData.MoveFirst()

            if Not rsData.EOF Then
                lTotal = rsData.RecordCount

                if lTotal > 65535 Then
                    gfMsgbox.Mensaje("No se puede guardar a Excel, más de 65535 registros.", vbExclamation, "CAIRO")
                    Exit Sub
                End if

                AppExcel = New Excel.Application
                AppExcel.Workbooks.Add()
                objXcel = AppExcel.Worksheets(1)

                if bTitulos Then
                    For iCol = 1 To rsData.Fields.Count
                        objXcel.Cells(lRow, iCol).Font.Bold = True
                        objXcel.Cells(lRow, iCol) = rsData(iCol - 1).Name
                    Next
                    lRow = lRow + 1
                End if

                While Not rsData.EOF
                    For iCol = 1 To rsData.Fields.Count
                        //Se agrega el if para evaluar si la celda es de tipo fecha Reporte Prescripcion Version 2 EBS 01/JULIO/2016
                        if IsDate(rsData(iCol - 1)) And InStr(1, sNombreArchivo, "PrescPolGpo", vbTextCompare) != 0 Then
                            objXcel.Cells(lRow, iCol) = FormatDateTime(rsData(iCol - 1).Value, vbShortDate)
                        else
                            //-- Inicio Reporte Saldo Vivienda EBS 08/MARZO/2017 -------------------------------------------------------------------
                            if rsData(iCol - 1).Name!= "Poliza" And IsNumeric(rsData(iCol - 1)) And InStr(1, sNombreArchivo, "SaldoVivienda", vbTextCompare) != 0 Then
                                objXcel.Cells(lRow, iCol) = FormatCurrency(rsData(iCol - 1), 2, vbFalse, vbFalse, vbTrue)
                            else
                                //-- Fin Reporte Saldo Vivienda EBS 08/MARZO/2017 -------------------------------------------------------------------
                                objXcel.Cells(lRow, iCol) = rsData(iCol - 1) //<-- esta linea ya existia EBS 01/JULIO/2016
                            End if
                        End if
                    Next
                    rsData.MoveNext()
                    //SetStatusBar("Exportando a Excel registro " & lRow - 1 & " de " & lTotal & " ...")
                    lRow = lRow + 1
                End While

                AppExcel.Workbooks(1).SaveAs(sNombreArchivo)
                AppExcel.Workbooks.Close()
                AppExcel.DisplayAlerts = True
                AppExcel.Quit()
                AppExcel = Nothing

                if bMensaje Then
                    gfMsgbox.Mensaje("Archivo generado exitosamente en " & sNombreArchivo, vbInformation, "CAIRO")
                End if
            End if

            //Screen.MousePointer = vbDefault
            Exit Sub
    msgerror:
            //Screen.MousePointer = vbDefault
            gsErrVB = Strings.Format$((Err.Number) & vbTab & Err.Source & vbTab & Err.Description)
            gfMsgbox.Mensaje("Error del Sistema Cairo", frmMensaje.ETipos.vbOKDetails + vbCritical, "CAIRO", , gsErrVB)
        End Sub


        // ------------------------------------------------------------------
        // Título:       Rutinas Generales de Prestamos ISSSTE
        // Modulo:       modGeneral.bas
        // Versión:      1.0
        // Fecha:        19/Nov/2009
        // Autor:        Silvia Gabriela Rodriguez Ruiz
        // Modificación:
        //
        // ------------------------------------------------------------------
        // Descripción:
        // Este modulo provee rutinas más utilizadas en forma general,
        // como son formatos de campos, obtención de datos, validaciones de
        // datos.


        public Function iUbicaElementoCombo(objCmb As Object, sClave As String) As Integer
            // Función:  iUbicaElementoCombo
            //
            // Descripción:
            //    Esta función obtiene la posición de un dato especifico, en un combo.
            //
            // Uso:
            //    cboLista.ListIndex = iUbicaElementoCombo(cboLista, "Elemento 5")
            //
            // Fecha: 19/Nov/2009
            // Autor: Silvia Gabriela Rodriguez Ruiz

            Dim iIndex As Integer
            Dim iPosicion As Integer
            Dim sError As String

            On Error GoTo iUbicaElementoCombo_Err

            iPosicion = -1
            iIndex = 0
            While iIndex<objCmb.ListCount And iPosicion = -1
                objCmb.ListIndex = iIndex
                if Mid(objCmb.Text, 1, Len(sClave)) = sClave Then
                    iPosicion = iIndex
                End if
                iIndex = iIndex + 1
            End While

    iUbicaElementoCombo_Sal:
            iUbicaElementoCombo = iPosicion
            Exit Function

    iUbicaElementoCombo_Err:
            iPosicion = -1
            sError = Err.Number & " " & Err.Description
            Resume iUbicaElementoCombo_Sal

        End Function

        public Function OnlyNumbers(intAsc As Integer) As Byte
            // Función:  OnlyNumbers
            //
            // Descripción:
            //    Esta función solo valida que el caracter presionado
            //    sea numerico.
            //
            // Uso:
            //    KeyAscii = OnlyNumbers(Asc(UCase(Chr(KeyAscii))))
            //
            // Fecha: 19/Nov/2009
            // Autor: Silvia Gabriela Rodriguez Ruiz

            Dim bytAscci As Byte
            On Error GoTo OnlyNumbers_Err

            if(intAsc >= 48 And intAsc <= 57) Then
                bytAscci = intAsc
            else
                if intAsc = 8 Then
                    bytAscci = intAsc
                else
                    bytAscci = 0
                End if
            End if

    OnlyNumbers_Sal:
            OnlyNumbers = bytAscci
            Exit Function

    OnlyNumbers_Err:
            bytAscci = 0
            Resume OnlyNumbers_Sal

        End Function

        public Function CaracterInvalido(intAsc As Integer) As Byte
            // Función:  OnlyNumbers
            //
            // Descripción:
            //    Funcion que permite limitar los caracteres que pueden ser tecleados en un text box, esta funcion
            //    no permite la captura de caracteres como ", `, // # , no se permite que se capturen estos caracteres
            //    ya que confunden la información con el uso de querys.
            //
            // Uso:
            //    KeyAscii = OnlyNumbers(Asc(UCase(Chr(KeyAscii))))
            //
            // Fecha: 19/Nov/2009
            // Autor: Silvia Gabriela Rodriguez Ruiz

            // Objetivo:
            Dim bytAscci As Byte
            On Error GoTo CaracterInvalido_Err

            if(intAsc >= 32 And intAsc <= 125) Or
           (intAsc = 8) Or
           (intAsc = 3) Or
           (intAsc = 22) Or
           (intAsc = Asc(vbCrLf)) Or
            (intAsc = Asc(vbCr)) Then
                if intAsc = 34 Or intAsc = 39 Or intAsc = 96 Or intAsc = 44 Then
                    bytAscci = 0
                else
                    bytAscci = intAsc
                End if
            else
                if intAsc = 193 Or intAsc = 201 Or intAsc = 205 Or
               intAsc = 211 Or intAsc = 218 Or intAsc = 225 Or
               intAsc = 233 Or intAsc = 237 Or intAsc = 243 Or
               intAsc = 250 Then
                    bytAscci = intAsc
                End if
            End if

    CaracterInvalido_Sal:
            CaracterInvalido = bytAscci
            Exit Function

    CaracterInvalido_Err:
            bytAscci = 0
            Resume CaracterInvalido_Sal
        End Function

        //02 de marzo 2011 RCS
        //se encarga de eliminar los campos repetidos en un rs

        public Function quitaCamposRepetidos(ByVal objRecordset As Recordset, Optional ByVal LockType As LockTypeenum = adLockBatchOptimistic, Optional ByVal CursorType As CursorTypeenum = adOpenDynamic) As Recordset
            On Error GoTo msgerror
            Dim objNewRS As ADODB.Recordset
            Dim objField As Object
            Dim objField1 As Object

            Dim lngCnt As Long
            Dim errVB As String

            objNewRS = New ADODB.Recordset
            objNewRS.CursorLocation = adUseClient
            objNewRS.LockType = LockType
            objNewRS.CursorType = CursorType

            Dim i, conta, Contador As Integer
            conta = 1

            Dim nameStr As String
            Dim typeSTR As String
            Dim sizeStr As String
            Dim attrib As String

            Dim camposNoRepetidos() As String
            Dim typeMATRIZ() As String
            Dim sizeMATRIZ() As String
            Dim attribMATRIZ() As String

            i = 0
            For Each objField In objRecordset.Fields

                Contador = 1
                For Each objField1 In objRecordset.Fields
                    if objField1.Name = objField.Name Then
                        if Contador = 1 Then
                            if InStr(nameStr, objField.Name) < 1 Then
                                nameStr = nameStr & "," & objField.Name
                                sizeStr = sizeStr & "," & objField.DefinedSize
                                typeSTR = typeSTR & "," & objField.Type
                                attrib = attrib & "," & objField.Attributes
                            End if
                        End if
                        Contador = Contador + 1
                    End if
                Next objField1
            Next objField

            camposNoRepetidos = Split(nameStr, ",")
            typeMATRIZ = Split(typeSTR, ",")
            sizeMATRIZ = Split(sizeStr, ",")
            attribMATRIZ = Split(attrib, ",")

            For i = 1 To UBound(camposNoRepetidos)
                objNewRS.Fields.Append(camposNoRepetidos(i), typeMATRIZ(i), sizeMATRIZ(i), attribMATRIZ(i))
            Next i

            objNewRS.Open()

            if Not objRecordset.RecordCount = 0 Then
                objRecordset.MoveFirst()

                While Not objRecordset.EOF
                    objNewRS.AddNew()

                    For lngCnt = 0 To UBound(camposNoRepetidos) - 1
                        objNewRS.Fields(lngCnt).Value = objRecordset.Fields(camposNoRepetidos(lngCnt + 1)).Value
                    Next lngCnt
                    objRecordset.MoveNext()
                End While
            End if
            quitaCamposRepetidos = objNewRS

            Exit Function

    msgerror:
            //Screen.MousePointer = vbDefault
            errVB = Strings.Format$((Err.Number) & vbTab & Err.Source & vbTab & Err.Description)
            gfMsgbox.Mensaje(" Error del Sistema Prospect ", frmMensaje.ETipos.vbOKDetails + vbCritical, "P R O S P E C T", , errVB)
        End Function
        public Function FindeMes(sMonth As String, sYear As String) As String

            Select Case sMonth
                Case Is = "01", "03", "05", "07", "08", "10", "12"
                    FindeMes = "31"
                Case Is = "04", "06", "09", "11"
                    FindeMes = "30"
                Case Is = "02"
                    if Int(sYear / 4) = sYear / 4 Then
                        FindeMes = "29"
                    else
                        FindeMes = "28"
                    End if
            End Select

        End Function


        //Inicio. Alexander Hdez 12/06/2012 Codigos Postales YA9A0F
        //----------------------------------------------------------------------------------------//
        //------------------   Funcion de validacion de entrada de datos, acepta lo que le llegue en ValidString   ----------------------//
        //----------------------------------------------------------------------------------------//
        Function Valida_TextoMinMay(KeyIn As Integer, ValidString As String, Editable As Boolean) As Integer
            Dim ValidList As String, KeyOut As Integer

            ValidList = Iif(Editable, ValidString & Chr(8), ValidString)

            if InStr(1, ValidList, Chr(KeyIn), 1) > 0 Then
                KeyOut = Iif(Editable, KeyIn, 0)
            else
                KeyOut = Iif(Editable, 0, Chr(KeyIn))
            End if

            Valida_TextoMinMay = KeyOut
        End Function


        Function Valida_Texto(KeyIn As Integer, ValidString As String, Editable As Boolean) As Integer
            Dim ValidList As String, KeyOut As Integer

            ValidList = Iif(Editable, UCase(ValidString) & Chr(8), UCase(ValidString))

            if InStr(1, ValidList, UCase(Chr(KeyIn)), 1) > 0 Then
                KeyOut = Iif(Editable, Asc(UCase(Chr(KeyIn))), 0)
            else
                KeyOut = Iif(Editable, 0, Asc(UCase(Chr(KeyIn))))
            End if

            Valida_Texto = KeyOut
        End Function
        //Fin. Alexander Hdez 12/06/2012

        //Inicio. Alexander Hdez 12/06/2012 Codigos Postales YA9A0F
        //----------------------------------------------------------------------------------------//
        //------------------   Funcion para Proceso de Codigos Postales  ----------------------//
        //----------------------------------------------------------------------------------------//

        public Function ExcelsePOMEX(rutaSEPOMEX As String)

            Dim errVB As String
            Dim Estados(33) As String
            Dim TamMatriz As Integer
            Dim i As Integer
            Dim ArchivoE As String

            Dim cnn As New ADODB.Connection
            Dim cnConn As New ADODB.Connection
            Dim lNumRegAfect As Long
            Dim strSQL As String
            Dim rs As New ADODB.Recordset
            Dim j As Integer
            Dim iCountCargaCP As Long

            On Error GoTo msgerror

            Dim ObjExcel As New Excel.Application
            Dim ObjW As Object
            ObjW = ObjExcel.Workbooks.Open(rutaSEPOMEX)


            TamMatriz = ObjW.Sheets.Count

            For i = 2 To ObjW.Sheets.Count
                Estados(i) = ObjW.Sheets(i).Name
            Next

            ObjW.Close
            ObjW = Nothing

            cnConn.CommandTimeout = 0
            cnConn.Open(gsConexion) //Conexion para Chairo

            TamMatriz = UBound(Estados)

            With cnn
                .CommandTimeout = 0
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = "Data Source=" & rutaSEPOMEX & " "
                //.Properties("Extended Properties") = "Excel 8.0" //This is a read only properti
                .Open()
            End With

            strSQL = ""
            strSQL = strSQL & " Delete from Cat_Sepomex  "
            cnConn.Execute(strSQL)

            //frmCatCodigosPostales.pgbCargaCP.Visible = False

            For j = 2 To TamMatriz //- 1
                //frmCatCodigosPostales.pgbCargaCP.Value = 0
                //iCountCargaCP = 1

                strSQL = ""
                strSQL = strSQL & "if  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N//[dbo].[TmpEstadoSEPOMEX]//) AND type in (N//U//)) "
                strSQL = strSQL & "DROP TABLE [dbo].[TmpEstadoSEPOMEX] "
                cnConn.Execute(strSQL)


                strSQL = ""
                strSQL = strSQL & "SELECT * INTO [ODBC;" & gsConexion & "].TmpEstadoSEPOMEX "
                strSQL = strSQL & "FROM [" & Estados(j) & "] "
                cnn.Execute(strSQL)


                //    strSQL = ""
                //    strSQL = "Select * from TmpEstadoSEPOMEX"
                //
                //    With rs
                //        .ActiveConnection = cnConn
                //        .CursorType = adOpenDynamic
                //        .CursorLocation = adUseClient
                //        .Open strSQL
                //    End With
                //
                //    frmCatCodigosPostales.pgbCargaCP.Max = rs.RecordCount
                //
                //    if Not rs.EOF Then
                //        rs.MoveFirst
                //        Do While Not rs.EOF
                //            strSQL = ""
                //            strSQL = strSQL & "Insert Into Cat_Sepomex(d_codigo,d_asenta,d_tipo_asenta,D_mnpio,d_estado,d_ciudad,d_CP,c_estado,c_oficina,c_CP,c_tipo_asenta,c_mnpio,id_asenta_cpcons,d_zona,c_cve_ciudad) "
                //            strSQL = strSQL & "Values(//" & rs.Fields("d_codigo") & "//,//" & rs.Fields("d_asenta") & "//,//" & rs.Fields("d_tipo_asenta") & "//,//" & rs.Fields("D_mnpio") & "//,//" & rs.Fields("d_estado") & "//,//" & rs.Fields("d_ciudad") & "//,//" & rs.Fields("d_CP") & "//,//" & rs.Fields("c_estado") & "//,//" & rs.Fields("c_oficina") & "//,//" & rs.Fields("c_CP") & "//,//" & rs.Fields("c_tipo_asenta") & "//,//" & rs.Fields("c_mnpio") & "//,//" & rs.Fields("id_asenta_cpcons") & "//,//" & rs.Fields("d_zona") & "//,//" & rs.Fields("c_cve_ciudad") & "//)"
                //            cnConn.Execute (strSQL)
                //            rs.MoveNext
                //
                //            frmCatCodigosPostales.pgbCargaCP.Value = iCountCargaCP
                //            iCountCargaCP = iCountCargaCP + 1
                //
                //        Loop
                //        rs.Close
                //    End if
                strSQL = ""
                strSQL = strSQL & "Insert Into Cat_Sepomex(d_codigo,d_asenta,d_tipo_asenta,D_mnpio,d_estado,d_ciudad,d_CP,c_estado,c_oficina,c_CP,c_tipo_asenta,c_mnpio,id_asenta_cpcons,d_zona,c_cve_ciudad) "
                strSQL = strSQL & "Select * from TmpEstadoSEPOMEX "
                cnConn.Execute(strSQL)

            Next



            strSQL = ""
            strSQL = strSQL & " exec sp_ActualizaAcentosSEPOMEX "
            cnConn.Execute(strSQL)

            strSQL = ""
            strSQL = strSQL & " exec sp_InsertColSEPOMEXvsCAIRO "
            cnConn.Execute(strSQL)


            rs = Nothing
            cnConn.Close()
            cnConn = Nothing


            Exit Function
    msgerror:
            errVB = Strings.Format((Err.Number) & vbTab & Err.Source & vbTab & Err.Description)
            //Set rsErrAdo = ErroresDLL(Nothing, errVB)

        End Function
        //Fin. Alexander Hdez 12/06/2012




        //Inicio. Alexander Hdez 22/10/2012 Codigos Postales YA9A0F
        //----------------------------------------------------------------------------------------//
        //------------------   Funcion para Proceso de Codigos Postales, valida Email  ----------------------//
        //----------------------------------------------------------------------------------------//

        public Function Validar_Email(ByVal Email As String) As Boolean

            Dim i As Integer, iLen As Integer, caracter As String
            Dim pos As Integer, bp As Boolean, ipos As Integer, iPos2 As Integer

            On Error GoTo Err_Sub

            Email = Trim$(Email)

            if Email = vbNullString Then
                Exit Function
            End if

            Email = LCase$(Email)
            iLen = Len(Email)


            For i = 1 To iLen
                caracter = Mid(Email, i, 1)

                if(Not (caracter Like "[a-z]")) And(Not (caracter Like "[0-9]")) Then

                    if InStr(1, "_-" & "." & "@", caracter) > 0 Then
                        if bp = True Then
                            Exit Function
                        else
                            bp = True

                            if i = 1 Or i = iLen Then
                                Exit Function
                            End if

                            if caracter = "@" Then
                                if ipos = 0 Then
                                    ipos = i
                                else

                                    Exit Function
                                End if
                            End if
                            if caracter = "." Then
                                iPos2 = i
                            End if

                        End if
                    else

                        Exit Function
                    End if
                else
                    bp = False
                End if
            Next i
            if ipos = 0 Or iPos2 = 0 Then
                Exit Function
            End if

            if iPos2<ipos Then
                Exit Function
            End if


            Validar_Email = True

            Exit Function
    Err_Sub:
            On Error Resume Next

            Validar_Email = False
        End Function
        //Fin. Alexander Hdez 22/10/2012


        //Alexander Hdez 2014-01-08, funcion que desabilita el menu contextuald e la cuenta de titulares
        //Inicio
        //Comienza el Hook

        // hace falta implementar esta funcion
        // 19 mayo 2023
        // RGB
        //public Sub Hook(Handle As Long)
        //    lpPrevWndProc = SetWindowLong(Handle, -4, AddressOf WinProc)
        //End Sub

        //Termina el Hook
        public Sub Unhook(Handle As Long)
            Call SetWindowLong(Handle, -4, lpPrevWndProc)
        End Sub

        //Procedimiento chequea los mensajes que llegan _
        //para ver si se despliega el menú contextual _
        //en el textbox indicado
        //Private ReadOnly DatosOK As Integer
        //Private ReadOnly NoHayDatos As Integer
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        public Function WinProc(ByVal hwnd As Long,
                                ByVal Msg As Long,
                                ByVal wParam As Long,
                                ByVal lParam As Long) As Long
            // Chequea si el mensaje  es WM_CONTEXTMENU ( el menú contextual )
            if Msg = WM_CONTEXTMENU Then
                WinProc = True
            else
                WinProc = CallWindowProc(lpPrevWndProc,
                                  hwnd, Msg, wParam, lParam)
            End if
        End Function
        //Fin Alexander Hdez 2014-01-08

        // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
        // Juan Martínez Díaz
        // 2022-11-23
        public Sub prObtieneConexionBase2019()
            vgNombre_ini = "c:\Cairo\Cairo2000.ini"
            if ExistFile(vgNombre_ini) Then
                ReadINI(vgNombre_ini, "SERVIDOR_2", "Servidor", gsServidor)
                ReadINI(vgNombre_ini, "ODBC_2", "DSN", gsDSN)
                ReadINI(vgNombre_ini, "USUARIO_2", "UID", gsUsrBD)
                ReadINI(vgNombre_ini, "PASSWORD_2", "PWD", gsPwdBD)
                // variables de .Net
                ReadINI(vgNombre_ini, "ODBC_NETCRYSTAL_2", "DSN", gsDSN_NetCrystal)
                ReadINI(vgNombre_ini, "SERVIDOR_NET_2", "Servidor", gsServidor_Net)
                ReadINI(vgNombre_ini, "ODBC_NET_2", "DSN", gsDSN_Net)
                ReadINI(vgNombre_ini, "USUARIO_NET_2", "UID", gsUsrBD_Net)
                ReadINI(vgNombre_ini, "PASSWORD_NET_2", "PWD", gsPwdBD_Net)

                ReadINI(vgNombre_ini, "SERVIDORMTS_2", "Servidormts", gsServidorMTS)
            End if
        End Sub

        // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
        // Juan Martínez Díaz
        // 2022-11-23
        public Sub prObtieneConexionBase2019Dualidad(ByRef prgsServidor As String, ByRef prgsDSN As String, ByRef prgsUsrBD As String, ByRef prgsPwdBD As String)
            vgNombre_ini = "c:\Cairo\Cairo2000.ini"
            if ExistFile(vgNombre_ini) Then
                ReadINI(vgNombre_ini, "SERVIDOR_2", "Servidor", prgsServidor)
                ReadINI(vgNombre_ini, "ODBC_2", "DSN", prgsDSN)
                ReadINI(vgNombre_ini, "USUARIO_2", "UID", prgsUsrBD)
                ReadINI(vgNombre_ini, "PASSWORD_2", "PWD", prgsPwdBD)
            End if
        End Sub

        public static string Midstr(string cadema, int inicia, int longitud)
        {
            return cadema.Substring(inicia, longitud);
            
        }

    }
}
