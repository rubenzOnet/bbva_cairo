using ADODB;
using System.Data;
using System.Data.OleDb;




namespace MTSCPolizas.Modulos
{
    public static class ModRecordset
    {
        //_______________________________________________________________
        //Título: Módulo Recordset
        //Módulo: ModRecordset.bas
        //Versión: 1.0
        //Fecha:    13/Julio/2000
        //Autor:    Roberto Reyes I.- Gerardo Acosta M.
        //Modificación:
        //Fecha de Modificacion:
        //_______________________________________________________________
        //Descripción:
        // Inicializacion :
        //           Crea el objeto MTSGenerico.ClsEjecucion con sus Metodos
        //     para el manejo de las Ejecucion de Inserts, Stored Procedure de Recordy
        //     Stored Procedure con Entradas y Salidas, este Objeto de Manejara directo en
        //     las demás Clases.
        //_______________________________________________________________


        public const string sNull = "null";

        public enum TipoResultado
        {
            DatosOK = 1,
            NoHayDatos = 2,
            ExisteError = 3
        }

        private static string sErrVB;                //Variable de Error
        private static string sSql;                  //Variable de Uso y Parametros

        private static int iArrCount;
        private static object vParams;              //Manejo de los Parametros de las Funciones de ModRecordset

        public const string gsINGLESA = "mm/dd/yyyy";   //Fecha en formato Inglesa para los Componentes MTS
        public const string gsFRANCESA = "dd/mmm/yyyy"; //Fecha en formato Francesa para los Componentes MTS

        private static Connection cnnCn;       //Variable de Conexión


        public static string SQLString(string strColumna)
        {
            //Regresa un string en el formato correcto del SQL, si este
            //contiene apostrofes.
            return "//" + strColumna.Replace("//", "////") + "//";
        }

        //__________________________________________________________
        //Descripción :
        //   La siguiente función ejecuta un comando sql que regresa un Recorset. La función
        //   regresa un valor enum con los siguientes valores :
        //   TipoResultado.DatosOK = 1       .- Indica que la función fue exitosa y el Rs  contiene datos.
        //   TipoResultado.NoHayDatos = 2  .- Indica que la función fue exitosa y el Rs  NO contiene datos.
        //   TipoResultado.ExisteError = 3  .- Indica que la función tuvo un error, los errores son regresados en Err1
        //
        //Uso:
        //     Res = EjecutaSql(rsData,sDSN, sSql, Err)
        //     Entrada :
        //               rsData.- Es el Rs que contiene la información.
        //               sDSN .- Es el string de conexión
        //               sSql   .- Es el comando sql que se va a ejecutar
        //     Salida  :
        //                Err1 .- Record que contiene los errores de ADO si el valor regresado
        //                         por la función es falso.
        //Fecha : 7 /Agosto/ 2000
        //Autor : Gerardo Acosta
        //__________________________________________________________

        public static int EjecutaSql(ref Recordset rsData, string sDsn, string sSql, Recordset err1, LockTypeEnum LockType = LockTypeEnum.adLockBatchOptimistic, CursorTypeEnum CursorType = CursorTypeEnum.adOpenDynamic, CursorLocationEnum CursorLocation = CursorLocationEnum.adUseClient)
        {

            RgbConn rgbConn = new RgbConn();
            rgbConn.conectase();


            Recordset rsdatos = new Recordset();
            Connection cnConn;
            TipoResultado EjecutaSqlRet;
            string sErrVB;
            //Obtenemos la conexion
            cnConn = new Connection();
            cnConn.Open(sDsn);

            try
            {
                // On Error GoTo EjecutaSQLError

                //SDC 2006-09-25 Por si las flies
                cnConn.CommandTimeout = 0;

                rsdatos.ActiveConnection = cnConn;
                rsdatos.CursorLocation = CursorLocation;
                rsdatos.CursorType = CursorType;
                rsdatos.LockType = LockType;
                //rsdatos.Source = sSql
                rsdatos.Open(sSql);

                if (rsdatos.BOF && rsdatos.EOF)
                    EjecutaSqlRet = TipoResultado.NoHayDatos;
                else
                    EjecutaSqlRet = TipoResultado.DatosOK;


                rsData = rsdatos;
                rsData.ActiveConnection = null;
                // rsdatos = null;
                // Cerramos la conexión
                cnConn.Close();
                // cnConn = null;
                //Exit Function

                return Convert.ToInt32( EjecutaSqlRet);
            }
            catch (Exception ex)
            {
                //EjecutaSQLError:
                EjecutaSqlRet = TipoResultado.ExisteError;
                sErrVB = ex.Source + "\t" + ex.Message;
                err1 = modErrores.ErroresDLL(cnConn, sErrVB);
                // rsData = null;
                // cnConn = null;
                return Convert.ToInt32(EjecutaSqlRet);
            }
        }


        public static DataTable ConvertRecordsetToDataTable(Recordset rs)
        {
            DataTable dataTable = new DataTable();

            for (int i = 0; i < rs.Fields.Count; i++)
            {
                dataTable.Columns.Add(rs.Fields[i].Name);
            }

            while (!rs.EOF)
            {
                DataRow row = dataTable.NewRow();
                for (int i = 0; i < rs.Fields.Count; i++)
                {
                    row[i] = rs.Fields[i].Value;
                }
                dataTable.Rows.Add(row);
                rs.MoveNext();
            }

            return dataTable;
        }

        //    //__________________________________________________________
        //    //Descripción :
        //    //   La siguiente función Inserta cualquier numero de Parametros a su Tabla especifica.
        //    //   con los siguientes valores :
        //    //Uso:
        //    //     bRes = Inserta(ErrAdo,VarDsn, Parametros)
        //    //     Entrada :
        //    //               VarDsn .- Es el string de conexión
        //    //               Parametros .- Es el arreglo con los datos
        //    //                                     En la posición Parametros(0)  se escribe el nombre de la Tabla
        //    //     Salida  :
        //    //                ErrAdo .- Record que contiene los errores de ADO si el valor regresado
        //    //                         por la función es falso.
        //    //Fecha :   13/Julio/2000
        //    //Autor :   Roberto Reyes I.
        //    //__________________________________________________________
        public static bool Inserta(ref Recordset errAdo, ref string varDSN, object[] parametros)
        {
            object[] _params;

            try
            {
                // On Error GoTo msgerror

                if (parametros.Length > 1)
                {
                    _params = parametros;
                }
                else
                {
                    if (parametros[0] is Array)
                        _params = (object[])parametros[0];
                    else
                        _params = parametros;
                }

                cnnCn = new Connection();
                cnnCn.Open(varDSN);

                sSql = "Insert into " + _params[0] + " Values(";

                iArrCount = 0;

                foreach (var vParams in _params)
                {
                    if (iArrCount != 0)
                    {
                        switch (Type.GetTypeCode(vParams.GetType()))
                        {
                            case TypeCode.Int32:  //Integer
                                sSql = sSql + vParams + ",";
                                break;
                            case TypeCode.Double: //Double
                                sSql = sSql + vParams + ",";
                                break;
                            //case TypeCode.Double:  //Money
                            //    sSql = sSql + vParams + ",";
                            //    break;
                            case TypeCode.DateTime: //Date
                                sSql = sSql + "//" + string.Format("{0:mm/dd/yyyy}", vParams) + "//,";
                                break;
                            case TypeCode.String:  //String
                                if (vParams == null)
                                    sSql = sSql + sNull + ",";
                                else
                                    sSql = sSql + SQLString(vParams.ToString()) + ",";
                                   break;
                        }
                    }

                    iArrCount = iArrCount + 1;
                }

                sSql = sSql.Substring(1, sSql.Length - 1) + ")";
                object recordAffected;
                cnnCn.Execute(sSql, out recordAffected);

                cnnCn.Close();

                return true;
                // Exit Function

            }
            catch (Exception Err)
            {
                //msgerror:
                //Inserta = False
                cnnCn.Close();
                sErrVB = Err.Source + "\t" + Err.Message;
                //errAdo = ErroresDLL(cnnCn, sErrVB)
                return false;
            }

        }

        //    //__________________________________________________________
        //    //Descripción :
        //    //   La siguiente función Ejecuta un Stored Procedure con parametros de Entrada y Salida
        //    //   con los siguientes valores :
        //    //Uso:
        //    //     bRes = ExecSPs(ErrAdo,VarDsn, Parametros, AdoCmd, NumParamIn, NumParamOut)
        //    //     Entrada :
        //    //               VarDsn .- Es el string de conexión
        //    //               Parametros .- Es el arreglo con los datos
        //    //                                     En la posición Parametros(0)  se escribe el nombre del Stored Procedure
        //    //               NumParamIn.- Indica el Número de Parametros de Entrada.
        //    //               NumParamOut.- Indica el Numero de Parametros de Salida.
        //    //     Salida  :
        //    //                ErrAdo .- Record que contiene los errores de ADO si el valor regresado
        //    //                         por la función es falso.
        //    //                AdoCmd .- Command que regresa los valores del Resultado del SP.
        //    //Fecha :   13/Julio/2000
        //    //Autor :   Roberto Reyes I.
        //    //__________________________________________________________

        //public bool ExecSPs(ref Recordset errAdo, ref string varDSN, ref object[] parametros, ref Command AdoCmd, ref int NumParamIn, ref int NumParamOut)
        //{

        //    try
        //    {
        //        // On Error GoTo msgerror
        //        bool blnFlag;
        //        object[] ArrObject;

        //        blnFlag = true; // Para los parámetros de entrada
        //        Connection cnnCn = new Connection();
        //        cnnCn.Open(varDSN);


        //        Command _AdoCmd = new Command();
        //        AdoCmd.ActiveConnection = cnnCn;
        //        AdoCmd.CommandType = CommandTypeEnum.adCmdStoredProc;
        //        AdoCmd.CommandText = parametros[0].ToString();


        //        for (int i = 0; i < parametros.Length; i++)
        //        {
        //            ArrObject = parametros[iArrCount];
        //            Select Case VarType(ArrObject(0))

        //                Case 2  //Integer
        //                    Select Case blnFlag
        //                        Case True
        //                            AdoCmd.Parameters.Append(AdoCmd.CreateParameter(ArrObject(1), adInteger, adParamInput))
        //                        Case False
        //                            AdoCmd.Parameters.Append(AdoCmd.CreateParameter(ArrObject(1), adInteger, adParamOutput))
        //                    End Select
        //                Case 5 //Double
        //                    Select Case blnFlag
        //                        Case True
        //                            AdoCmd.Parameters.Append(AdoCmd.CreateParameter(ArrObject(1), adDouble, adParamInput))
        //                        Case False
        //                            AdoCmd.Parameters.Append(AdoCmd.CreateParameter(ArrObject(1), adDouble, adParamInputOutput))
        //                    End Select

        //                Case 7 //Date
        //                    Select Case blnFlag
        //                        Case True
        //                            AdoCmd.Parameters.Append(AdoCmd.CreateParameter(ArrObject(1), adDate, adParamInput))
        //                        Case False
        //                            AdoCmd.Parameters.Append(AdoCmd.CreateParameter(ArrObject(1), adDate, adParamInputOutput))
        //                    End Select
        //                Case 8  //String
        //                    Select Case blnFlag
        //                        Case True
        //                            AdoCmd.Parameters.Append(AdoCmd.CreateParameter(ArrObject(1), adChar, adParamInput, ArrObject(2)))
        //                        Case False
        //                            AdoCmd.Parameters.Append(AdoCmd.CreateParameter(ArrObject(1), adVarChar, adParamOutput, ArrObject(2)))
        //                    End Select
        //            End Select
        //            if iArrCount = NumParamIn Then blnFlag = False
        //            if iArrCount = (NumParamIn + NumParamOut) Then Exit For
        //        }

        //        For iArrCount = 1 To UBound(parametros)

        //        Next iArrCount

        //        For iArrCount = 0 To UBound(parametros)
        //            ArrObject = parametros(iArrCount + 1)
        //            if iArrCount <= NumParamIn - 1 Then
        //                AdoCmd(iArrCount).Value = ArrObject(0)
        //            Else
        //                Exit For
        //            End if
        //        Next iArrCount


        //        AdoCmd.Execute()
        //        AdoCmd.ActiveConnection = Nothing

        //        cnnCn.Close()
        //        cnnCn = Nothing
        //        ExecSPs = True

        //        Exit Function

        //        }
        //        catch (Exception)
        //        {
        //        msgerror:

        //            ExecSPs = False
        //            sErrVB = String.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //            errAdo = ErroresDLL(cnnCn, sErrVB)

        //            throw;
        //        }

        //    }



        //    //__________________________________________________________
        //    //Descripción :
        //    //   La siguiente función Ejecuta un Stored Procedure Regresando un Recordset
        //    //   con los siguientes valores :
        //    //Uso:
        //    //     bRes = EjecutaSPs(ErrAdo,VarDsn, RsAdo, Parametros)
        //    //     Entrada :
        //    //               VarDsn .- Es el string de conexión
        //    //               Parametros .- Es el arreglo con los datos
        //    //                                     En la posición Parametros(0)  se escribe el nombre del Stored Procedure
        //    //               RsAdo.- Recordque regresa los registros al componente.
        //    //
        //    //     Salida  :
        //    //                ErrAdo .- Record que contiene los errores de ADO si el valor regresado
        //    //                         por la función es falso.
        //    //Fecha :   13/Julio/2000
        //    //Autor :   Roberto Reyes I.
        //    //__________________________________________________________
        //    public Function EjecutaSPs(ByRef errAdo As ADODB.Recordset, ByRef varDSN As String, ByRef RsAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        On Error GoTo msgerror

        //        cnnCn = new ADODB.Connection
        //        cnnCn.Open(varDSN)
        //        cnnCn.CommandTimeout = 0 //20200911. Alexander Hdez se cambio el timeout de 300 a 0

        //        sSql = parametros(0) & " "
        //        iArrCount = 0
        //        For Each vParams In parametros
        //            if iArrCount<> 0 Then
        //                Select Case VarType(vParams)

        //                    Case 2  //Integer
        //                        sSql = sSql & vParams & ","

        //                    Case 5 //Double
        //                        sSql = sSql & vParams & ","

        //                    Case 6  //Money
        //                        sSql = sSql & vParams & ","
        //                    Case 7 //Date
        //                        sSql = sSql & "//" & String.Format(vParams, "mm/dd/yyyy") & "//,"
        //                    Case 8  //String
        //                        sSql = sSql & "//" & vParams & "//,"

        //                End Select
        //            End if
        //            iArrCount = iArrCount + 1
        //        Next
        //        sSql = Mid(sSql, 1, Len(sSql) - 1)
        //        RsAdo = ADORecordset(cnnCn, errAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)
        //        RsAdo.ActiveConnection = Nothing

        //        cnnCn.Close()
        //        cnnCn = Nothing
        //        EjecutaSPs = True

        //        Exit Function

        //msgerror:
        //        EjecutaSPs = False
        //        sErrVB = String.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        errAdo = ErroresDLL(cnnCn, sErrVB)
        //        cnnCn.Close()
        //        cnnCn = Nothing
        //    End Function
        //    //__________________________________________________________
        //    //Descripción :
        //    //   La siguiente función se usa para la ejecución de los Recordsets anidados
        //    //   con los siguientes valores :
        //    //Uso:
        //    //    rsAux = AdoRecordset(objConnection,ErrAdo, lockType,CursorType,CursorLocation,SQL)
        //    //     Entrada :
        //    //                        objConnection.- Conexión activada desde el componente.
        //    //                        LockType .- tipo de Recordset.
        //    //                        CursorType .- Tipo de Cursor.
        //    //                        CursorLocation .- En donde se ejecutara el Record(Servidor, Cliente)
        //    //                        SQL.- Strings de losQuerys.
        //    //     Salida  :
        //    //                ErrAdo .- Record que contiene los errores de ADO si el valor regresado
        //    //                         por la función es falso.
        //    //      NOTA : No se manejan los errores, debido a que estos suben directos al componente y este los manipula.
        //    //Fecha :   13/Julio/2000
        //    //Autor :   Roberto Reyes I.
        //    //__________________________________________________________
        //    public Function ADORecordset(ByVal objConnection As ADODB.Connection, ByRef errAdo As ADODB.Recordset, Optional ByVal LockType As LockTypeEnum = adLockBatchOptimistic, Optional ByVal CursorType As CursorTypeEnum = adOpenDynamic, Optional CursorLocation As CursorLocationEnum = adUseClient, Optional ByVal SQL As String = "") As ADODB.Recordset
        //        Dim adoInRs As ADODB.Recordset
        //        adoInRs = new ADODB.Recordset
        //        adoInRs.ActiveConnection = objConnection
        //        adoInRs.CursorLocation = CursorLocation
        //        adoInRs.CursorType = CursorType
        //        adoInRs.LockType = LockType
        //        adoInRs.Source = SQL
        //        adoInRs.Open()
        //        adoInRs.ActiveConnection = Nothing
        //        //ADORecord = adoInRs
        //    End Function

        //    //__________________________________________________________
        //    //Descripción :
        //    //   La siguiente función ejecuta un comando sql que no regresa resultados. La función
        //    //   regresa TRUE si el sql se ejecuto con exito, en caso contrario regresa FALSE
        //    //Uso:
        //    //     Res = ExecSql(sDSN, sSql, Err)
        //    //     Entrada :
        //    //               sDSN .- Es el string de conexión
        //    //               sSql   .- Es el comando sql que se va a ejecutar
        //    //     Salida  :
        //    //                Err .- Record que contiene los errores de ADO si el valor regresado
        //    //                         por la función es falso.
        //    //Fecha : 7 /Agosto/ 2000
        //    //Autor : Gerardo Acosta
        //    //__________________________________________________________
        //    public Function ExecSql(sDsn As String, sSql As String, rsErr1 As ADODB.Recordset) As Boolean
        //        Dim cmdComan As ADODB.Command
        //        Dim cnnConn As new ADODB.Connection

        //        On Error GoTo msgerror

        //        //Obtenemos la conexion
        //        cnnConn.Open(sDsn)
        //        cnnConn.CommandTimeout = 0 //Alexander Hernandez 2017-08-10

        //        cmdComan = new ADODB.Command

        //        ExecSql = True

        //        With cmdComan
        //            .CommandTimeout = 0 //Alexander Hernandez 2017-08-10
        //            .ActiveConnection = cnnConn
        //            .CommandText = sSql
        //            .CommandType = adCmdText
        //            .Execute()
        //        End With

        //        //Cierro la conexión
        //        cnnConn.Close()

        //        cmdComan = Nothing
        //        cnnConn = Nothing

        //        Exit Function

        //msgerror:
        //        ExecSql = False
        //        sErrVB = String.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErr1 = ErroresDLL(cnnConn, sErrVB)
        //        cmdComan = Nothing
        //        cnnConn = Nothing

        //    End Function




        //    //__________________________________________________________
        //    //Descripción :
        //    //   La siguiente función inserta en la tabla de Bitacora
        //    //   una acción registrada en sistema Cairo
        //    //Uso:
        //    //     Res = ExecSql(sDSN, sSql, Err)
        //    //     Entrada :
        //    //               sDSN .- Es el string de conexión
        //    //               sComando   .- Es el comando sql que se ejecuto
        //    //               IDAccion .- Es el numero de la acción ejecutada
        //    //               IDUsuario.- Es el ID del usuario que realizo la acción
        //    //               sUsrRed .-    Es el usuario de la red
        //    //               sHostName .- es el nombre de la pc donde se ejecuta la acción
        //    //     Salida  :
        //    //                Err .- Record que contiene los errores de ADO si el valor regresado
        //    //                         por la función es falso.
        //    //Fecha : 24 /Septiembre/ 2002
        //    //Autor : Gerardo Acosta
        //    //__________________________________________________________
        //    public Function InsertaBitacora(sDsn As String, sComando As String, rsDatosUser As ADODB.Recordset, rsErr1 As ADODB.Recordset) As Boolean
        //        Dim sErrVB As String
        //        Dim cmdComan As ADODB.Command
        //        Dim cnnConn As new ADODB.Connection
        //        Dim sSql As String

        //        On Error GoTo msgerror

        //        //Obtenemos la conexion
        //        cnnConn.Open(sDsn)
        //        cnnConn.CommandTimeout = 0 //Alexander Hernandez 2017-08-10

        //        cmdComan = new ADODB.Command

        //        InsertaBitacora = True

        //        sSql = " insert Bit_Acciones values ("
        //        sSql = sSql & rsDatosUser("IDAccion").Value & ","
        //        sSql = sSql & rsDatosUser("IDUsuario").Value & ","
        //        sSql = sSql & SQLString(sComando)
        //        sSql = sSql & ",//" & rsDatosUser("UsrRed").Value & "//,//"
        //        sSql = sSql & rsDatosUser("HostName").Value & "//, getdate(), getdate() )"

        //        With cmdComan
        //            .CommandTimeout = 0 //Alexander Hernandez Perez 2017-08-10
        //            .ActiveConnection = cnnConn
        //            .CommandText = sSql
        //            .CommandType = adCmdText
        //            .Execute()
        //        End With

        //        cmdComan = Nothing
        //        cnnConn = Nothing

        //        Exit Function

        //msgerror:
        //        InsertaBitacora = False
        //        sErrVB = String.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErr1 = ErroresDLL(cnnConn, sErrVB)
        //        cmdComan = Nothing
        //        cnnConn = Nothing

        //    End Function

        //    public Function EjecutaStoredProcedure(ByRef errAdo As ADODB.Recordset, ByRef varDSN As String, ByRef RsAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        On Error GoTo msgerror
        //        cnnCn = new ADODB.Connection
        //        cnnCn.Open(varDSN)
        //        cnnCn.CommandTimeout = 0

        //        sSql = parametros(0) & " "
        //        iArrCount = 0
        //        For Each vParams In parametros
        //            if iArrCount<> 0 Then
        //                Select Case VarType(vParams)
        //                    Case 1  //NULL
        //                        sSql = sSql & "null,"
        //                    Case 2  //Integer
        //                        sSql = sSql & vParams & ","
        //                    Case 3  //Long
        //                        sSql = sSql & vParams & ","
        //                    Case 5 //Double
        //                        sSql = sSql & vParams & ","
        //                    Case 6  //Money
        //                        sSql = sSql & vParams & ","
        //                    Case 7 //Date
        //                        sSql = sSql & "//" & String.Format(vParams, "mm/dd/yyyy") & "//,"
        //                    Case 8  //String
        //                        sSql = sSql & "//" & vParams & "//,"
        //                    Case 14 //Decimal
        //                        sSql = sSql & vParams & ","
        //                End Select
        //            End if
        //            iArrCount = iArrCount + 1
        //        Next
        //        sSql = Mid(sSql, 1, Len(sSql) - 1)
        //        RsAdo = ADORecordset(cnnCn, errAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)
        //        RsAdo.ActiveConnection = Nothing

        //        cnnCn.Close()
        //        cnnCn = Nothing
        //        EjecutaStoredProcedure = True

        //        Exit Function

        //msgerror:
        //        EjecutaStoredProcedure = False
        //        sErrVB = String.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        errAdo = ErroresDLL(cnnCn, sErrVB)
        //        cnnCn.Close()
        //        cnnCn = Nothing
        //    End Function

        //    public Function EjecutaStoredProcedure2(ByRef errAdo As ADODB.Recordset, ByRef conn As ADODB.Connection, ByRef RsAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        On Error GoTo msgerror

        //        conn.CommandTimeout = 0

        //        sSql = parametros(0) & " "
        //        iArrCount = 0
        //        For Each vParams In parametros
        //            if iArrCount<> 0 Then
        //                Select Case VarType(vParams)
        //                    Case 1  //NULL
        //                        sSql = sSql & "null,"
        //                    Case 2  //Integer
        //                        sSql = sSql & vParams & ","
        //                    Case 3  //Long
        //                        sSql = sSql & vParams & ","
        //                    Case 5 //Double
        //                        sSql = sSql & vParams & ","
        //                    Case 6  //Money
        //                        sSql = sSql & vParams & ","
        //                    Case 7 //Date
        //                        sSql = sSql & "//" & String.Format(vParams, "mm/dd/yyyy") & "//,"
        //                    Case 8  //String
        //                        sSql = sSql & "//" & vParams & "//,"
        //                    Case 14 //Decimal
        //                        sSql = sSql & vParams & ","
        //                End Select
        //            End if
        //            iArrCount = iArrCount + 1
        //        Next
        //        sSql = Mid(sSql, 1, Len(sSql) - 1)
        //        RsAdo = ADORecordset(conn, errAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)
        //        RsAdo.ActiveConnection = Nothing

        //        EjecutaStoredProcedure2 = True

        //        Exit Function

        //msgerror:
        //        EjecutaStoredProcedure2 = False
        //        sErrVB = String.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        errAdo = ErroresDLL(conn, sErrVB)
        //    End Function




        //    //-----Modificacion a ROP JCMN 30/09/2014
        //    //__________________________________________________________
        //    //Descripción :
        //    //   La siguiente función ejecuta comandos sql que no regresa resultados. La función
        //    //   regresa TRUE si el sql se ejecuto con exito, en caso contrario regresa FALSE
        //    //Uso:
        //    //     Res = ExecSqlROPC(sDSN, sSql, Err)
        //    //     Entrada :
        //    //               sDSN .- Es el string de conexión
        //    //               sSqlTable   .- Es el 1er comando sql que se va a ejecutar
        //    //               sSqlConta   .- Es el 1er comando sql que se va a ejecutar
        //    //     Salida  :
        //    //                rsErr1 .- Record que contiene los errores de ADO si el valor regresado
        //    //                         por la función es falso.
        //    //Fecha : 30/09/2014
        //    //Autor : Julio Cesar Mtz
        //    //__________________________________________________________
        //    public Function ExecSqlROPC(sDsn As String, sSqlConta As String, rsErr1 As ADODB.Recordset) As Boolean
        //        Dim cmdComan As ADODB.Command
        //        Dim cnnConn As new ADODB.Connection

        //        On Error GoTo msgerror

        //        //Obtenemos la conexion
        //        cnnConn.Open(sDsn)

        //        //cnnConn.BeginTrans

        //        cmdComan = new ADODB.Command

        //        ExecSqlROPC = True

        //        With cmdComan
        //            .ActiveConnection = cnnConn
        //            .CommandText = sSqlConta
        //            .CommandType = adCmdText
        //            //.Execute
        //            //.CommandText = sSqlConta
        //            .Execute()
        //        End With

        //        //cnnConn.CommitTrans

        //        //Cierro la conexión
        //        cnnConn.Close()

        //        cmdComan = Nothing
        //        cnnConn = Nothing

        //        Exit Function

        //msgerror:
        //        //cnnConn.RollbackTrans
        //        cnnConn.Close()
        //        //ExecSql = False
        //        sErrVB = String.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErr1 = ErroresDLL(cnnConn, sErrVB)
        //        cmdComan = Nothing
        //        cnnConn = Nothing

        //    End Function


    }
}
