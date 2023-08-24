using ADODB;

namespace MTSEndososCET.Modulos
{
    public static class ModRecordset
    {

        public const string sNull = null;
        public const string gsINGLESA = "mm/dd/yyyy";   //Fecha en formato Inglesa para los Componentes MTS
        public const string gsFRANCESA = "dd/mmm/yyyy"; //Fecha en formato Francesa para los Componentes MTS


        public enum TipoResultado
        {
            DatosOK = 1,
            NoHayDatos = 2,
            ExisteError = 3
        }

        private static string sErrVB;                //Variable de Error
        private static string sSql;                  //Variable de Uso y Parametros

        private static int iArrCount;
        private static object vParams;             //Manejo de los Parametros de las Funciones de ModRecordset


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
        //   DatosOK = 1       .- Indica que la función fue exitosa y el Rs  contiene datos.
        //   NoHayDatos = 2  .- Indica que la función fue exitosa y el Rs  NO contiene datos.
        //   ExisteError = 3  .- Indica que la función tuvo un error, los errores son regresados en Err1
        //
        //Uso:
        //     Res = EjecutaSql(rsData,sDSN, sSql, Err)
        //     Entrada :
        //               rsData.- Es el Rs que contiene la información.
        //               sDSN .- Es el string de conexión
        //               sSql   .- Es el comando sql que se va a ejecutar
        //     Salida  :
        //                Err1 .- Recordset  que contiene los errores de ADO si el valor regresado
        //                         por la función es falso.
        //Fecha : 7 /Agosto/ 2000
        //Autor : Gerardo Acosta
        //__________________________________________________________

        public static TipoResultado EjecutaSql(ref Recordset rsData, string sDsn, string sSql, Recordset err1, LockTypeEnum LockType = LockTypeEnum.adLockBatchOptimistic, CursorTypeEnum CursorType = CursorTypeEnum.adOpenDynamic, CursorLocationEnum CursorLocation = CursorLocationEnum.adUseClient)
        {

            try
            {
                Recordset rsdatos = new Recordset();
                Connection cnConn = new Connection();
                TipoResultado strEjecutaSql = TipoResultado.DatosOK;

                string sErrVB;
                //Obtenemos la conexion
                cnConn.Open(sDsn);

                //On Error GoTo EjecutaSQLError

                //SDC 2006-09-25 Por si las flies
                cnConn.CommandTimeout = 0;

                rsdatos.ActiveConnection = cnConn;
                rsdatos.CursorLocation = CursorLocation;
                rsdatos.CursorType = CursorType;
                rsdatos.LockType = LockType;
                rsdatos.Source = sSql;
                rsdatos.Open();

                if (rsdatos.BOF && rsdatos.EOF)
                    strEjecutaSql = TipoResultado.NoHayDatos;
                else
                    strEjecutaSql = TipoResultado.DatosOK;


                //rsData = rsdatos;
                //rsData.ActiveConnection = null;
                rsdatos = null;
                //Cerramos la conexión
                cnConn.Close();
                cnConn = null;

                return strEjecutaSql;

            }
            catch (Exception)
            {
                //EjecutaSQLError:
                //sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
                //err1 = ErroresDLL(cnConn, sErrVB)
                //rsData = Nothing
                //cnConn = Nothing

                return TipoResultado.ExisteError;
            }

        }

        //__________________________________________________________
        //Descripción :
        //   La siguiente función Inserta cualquier numero de Parametros a su Tabla especifica.
        //   con los siguientes valores :
        //Uso:
        //     bRes = Inserta(ErrAdo,VarDsn, Parametros)
        //     Entrada :
        //               VarDsn .- Es el string de conexión
        //               Parametros .- Es el arreglo con los datos
        //                                     En la posición Parametros(0)  se escribe el nombre de la Tabla
        //     Salida  :
        //                ErrAdo .- Recordset  que contiene los errores de ADO si el valor regresado
        //                         por la función es falso.
        //Fecha :   13/Julio/2000
        //Autor :   Roberto Reyes I.
        //__________________________________________________________

        //public static bool Inserta(ref Recordset errAdo, ref string varDSN, object parametros)
        //{
        //    object params;

        //    //On Error GoTo msgerror

        //    if ((parametros.Max() - parametros.Min()) > 1)
        //    { 
        //        params = parametros;
        //    }
        //    else
        //    { 
        //         if IsArray(parametros[0]) Then
        //             params = parametros[0]
        //         else
        //            params = parametros
        //        End if
        //    }

        //    set cnnCn = New ADODB.Connection
        //    cnnCn.Open varDSN


        //       sSql = "Insert into " & params (0) & " Values("
        //       iArrCount = 0
        //                   For Each vParams In params
        //                     if iArrCount<> 0 Then
        //                         Select Case VarType(vParams)
        //                         Case 1:  //Null
        //                                        sSql = sSql & vParams & ","


        //                          Case 2:  //Integer
        //                                        sSql = sSql & vParams & ","


        //                         Case 5: //Double
        //                                       sSql = sSql & vParams & ","


        //                         Case 6:  //Money
        //                                       sSql = sSql & vParams & ","
        //                         Case 7: //Date
        //                                       sSql = sSql & "//" & Format(vParams, "mm/dd/yyyy") & "//,"
        //                         Case 8:  //String
        //                                      if vParams = sNull Then
        //                                           sSql = sSql & sNull & ","
        //                                      else
        //            sSql = sSql & SQLString(vParams) & ","
        //                                      End if


        //                         End Select
        //                      End if
        //                      iArrCount = iArrCount + 1
        //                   Next

        //      sSql = Mid(sSql, 1, Len(sSql) - 1) & ")"


        //       cnnCn.Execute sSql, 64


        //       cnnCn.Close
        //       set cnnCn = Nothing
        //       Inserta = True

        //}

        //msgerror:
        //  Inserta = False
        //  cnnCn.Close
        //  set cnnCn = Nothing
        //   sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //   set errAdo = ErroresDLL(cnnCn, sErrVB)

        //End Function

        //__________________________________________________________
        //Descripción :
        //   La siguiente función Ejecuta un Stored Procedure con parametros de Entrada y Salida
        //   con los siguientes valores :
        //Uso:
        //     bRes = ExecSPs(ErrAdo,VarDsn, Parametros, AdoCmd, NumParamIn, NumParamOut)
        //     Entrada :
        //               VarDsn .- Es el string de conexión
        //               Parametros .- Es el arreglo con los datos
        //                                     En la posición Parametros(0)  se escribe el nombre del Stored Procedure
        //               NumParamIn.- Indica el Número de Parametros de Entrada.
        //               NumParamOut.- Indica el Numero de Parametros de Salida.
        //     Salida  :
        //                ErrAdo .- Recordset  que contiene los errores de ADO si el valor regresado
        //                         por la función es falso.
        //                AdoCmd .- Command que regresa los valores del Resultado del SP.
        //Fecha :   13/Julio/2000
        //Autor :   Roberto Reyes I.
        //__________________________________________________________

        //public Function ExecSPs(ByRef errAdo As ADODB.Recordset, ByRef varDSN As String, ByRef parametros As Variant, ByRef AdoCmd As ADODB.Command, ByRef NumParamIn As Integer, ByRef NumParamOut As Integer) As Boolean
        //On Error GoTo msgerror

        //Dim blnFlag As Boolean
        //Dim ArrVariant As Variant

        //blnFlag = True // Para los parámetros de entrada
        //set cnnCn = New ADODB.Connection
        //cnnCn.Open varDSN


        //set AdoCmd = New ADODB.Command
        //AdoCmd.ActiveConnection = cnnCn
        //AdoCmd.CommandType = adCmdStoredProc
        //AdoCmd.CommandText = parametros(0)

        //             For iArrCount = 1 To UBound(parametros)
        //                     ArrVariant = parametros(iArrCount)
        //                     Select Case VarType(ArrVariant(0))

        //                      Case 2:  //Integer
        //                                   Select Case blnFlag
        //                                     Case True:
        //                                                     AdoCmd.Parameters.Append AdoCmd.CreateParameter(ArrVariant(1), adInteger, adParamInput)
        //                                     Case False:
        //                                                     AdoCmd.Parameters.Append AdoCmd.CreateParameter(ArrVariant(1), adInteger, adParamOutput)
        //                                   End Select
        //                     Case 5: //Double
        //                                   Select Case blnFlag
        //                                     Case True:
        //                                                     AdoCmd.Parameters.Append AdoCmd.CreateParameter(ArrVariant(1), adDouble, adParamInput)
        //                                     Case False:
        //                                                     AdoCmd.Parameters.Append AdoCmd.CreateParameter(ArrVariant(1), adDouble, adParamInputOutput)
        //                                   End Select


        //                     Case 7: //Date
        //                                   Select Case blnFlag
        //                                     Case True:
        //                                                     AdoCmd.Parameters.Append AdoCmd.CreateParameter(ArrVariant(1), adDate, adParamInput)
        //                                     Case False:
        //                                                     AdoCmd.Parameters.Append AdoCmd.CreateParameter(ArrVariant(1), adDate, adParamInputOutput)
        //                                   End Select
        //                     Case 8:  //String
        //                                   Select Case blnFlag
        //                                     Case True:
        //                                                     AdoCmd.Parameters.Append AdoCmd.CreateParameter(ArrVariant(1), adChar, adParamInput, ArrVariant(2))
        //                                     Case False:
        //                                                     AdoCmd.Parameters.Append AdoCmd.CreateParameter(ArrVariant(1), adVarChar, adParamOutput, ArrVariant(2))
        //                                   End Select
        //                     End Select
        //                              if iArrCount = NumParamIn Then blnFlag = False
        //                              if iArrCount = (NumParamIn + NumParamOut) Then Exit For
        //               Next iArrCount

        //              For iArrCount = 0 To UBound(parametros)
        //              ArrVariant = parametros(iArrCount + 1)
        //                     if iArrCount <= NumParamIn - 1 Then
        //                        AdoCmd(iArrCount).Value = ArrVariant(0)
        //                      else
        //                        Exit For
        //                     End if
        //              Next iArrCount


        //AdoCmd.Execute
        //set AdoCmd.ActiveConnection = Nothing

        //   cnnCn.Close
        //   set cnnCn = Nothing
        //   ExecSPs = True

        //Exit Function

        //msgerror:

        //  ExecSPs = False
        //   sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //   set errAdo = ErroresDLL(cnnCn, sErrVB)

        //End Function
        ////__________________________________________________________
        ////Descripción :
        ////   La siguiente función Ejecuta un Stored Procedure Regresando un Recordset
        ////   con los siguientes valores :
        ////Uso:
        ////     bRes = EjecutaSPs(ErrAdo,VarDsn, RsAdo, Parametros)
        ////     Entrada :
        ////               VarDsn .- Es el string de conexión
        ////               Parametros .- Es el arreglo con los datos
        ////                                     En la posición Parametros(0)  se escribe el nombre del Stored Procedure
        ////               RsAdo.- Recordset que regresa los registros al componente.
        ////
        ////     Salida  :
        ////                ErrAdo .- Recordset  que contiene los errores de ADO si el valor regresado
        ////                         por la función es falso.
        ////Fecha :   13/Julio/2000
        ////Autor :   Roberto Reyes I.
        ////__________________________________________________________
        //public Function EjecutaSPs(ByRef errAdo As ADODB.Recordset, ByRef varDSN As String, ByRef RsAdo As ADODB.Recordset, parametros As Variant) As Boolean
        //On Error GoTo msgerror

        //set cnnCn = New ADODB.Connection
        //cnnCn.Open varDSN
        //cnnCn.CommandTimeout = 0 //20200911. Alexander Hdez se cambio el timeout de 300 a 0

        //   sSql = parametros(0) & " "
        //   iArrCount = 0
        //               For Each vParams In parametros
        //                 if iArrCount<> 0 Then
        //                     Select Case VarType(vParams)

        //                      Case 2:  //Integer
        //                                    sSql = sSql & vParams & ","


        //                     Case 5: //Double
        //                                   sSql = sSql & vParams & ","


        //                     Case 6:  //Money
        //                                   sSql = sSql & vParams & ","
        //                     Case 7: //Date
        //                                   sSql = sSql & "//" & Format(vParams, "mm/dd/yyyy") & "//,"
        //                     Case 8:  //String
        //                                   sSql = sSql & "//" & vParams & "//,"


        //                     End Select
        //                  End if
        //                  iArrCount = iArrCount + 1
        //               Next
        //    sSql = Mid(sSql, 1, Len(sSql) - 1)
        //  set RsAdo = ADORecordset(cnnCn, errAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)
        //  set RsAdo.ActiveConnection = Nothing


        //   cnnCn.Close
        //   set cnnCn = Nothing
        //   EjecutaSPs = True

        //Exit Function

        //msgerror:
        //  EjecutaSPs = False
        //  sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //  set errAdo = ErroresDLL(cnnCn, sErrVB)
        //  cnnCn.Close
        //  set cnnCn = Nothing
        //End Function
        ////__________________________________________________________
        ////Descripción :
        ////   La siguiente función se usa para la ejecución de los Recordsets anidados
        ////   con los siguientes valores :
        ////Uso:
        ////    rsAux = AdoRecordset(objConnection,ErrAdo, lockType,CursorType,CursorLocation,SQL)
        ////     Entrada :
        ////                        objConnection.- Conexión activada desde el componente.
        ////                        LockType .- tipo de Recordset.
        ////                        CursorType .- Tipo de Cursor.
        ////                        CursorLocation .- En donde se ejecutara el Recordset (Servidor, Cliente)
        ////                        SQL.- Strings de losQuerys.
        ////     Salida  :
        ////                ErrAdo .- Recordset  que contiene los errores de ADO si el valor regresado
        ////                         por la función es falso.
        ////      NOTA : No se manejan los errores, debido a que estos suben directos al componente y este los manipula.
        ////Fecha :   13/Julio/2000
        ////Autor :   Roberto Reyes I.
        ////__________________________________________________________
        //public Function ADORecordset(ByVal objConnection As ADODB.Connection, ByRef errAdo As ADODB.Recordset, Optional ByVal LockType As LockTypeEnum = adLockBatchOptimistic, Optional ByVal CursorType As CursorTypeEnum = adOpenDynamic, Optional CursorLocation As CursorLocationEnum = adUseClient, Optional ByVal SQL As String) As ADODB.Recordset
        //   Dim adoInRs As ADODB.Recordset
        //      set adoInRs = New ADODB.Recordset
        //      set adoInRs.ActiveConnection = objConnection
        //              adoInRs.CursorLocation = CursorLocation
        //              adoInRs.CursorType = CursorType
        //             adoInRs.LockType = LockType
        //             adoInRs.Source = SQL
        //             adoInRs.Open
        //        set adoInRs.ActiveConnection = Nothing
        //        set ADORecordset = adoInRs
        //End Function

        ////__________________________________________________________
        ////Descripción :
        ////   La siguiente función ejecuta un comando sql que no regresa resultados. La función
        ////   regresa TRUE si el sql se ejecuto con exito, en caso contrario regresa FALSE
        ////Uso:
        ////     Res = ExecSql(sDSN, sSql, Err)
        ////     Entrada :
        ////               sDSN .- Es el string de conexión
        ////               sSql   .- Es el comando sql que se va a ejecutar
        ////     Salida  :
        ////                Err .- Recordset  que contiene los errores de ADO si el valor regresado
        ////                         por la función es falso.
        ////Fecha : 7 /Agosto/ 2000
        ////Autor : Gerardo Acosta
        ////__________________________________________________________
        //public Function ExecSql(sDsn As String, sSql As String, rsErr1 As ADODB.Recordset) As Boolean
        //Dim cmdComan As ADODB.Command
        //Dim cnnConn As New ADODB.Connection

        //On Error GoTo msgerror

        ////Obtenemos la conexion
        // cnnConn.Open sDsn
        //cnnConn.CommandTimeout = 0 //Alexander Hernandez 2017-08-10

        //set cmdComan = New ADODB.Command

        //ExecSql = True

        //With cmdComan
        //    .CommandTimeout = 0 //Alexander Hernandez 2017-08-10
        //    .ActiveConnection = cnnConn
        //    .CommandText = sSql
        //    .CommandType = adCmdText
        //    .Execute
        //End With

        ////Cierro la conexión
        //cnnConn.Close

        //set cmdComan = Nothing
        //set cnnConn = Nothing


        //Exit Function

        //msgerror:
        //  ExecSql = False
        //  sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //  set rsErr1 = ErroresDLL(cnnConn, sErrVB)
        //  set cmdComan = Nothing
        //  set cnnConn = Nothing


        //End Function




        ////__________________________________________________________
        ////Descripción :
        ////   La siguiente función inserta en la tabla de Bitacora
        ////   una acción registrada en sistema Cairo
        ////Uso:
        ////     Res = ExecSql(sDSN, sSql, Err)
        ////     Entrada :
        ////               sDSN .- Es el string de conexión
        ////               sComando   .- Es el comando sql que se ejecuto
        ////               IDAccion .- Es el numero de la acción ejecutada
        ////               IDUsuario.- Es el ID del usuario que realizo la acción
        ////               sUsrRed .-    Es el usuario de la red
        ////               sHostName .- es el nombre de la pc donde se ejecuta la acción
        ////     Salida  :
        ////                Err .- Recordset  que contiene los errores de ADO si el valor regresado
        ////                         por la función es falso.
        ////Fecha : 24 /Septiembre/ 2002
        ////Autor : Gerardo Acosta
        ////__________________________________________________________
        //public Function InsertaBitacora(sDsn As String, sComando As String, rsDatosUser As ADODB.Recordset, rsErr1 As ADODB.Recordset) As Boolean
        //Dim sErrVB As String
        //Dim cmdComan As ADODB.Command
        //Dim cnnConn As New ADODB.Connection
        //Dim sSql As String

        //On Error GoTo msgerror

        ////Obtenemos la conexion
        // cnnConn.Open sDsn
        // cnnConn.CommandTimeout = 0 //Alexander Hernandez 2017-08-10

        //set cmdComan = New ADODB.Command

        //InsertaBitacora = True

        //sSql = " insert Bit_Acciones values ("
        //sSql = sSql & rsDatosUser("IDAccion").Value & ","
        //sSql = sSql & rsDatosUser("IDUsuario").Value & ","
        //sSql = sSql & SQLString(sComando)
        //sSql = sSql & ",//" & rsDatosUser("UsrRed").Value & "//,//"
        //sSql = sSql & rsDatosUser("HostName").Value & "//, getdate(), getdate() )"

        //With cmdComan
        //    .CommandTimeout = 0 //Alexander Hernandez Perez 2017-08-10
        //    .ActiveConnection = cnnConn
        //    .CommandText = sSql
        //    .CommandType = adCmdText
        //    .Execute
        //End With


        //set cmdComan = Nothing
        //set cnnConn = Nothing


        //Exit Function

        //msgerror:
        //  InsertaBitacora = False
        //  sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //  set rsErr1 = ErroresDLL(cnnConn, sErrVB)
        //  set cmdComan = Nothing
        //  set cnnConn = Nothing


        //End Function

        //public Function EjecutaStoredProcedure(ByRef errAdo As ADODB.Recordset, ByRef varDSN As String, ByRef RsAdo As ADODB.Recordset, parametros As Variant) As Boolean
        //On Error GoTo msgerror
        //set cnnCn = New ADODB.Connection
        //cnnCn.Open varDSN
        //cnnCn.CommandTimeout = 0

        //sSql = parametros(0) & " "
        //iArrCount = 0
        //For Each vParams In parametros
        //    if iArrCount<> 0 Then
        //        Select Case VarType(vParams)
        //            Case 1:  //NULL
        //                sSql = sSql & "null,"
        //            Case 2:  //Integer
        //                sSql = sSql & vParams & ","
        //            Case 3:  //Long
        //                sSql = sSql & vParams & ","
        //            Case 5: //Double
        //                sSql = sSql & vParams & ","
        //            Case 6:  //Money
        //                sSql = sSql & vParams & ","
        //            Case 7: //Date
        //                sSql = sSql & "//" & Format(vParams, "mm/dd/yyyy") & "//,"
        //            Case 8:  //String
        //                sSql = sSql & "//" & vParams & "//,"
        //            Case 14: //Decimal
        //                sSql = sSql & vParams & ","
        //        End Select
        //    End if
        //    iArrCount = iArrCount + 1
        //Next
        //sSql = Mid(sSql, 1, Len(sSql) - 1)
        //set RsAdo = ADORecordset(cnnCn, errAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)
        //set RsAdo.ActiveConnection = Nothing

        //cnnCn.Close
        //set cnnCn = Nothing
        //EjecutaStoredProcedure = True

        //Exit Function

        //msgerror:
        //    EjecutaStoredProcedure = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    set errAdo = ErroresDLL(cnnCn, sErrVB)
        //    cnnCn.Close
        //    set cnnCn = Nothing
        //End Function

        //public Function EjecutaStoredProcedure2(ByRef errAdo As ADODB.Recordset, ByRef conn As ADODB.Connection, ByRef RsAdo As ADODB.Recordset, parametros As Variant) As Boolean
        //On Error GoTo msgerror

        //conn.CommandTimeout = 0

        //sSql = parametros(0) & " "
        //iArrCount = 0
        //For Each vParams In parametros
        //    if iArrCount<> 0 Then
        //        Select Case VarType(vParams)
        //            Case 1:  //NULL
        //                sSql = sSql & "null,"
        //            Case 2:  //Integer
        //                sSql = sSql & vParams & ","
        //            Case 3:  //Long
        //                sSql = sSql & vParams & ","
        //            Case 5: //Double
        //                sSql = sSql & vParams & ","
        //            Case 6:  //Money
        //                sSql = sSql & vParams & ","
        //            Case 7: //Date
        //                sSql = sSql & "//" & Format(vParams, "mm/dd/yyyy") & "//,"
        //            Case 8:  //String
        //                sSql = sSql & "//" & vParams & "//,"
        //            Case 14: //Decimal
        //                sSql = sSql & vParams & ","
        //        End Select
        //    End if
        //    iArrCount = iArrCount + 1
        //Next
        //sSql = Mid(sSql, 1, Len(sSql) - 1)
        //set RsAdo = ADORecordset(conn, errAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)
        //set RsAdo.ActiveConnection = Nothing

        //EjecutaStoredProcedure2 = True

        //Exit Function

        //msgerror:
        //    EjecutaStoredProcedure2 = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    set errAdo = ErroresDLL(conn, sErrVB)
        //End Function




        ////-----Modificacion a ROP JCMN 30/09/2014
        ////__________________________________________________________
        ////Descripción :
        ////   La siguiente función ejecuta comandos sql que no regresa resultados. La función
        ////   regresa TRUE si el sql se ejecuto con exito, en caso contrario regresa FALSE
        ////Uso:
        ////     Res = ExecSqlROPC(sDSN, sSql, Err)
        ////     Entrada :
        ////               sDSN .- Es el string de conexión
        ////               sSqlTable   .- Es el 1er comando sql que se va a ejecutar
        ////               sSqlConta   .- Es el 1er comando sql que se va a ejecutar
        ////     Salida  :
        ////                rsErr1 .- Recordset  que contiene los errores de ADO si el valor regresado
        ////                         por la función es falso.
        ////Fecha : 30/09/2014
        ////Autor : Julio Cesar Mtz
        ////__________________________________________________________
        //public Function ExecSqlROPC(sDsn As String, sSqlConta As String, rsErr1 As ADODB.Recordset) As Boolean
        //Dim cmdComan As ADODB.Command
        //Dim cnnConn As New ADODB.Connection

        //On Error GoTo msgerror

        ////Obtenemos la conexion
        // cnnConn.Open sDsn

        // //cnnConn.BeginTrans

        //set cmdComan = New ADODB.Command

        //ExecSqlROPC = True

        //With cmdComan
        //    .ActiveConnection = cnnConn
        //    .CommandText = sSqlConta
        //    .CommandType = adCmdText
        //    //.Execute
        //    //.CommandText = sSqlConta
        //    .Execute
        //End With

        ////cnnConn.CommitTrans

        ////Cierro la conexión
        //cnnConn.Close

        //set cmdComan = Nothing
        //set cnnConn = Nothing


        //Exit Function

        //msgerror:
        //  //cnnConn.RollbackTrans
        //  cnnConn.Close
        //  //ExecSql = False
        //  sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //  set rsErr1 = ErroresDLL(cnnConn, sErrVB)
        //  set cmdComan = Nothing
        //  set cnnConn = Nothing


        //End Function






    }
}
