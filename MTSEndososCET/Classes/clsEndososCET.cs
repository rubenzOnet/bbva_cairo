using System.Collections.Generic;
using System.Security.AccessControl;
using System.Xml.Linq;
using ADODB;


namespace MTSEndososCET.Classes
{
    public class clsEndososCET
    {
        private Connection cnnConexion;      //Variable de Conexión
        private Recordset rsAux;             //Variable Auxiliar para el Recordset
        private string sSql;                       //Variable de Sentencias SQL
        private string sErrVB;                     //Variable Descripción de Error
        //private ctxObject As ObjectContext
        private object ctxObject;

        //_______________________________________________________________
        //Título: Consulta los Criterios de Busqueda de Pólizas.
        //Clase: clsEndososCET.iConsulta_Poliza
        //Versión: 1.0
        //Fecha:    11/04/2003
        //Autor:    Samuel Dueñas
        //Modificación:
        //Fecha de Modificacion:
        //_______________________________________________________________
        //Descripción:
        //           Realiza la consulta a la base de datos según los criterios
        //           de búsqueda.
        //_______________________________________________________________
        //Public Function iConsulta_Poliza(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    rsRecord.CursorLocation = adUseClient


        //    sSql = "declare  @fh_inider datetime, @id_poliza int " & Chr(13)
        //    sSql = sSql + "select   @id_poliza = ID_Poliza from Polizas "
        //    sSql = sSql + "where    ID_Empresa = " & vParametros(2)
        //    if vParametros(0) <> 0 Then sSql = sSql + "     and Fol_Poliza = " & vParametros(0)
        //    if vParametros(1) <> Empty Then sSql = sSql + "     and Num_SegSocial like //" & vParametros(1) & "//"
        //    sSql = sSql + "         and ID_InstitucionSS in (1,2)" //CEFB CAIRO-10-0138 Solo IMMS97 e IMSS08
        //    sSql = sSql + Chr(13)
        //    sSql = sSql + "select   @fh_inider = Fecha_Aplicacion "
        //    sSql = sSql + "from     Endosos e, "
        //    sSql = sSql + "         DEndosos d "
        //    sSql = sSql + "where    e.ID_HisEndoso = d.ID_HisEndoso "
        //    sSql = sSql + "         and ID_PolBenef in (select ID_PolBenef from Pol_Benefs where ID_Parentesco = 1) "
        //    sSql = sSql + "         and ID_StaResolEndoso = 1 "
        //    sSql = sSql + "         and ID_Endoso = 8 "
        //    sSql = sSql + "         and ID_Poliza = @id_poliza "
        //    sSql = sSql + "select   @fh_inider = isnull(@fh_inider,Fecha_IniDer) from Polizas where ID_Poliza = @id_poliza " & Chr(13)
        //    //NOTA IMPORTANTE, CUALUQIER CAMPO QUE SE AÑADA
        //    //DEBE IR AL FINAL DE ESTE SELECT
        //    sSql = sSql + "select   Ramo,"
        //    sSql = sSql + "         Pension,"
        //    sSql = sSql + "         RTRIM(Nom_Aseg) + // // + RTRIM(ApP_Aseg) + // // + RTRIM(ApM_Aseg) as Nombre_Aseg,"
        //    sSql = sSql + "         Fecha_Solic,"
        //    sSql = sSql + "         pol.Fecha_Resol,"
        //    sSql = sSql + "         Fecha_IniDer,"
        //    sSql = sSql + "         CURP,"
        //    sSql = sSql + "         Sexo," //Campo utilizado en frmEndosoCET.bValidaAlta()
        //    sSql = sSql + "         Fecha_Nacto," //Campo utilizado en frmEndosoCET.bValidaAlta()
        //    sSql = sSql + "         RTRIM(Nom_Solic) + // // + RTRIM(ApP_Solic) + // // + RTRIM(ApM_Solic) as Nombre_Solic,"
        //    sSql = sSql + "         Domicilio, "
        //    sSql = sSql + "         Salario_IV, "
        //    sSql = sSql + "         Salario_RT, "
        //    sSql = sSql + "         Pje_Ayuda, "
        //    sSql = sSql + "         Pje_Valuacion, "
        //    sSql = sSql + "         ID_Poliza, "
        //    sSql = sSql + "         Fol_Poliza, "
        //    sSql = sSql + "         r.ID_Ramo, "
        //    sSql = sSql + "         p.ID_Pension, "
        //    sSql = sSql + "         Fecha_IniDerAlta = @fh_inider, "
        //    sSql = sSql + "         Sobrevivencia = case when @fh_inider <> Fecha_IniDer then 1 else 0 end, "
        //    sSql = sSql + "         Concurrencia = (select count(*) from Tmp_EndosoCET t where t.ID_Poliza = pol.ID_Poliza), "
        //    sSql = sSql + "         Fecha_IniVig "
        //    sSql = sSql + "         ,re.Fecha_ABaseResol " //EBS BAJAS POR FALLECIMIENTO 12/02/2016
        //    sSql = sSql + "         ,pol.Fecha_Emision " //EBS BAJAS POR FALLECIMIENTO 12/02/2016
        //    //NOTA IMPORTANTE, CUALUQIER CAMPO QUE SE AÑADA
        //    //DEBE IR AL FINAL DE ESTE SELECT

        //    // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
        //    // Juan Martínez Díaz
        //    // 2022-07-11
        //    // [Operadores *=, =*]
        //    sSql = sSql + " from Polizas pol "
        //    sSql = sSql + " inner join Cat_Ramos r on pol.ID_Ramo = r.ID_Ramo "
        //    sSql = sSql + " inner join Cat_Pensiones p on pol.ID_Pension = p.ID_Pension "
        //    sSql = sSql + " inner join Cat_Sexos s on pol.ID_Sexo = s.ID_Sexo and ID_Poliza= @id_poliza and pol.ID_InstitucionSS in (1,2) " //CEFB BAU Solo IMMS97 e IMSS08
        //    sSql = sSql + " left join Resoluciones re on pol.Num_SegSocial = re.Num_SegSocial " //EBS BAJAS POR FALLECIMIENTO 12/02/2016

        //    // Reemplaza

        ////    sSql = sSql + "from     Polizas pol,"
        ////    sSql = sSql + "         Cat_Ramos r,"
        ////    sSql = sSql + "         Cat_Pensiones p,"
        ////    sSql = sSql + "         Cat_Sexos s "
        ////    sSql = sSql + "         ,Resoluciones re " //EBS BAJAS POR FALLECIMIENTO 12/02/2016
        ////    sSql = sSql + "where    pol.ID_Ramo = r.ID_Ramo"
        ////    sSql = sSql + "         and pol.ID_Pension = p.ID_Pension"
        ////    sSql = sSql + "         and pol.ID_Sexo = s.ID_Sexo"
        ////    sSql = sSql + "         and ID_Poliza = @id_poliza"
        ////    sSql = sSql + "         and pol.ID_InstitucionSS in (1,2)" //CEFB BAU Solo IMMS97 e IMSS08
        ////    sSql = sSql + "         and pol.Num_SegSocial *= re.Num_SegSocial" //EBS BAJAS POR FALLECIMIENTO 12/02/2016

        //    // [Operadores *=, =*]


        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic
        //    Set rsRecord = rsRecord.NextRecordset
        //    Set rsRecord = rsRecord.NextRecordset
        //    Set rsRecord = rsRecord.NextRecordset


        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iConsulta_Poliza = NoHayDatos
        //    Else
        //        iConsulta_Poliza = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iConsulta_Poliza = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        //Public Function iConsulta_Beneficiarios(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    rsRecord.CursorLocation = adUseClient


        //    sSql = "select  StaPolBenef = case ID_StaPolBenef when 1 then //ACTIVO// "
        //    sSql = sSql + "         when 2 then //BAJA// "
        //    sSql = sSql + "         else //ALTA// end, "
        //    sSql = sSql + " ID_StaPolBenef, "
        //    sSql = sSql + " ID_PolBenef, "
        //    sSql = sSql + " ID_Ramo, "
        //    sSql = sSql + " p.ID_Pension, "
        //    sSql = sSql + " cpen.Pension, "
        //    sSql = sSql + " te.ID_Parentesco, "
        //    sSql = sSql + " Abv_Parentesco, "
        //    sSql = sSql + " te.ID_Sexo, "
        //    sSql = sSql + " Abv_Sexo, "
        //    sSql = sSql + " te.ID_Orfandad, "
        //    sSql = sSql + " Cve_Orfandad, "
        //    sSql = sSql + " te.ID_Invalidez, "
        //    sSql = sSql + " Invalidez, "
        //    sSql = sSql + " te.Fecha_Nacto, "
        //    sSql = sSql + " Fecha_Vento, "
        //    sSql = sSql + " te.Fecha_IniDer, "
        //    sSql = sSql + " ID_Grupo, "
        //    sSql = sSql + " RTRIM(Nom_Benef) + // // + RTRIM(ApP_Benef) + // // + RTRIM(ApM_Benef) as Nombre_Benef "
        //    sSql = sSql + "from Polizas p, "
        //    sSql = sSql + " Pol_Benefs te, "
        //    sSql = sSql + " Cat_Parentescos cp, "
        //    sSql = sSql + " Cat_Orfandades co, "
        //    sSql = sSql + " Cat_Pensiones cpen, "
        //    sSql = sSql + " Cat_Sexos cs, "
        //    sSql = sSql + " Cat_Invalidez ci "
        //    sSql = sSql + "where p.ID_Poliza = te.ID_Poliza "
        //    sSql = sSql + " and te.ID_Parentesco = cp.ID_Parentesco "
        //    sSql = sSql + " and te.ID_Orfandad = co.ID_Orfandad "
        //    sSql = sSql + " and te.ID_Sexo = cs.ID_Sexo "
        //    sSql = sSql + " and te.ID_Invalidez = ci.ID_Invalidez "
        //    sSql = sSql + " and cpen.ID_Pension = p.ID_Pension "
        //    sSql = sSql + " and te.ID_Parentesco <> 5 "
        //    sSql = sSql + " and ID_StaPolBenef = 1 "
        //    sSql = sSql + " and te.ID_Poliza = " & vParametros(0)
        //    sSql = sSql + "order by te.ID_Parentesco "
        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic

        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iConsulta_Beneficiarios = NoHayDatos
        //    Else
        //        iConsulta_Beneficiarios = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iConsulta_Beneficiarios = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function
        ////_______________________________________________________________
        ////Título: Consulta Beneficiarios de la poliza.
        ////Clase: clsEndososCET.bsp_EndosoCET
        ////Versión: 1.0
        ////Fecha:    15/04/2003
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Realiza la consulta de los beneficiarios de la póliza.
        ////_______________________________________________________________
        //Public Function bsp_EndosoCET(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Boolean
        //On Error GoTo msgerror


        //    bsp_EndosoCET = EjecutaStoredProcedure(rsErrAdo, gDSN, rsRecord, vParametros)


        //Exit Function
        //msgerror:
        //    bsp_EndosoCET = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function
        ////_______________________________________________________________
        ////Título: Hisotoria de la Póliza.
        ////Clase: clsEndososCET.bHistoriaPoliza
        ////Versión: 1.0
        ////Fecha:    11/04/2003
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Llama a la stored procedure que regenera la historia
        ////           de la Póliza
        ////_______________________________________________________________
        //Public Function bCalculaEndoso(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional ByRef rsAjuste As ADODB.Recordset, Optional ByRef vParametros As Variant) As Boolean

        //    Select Case vParametros(2)
        //    Case 8:
        //        bCalculaEndoso = bFallecimiento(rsErrAdo, gDSN, rsRecord, rsAjuste, vParametros)
        //    Case 9:
        //        bCalculaEndoso = bSegundasNupcias(rsErrAdo, gDSN, rsRecord, rsAjuste, vParametros)
        //    Case 10:
        //        bCalculaEndoso = bImprocedencia(rsErrAdo, gDSN, rsRecord, rsAjuste, vParametros)
        //    Case 27: //CEFB YA9A0E Sep2012
        //        bCalculaEndoso = bCancelacionSinDevReserva(rsErrAdo, gDSN, rsRecord, vParametros) //CEFB YA9A0E Sep2012
        //    End Select


        //End Function

        ////_______________________________________________________________
        ////Título: Consulta los Criterios de Busqueda de Pólizas.
        ////Clase: clsEndososCET.iConsulta_Poliza
        ////Versión: 1.0
        ////Fecha:    11/04/2003
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Realiza la consulta a la base de datos según los criterios
        ////           de búsqueda.
        ////_______________________________________________________________
        //Public Function bGuardaEndoso(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Boolean
        //On Error GoTo msgerror


        //    bGuardaEndoso = EjecutaStoredProcedure(rsErrAdo, gDSN, rsRecord, vParametros)

        //Exit Function
        //msgerror:
        //    bGuardaEndoso = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function


        ////_______________________________________________________________
        ////Título: bLlenaCatalogos
        ////Clase:  clsEndososCET.bLlenaCatalogos
        ////Versión: 1.0
        ////Fecha:    14/04/2003
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////
        ////           Consulta de Catálogos.
        ////_______________________________________________________________
        //Public Function bLlenaCatalogos(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional iNumCatalogo As Integer) As Boolean
        //On Error GoTo msgerror
        // Set cnnConexion = New ADODB.Connection
        //     cnnConexion.Open gDSN

        //    if iNumCatalogo = 0 Then
        //        sSql = "select * "
        //        sSql = sSql + "from     Cat_Endosos "
        //        sSql = sSql + "where   En_Pantalla = 1 "
        //        //sSql = sSql + "where    ID_Endoso in (7,8,9,10)"
        //    Elseif iNumCatalogo = 1 Then
        //        sSql = "select          ID_Ramo,"
        //        sSql = sSql + "         Ramo "
        //        sSql = sSql + "from     Cat_Ramos "
        //        sSql = sSql + "where    ID_Ramo in (1,2,3,4,5,6)"
        //    Elseif iNumCatalogo = 2 Then
        //        sSql = "select          ID_Parentesco,"
        //        sSql = sSql + "         Parentesco "
        //        sSql = sSql + "from     Cat_Parentesco "
        //    End if


        //    Set rsRecord = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)

        //cnnConexion.Close
        //Set cnnConexion = Nothing
        //bLlenaCatalogos = True

        //Exit Function
        //msgerror:
        //    bLlenaCatalogos = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        //Public Function bSegundasNupcias(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional ByRef rsAjuste As ADODB.Recordset, Optional ByRef vParametros As Variant) As Boolean
        //Dim vParamHistoria(6) As Variant      //Nuevo variant para ejecutar Stored Procedure
        //Dim vParamPtmo(10) As Variant      //Nuevo variant para ejecutar Stored Procedure //CEFB 2009-08-14 Nueva firma de sp_EstadoActualPrestamosPB
        //Dim iCont As Integer
        //Dim dPago_Indebido As Double
        //Dim dID_Prestamo As Long
        //Dim dPrestamo As Double
        //Dim tFecha_Matrimonio As Date
        //Dim tFecha_Valuacion As Date
        //Dim dFiniquitoNeto As Double
        //Dim dCuanMensViuda As Double
        //Dim dPagosARedistribuir As Double
        //Dim dPagosARedistribuirInc As Double
        //Dim dBasico As Double
        //Dim dArticulo14 As Double
        //Dim dAguinaldo As Double
        //Dim dAgArticulo14 As Double
        //On Error GoTo msgerror

        //    dPrestamo = 0
        //    dID_Prestamo = 0
        //    dPago_Indebido = 0
        //    dPagosARedistribuir = 0
        //    dBasico = 0
        //    dArticulo14 = 0
        //    dAguinaldo = 0
        //    dAgArticulo14 = 0

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN

        ////Traer la historia 0de la poliza
        //    vParamHistoria(0) = "sp_HistoriaPoliza"   //Nombre del stored procedure
        //    vParamHistoria(1) = vParametros(0)        //ID_Poliza
        //    vParamHistoria(2) = Format(vParametros(7), "yyyy-mm-dd") //Fecha Inicial
        //    vParamHistoria(3) = Format(vParametros(8), "yyyy-mm-dd") //Fecha Final
        //    vParamHistoria(4) = 9 //ID de Endoso
        //    vParamHistoria(5) = Format(vParametros(10), "yyyy-mm-dd") //Fecha System

        //    tFecha_Matrimonio = vParametros(7)


        //    if EjecutaStoredProcedure2(rsErrAdo, cnnConexion, rsRecord, vParamHistoria) = True Then
        //        //Traer Datos de Ajuste de ROPC
        //        //SDC 2006-08-15 Ya no se utiliza el grid de ajuste, los movimientos contables se hacen al
        //        //momento de guardar el endoso
        //        //SDC 2006-11-07 Modificaciones al cálculo de PI, Retro y Finiquito de Viuda.
        //        //El cálculo se hace en el stored procedure


        //        dPago_Indebido = rsRecord!Indebido
        //        dCuanMensViuda = rsRecord!Cuantia
        //        dFiniquitoNeto = dCuanMensViuda * 36


        //        Do While Not rsRecord.EOF
        //            //Total de Básico
        //            if rsRecord!ID_Beneficio = 76 Then
        //                dBasico = dBasico + rsRecord!Diferencia
        //            //Total de Artículo 14
        //            Elseif rsRecord!ID_Beneficio = 78 Then
        //                dArticulo14 = dArticulo14 + rsRecord!Diferencia
        //            //Total de Aguinaldo
        //            Elseif rsRecord!ID_Beneficio = 79 Then
        //                dAguinaldo = dAguinaldo + rsRecord!Diferencia
        //            //Total de Aguinaldo de Artículo 14
        //            Elseif rsRecord!ID_Beneficio = 80 Then
        //                dAgArticulo14 = dAgArticulo14 + rsRecord!Diferencia
        //            End if
        //            rsRecord.MoveNext
        //        Loop
        //        rsRecord.MoveFirst

        //        //Préstamo con Seguros Bancomer
        //        sSql = "select  ID_FolPrestamo "
        //        sSql = sSql + "from Prestamos_Bancomer pr, "
        //        sSql = sSql + "     Tmp_EndosoCET t "
        //        sSql = sSql + "where pr.ID_Poliza = t.ID_Poliza "
        //        sSql = sSql + "     and t.ID_Grupo = pr.ID_Grupo "
        //        sSql = sSql + "     and t.ID_Poliza = " & vParametros(0)
        //        sSql = sSql + "     and ID_StaPolBenef = 2 " //Beneficiario de baja
        //        sSql = sSql + "     and ID_StaPtoSeg = 4 " //Préstamo Activo
        //        Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)


        //        if Not rsAux.EOF Then
        //            dID_Prestamo = rsAux!ID_FolPrestamo
        //        End if

        //        //Calcular la cantidad a liquidar
        //        if dID_Prestamo > 0 Then
        //            vParamPtmo(0) = "sp_EstadoActualPrestamo"   //Nombre del stored procedure
        //            vParamPtmo(1) = dID_Prestamo        //Folio
        //            vParamPtmo(2) = 1 //Liquidación
        //            vParamPtmo(3) = Null //Fecha Liquidacion --CEFB 18-06-2009 Se agrega null al parametro
        //            if EjecutaStoredProcedure(rsErrAdo, gDSN, rsAux, vParamPtmo) = True Then
        //                //dPrestamo = Round(rsAux!cantidadParaLiquidar, 2) + Round(rsAux!next_PagoFijo, 2)
        //                dPrestamo = Round(rsAux!cantidadParaLiquidar, 2)
        //            End if
        //        End if

        //        //Préstamo con Pensiones Bancomer
        //        //Alexander Hdez 27/Jul/2010. COMENTAR Se comenta para que calcule bien los pagos de endosos en cuanto a los PB y PSeguros
        ////        sSql = "select  ID_FolPrestamo "
        ////        sSql = sSql + "from Prestamos_PB pr, "
        ////        sSql = sSql + "     Tmp_EndosoCET t "
        ////        sSql = sSql + "where pr.ID_Poliza = t.ID_Poliza "
        ////        sSql = sSql + "     and t.ID_Grupo = pr.ID_Grupo "
        ////        sSql = sSql + "     and t.ID_Poliza = " & vParametros(0)
        ////        sSql = sSql + "     and ID_StaPolBenef = 2 " //Beneficiario de baja
        ////        sSql = sSql + "     and ID_StaPtoPB = 4 " //Préstamo Activo
        ////        Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)
        ////
        ////        if Not rsAux.EOF Then
        ////            dID_Prestamo = rsAux!ID_FolPrestamo
        ////        End if

        //        //Calcular la cantidad a liquidar
        //        //Alexander 26/Ago/2010, comente porque en la pantalla de endosos aparecio a liquidar el monto de pensiones en lugar del de seguros
        //        //Inicio
        ////        if dID_Prestamo > 0 Then
        ////            vParamPtmo(0) = "sp_EstadoActualPrestamoPB"   //Nombre del stored procedure
        ////            vParamPtmo(1) = dID_Prestamo        //Folio
        ////            vParamPtmo(2) = 1 //Liquidación
        ////            vParamPtmo(3) = Format(vParametros(10), "yyyy-mm-dd") //Fecha del EndosoLiquidación
        ////            vParamPtmo(4) = Null //CEFB 19-08-2009 El parametro @diasInteres debe ser null
        ////            vParamPtmo(5) = Null //CEFB 19-08-2009 El parametro @capitalInsoluto debe ser null
        ////            vParamPtmo(6) = Null //CEFB 19-08-2009 El parametro @intereses debe ser null
        ////            vParamPtmo(7) = Null //CEFB 19-08-2009 El parametro @ivaInteresesReales debe ser null
        ////            vParamPtmo(8) = Null //CEFB 19-08-2009 El parametro @seguro debe ser null
        ////            vParamPtmo(9) = Null //CEFB 19-08-2009 El parametro @totalDeuda debe ser null
        ////            if EjecutaStoredProcedure(rsErrAdo, gDSN, rsAux, vParamPtmo) = True Then
        ////                dPrestamo = Round(rsAux!totalDeuda, 2)
        ////            End if
        ////        End if
        //        //fin

        //        //SDC 2007-02-23 Si la viuda tiene derecho al art14 2004 los PR lo incluyen, si no,
        //        //se manejan como PR incobrables y no se restan al finiquito
        //        rsRecord.Filter = "ID_Parentesco = 2"
        //        vParametros(24) = rsRecord!StatusArt14
        //        if rsRecord!StatusArt14 = 2 Then
        //            dPagosARedistribuir = dBasico + dArticulo14 + dAguinaldo + dAgArticulo14
        //            dPagosARedistribuirInc = 0
        //        Else
        //            dPagosARedistribuir = dBasico + dAguinaldo
        //            dPagosARedistribuirInc = dArticulo14 + dAgArticulo14
        //        End if
        //        rsRecord.Filter = adFilterNone

        //        //Cuantia Mensual Viuda a FhM
        //        vParametros(12) = dCuanMensViuda
        //        //Finiquito por Nuevas Nupcias
        //        vParametros(13) = dFiniquitoNeto
        //        //Total Pagos Indebidos
        //        vParametros(14) = dPago_Indebido
        //        //Total Pagos Indebidos
        //        vParametros(15) = dPrestamo

        //        //Finiquito Total Viuda
        //        vParametros(16) = dMaximo(dFiniquitoNeto - dPago_Indebido - dPagosARedistribuir, 0)

        //        //Pagos a redistribuir
        //        vParametros(17) = dPagosARedistribuir
        //        vParametros(18) = dBasico
        //        vParametros(19) = dArticulo14
        //        vParametros(20) = dAguinaldo
        //        vParametros(21) = dAgArticulo14
        //        vParametros(22) = dPagosARedistribuirInc

        //        //SDC 2006-08-21 Revisar si hay descuentos en la ROPC que se está ajustando
        //        //para avisar al usuario
        //        //SDC 2007-02-23 Ya que este endoso no modifica ROPC solo consultamos el total
        //        //de beneficio básico en la ROPC
        //        sSql = "select  Total = isnull(sum(Imp_PBenef), 0) "
        //        sSql = sSql + "from Pagos pg, "
        //        sSql = sSql + "     MPagos m, "
        //        sSql = sSql + "     Cat_TipoPagos ct, "
        //        sSql = sSql + "     Cat_Beneficios cb "
        //        sSql = sSql + "where pg.ID_Pago = m.ID_Pago "
        //        sSql = sSql + "     and pg.ID_TipoPago = ct.ID_TipoPago "
        //        sSql = sSql + "     and m.ID_Beneficio = cb.ID_Beneficio "
        //        sSql = sSql + "     and pg.ID_Grupo = (select ID_Grupo from Pol_Benefs pb where ID_Poliza = pg.ID_Poliza and ID_Parentesco = 2 and ID_StaPolBenef = 1) "
        //        sSql = sSql + "     and ct.ID_TpoROPC = 3 " //ROPC Grupo
        //        sSql = sSql + "     and pg.ID_StaPago = 2 " //En ROPC
        //        sSql = sSql + "     and (cb.ID_TpoBenef = 1 or cb.ID_Beneficio in (130,131)) " //Solo Básicos //CEFB Se agrega BAU ROPC-BAU
        //        sSql = sSql + "     and pg.ID_Poliza = " & vParametros(0)
        //        Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)

        //        if Not rsAux.EOF Then
        //            vParametros(23) = rsAux!Total
        //        End if

        //        //////////////vParametros (24) Ya tiene el status de articulo 14 de la viuda

        //        bSegundasNupcias = True
        //    Else
        //        bSegundasNupcias = False
        //    End if

        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    bSegundasNupcias = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //End Function
        //Public Function bFallecimiento(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional ByRef rsAjuste As ADODB.Recordset, Optional ByRef vParametros As Variant) As Boolean
        //Dim vNParametros(6) As Variant      //Nuevo variant para ejecutar Stored Procedure
        //Dim tFecha_Fallecimiento As Date
        //On Error GoTo msgerror


        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN

        ////Traer la historia de la poliza
        //    vNParametros(0) = "sp_HistoriaPoliza"   //Nombre del stored procedure
        //    vNParametros(1) = vParametros(0)        //ID_Poliza
        //    vNParametros(2) = Format(vParametros(4), "yyyy-mm-dd") //Fecha Inicial
        //    vNParametros(3) = Format(vParametros(5), "yyyy-mm-dd") //Fecha Final
        //    vNParametros(4) = 8 //ID de Endoso
        //    vNParametros(5) = Format(vParametros(10), "yyyy-mm-dd") //Fecha System

        //    tFecha_Fallecimiento = vParametros(4)
        //    iSobrevivencia = vParametros(6)

        //    if EjecutaStoredProcedure2(rsErrAdo, cnnConexion, rsRecord, vNParametros) = True Then
        //        //Traer Datos de Ajuste de ROPC
        //        //SDC 2006-08-15 Ya no se utiliza el grid de ajuste, los movimientos contables se hacen al
        //        //momento de guardar el endoso

        //        //Traer el pago al mes de fallecimiento
        //        sSql = "select Importe = isnull(sum(Importe), 0) "
        //        sSql = sSql + "from Tmp_PagosEndosos "
        //        sSql = sSql + "where ID_Poliza = " & vParametros(0)
        //        sSql = sSql + "     and ID_Beneficio = 85 "
        //        Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)


        //        if Not rsAux.EOF Then
        //            vParametros(12) = rsAux!Importe
        //        End if

        //        //Traer el aguinaldo al año de fallecimiento
        //        sSql = "select Importe = isnull(sum(Importe), 0) "
        //        sSql = sSql + "from Tmp_PagosEndosos "
        //        sSql = sSql + "where ID_Poliza = " & vParametros(0)
        //        sSql = sSql + "     and ID_Beneficio in (112, 113) "
        //        Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)


        //        if Not rsAux.EOF Then
        //            vParametros(13) = rsAux!Importe
        //        End if

        //        //INI CEFB Oct2011 No sumaba el incremento cuando era ROPC HS
        //        //Traer el pago al mes de fallecimiento de Hijos Suspendidos
        //        sSql = "select Importe = isnull(sum(Monto_Ajuste), 0) "
        //        sSql = sSql + "from Tmp_AjusteROPCEndoso "
        //        sSql = sSql + "where ID_Poliza = " & vParametros(0)
        //        sSql = sSql + "     and ID_Beneficio = 85 "
        //        Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)


        //        if Not rsAux.EOF Then
        //            vParametros(12) = vParametros(12) + rsAux!Importe
        //        End if
        //        //FIN CEFB Oct2011 No sumaba el incremento cuando era ROPC HS

        //        //SDC 2006-08-21 Revisar si hay descuentos en la ROPC que se está ajustando
        //        //para avisar al usuario
        //        sSql = "select  Importe = isnull(sum(abs(Imp_PBenef)), 0) "
        //        sSql = sSql + "from Pagos pg, "
        //        sSql = sSql + "     MPagos m, "
        //        sSql = sSql + "     Cat_TipoPagos ct "
        //        sSql = sSql + "where pg.ID_Pago = m.ID_Pago "
        //        sSql = sSql + "     and pg.ID_TipoPago = ct.ID_TipoPago "
        //        sSql = sSql + "     and pg.ID_Poliza = " & vParametros(0)
        //        sSql = sSql + "     and ct.ID_TpoROPC in (2,3) "
        //        sSql = sSql + "     and pg.ID_StaPago in (2,8) "
        //        sSql = sSql + "     and m.Imp_PBenef < 0 "
        //        Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)


        //        if Not rsAux.EOF Then
        //            vParametros(14) = rsAux!Importe
        //        End if

        //        bFallecimiento = True
        //    Else
        //        bFallecimiento = False
        //    End if

        //    cnnConexion.Close
        //    Set cnnConexion = Nothing

        //Exit Function
        //msgerror:
        //    bFallecimiento = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //End Function
        //Public Function bImprocedencia(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional ByRef rsAjuste As ADODB.Recordset, Optional ByRef vParametros As Variant) As Boolean
        //Dim vNParametros(6) As Variant      //Nuevo variant para ejecutar Stored Procedure
        //On Error GoTo msgerror


        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN

        //    //Genera la historia de la poliza
        //    vNParametros(0) = "sp_HistoriaPoliza"   //Nombre del stored procedure
        //    vNParametros(1) = vParametros(0)        //ID_Poliza
        //    vNParametros(2) = Format(vParametros(4), "yyyy-mm-dd") //Fecha Inicial
        //    vNParametros(3) = Format(vParametros(5), "yyyy-mm-dd") //Fecha Final
        //    vNParametros(4) = 10 //ID de Endoso
        //    vNParametros(5) = Format(vParametros(10), "yyyy-mm-dd") //Fecha System

        //    if EjecutaStoredProcedure2(rsErrAdo, cnnConexion, rsRecord, vNParametros) = True Then
        //        //Traer Datos de Ajuste de ROPC
        //        //SDC 2006-08-15 Ya no se utiliza el grid de ajuste, los movimientos contables se hacen al
        //        //momento de guardar el endoso

        //        //SDC 2006-08-21 Revisar si hay descuentos en la ROPC que se está ajustando
        //        //para avisar al usuario
        //        sSql = "select  Importe = isnull(sum(abs(Imp_PBenef)), 0) "
        //        sSql = sSql + "from Pagos pg, "
        //        sSql = sSql + "     MPagos m, "
        //        sSql = sSql + "     Cat_TipoPagos ct "
        //        sSql = sSql + "where pg.ID_Pago = m.ID_Pago "
        //        sSql = sSql + "     and pg.ID_TipoPago = ct.ID_TipoPago "
        //        sSql = sSql + "     and pg.ID_Poliza = " & vParametros(0)
        //        sSql = sSql + "     and ct.ID_TpoROPC in (2,3) "
        //        sSql = sSql + "     and pg.ID_StaPago in (2,8) "
        //        sSql = sSql + "     and m.Imp_PBenef < 0 "
        //        Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)


        //        if Not rsAux.EOF Then
        //            vParametros(12) = rsAux!Importe
        //        End if

        //        bImprocedencia = True
        //    Else
        //        bImprocedencia = False
        //    End if

        //    cnnConexion.Close
        //    Set cnnConexion = Nothing

        //Exit Function
        //msgerror:
        //    bImprocedencia = False
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //End Function

        ////////Public Function bAlta(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional ByRef vParametros As Variant) As Boolean
        ////////Dim vNParametros(6) As Variant      //Nuevo variant para ejecutar Stored Procedure
        ////////Dim tFecha_IniDer As Date
        ////////Dim tFecha_Proceso As Date
        ////////Dim dProp_Renta As Double
        ////////Dim dProp_Aguin1 As Double //Proporcional de aguinaldo al año de fallecimiento
        ////////Dim dProp_Aguin2 As Double //Proporcional de los años subsecuentes al fallecimiento
        ////////Dim dPagos_Iniciales As Double
        ////////Dim dRentas_Mensuales As Double
        ////////Dim dUltima_Renta As Double
        ////////Dim dPagos_Indebidos As Double
        ////////Dim dDiferencial_Prima As Double
        ////////Dim iDias_PropAntes As Integer
        ////////Dim iDias_PropDesp As Integer
        ////////Dim iSobrevivencia As Integer
        ////////On Error GoTo msgerror
        ////////
        ////////    dProp_Renta = 0
        ////////    dProp_Aguin1 = 0
        ////////    dProp_Aguin2 = 0
        ////////    dRentas_Mensuales = 0
        ////////    dUltima_Renta = 0
        ////////    dPagos_Indebidos = 0
        //////////Traer la historia de la poliza
        ////////    vNParametros(0) = "sp_HistoriaPoliza"   //Nombre del stored procedure
        ////////    vNParametros(1) = vParametros(0)        //ID_Poliza
        ////////    vNParametros(2) = Format(vParametros(5), "yyyy-mm-dd") //Fecha Inicial
        ////////    vNParametros(3) = Format(vParametros(6), "yyyy-mm-dd") //Fecha Final
        ////////    vNParametros(4) = vParametros(1)        //ID_PolBenef
        ////////    vNParametros(5) = vParametros(2)        //ID_Endoso
        ////////    tFecha_IniDer = vParametros(5)
        ////////    tFecha_Proceso = vParametros(6)
        ////////    dDiferencial_Prima = vParametros(4)
        ////////    iSobrevivencia = vParametros(7)
        ////////    iDias_PropAntes = DateDiff("d", "01/01/" & Year(tFecha_IniDer), tFecha_IniDer) + 1
        ////////    iDias_PropDesp = DateDiff("d", tFecha_IniDer, "12/31/" & Year(tFecha_IniDer))
        ////////
        ////////    if EjecutaStoredProcedure(rsErrAdo, gDSN, rsRecord, vNParametros) = True Then
        ////////        rsRecord.MoveFirst
        ////////        Do While Not rsRecord.EOF
        ////////            if iSobrevivencia = 1 Then
        ////////                if rsRecord!Orden = 0 Or rsRecord!Orden = 3 Then
        ////////                    if rsRecord!Tipo = "RENTA" And Month(tFecha_IniDer) = Month(rsRecord!Periodo) And Year(tFecha_IniDer) = Year(rsRecord!Periodo) Then
        ////////                        dProp_Renta = dProp_Renta + Round(rsRecord!Se_Debio, 2)
        ////////                    End if
        ////////                    if rsRecord!Tipo = "AGUINALDO" Then
        ////////                        if Year(tFecha_IniDer) = Year(rsRecord!Periodo) Then
        ////////                            if Year(tFecha_Proceso) = Year(rsRecord!Periodo) Then
        ////////                                dProp_Aguin1 = dProp_Aguin1 + Round(rsRecord!Se_Debio * iDias_PropDesp / 365, 2)
        ////////                            Else
        ////////                                dProp_Aguin1 = dProp_Aguin1 + Round(rsRecord!Se_Pago * iDias_PropAntes / 365, 2) + Round(rsRecord!Se_Debio * iDias_PropDesp / 365, 2)
        ////////                            End if
        ////////                        Else
        ////////                            dProp_Aguin2 = dProp_Aguin2 + Round(rsRecord!Se_Debio, 2)
        ////////                        End if
        ////////                    End if
        ////////                End if
        ////////            End if
        ////////            if rsRecord!Orden = 1 Or rsRecord!Orden = 4 Then //Totales Pension
        ////////                dRentas_Mensuales = dRentas_Mensuales + rsRecord!Diferencia
        ////////                dPagos_Indebidos = dPagos_Indebidos + rsRecord!Pago_Indebido
        ////////            Elseif rsRecord!Orden = 2 Or rsRecord!Orden = 5 Then //Totales Aguinaldo
        ////////                dPagos_Indebidos = dPagos_Indebidos + rsRecord!Pago_Indebido
        ////////                if iSobrevivencia = 0 Then
        ////////                    dProp_Aguin1 = dProp_Aguin1 + rsRecord!Se_Debio
        ////////                End if
        ////////            End if
        ////////            if rsRecord!Orden = 0 Or rsRecord!Orden = 3 Then
        ////////                if Month(tFecha_Proceso) = Month(rsRecord!Periodo) And Year(tFecha_Proceso) = Year(rsRecord!Periodo) And rsRecord!Tipo = "RENTA" Then
        ////////                    dUltima_Renta = dUltima_Renta + rsRecord!Se_Debio
        ////////                End if
        ////////            End if
        ////////            rsRecord.MoveNext
        ////////        Loop
        ////////        bAlta = True
        ////////    Else
        ////////        bAlta = False
        ////////    End if
        ////////
        ////////    //Diferencial de Prima
        ////////    vParametros(12) = dDiferencial_Prima
        ////////    //Pagos Iniciales
        ////////    vParametros(13) = 0 //dPagos_Iniciales(vParametros)
        ////////    //Rentas Mensuales
        ////////    vParametros(14) = dRentas_Mensuales
        ////////    //Aguinaldo
        ////////    vParametros(15) = dProp_Aguin1 + dProp_Aguin2
        ////////    //A favor de la aseguradora
        ////////    vParametros(16) = dDiferencial_Prima + vParametros(13) + dRentas_Mensuales + vParametros(15)
        ////////    //Pagos Indebidos
        ////////    vParametros(17) = Iif(vParametros(3) = 1, dPagos_Indebidos, 0) //vParametros(3) trae el maximo numero de grupos
        ////////    //Total a transferir por el IMSS
        ////////    vParametros(18) = Iif(vParametros(3) = 1, vParametros(16) - dPagos_Indebidos, 0) //vParametros(3) trae el maximo numero de grupos
        ////////    //Total a transferir por la aseguradora
        ////////    vParametros(19) = 0
        ////////    //Total de PI a descontar por la aseguradora
        ////////    vParametros(20) = Iif(vParametros(3) > 1, dPagos_Indebidos, 0) //vParametros(3) trae el maximo numero de grupos
        ////////    //Proporcional de Renta
        ////////    vParametros(21) = dProp_Renta
        ////////    //Proporcional de aguinaldo al año de fallecimiento
        ////////    vParametros(22) = Iif(iSobrevivencia = 0, 0, dProp_Aguin1)
        ////////    //Proporcional de los años subsecuentes al fallecimiento
        ////////    vParametros(23) = Iif(iSobrevivencia = 0, 0, dProp_Aguin1 + dProp_Aguin2)
        ////////
        ////////Exit Function
        ////////msgerror:
        ////////    bAlta = False
        ////////    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        ////////    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        ////////    Set cnnConexion = Nothing
        ////////End Function

        //private Function dPagos_Iniciales(Optional ByRef vParametros As Variant) As Double
        //    dPagos_Iniciales = 0
        //End Function

        //Public Function iConsulta_PolizaImagen(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    rsRecord.CursorLocation = adUseClient
        //    rsRecord.CursorType = adOpenKeyset
        //    rsRecord.LockType = adLockOptimistic


        //    if vParametros(0) = 0 Then //Datos Generales
        //        sSql = "select  Poliza = case p.ID_Empresa when 1 then p.Fol_Poliza else p.Fol_Poliza + 90000000 end, i.Fecha_IniDer, i.Fecha_IniVig, Fecha_Calculo, Fecha_CalculoInc, i.Fecha_ABase, i.Pje_Valuacion, i.Pje_AyudaAsist, i.Salario_RT, i.Salario_IV, Fecha_NactoAseg = i.Fecha_Nacto, Sexo = Abv_Sexo "
        //        sSql = sSql + "from  Img_Polizas i, Polizas p, Cat_Sexos cs "
        //        sSql = sSql + "where i.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + "  and i.ID_Sexo = cs.ID_Sexo "
        //        sSql = sSql + "  and p.Fol_Poliza = " & vParametros(1)
        //        sSql = sSql + "  and ID_Empresa = " & vParametros(2)
        //    Elseif vParametros(0) = 1 Then //Datos de Beneficiarios
        //        sSql = "select  Poliza = case p.ID_Empresa when 1 then p.Fol_Poliza else p.Fol_Poliza + 90000000 end, Num_Benef, BenefIni, Abv_Parentesco, Cve_Orfandad, Abv_Sexo, Invalidez, Sta_Incremento = case when i.ID_StaIncremento = 0 then //NO// else //SI// end, i.Fecha_Nacto "
        //        sSql = sSql + "from     Img_PolBenefs i,  "
        //        sSql = sSql + " Polizas p, Cat_BenefIni cb, Cat_Parentescos cp, Cat_Orfandades co, Cat_Sexos cs, Cat_Invalidez ci "
        //        sSql = sSql + "where    i.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + " and i.ID_BenefIni = cb.ID_BenefIni "
        //        sSql = sSql + " and i.ID_Parentesco = cp.ID_Parentesco "
        //        sSql = sSql + " and i.ID_Orfandad = co.ID_Orfandad "
        //        sSql = sSql + " and i.ID_Sexo = cs.ID_Sexo "
        //        sSql = sSql + " and i.ID_Invalidez = ci.ID_Invalidez "
        //        sSql = sSql + " and p.Fol_Poliza = " & vParametros(1)
        //        sSql = sSql + " and ID_Empresa = " & vParametros(2)
        //        sSql = sSql + " order by Num_Benef "
        //    Elseif vParametros(0) = 2 Then //Datos de Beneficio Adicional
        //        sSql = "select  Poliza = case p.ID_Empresa when 1 then p.Fol_Poliza else p.Fol_Poliza + 90000000 end, Abv_Beneficio, Proporcional, Opcion, Cuantia_AyEscolar, Pago_Art14 "
        //        sSql = sSql + "from  Img_PolBASelec i, "
        //        sSql = sSql + " Polizas p, "
        //        sSql = sSql + " Cat_Beneficios cb "
        //        sSql = sSql + "where    i.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + " and i.ID_Beneficio = cb.ID_Beneficio "
        //        ////sSql = sSql + " and ( ID_TipoBA in (1,2,3,4) or i.ID_Beneficio = 9 ) "
        //        sSql = sSql + " and p.Fol_Poliza = " & vParametros(1)
        //        sSql = sSql + " and ID_Empresa = " & vParametros(2)
        //    End if


        //    rsRecord.Open sSql, cnnConexion
        //    if rsRecord.EOF Then
        //        iConsulta_PolizaImagen = NoHayDatos
        //    Else
        //        iConsulta_PolizaImagen = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iConsulta_PolizaImagen = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        public ModRecordset.TipoResultado iDatosSUC(ref Recordset rsErrAdo, ref string gDSN, ref Recordset rsRecord, object[] vParametros)
        {

            ModRecordset.TipoResultado _iDatosSUC;

            try
            {
                //On Error GoTo msgerror

                Connection cnnConexion = new Connection();
                cnnConexion.Open(gDSN);
                rsRecord = new Recordset();
                //rsRecord.CursorLocation = adUseClient;
                //rsRecord.CursorType = adOpenKeyset;
                //rsRecord.LockType = adLockOptimistic;
                string sSql = "";
                

                if (Convert.ToInt32(vParametros[0]) == 0)  //Datos de Asegurado
                {
                    sSql = "select   Texto = isnull(convert(varchar(4), year(o.Fecha_ABase)) + right(//00// + convert(varchar(2), month(o.Fecha_ABase)),2) +  right(//00// + convert(varchar(2), day(o.Fecha_ABase)),2), //00010101//) ";
                    sSql = sSql + " + isnull(convert(varchar(4), year(o.Fecha_DEA)) + right(//00// + convert(varchar(2), month(o.Fecha_DEA)),2) +  right(//00// + convert(varchar(2), day(o.Fecha_DEA)),2), //00010101//) ";
                    sSql = sSql + " + isnull(Tpo_Registro, // //) ";
                    sSql = sSql + " + left(rtrim(isnull(o.ApPAseg, ////))+// //+rtrim(isnull(o.ApMAseg, ////))+// //+rtrim(isnull(o.NomAseg, ////)) + space(60), 60) ";
                    sSql = sSql + " + right(//00000000000// + p.Num_SegSocial, 11) ";
                    sSql = sSql + " + right(//00// + o.Num_Solicitud, 2) ";
                    sSql = sSql + " + isnull(convert(varchar(4), year(p.Fecha_Nacto)) + right(//00// + convert(varchar(2), month(p.Fecha_Nacto)),2) +  right(//00// + convert(varchar(2), day(p.Fecha_Nacto)),2), //00010101//) ";
                    sSql = sSql + " + convert(char(1), Abv_Sexo) ";
                    sSql = sSql + " + right(space(18) + isnull(p.CURP, ////), 18) ";
                    sSql = sSql + " + right(//00// + p.Deleg, 2) ";
                    sSql = sSql + " + right(//000// + p.Subdeleg, 3) ";
                    sSql = sSql + " + isnull(convert(varchar(4), year(o.Fecha_BajaRO)) + right(//00// + convert(varchar(2), month(o.Fecha_BajaRO)),2) +  right(//00// + convert(varchar(2), day(o.Fecha_BajaRO)),2), //00010101//) ";
                    sSql = sSql + " + isnull(convert(varchar(4), year(p.Fecha_IniDer)) + right(//00// + convert(varchar(2), month(p.Fecha_IniDer)),2) +  right(//00// + convert(varchar(2), day(p.Fecha_IniDer)),2), //00010101//) ";
                    sSql = sSql + " + right(//00000// + convert(varchar(5),case when p.Pje_Valuacion in (0,100) then //// else //0// + replace(convert(varchar(5),p.Pje_Valuacion),//.//,////) end), 5) ";
                    sSql = sSql + " + case p.ID_Ramo when 1 then //RT// else //IM// end ";
                    sSql = sSql + " + case p.ID_Pension when 1 then //IN// when 2 then //VI// when 3 then //VO// when 4 then //OR// when 5 then //AS// when 6 then //IP// end ";
                    sSql = sSql + " + isnull(right(//0000// + convert(varchar(5), p.Semanas_Cot), 4), //0000//) ";
                    sSql = sSql + " + right(//0000000000000// + replace(convert(varchar,p.Salario_RT), //.//, ////), 13) ";
                    sSql = sSql + " + right(//0000000000000// + replace(convert(varchar,p.Salario_IV), //.//, ////), 13) ";
                    sSql = sSql + " + right(//0000000000000// + replace(convert(varchar,p.Cuantia_BaseFC), //.//, ////), 13) ";
                    sSql = sSql + " + right(//00000// + convert(varchar(5),case when p.Pje_Ayuda = 0 then //// else //0// + replace(convert(varchar(5),p.Pje_Ayuda),//.//,////) end), 5) ";
                    sSql = sSql + " + right(//0000000000000// + replace(convert(varchar,p.Pension_MensualFC), //.//, ////), 13) ";
                    sSql = sSql + " + left(rtrim(o.ApPSolic)+// //+rtrim(o.ApMSolic)+// //+rtrim(o.NomSolic) + space(60), 60) ";
                    sSql = sSql + " + convert(varchar(4), year(isnull(p.Fecha_Solic,p.Fecha_IniDer))) + right(//00// + convert(varchar(2), month(isnull(p.Fecha_Solic,p.Fecha_IniDer))),2) +  right(//00// + convert(varchar(2), day(isnull(p.Fecha_Solic,p.Fecha_IniDer))),2) ";
                    sSql = sSql + " + left(isnull(p.Domicilio, ////) + space(60), 60) ";
                    if (vParametros[2].ToString() == "9999-12-31")
                    {
                        sSql = sSql + " + convert(varchar(4), year(Fecha_IniVig)) + right(//00// + convert(varchar(2), month(Fecha_IniVig)),2) +  right(//00// + convert(varchar(2), day(Fecha_IniVig)),2) ";
                    }
                    else
                    { 
                        sSql = sSql + " + convert(varchar(4), year(//" + vParametros[2] + "//)) + right(//00// + convert(varchar(2), month(//" + vParametros[2] + "//)),2) +  right(//00// + convert(varchar(2), day(//" + vParametros[2] + "//)),2) ";
                    }
                    sSql = sSql + " + //0000000000000// ";
                    sSql = sSql + " + //0000000000000// ";
                    sSql = sSql + " + right(//0000000000000// + replace(convert(varchar,p.Pension_MensualFC), //.//, ////), 13) ";

                    // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
                    // Juan Martínez Díaz
                    // 2022-07-11
                    // [Operadores *=, =*]
                    sSql = sSql + " from Polizas p ";
                    sSql = sSql + " inner join Ofertas o on p.ID_Oferta = o.ID_Oferta ";
                    sSql = sSql + " left join Resoluciones r on o.ID_Oferta = r.ID_Oferta ";
                    sSql = sSql + " inner join Cat_Ramos cr on p.ID_Ramo = cr.ID_Ramo ";
                    sSql = sSql + " inner join Cat_Pensiones cp on p.ID_Pension = cp.ID_Pension ";
                    sSql = sSql + " inner join Cat_Sexos cs on p.ID_Sexo = cs.ID_Sexo ";

                    // Reemplaza

                    //        sSql = sSql + "from Polizas p,  "
                    //        sSql = sSql + " Ofertas o,  "
                    //        sSql = sSql + " Cat_Ramos cr,  "
                    //        sSql = sSql + " Cat_Pensiones cp,  "
                    //        sSql = sSql + " Cat_Sexos cs, "
                    //        sSql = sSql + " Resoluciones r "
                    //        sSql = sSql + "where    p.ID_Oferta = o.ID_Oferta "
                    //        sSql = sSql + " and o.ID_Oferta *= r.ID_Oferta "
                    //        sSql = sSql + " and p.ID_Ramo = cr.ID_Ramo "
                    //        sSql = sSql + " and p.ID_Pension = cp.ID_Pension "
                    //        sSql = sSql + " and p.ID_Sexo = cs.ID_Sexo "

                    // [Operadores *=, =*]

                    sSql = sSql + " and p.ID_StaPoliza = 2 ";
                    sSql = sSql + " and p.Fol_Poliza in (" + vParametros[1] + ")";
                    if ( Convert.ToInt32(vParametros[3]) != 99)
                    {
                        sSql = sSql + " and p.ID_Empresa = " + vParametros[3];
                    }

                    sSql = sSql + "order by p.ID_Empresa, p.Num_SegSocial ";
                }
                else if ( Convert.ToInt32(vParametros[0]) == 1) //Datos de Beneficiarios
                {
                    sSql = "select   Texto = right(//00000000000// + p.Num_SegSocial, 11) ";
                    sSql = sSql + " + right(//00// + o.Num_Solicitud, 2) ";
                    sSql = sSql + " + left(rtrim(isnull(ApP_Benef, ////))+// //+rtrim(isnull(ApM_Benef, ////))+// //+rtrim(isnull(Nom_Benef, ////)) + space(60), 60) ";
                    sSql = sSql + " + Cve_Parentesco ";
                    sSql = sSql + " + Abv_Sexo ";
                    sSql = sSql + " + convert(varchar(4), year(pb.Fecha_Nacto)) + right(//00// + convert(varchar(2), month(pb.Fecha_Nacto)),2) +  right(//00// + convert(varchar(2), day(pb.Fecha_Nacto)),2) ";
                    sSql = sSql + " + convert(varchar(4), year(pb.Fecha_IniDer)) + right(//00// + convert(varchar(2), month(pb.Fecha_IniDer)),2) +  right(//00// + convert(varchar(2), day(pb.Fecha_IniDer)),2) ";
                    sSql = sSql + " + case when ID_Invalidez = 2 or pb.ID_Parentesco <> 3 then //00010101// else convert(varchar(4), year(pb.Fecha_Vento)) + right(//00// + convert(varchar(2), month(pb.Fecha_Vento)),2) +  right(//00// + convert(varchar(2), day(pb.Fecha_Vento)),2) end ";
                    sSql = sSql + " + case when Cve_Orfandad = //-// then //N// else Cve_Orfandad end ";
                    sSql = sSql + "from Polizas p,  ";
                    sSql = sSql + " Pol_Benefs pb, ";
                    sSql = sSql + " Ofertas o,  ";
                    sSql = sSql + " Cat_Sexos cs, ";
                    sSql = sSql + " Cat_Parentescos cp, ";
                    sSql = sSql + " Cat_Orfandades co ";
                    sSql = sSql + "where    p.ID_Poliza = pb.ID_Poliza ";
                    sSql = sSql + " and pb.ID_Sexo = cs.ID_Sexo ";
                    sSql = sSql + " and pb.ID_Parentesco = cp.ID_Parentesco ";
                    sSql = sSql + " and pb.ID_Orfandad = co.ID_Orfandad ";
                    sSql = sSql + " and p.ID_Oferta = o.ID_Oferta ";
                    sSql = sSql + " and p.ID_StaPoliza = 2 ";
                    sSql = sSql + " and pb.ID_StaPolBenef = 1 ";
                    sSql = sSql + " and pb.ID_Parentesco not in (1,5) ";
                    sSql = sSql + " and p.Fol_Poliza in (" + vParametros[1] + ")";
                    if (Convert.ToInt32(vParametros[3]) != 99)
                    {
                        sSql = sSql + " and p.ID_Empresa = " + vParametros[3];
                    }

                    sSql = sSql + "order by p.ID_Empresa, p.Num_SegSocial ";
                }


                rsRecord.Open(sSql, cnnConexion);

                if (rsRecord.EOF)
                    _iDatosSUC = ModRecordset.TipoResultado.NoHayDatos;
                else
                    _iDatosSUC = ModRecordset.TipoResultado.DatosOK;


                rsRecord.ActiveConnection = null;
                cnnConexion.Close();
                //cnnConexion = PrivilegeNotHeldException;

                return _iDatosSUC;
            }
            catch (Exception)
            {
                _iDatosSUC = ModRecordset.TipoResultado.ExisteError;
                //sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
                //Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
                //Set cnnConexion = Nothing
                return _iDatosSUC;
            }

        }

        //Public Function iConsultaEndosos(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, ByRef rsRecord As ADODB.Recordset, lID_Poliza As Long) As Integer
        //On Error GoTo msgerror

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    rsRecord.CursorLocation = adUseClient

        //    //SDC 2007-05-04 Para que no se pueda imprimir una Orden de Trabajo cuando exista un endoso
        //    //de Sobrevivencia y no se haya capturado la resolucion en el endoso de fallecimiento del asegurado.
        //    //Resolucion_SS = 0 Existe endoso de SS pero no se ha capturado
        //    //Resolucion_SS = 1 Existe endoso de SS y está capturada
        //    //Resolucion_SS = 2 No Existe endoso de SS
        //    sSql = "select ID_Empresa, Fol_Poliza, e.ID_HisEndoso, Endoso, Nombre = convert(varchar(2), Num_Benef) + //-// + Nom_Benef + // // + ApP_Benef + // // + ApM_Benef, d.ID_StaResolEndoso, Sta_ResolEndoso, d.Num_Resolucion, d.Fecha_Resol, e.Fecha_System, "
        //    sSql = sSql + vbCrLf & "        Resolucion_SS = isnull((select case when Num_Resolucion is null then 0 else 1 end from Pol_Benefs pb, Endosos e, DEndosos d where pb.ID_Poliza = p.ID_Poliza and pb.ID_Poliza = e.ID_Poliza and pb.ID_PolBenef = e.ID_PolBenef and e.ID_HisEndoso = d.ID_HisEndoso and pb.ID_Parentesco = 1 and e.ID_Endoso = 8), 2) "
        //    sSql = sSql + vbCrLf & "from    Polizas p, Endosos e left join DEndosos d on e.ID_HisEndoso = d.ID_HisEndoso left join Cat_StaResolEndoso cs  on d.ID_StaResolEndoso = cs.ID_StaResolEndoso left join Pol_Benefs pb on e.ID_PolBenef = pb.ID_PolBenef, "
        //    sSql = sSql + vbCrLf & "        Cat_Endosos c "
        //    sSql = sSql + vbCrLf & "where   p.ID_Poliza = e.ID_Poliza  "
        //    sSql = sSql + vbCrLf & "        and e.ID_Endoso = c.ID_Endoso  "
        //    sSql = sSql + vbCrLf & "        and e.ID_Poliza = " & lID_Poliza & " "
        //    sSql = sSql + vbCrLf & "order by e.Fecha_System "


        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic
        //    if rsRecord.EOF Then
        //        iConsultaEndosos = NoHayDatos
        //    Else
        //        iConsultaEndosos = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iConsultaEndosos = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        //Public Function bCapturaResolucion(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsDatosUser As ADODB.Recordset, Optional vParametros As Variant) As Boolean
        //Dim bBool As Boolean
        //Dim rsAux As ADODB.Recordset
        //On Error GoTo msgerror

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN


        //    Set ctxObject = GetObjectContext()
        //    bBool = True

        //    //Actualizamos el endoso
        //    sSql = "update d set ID_StaResolEndoso = 2, Fecha_Status = dbo.getdateusr(), "
        //    sSql = sSql + vbCrLf & "    Num_Resolucion = //" & vParametros(1) & "//, "
        //    sSql = sSql + vbCrLf & "    Fecha_Resol = //" & vParametros(2) & "// "
        //    sSql = sSql + vbCrLf & "from DEndosos d "
        //    sSql = sSql + vbCrLf & "where ID_HisEndoso = " & vParametros(0)
        //    cnnConexion.Execute sSql

        //    //Activar los pagos al mes de fallecimiento
        //    sSql = "update p set ID_StaPagoEndoso = 2, Fecha_StaPagoEndoso = dbo.getdateusr() "
        //    sSql = sSql + vbCrLf & "from Pagos_Endosos p "
        //    sSql = sSql + vbCrLf & "where ID_HisEndoso = " & vParametros(0)
        //    sSql = sSql + vbCrLf & "    and ID_StaPagoEndoso = 1 "
        //    cnnConexion.Execute sSql

        //    //SDC 2006-09-15 Insertar los registros de los componentes para
        //    //solicitar recursos de artículo 14 2004.
        //    Set rsAux = New ADODB.Recordset

        //    sSql = "select Fecha_Aplicacion from Endosos e, Pol_Benefs pb where e.ID_Poliza = pb.ID_Poliza and e.ID_PolBenef = pb.ID_PolBenef and pb.ID_Parentesco = 1 and e.ID_HisEndoso = " & vParametros(0)


        //    rsAux.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic

        //    if Not rsAux.EOF Then
        //        sSql = "execute sp_InsertaIncrementos " & vParametros(3) & ", //" & Format(Now, "yyyy-mm-dd") & "//, //" & Format(rsAux!Fecha_Aplicacion, "yyyy-mm-dd") & "//, 0, 0, 0, 3 "
        //        cnnConexion.Execute sSql
        //        rsAux.Close
        //    End if

        //    Set rsErrAdo = Nothing
        //    sSql = "Captura de Resolución de Endoso " & vParametros(0) & " Resolución " & vParametros(1) & " Fecha de Resolución " & vParametros(2)
        //    InsertaBitacora gDSN, sSql, rsDatosUser, rsErrAdo
        //    if Not rsErrAdo Is Nothing Then
        //        bBool = False
        //        GoTo msgerror
        //    End if


        //    ctxObject.SetComplete
        //    cnnConexion.Close
        //    Set ctxObject = Nothing
        //    Set cnnConexion = Nothing
        //    bCapturaResolucion = True

        //Exit Function
        //msgerror:
        //    bCapturaResolucion = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    if bBool Then
        //        Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    End if
        //    ctxObject.SetAbort
        //    cnnConexion.Close
        //    Set ctxObject = Nothing
        //    Set cnnConexion = Nothing
        //End Function

        //private Function tUltimoDiaDeMes(ByRef Fecha As Date) As Date
        //    On Error Resume Next
        //    tUltimoDiaDeMes = Format(Fecha, "yyyy-mm-") & "31"
        //    if Err.Number = 13 Then
        //        On Error Resume Next
        //        tUltimoDiaDeMes = Format(Fecha, "yyyy-mm-") & "30"
        //        if Err.Number = 13 Then
        //            On Error Resume Next
        //            tUltimoDiaDeMes = Format(Fecha, "yyyy-mm-") & "29"
        //            if Err.Number = 13 Then
        //                tUltimoDiaDeMes = Format(Fecha, "yyyy-mm-") & "28"
        //            End if
        //        End if
        //    End if
        //End Function

        //private Function dMaximo(ByRef dUno As Double, ByRef dDos As Double) As Double
        //    if dUno > dDos Then
        //        dMaximo = dUno
        //    Else
        //        dMaximo = dDos
        //    End if
        //End Function

        //Public Function bInsertaEnc(ByRef grsErrADO As ADODB.Recordset, ByRef gDSN As String, vParams As Variant) As Boolean
        //Dim sSql As String

        //On Error GoTo msgerror
        //    sSql = "exec sp_InsertaEncXAltaBenef "
        //    sSql = sSql + vParams(0)
        //    sSql = sSql + ", //" & vParams(1) & "//"

        //    bInsertaEnc = ExecSql(gDSN, sSql, grsErrADO)

        //Exit Function
        //msgerror:
        //    bInsertaEnc = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing

        //End Function


        //Public Function ObtenProrrogas(ByRef rsData As ADODB.Recordset, gDSN As String, grsErrADO As ADODB.Recordset, vParametros As Variant) As Integer
        //Dim sSql As String
        //Dim errVB As String
        //On Error GoTo msgerror

        //    //SDC 2006-09-20 Para traer el nombre del asegurado, la fecha de nacimiento y el grupo familiar
        //    //SDC 2007-06-08 Para traer el tipo de Prórroga del catálogo Cat_TpoComprobacionEstudios
        //    sSql = "select CE.ID_Comprobacion, E.Abv_Empresa, P.Fol_Poliza ,P.Num_SegSocial, PB.Num_Benef, PB.ID_Grupo, "
        //    sSql = sSql + "     Nombre_Aseg = isnull(P.ApP_Aseg, ////) + // // + isnull(P.ApM_Aseg, ////) + // // + isnull(P.Nom_Aseg, ////),  "
        //    sSql = sSql + "     Nombre = isnull(PB.ApP_Benef, ////) + // // + isnull(ApM_Benef, ////) + // // + isnull(PB.Nom_Benef, ////),  "
        //    sSql = sSql + "     PB.Fecha_Nacto, CE.Fecha_Inicial ,  CE.Fecha_Final , CE.Fecha_Server, C.Tpo_ComprobacionEstudios "
        //    sSql = sSql + "from Comprobacion_Estudios CE, Pol_Benefs PB , Polizas P, Empresas E, Cat_TpoComprobacionEstudios C "
        //    sSql = sSql + "where PB.ID_PolBenef = CE.ID_PolBenef "
        //    sSql = sSql + "     and P.ID_Poliza = PB.ID_Poliza "
        //    sSql = sSql + "     and P.ID_Empresa = E.ID_Empresa "
        //    sSql = sSql + "     and P.ID_InstitucionSS in (1,2) " //CEFB 08/09/2014
        //    sSql = sSql + "     and CE.ID_TpoComprobacionEstudios = C.ID_TpoComprobacionEstudios "
        //    if vParametros(0) <> 99 Then
        //        sSql = sSql + "and P.ID_Empresa = " & vParametros(0)
        //    End if
        //    if vParametros(1) <> Empty Then
        //        sSql = sSql + " and P.Fol_Poliza = " & vParametros(1)
        //    End if
        //    if vParametros(2) <> Empty Then
        //        sSql = sSql + " and P.Num_SegSocial = " & vParametros(2)
        //    End if
        //    if vParametros(3) <> Empty Then
        //        sSql = sSql + " and CE.Fecha_Server >= //" & vParametros(3) & "// "
        //    End if
        //    if vParametros(4) <> Empty Then
        //        sSql = sSql + " and CE.Fecha_Server < dateadd(dd, 1, //" & vParametros(4) & "//) "
        //    End if
        //    sSql = sSql + " order by CE.Fecha_Server "

        //    ObtenProrrogas = EjecutaSql(rsData, gDSN, sSql, grsErrADO)

        //Exit Function


        //msgerror:
        //    errVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set grsErrADO = ErroresDLL(Nothing, errVB)
        //End Function

        ////_______________________________________________________________
        ////Título: Consulta los Criterios de Busqueda de Contabilidades de ROPC por Endoso.
        ////Clase: clsEndososCET.iConsulta_Contabilidad
        ////Versión: 1.0
        ////Fecha:    22/08/2006
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////   Consulta los endosos con remesas dependiendo de los criterios de búsqueda
        ////_______________________________________________________________
        //Public Function iConsulta_Contabilidad(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    Set rsErrAdo = Nothing
        //    rsRecord.CursorLocation = adUseClient


        //    sSql = "select   distinct emp.Abv_Empresa, "
        //    sSql = sSql + "     p.Fol_Poliza, "
        //    sSql = sSql + "     e.ID_HisEndoso, "
        //    sSql = sSql + "     r.Fol_Remesa, "
        //    sSql = sSql + "     e.Fecha_System, "
        //    sSql = sSql + "     ce.Endoso "
        //    sSql = sSql + "from Polizas p inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        //    sSql = sSql + "     inner join Empresas emp on p.ID_Empresa = emp.ID_Empresa "
        //    sSql = sSql + "     inner join Cat_Endosos ce on e.ID_Endoso = ce.ID_Endoso "
        //    sSql = sSql + "     inner join Remesas r on r.Fecha_System = e.Fecha_System and r.ID_TpoRem = 27 and p.ID_Empresa = r.ID_Empresa "
        //    sSql = sSql + "     inner join MRemesas mr on r.ID_Remesa = mr.ID_Remesa and convert(int, mr.Fol_Docto) = p.ID_Poliza and convert(int, mr.Folio) = e.ID_HisEndoso and mr.ID_TpoDocRem = 6 "
        //    if vParametros(0) <> 99 Then
        //        sSql = sSql + "where    p.ID_Empresa = " & vParametros(0)
        //    Else
        //        sSql = sSql + "where    p.ID_Empresa >= 0 "
        //    End if
        //    if vParametros(1) <> Empty Then
        //        sSql = sSql + "     and p.Fol_Poliza = " & vParametros(1)
        //    End if
        //    if Not IsNull(vParametros(2)) Then
        //        sSql = sSql + "     and e.Fecha_System >= //" & Format(vParametros(2), "yyyy-mm-dd") & "// "
        //    End if
        //    if Not IsNull(vParametros(3)) Then
        //        sSql = sSql + "     and e.Fecha_System <= //" & Format(vParametros(3), "yyyy-mm-dd") & "// "
        //    End if
        //    sSql = sSql + "order by e.ID_HisEndoso "

        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic


        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iConsulta_Contabilidad = NoHayDatos
        //    Else
        //        iConsulta_Contabilidad = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iConsulta_Contabilidad = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        ////_______________________________________________________________
        ////Título: Consulta los Detalles de Contabilidades de ROPC por Endoso.
        ////Clase: clsEndososCET.iConsulta_DetalleContabilidad
        ////Versión: 1.0
        ////Fecha:    22/08/2006
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////   Consulta los registros de Contabilidad_Endosos para un endoso dado
        ////_______________________________________________________________
        //Public Function iConsulta_DetalleContabilidad(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //Dim rsAux As ADODB.Recordset
        //On Error GoTo msgerror

        //    //SDC 2007-07-11 Para traer los datos según se utilizaba con ID_Concpeto o ahora con ID_ConceptoContable, ID_TpoROPC
        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    Set rsAux = New ADODB.Recordset
        //    Set rsErrAdo = Nothing
        //    rsRecord.CursorLocation = adUseClient


        //    sSql = "select ID_ConceptoContable = min(ID_ConceptoContable) from Contabilidad_Endosos where ID_HisEndoso = " & vParametros(0)
        //    rsAux.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic

        //    if rsAux("ID_ConceptoContable") = 0 Then
        //        sSql = "select   distinct ce.ID_HisEndoso, Mov_Conta = case when ce.ID_MovConta = 1 then //C// else //A// end, "
        //        sSql = sSql + "     Nombre = case when ce.ID_MovRem = 10 then //ROPC // + cc.Concepto else ccmr.Nombre end, "
        //        sSql = sSql + "     cra.Cve_Ramo, cri.Abv_Riesgo, Concepto = cc.Abv_Concepto, Abv_TpoROPC = null, ce.Monto_Mov, ce.ID_Concepto, ce.ID_MovConta "
        //        sSql = sSql + "from Contabilidad_Endosos ce inner join Cat_Beneficios cb on ce.ID_Concepto = cb.ID_Concepto "
        //        sSql = sSql + "     inner join Cat_ConMovRem ccmr on cb.ID_ConceptoContable = ccmr.ID_ConceptoContable and ce.ID_MovRem = ccmr.ID_MovRem and ce.ID_NivMovRem = ccmr.ID_NivMovRem "
        //        sSql = sSql + "     inner join Cat_Conceptos cc on ce.ID_Concepto = cc.ID_Concepto "
        //        sSql = sSql + "     inner join Cat_RamoContable crc on ce.ID_RamoContable = crc.ID_RamoContable "
        //        sSql = sSql + "     inner join Cat_Ramos cra on crc.ID_Ramo = cra.ID_Ramo "
        //        sSql = sSql + "     inner join Cat_Riesgos cri on crc.ID_Riesgo = cri.ID_Riesgo "
        //        sSql = sSql + "where    ce.ID_HisEndoso = " & vParametros(0)
        //        sSql = sSql + "order by ce.ID_Concepto, ce.ID_MovConta "
        //    Else
        //        sSql = "select   ce.ID_HisEndoso, Mov_Conta = case when ce.ID_MovConta = 1 then //C// else //A// end, ccmr.Nombre, cra.Cve_Ramo, cri.Abv_Riesgo, Concepto = cc.Abv_ConceptoContable, ctr.Abv_TpoROPC, ce.Monto_Mov "
        //        sSql = sSql + "from Contabilidad_Endosos ce inner join Cat_ConMovRem ccmr on ce.ID_ConceptoContable = ccmr.ID_ConceptoContable and ce.ID_TpoROPC = ccmr.ID_TpoROPC and ce.ID_MovRem = ccmr.ID_MovRem and ce.ID_NivMovRem = ccmr.ID_NivMovRem "
        //        sSql = sSql + "     inner join Cat_ConceptoContable cc on ce.ID_ConceptoContable = cc.ID_ConceptoContable "
        //        sSql = sSql + "     inner join Cat_TpoROPC ctr on ce.ID_TpoROPC = ctr.ID_TpoROPC "
        //        sSql = sSql + "     inner join Cat_RamoContable crc on ce.ID_RamoContable = crc.ID_RamoContable "
        //        sSql = sSql + "     inner join Cat_Ramos cra on crc.ID_Ramo = cra.ID_Ramo "
        //        sSql = sSql + "     inner join Cat_Riesgos cri on crc.ID_Riesgo = cri.ID_Riesgo "
        //        sSql = sSql + "where    ce.ID_HisEndoso = " & vParametros(0)
        //        sSql = sSql + "order by ce.ID_ConceptoContable, ce.ID_MovConta "
        //    End if


        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic


        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iConsulta_DetalleContabilidad = NoHayDatos
        //    Else
        //        iConsulta_DetalleContabilidad = DatosOK
        //    End if

        //    rsAux.Close
        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iConsulta_DetalleContabilidad = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        ////_______________________________________________________________
        ////Título: Consulta en forma de PgoNom los Detalles de Contabilidades de ROPC por Endoso.
        ////Clase: clsEndososCET.iReporte_DetalleContabilidad
        ////Versión: 1.0
        ////Fecha:    22/08/2006
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////   Consulta los registros de Contabilidad_Endosos segun los criterios de
        ////   busqueda en forma de columnas
        ////_______________________________________________________________
        //Public Function iReporte_DetalleContabilidad(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror

        //    //SDC 2007-07-11 Para traer los datos según se utilizaba con ID_Concpeto o ahora con ID_ConceptoContable, ID_TpoROPC
        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    Set rsErrAdo = Nothing
        //    rsRecord.CursorLocation = adUseClient

        //    //Movimientos anteriores con ID_Concepto
        //    sSql = "select Empresa = emp.Abv_Empresa, Remesa = r.Fol_Remesa, Fecha = convert(varchar(10), e.Fecha_System, 120), Poliza = p.Fol_Poliza, Ramo = cra.Cve_Ramo, Tipo = cri.Abv_Riesgo, Endoso = e.ID_HisEndoso, Cargo_Abono = case when ce.ID_MovConta = 1 then //C// else //A// end, cmr.Nombre_Cuenta, ROPC = ctr.Abv_TpoROPC, "
        //    sSql = sSql + vbCrLf & "     Basicos = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_Concepto = 3), 0), "
        //    sSql = sSql + vbCrLf & "     Aguinaldo = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_Concepto = 6), 0), "
        //    sSql = sSql + vbCrLf & "     Pension_Adicional = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_Concepto = 25), 0), "
        //    sSql = sSql + vbCrLf & "     Aguinaldo_Adicional = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_Concepto = 26), 0), "
        //    sSql = sSql + vbCrLf & "     Ayuda_Escolar = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_Concepto = 27), 0), "
        //    sSql = sSql + vbCrLf & "     BAMI = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_Concepto = 28), 0), "
        //    sSql = sSql + vbCrLf & "     BAMI_Aguinaldo = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_Concepto = 34), 0) "
        //    sSql = sSql + vbCrLf & "from Polizas p inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        //    sSql = sSql + vbCrLf & "     inner join Empresas emp on p.ID_Empresa = emp.ID_Empresa "
        //    sSql = sSql + vbCrLf & "     inner join Remesas r on r.Fecha_System = e.Fecha_System and ID_TpoRem = 27 and p.ID_Empresa = r.ID_Empresa "
        //    sSql = sSql + vbCrLf & "     inner join MRemesas mr on r.ID_Remesa = mr.ID_Remesa and Fol_Docto = p.ID_Poliza and Folio = e.ID_HisEndoso "
        //    sSql = sSql + vbCrLf & "     inner join Contabilidad_Endosos ce on ce.ID_HisEndoso = e.ID_HisEndoso "
        //    sSql = sSql + vbCrLf & "     inner join Cat_MovRem cmr on ce.ID_MovRem = cmr.ID_MovRem and ce.ID_NivMovRem = cmr.ID_NivMovRem "
        //    sSql = sSql + vbCrLf & "     inner join Cat_RamoContable crc on ce.ID_RamoContable = crc.ID_RamoContable "
        //    sSql = sSql + vbCrLf & "     inner join Cat_Ramos cra on crc.ID_Ramo = cra.ID_Ramo "
        //    sSql = sSql + vbCrLf & "     inner join Cat_Riesgos cri on crc.ID_Riesgo = cri.ID_Riesgo "
        //    sSql = sSql + vbCrLf & "     inner join Cat_TpoROPC ctr on ce.ID_TpoROPC = ctr.ID_TpoROPC "
        //    if vParametros(0) <> 99 Then
        //        sSql = sSql + vbCrLf & "where    p.ID_Empresa = " & vParametros(0)
        //    Else
        //        sSql = sSql + vbCrLf & "where    p.ID_Empresa >= 0 "
        //    End if
        //    if vParametros(1) <> Empty Then
        //        sSql = sSql + vbCrLf & "     and p.Fol_Poliza = " & vParametros(1)
        //    End if
        //    if Not IsNull(vParametros(2)) Then
        //        sSql = sSql + vbCrLf & "     and e.Fecha_System >= //" & Format(vParametros(2), "yyyy-mm-dd") & "// "
        //    End if
        //    if Not IsNull(vParametros(3)) Then
        //        sSql = sSql + vbCrLf & "     and e.Fecha_System <= //" & Format(vParametros(3), "yyyy-mm-dd") & "// "
        //    End if
        //    sSql = sSql + vbCrLf & "     and ce.ID_ConceptoContable = 0 "
        //    sSql = sSql + vbCrLf & "group by emp.Abv_Empresa, r.Fol_Remesa, e.Fecha_System, p.Fol_Poliza, e.ID_HisEndoso, cmr.Nombre_Cuenta, ce.ID_MovConta, ce.ID_MovRem, ce.ID_NivMovRem, ce.ID_RamoContable, cra.Cve_Ramo, cri.Abv_Riesgo, ctr.Abv_TpoROPC "

        //    sSql = sSql + vbCrLf & "union "

        //    //Movimientos nuevos con ID_ConceptoContable y ID_TpoROPC
        //    sSql = sSql + vbCrLf & "select Empresa = emp.Abv_Empresa, Remesa = r.Fol_Remesa, Fecha = convert(varchar(10), e.Fecha_System, 120), Poliza = p.Fol_Poliza, Ramo = cra.Cve_Ramo, Tipo = cri.Abv_Riesgo, Endoso = e.ID_HisEndoso, Cargo_Abono = case when ce.ID_MovConta = 1 then //C// else //A// end, cmr.Nombre_Cuenta, ROPC = ctr.Abv_TpoROPC, "
        //    sSql = sSql + vbCrLf & "     Basicos = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_TpoROPC = ce.ID_TpoROPC and ce2.ID_ConceptoContable in (1,2,3,4)), 0), "
        //    sSql = sSql + vbCrLf & "     Aguinaldo = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_TpoROPC = ce.ID_TpoROPC and ce2.ID_ConceptoContable = 5), 0), "
        //    sSql = sSql + vbCrLf & "     Pension_Adicional = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_TpoROPC = ce.ID_TpoROPC and ce2.ID_ConceptoContable = 6), 0), "
        //    sSql = sSql + vbCrLf & "     Aguinaldo_Adicional = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_TpoROPC = ce.ID_TpoROPC and ce2.ID_ConceptoContable = 7), 0), "
        //    sSql = sSql + vbCrLf & "     Ayuda_Escolar = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_TpoROPC = ce.ID_TpoROPC and ce2.ID_ConceptoContable = 8), 0), "
        //    sSql = sSql + vbCrLf & "     BAMI = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_TpoROPC = ce.ID_TpoROPC and ce2.ID_ConceptoContable = 9), 0), "
        //    sSql = sSql + vbCrLf & "     BAMI_Aguinaldo = isnull((select sum(ce2.Monto_Mov) from Contabilidad_Endosos ce2 where ce2.ID_HisEndoso = e.ID_HisEndoso and ce2.ID_MovConta = ce.ID_MovConta and ce2.ID_MovRem = ce.ID_MovRem and ce2.ID_NivMovRem = ce.ID_NivMovRem and ce2.ID_RamoContable = ce.ID_RamoContable and ce2.ID_TpoROPC = ce.ID_TpoROPC and ce2.ID_ConceptoContable = 10), 0) "
        //    sSql = sSql + vbCrLf & "from Polizas p inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        //    sSql = sSql + vbCrLf & "     inner join Empresas emp on p.ID_Empresa = emp.ID_Empresa "
        //    sSql = sSql + vbCrLf & "     inner join Remesas r on r.Fecha_System = e.Fecha_System and ID_TpoRem = 27 and p.ID_Empresa = r.ID_Empresa "
        //    sSql = sSql + vbCrLf & "     inner join MRemesas mr on r.ID_Remesa = mr.ID_Remesa and Fol_Docto = p.ID_Poliza and Folio = e.ID_HisEndoso "
        //    sSql = sSql + vbCrLf & "     inner join Contabilidad_Endosos ce on ce.ID_HisEndoso = e.ID_HisEndoso "
        //    sSql = sSql + vbCrLf & "     inner join Cat_MovRem cmr on ce.ID_MovRem = cmr.ID_MovRem and ce.ID_NivMovRem = cmr.ID_NivMovRem "
        //    sSql = sSql + vbCrLf & "     inner join Cat_RamoContable crc on ce.ID_RamoContable = crc.ID_RamoContable "
        //    sSql = sSql + vbCrLf & "     inner join Cat_Ramos cra on crc.ID_Ramo = cra.ID_Ramo "
        //    sSql = sSql + vbCrLf & "     inner join Cat_Riesgos cri on crc.ID_Riesgo = cri.ID_Riesgo "
        //    sSql = sSql + vbCrLf & "     inner join Cat_TpoROPC ctr on ce.ID_TpoROPC = ctr.ID_TpoROPC "
        //    if vParametros(0) <> 99 Then
        //        sSql = sSql + vbCrLf & "where    p.ID_Empresa = " & vParametros(0)
        //    Else
        //        sSql = sSql + vbCrLf & "where    p.ID_Empresa >= 0 "
        //    End if
        //    if vParametros(1) <> Empty Then
        //        sSql = sSql + vbCrLf & "     and p.Fol_Poliza = " & vParametros(1)
        //    End if
        //    if Not IsNull(vParametros(2)) Then
        //        sSql = sSql + vbCrLf & "     and e.Fecha_System >= //" & Format(vParametros(2), "yyyy-mm-dd") & "// "
        //    End if
        //    if Not IsNull(vParametros(3)) Then
        //        sSql = sSql + vbCrLf & "     and e.Fecha_System <= //" & Format(vParametros(3), "yyyy-mm-dd") & "// "
        //    End if
        //    sSql = sSql + vbCrLf & "     and ce.ID_ConceptoContable > 0 "
        //    sSql = sSql + vbCrLf & "group by emp.Abv_Empresa, r.Fol_Remesa, e.Fecha_System, p.Fol_Poliza, e.ID_HisEndoso, cmr.Nombre_Cuenta, ce.ID_MovConta, ce.ID_MovRem, ce.ID_NivMovRem, ce.ID_RamoContable, cra.Cve_Ramo, cri.Abv_Riesgo, ce.ID_TpoROPC, ctr.Abv_TpoROPC "

        //    sSql = sSql + vbCrLf & "order by emp.Abv_Empresa, r.Fol_Remesa, e.ID_HisEndoso, case when ce.ID_MovConta = 1 then //C// else //A// end, cmr.Nombre_Cuenta, ctr.Abv_TpoROPC "
        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic


        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iReporte_DetalleContabilidad = NoHayDatos
        //    Else
        //        iReporte_DetalleContabilidad = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iReporte_DetalleContabilidad = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        //Public Function bActualizaPagoRedistribuir(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional vParametros As Variant) As Boolean
        //Dim rsAux As ADODB.Recordset
        //Dim dPje_Total As Double
        //Dim iBenef As Byte
        //On Error GoTo msgerror

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN


        //    Set ctxObject = GetObjectContext()

        //    //Barrer los 4 beneficios
        //    For x = 1 To 4
        //        Select Case x
        //            Case 1:
        //                iBenef = 76 //Básico
        //            Case 2:
        //                iBenef = 78 //Artículo 14
        //            Case 3:
        //                iBenef = 79 //Aguinaldo
        //            Case 4:
        //                iBenef = 80 //Aguinaldo Artículo 14
        //        End Select


        //        dPje_Total = 0

        //        //Si es mayor que cero, ajusto
        //        if vParametros(x) > 0 Then
        //            //NOTA: Dado que en estos casos la póliza siempre queda de orfandad
        //            //tomo los porcentajes dependiendo solo si son hijos Sencillos o Dobles de Tmp_EndosoCET
        //            sSql = "select Pje_Total = sum(ID_Orfandad) "
        //            sSql = sSql + vbCrLf & "from Tmp_PagosEndosos tp, "
        //            sSql = sSql + vbCrLf & "     Tmp_EndosoCET te "
        //            sSql = sSql + vbCrLf & "where te.ID_Poliza = tp.ID_Poliza "
        //            sSql = sSql + vbCrLf & "    and te.ID_PolBenef = tp.ID_PolBenef "
        //            sSql = sSql + vbCrLf & "    and te.ID_Poliza = " & vParametros(0)
        //            sSql = sSql + vbCrLf & "    and tp.ID_Beneficio = " & iBenef
        //            sSql = sSql + vbCrLf & "    and ID_StaPolBenef = 1 " //Activos
        //            sSql = sSql + vbCrLf & "    and ID_Parentesco = 3 " //Hijo
        //            Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)


        //            if Not rsAux.EOF Then
        //                dPje_Total = rsAux!Pje_Total

        //                //Ajustar por componente
        //                sSql = "update tp set Importe = round(ID_Orfandad / " & Format(dPje_Total, "#.00") & " * " & vParametros(x) & ", 2) "
        //                sSql = sSql + vbCrLf & "from Tmp_PagosEndosos tp, "
        //                sSql = sSql + vbCrLf & "     Tmp_EndosoCET te "
        //                sSql = sSql + vbCrLf & "where te.ID_Poliza = tp.ID_Poliza "
        //                sSql = sSql + vbCrLf & "    and te.ID_PolBenef = tp.ID_PolBenef "
        //                sSql = sSql + vbCrLf & "    and te.ID_Poliza = " & vParametros(0)
        //                sSql = sSql + vbCrLf & "    and tp.ID_Beneficio = " & iBenef
        //                cnnConexion.Execute sSql


        //                sSql = "select Total = sum(Importe) from Tmp_PagosEndosos t where ID_Poliza = " & vParametros(0) & " and ID_Beneficio = " & iBenef
        //                Set rsAux = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)


        //                if rsAux!Total<> vParametros(x) Then
        //                    //Ajustar centavos al primer componente
        //                    sSql = "update tp set Importe = Importe - " & rsAux!Total - vParametros(x) & " "
        //                    sSql = sSql + vbCrLf & "from Tmp_PagosEndosos tp "
        //                    sSql = sSql + vbCrLf & "where ID_Poliza = " & vParametros(0)
        //                    sSql = sSql + vbCrLf & "    and tp.ID_Beneficio = " & iBenef
        //                    sSql = sSql + vbCrLf & "    and tp.ID_PolBenef = (select min(ID_PolBenef) from Tmp_PagosEndosos where ID_Poliza = tp.ID_Poliza and ID_Beneficio = tp.ID_Beneficio) "
        //                    cnnConexion.Execute sSql
        //                End if
        //            End if
        //        //Si es cero actualizo todo a cero, no borro los registros por si el usuario
        //        //se equivocó y luego corrige, sp_GuardarEndoso los ignorará
        //        Elseif vParametros(x) = 0 Then
        //            sSql = "update Tmp_PagosEndosos set Importe = 0 where ID_Poliza = " & vParametros(0) & " and ID_Beneficio = " & iBenef
        //            cnnConexion.Execute sSql
        //        End if
        //    Next

        //    //SDC 2007-02-28 Ya no se utiliza el total, ya que se checa si la viuda tiene derecho
        //    //al artículo 14, esto se ve en la forma.

        //    ctxObject.SetComplete
        //    cnnConexion.Close
        //    Set ctxObject = Nothing
        //    Set cnnConexion = Nothing
        //    bActualizaPagoRedistribuir = True

        //Exit Function
        //msgerror:
        //    bActualizaPagoRedistribuir = False
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    if bBool Then
        //        Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    End if
        //    ctxObject.SetAbort
        //    cnnConexion.Close
        //    Set ctxObject = Nothing
        //    Set cnnConexion = Nothing
        //End Function

        ////_______________________________________________________________
        ////Título: Catálogo Baja/Alta
        ////Clase: clsEndososCET.iCatalogo_BajaAlta
        ////Versión: 1.0
        ////Fecha:    19/12/2006
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////   Trae el catalogo de Causa de Baja o Alta para Estadística Anual
        ////_______________________________________________________________
        //Public Function iCatalogo_CausaBajaAlta(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, ByRef rsRecord As ADODB.Recordset, ByRef vParametros As Variant) As Integer
        //On Error GoTo msgerror

        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    Set rsErrAdo = Nothing
        //    rsRecord.CursorLocation = adUseClient


        //    if vParametros(2) = 0 Then
        //        sSql = "select      ID_Causa = ca.ID_CausaAlta, "
        //        sSql = sSql + "     Causa = ca.Causa_Alta, "
        //        sSql = sSql + "     ID_Default = case when t.ID_Parentesco = 3 then case when datediff(dd, t.Fecha_Nacto, dbo.getdateusr()) > 365 then 2 else 1 end "
        //        sSql = sSql + "                 when t.ID_Parentesco in (2,5) then case when t.ID_Pension in (1,6) then 3 else 2 end "
        //        sSql = sSql + "                 when t.ID_Parentesco in (4,6) then 2 "
        //        sSql = sSql + "                 when t.ID_Parentesco = 1 then 9 end "
        //        sSql = sSql + "from Cat_CausaAltaEA ca, "
        //        sSql = sSql + "     Tmp_EndosoCET t "
        //        sSql = sSql + "where t.ID_Poliza = " & vParametros(0)
        //        sSql = sSql + "     and t.ID_PolBenef = (select min(ID_PolBenef) from Tmp_EndosoCET where ID_Poliza = t.ID_Poliza and ID_StaPolBenef = 3) "
        //        sSql = sSql + "order by ID_Causa "
        //    Elseif vParametros(2) = 1 Then
        //        sSql = "select      ID_Causa = ca.ID_CausaBaja, "
        //        sSql = sSql + "     Causa = ca.Causa_Baja, "
        //        sSql = sSql + "     ID_Default = case " & vParametros(1) & " "
        //        sSql = sSql + "                 when 8 then 1 "
        //        sSql = sSql + "                 when 9 then 3 "
        //        sSql = sSql + "                 when 10 then case when t.ID_Parentesco in (2,5) then case when t.ID_Pension in (1,6) then 4 else 13 end "
        //        sSql = sSql + "                                 when t.ID_Parentesco in (1,3,4,6) then 13 end "
        //        sSql = sSql + "                 when 24 then 13 " //CEFB YA9A0E Sep2012
        //        sSql = sSql + "                 when 27 then 14 end " //CEFB YA9A0E Sep2012
        //        sSql = sSql + "from Cat_CausaBajaEA ca "
        //        if vParametros(1) <> 24 And vParametros(1) <> 27 Then //CEFB YA9A0E Sep2012
        //            sSql = sSql + " inner "
        //        Else
        //            sSql = sSql + " left "
        //        End if
        //        sSql = sSql + "     join Tmp_EndosoCET t on t.ID_Poliza = " & vParametros(0) & " and t.ID_StaPolBenef = 2 "
        //        sSql = sSql + "order by ID_Causa "
        //    Else
        //        iCatalogo_CausaBajaAlta = NoHayDatos
        //        GoTo fin
        //    End if


        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic


        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iCatalogo_CausaBajaAlta = NoHayDatos
        //    Else
        //        iCatalogo_CausaBajaAlta = DatosOK
        //    End if

        //fin:
        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //    Exit Function
        //msgerror:
        //    iCatalogo_CausaBajaAlta = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        ////_______________________________________________________________
        ////Título: Reporte de Endosos
        ////Clase: clsEndososCET.iReportes
        ////Versión: 1.0
        ////Fecha:    22/03/2007
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Consulta de endosos CET por rango de fechas
        ////_______________________________________________________________
        //Public Function iReportes(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror
        //    //SDC 2007-04-23 Añadir Rentas Mensuales y Aguinaldo, corregir Prima Riesgo en Cambios en el Status
        //    //Añadir parentesco en SS.
        //    //Añadir columnas propias de Segundas Nupcias en Devoluciones


        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    rsRecord.CursorLocation = adUseClient


        //    Select Case vParametros(0)


        //    Case 0:
        //        //Cambios en el Status
        //        sSql = "select  [No. Endoso] = e.ID_HisEndoso, "
        //        sSql = sSql + " Poliza = dbo.Poliza_Tecnica(p.ID_Poliza), "
        //        sSql = sSql + " p.Num_SegSocial, "
        //        sSql = sSql + " ce.Endoso, "
        //        sSql = sSql + " d.Dif_Prima, "
        //        sSql = sSql + " Prima_Emitida = d.Tot_TransAlAseg, "
        //        sSql = sSql + " Prima_Riesgo = round(d.Dif_Prima / 1.03, 2), "
        //        sSql = sSql + " Pago_Vencido = d.Pagos_Venc, "
        //        sSql = sSql + " Rentas_Mensuales = d.Rentas_Mensuales, "
        //        sSql = sSql + " Aguinaldo = d.Aguinaldo, "
        //        sSql = sSql + " Fecha_Aplicacion = convert(varchar(10), e.Fecha_System, 120), "
        //        sSql = sSql + " Derecho_Art14 = isnull(Sta_Incremento, //SIN DERECHO//), "
        //        sSql = sSql + " e.Descripcion "
        //        sSql = sSql + "from Polizas p inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        //        sSql = sSql + "     inner join Cat_Endosos ce on e.ID_Endoso = ce.ID_Endoso "
        //        sSql = sSql + "     inner join DEndosos d on e.ID_HisEndoso = d.ID_HisEndoso "
        //        sSql = sSql + "     inner join Pol_Benefs pb on p.ID_Poliza = pb.ID_Poliza and pb.ID_PolBenef = e.ID_PolBenef "
        //        sSql = sSql + "     left join Incremento_Pensiones i on pb.ID_PolBenef = i.ID_PolBenef "
        //        sSql = sSql + "     left join Cat_StaIncremento csi on i.ID_StaIncremento = csi.ID_StaIncremento "
        //        sSql = sSql + "where ce.ID_TpoEndoso = 1 "
        //        sSql = sSql + "     and e.ID_Endoso <> 8 "
        //        sSql = sSql + "     and d.Devolucion = 0 "
        //        sSql = sSql + "     and not exists (select * from Pol_Benefs pb1, Endosos e1 where pb1.ID_Poliza = p.ID_Poliza and pb1.ID_Poliza = e1.ID_Poliza and pb1.ID_PolBenef = e1.ID_PolBenef and pb1.ID_Parentesco = 1 and ID_Endoso = 8) "
        //        sSql = sSql + "     and e.Fecha_System between " & vParametros(1) & " and " & vParametros(2)
        //        sSql = sSql + "     and p.ID_InstitucionSS in (1,2) " //CEFB YA9A0E Sep2012
        //        sSql = sSql + "order by e.ID_HisEndoso "

        //    Case 1:
        //        //Cambios en el Status por Sobrevivencia
        //        sSql = "select  [No. Endoso] = e.ID_HisEndoso, "
        //        sSql = sSql + " Poliza = dbo.Poliza_Tecnica(p.ID_Poliza), "
        //        sSql = sSql + " p.Num_SegSocial, "
        //        sSql = sSql + " ce.Endoso, "
        //        sSql = sSql + " d.Dif_Prima, "
        //        sSql = sSql + " Prima_Emitida = d.Tot_TransAlAseg, "
        //        sSql = sSql + " Prima_Riesgo = round(d.Tot_TransAlAseg / 1.03, 2), "
        //        sSql = sSql + " Pago_Vencido = d.Pagos_Venc, "
        //        sSql = sSql + " Rentas_Mensuales = d.Rentas_Mensuales, "
        //        sSql = sSql + " Aguinaldo = d.Aguinaldo, "
        //        sSql = sSql + " Fecha_Aplicacion = convert(varchar(10), e.Fecha_System, 120), "
        //        sSql = sSql + " Derecho_Art14 = isnull(Sta_Incremento, //SIN DERECHO//), "
        //        sSql = sSql + " e.Descripcion "
        //        sSql = sSql + "from Polizas p inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        //        sSql = sSql + "     inner join Cat_Endosos ce on e.ID_Endoso = ce.ID_Endoso "
        //        sSql = sSql + "     inner join DEndosos d on e.ID_HisEndoso = d.ID_HisEndoso "
        //        sSql = sSql + "     inner join Pol_Benefs pb on p.ID_Poliza = pb.ID_Poliza and pb.ID_PolBenef = e.ID_PolBenef "
        //        sSql = sSql + "     left join Incremento_Pensiones i on pb.ID_PolBenef = i.ID_PolBenef "
        //        sSql = sSql + "     left join Cat_StaIncremento csi on i.ID_StaIncremento = csi.ID_StaIncremento "
        //        sSql = sSql + "where ce.ID_TpoEndoso = 1 "
        //        sSql = sSql + "     and e.ID_Endoso <> 8 "
        //        sSql = sSql + "     and d.Devolucion = 0 "
        //        sSql = sSql + "     and exists (select * from Pol_Benefs pb1, Endosos e1 where pb1.ID_Poliza = p.ID_Poliza and pb1.ID_Poliza = e1.ID_Poliza and pb1.ID_PolBenef = e1.ID_PolBenef and pb1.ID_Parentesco = 1 and ID_Endoso = 8) "
        //        sSql = sSql + "     and e.Fecha_System between " & vParametros(1) & " and " & vParametros(2)
        //        sSql = sSql + "     and p.ID_InstitucionSS in (1,2) " //CEFB YA9A0E Sep2012
        //        sSql = sSql + "order by e.ID_HisEndoso "

        //    Case 2:
        //        //Fallecimientos
        //        sSql = "select  [No. Endoso] = e.ID_HisEndoso, "
        //        sSql = sSql + " Poliza = dbo.Poliza_Tecnica(p.ID_Poliza), "
        //        sSql = sSql + " p.Num_SegSocial, "
        //        sSql = sSql + " Nombre = isnull(pb.Nom_Benef, ////) + // // + isnull(pb.ApP_Benef, ////) + // // + isnull(pb.ApM_Benef, ////), "
        //        sSql = sSql + " Parentesco = cp.Cve_Parentesco, "
        //        sSql = sSql + " Prima_Emitida = d.Tot_TransAlAseg, "
        //        sSql = sSql + " Pago_Vencido = d.Pagos_Venc, "
        //        sSql = sSql + " Rentas_Mensuales = d.Rentas_Mensuales, "
        //        sSql = sSql + " Aguinaldo = d.Aguinaldo, "
        //        sSql = sSql + " cs.Sexo, "
        //        sSql = sSql + " Fecha_Nacto = convert(varchar(10), pb.Fecha_Nacto, 120), "
        //        sSql = sSql + " Fecha_Defuncion = convert(varchar(10), e.Fecha_Aplicacion, 120), "
        //        sSql = sSql + " Edad = dbo.Age(pb.Fecha_Nacto, e.Fecha_Aplicacion), "
        //        sSql = sSql + " Fecha_Aplicacion = convert(varchar(10), e.Fecha_System, 120), "
        //        sSql = sSql + " Tipo_Seguro = cr.Ramo, "
        //        sSql = sSql + " Tipo_PensionAntes = cp1.Pension, "
        //        sSql = sSql + " Tipo_PensionDespues = cp2.Pension, "
        //        sSql = sSql + " Reserva_Liberada = d.Res_Liberada, "
        //        sSql = sSql + " Reserva_LiberadaArt1402 = d.Res_LiberadaArt14, "
        //        sSql = sSql + " Reserva_LiberadaArt1404 = d.Res_LiberadaInc04, "
        //        sSql = sSql + " Status_Poliza = csp.Sta_Poliza "
        //        sSql = sSql + "from Polizas p inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        //        sSql = sSql + "     inner join Cat_Endosos ce on e.ID_Endoso = ce.ID_Endoso "
        //        sSql = sSql + "     inner join DEndosos d on e.ID_HisEndoso = d.ID_HisEndoso "
        //        sSql = sSql + "     inner join Pol_Benefs pb on p.ID_Poliza = pb.ID_Poliza and pb.ID_PolBenef = e.ID_PolBenef "
        //        sSql = sSql + "     inner join Cat_Parentescos cp on pb.ID_Parentesco = cp.ID_Parentesco "
        //        sSql = sSql + "     inner join Cat_Sexos cs on pb.ID_Sexo = cs.ID_Sexo "
        //        sSql = sSql + "     inner join Cat_Ramos cr on p.ID_Ramo = cr.ID_Ramo "
        //        sSql = sSql + "     inner join Cat_Pensiones cp2 on p.ID_Pension = cp2.ID_Pension "
        //        sSql = sSql + "     inner join Cat_StaPolizas csp on p.ID_StaPoliza = csp.ID_StaPoliza "
        //        sSql = sSql + "     left join DBHISCAIRO.dbo.Img_PolizasEst i on p.ID_Poliza = i.ID_Poliza and i.Fecha_Valuacion = dbo.Ultimo_DiaMes(dateadd(mm, -1, e.Fecha_System)) "
        //        sSql = sSql + "     left join Cat_Pensiones cp1 on i.ID_Pension = cp1.ID_Pension "
        //        sSql = sSql + "where ce.ID_TpoEndoso = 1 "
        //        sSql = sSql + "     and e.ID_Endoso = 8 "
        //        sSql = sSql + "     and e.Fecha_System between " & vParametros(1) & " and " & vParametros(2)
        //        sSql = sSql + "     and p.ID_InstitucionSS in (1,2) " //CEFB YA9A0E Sep2012
        //        sSql = sSql + "order by e.ID_HisEndoso "

        //    Case 3:
        //        //Devoluciones al IMSS
        //        sSql = "select  [No. Endoso] = e.ID_HisEndoso, "
        //        sSql = sSql + " Poliza = dbo.Poliza_Tecnica(p.ID_Poliza), "
        //        sSql = sSql + " p.Num_SegSocial, "
        //        sSql = sSql + " ce.Endoso, "
        //        sSql = sSql + " d.Devolucion, "
        //        sSql = sSql + " Pago_Vencido = d.Pagos_Venc, "
        //        sSql = sSql + " Rentas_Mensuales = d.Rentas_Mensuales, "
        //        sSql = sSql + " Aguinaldo = d.Aguinaldo, "
        //        sSql = sSql + " Cuantia_MensualViuda = round(d.Finiq_Neto / 36.00, 2), "
        //        sSql = sSql + " Finiquito_Neto = d.Finiq_Neto, "
        //        sSql = sSql + " d.Pagos_Indebidos, "
        //        sSql = sSql + " Pagos_Redistribucion = d.Pag_RetroPension, "
        //        sSql = sSql + " Finiquito_Total = d.Finiq_Total, "
        //        sSql = sSql + " Pagos_RedistribucionIncob = pe.Imp_Pago - d.Finiq_Neto + d.Pagos_Indebidos + d.Liquidacion_Prestamo, "
        //        sSql = sSql + " Prestamo = d.Liquidacion_Prestamo, "
        //        sSql = sSql + " Pago_Neto = pe.Imp_Pago, "
        //        sSql = sSql + " Fecha_Aplicacion = convert(varchar(10), e.Fecha_System, 120), "
        //        sSql = sSql + " e.Descripcion "
        //        sSql = sSql + "from Polizas p inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        //        sSql = sSql + "     inner join Cat_Endosos ce on e.ID_Endoso = ce.ID_Endoso "
        //        sSql = sSql + "     inner join DEndosos d on e.ID_HisEndoso = d.ID_HisEndoso "
        //        sSql = sSql + "     inner join Pol_Benefs pb on p.ID_Poliza = pb.ID_Poliza and pb.ID_PolBenef = e.ID_PolBenef "
        //        sSql = sSql + "     left join Incremento_Pensiones i on pb.ID_PolBenef = i.ID_PolBenef "
        //        sSql = sSql + "     left join Pagos_Endosos pe on pe.ID_Poliza = p.ID_Poliza and pe.ID_HisEndoso = case when e.ID_Endoso <> 9 then 0 else e.ID_HisEndoso end "
        //        sSql = sSql + "where ce.ID_TpoEndoso = 1 "
        //        sSql = sSql + "     and d.Devolucion > 0 "
        //        sSql = sSql + "     and e.Fecha_System between " & vParametros(1) & " and " & vParametros(2)
        //        sSql = sSql + "     and p.ID_InstitucionSS in (1,2) " //CEFB YA9A0E Sep2012
        //        sSql = sSql + "order by e.ID_HisEndoso "

        //    //INI CEFB YA9A0E Sep2012
        //    Case 4:
        //        //Cancelacion de Pol/Gpo sin Devolución de Reserva
        ////        sSql = "select  [No. Endoso] = e.ID_HisEndoso, "
        ////        sSql = sSql + " Poliza = dbo.Poliza_Tecnica(p.ID_Poliza), "
        ////        sSql = sSql + " p.Num_SegSocial, "
        ////        sSql = sSql + " Nombre = isnull(pb.Nom_Benef, ////) + // // + isnull(pb.ApP_Benef, ////) + // // + isnull(pb.ApM_Benef, ////), "
        ////        sSql = sSql + " Parentesco = cp.Cve_Parentesco, "
        ////        sSql = sSql + " Prima_Emitida = d.Tot_TransAlAseg, "
        ////        sSql = sSql + " Pago_Vencido = d.Pagos_Venc, "
        ////        sSql = sSql + " Rentas_Mensuales = d.Rentas_Mensuales, "
        ////        sSql = sSql + " Aguinaldo = d.Aguinaldo, "
        ////        sSql = sSql + " cs.Sexo, "
        ////        sSql = sSql + " Fecha_Nacto = convert(varchar(10), pb.Fecha_Nacto, 120), "
        ////        sSql = sSql + " Fecha_Defuncion = convert(varchar(10), e.Fecha_Aplicacion, 120), "
        ////        sSql = sSql + " Edad = dbo.Age(pb.Fecha_Nacto, e.Fecha_Aplicacion), "
        ////        sSql = sSql + " Fecha_Aplicacion = convert(varchar(10), e.Fecha_System, 120), "
        ////        sSql = sSql + " Tipo_Seguro = cr.Ramo, "
        ////        sSql = sSql + " Tipo_PensionAntes = cp1.Pension, "
        ////        sSql = sSql + " Tipo_PensionDespues = cp2.Pension, "
        ////        sSql = sSql + " Reserva_Liberada = d.Res_Liberada, "
        ////        sSql = sSql + " Reserva_LiberadaArt1402 = d.Res_LiberadaArt14, "
        ////        sSql = sSql + " Reserva_LiberadaArt1404 = d.Res_LiberadaInc04, "
        ////        sSql = sSql + " Status_Poliza = csp.Sta_Poliza "
        ////        sSql = sSql + "from Polizas p inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        ////        sSql = sSql + "     inner join Cat_Endosos ce on e.ID_Endoso = ce.ID_Endoso "
        ////        sSql = sSql + "     inner join DEndosos d on e.ID_HisEndoso = d.ID_HisEndoso "
        ////        sSql = sSql + "     inner join Pol_Benefs pb on p.ID_Poliza = pb.ID_Poliza and pb.ID_PolBenef = e.ID_PolBenef "
        ////        sSql = sSql + "     inner join Cat_Parentescos cp on pb.ID_Parentesco = cp.ID_Parentesco "
        ////        sSql = sSql + "     inner join Cat_Sexos cs on pb.ID_Sexo = cs.ID_Sexo "
        ////        sSql = sSql + "     inner join Cat_Ramos cr on p.ID_Ramo = cr.ID_Ramo "
        ////        sSql = sSql + "     inner join Cat_Pensiones cp2 on p.ID_Pension = cp2.ID_Pension "
        ////        sSql = sSql + "     inner join Cat_StaPolizas csp on p.ID_StaPoliza = csp.ID_StaPoliza "
        ////        sSql = sSql + "     left join Img_PolizasEst i on p.ID_Poliza = i.ID_Poliza and i.Fecha_Valuacion = dbo.Ultimo_DiaMes(dateadd(mm, -1, e.Fecha_System)) "
        ////        sSql = sSql + "     left join Cat_Pensiones cp1 on i.ID_Pension = cp1.ID_Pension "
        ////        sSql = sSql + "where ce.ID_TpoEndoso = 2 "
        ////        sSql = sSql + "     and e.ID_Endoso = 27 "
        ////        sSql = sSql + "     and e.Fecha_System between " & vParametros(1) & " and " & vParametros(2)
        ////        sSql = sSql + "     and p.ID_InstitucionSS in (1,2) "
        ////        sSql = sSql + "order by e.ID_HisEndoso "
        //        sSql = " select   Poliza = dbo.Poliza_Tecnica(p.ID_Poliza), "
        //        sSql = sSql + " Ramo = cr.Cve_Ramo, "
        //        sSql = sSql + " Riesgo = cr2.Abv_Riesgo, "
        //        sSql = sSql + " Pension = cp.Cve_Pension, "
        //        sSql = sSql + " Pagos_Vencidos = p.Pagos_VTotal, "
        //        sSql = sSql + " Monto_Constitutivo = p.Monto_ConstTotal, "
        //        sSql = sSql + " Res_Matematica = ISNULL(r.Res_Matematica, 0), "
        //        sSql = sSql + " Fecha_IniVig = p.Fecha_IniVig, "
        //        sSql = sSql + " Fecha_Emision = p.Fecha_Emision, "
        //        sSql = sSql + " Fecha_Aplicacion = convert(varchar(10), e.Fecha_System, 120), "
        //        sSql = sSql + " Usuario = u.Usuario "
        //        sSql = sSql + " from Polizas p "
        //        sSql = sSql + "      inner join Endosos e on p.ID_Poliza = e.ID_Poliza "
        //        sSql = sSql + "      inner join DEndosos d on e.ID_HisEndoso = d.ID_HisEndoso "
        //        sSql = sSql + "      inner join Cat_Ramos cr on p.ID_Ramo = cr.ID_Ramo "
        //        sSql = sSql + "      inner join Cat_Riesgos cr2 on p.ID_Riesgo = cr2.ID_Riesgo "
        //        sSql = sSql + "      inner join Cat_Pensiones cp on p.ID_Pension = cp.ID_Pension "
        //        sSql = sSql + "      left  join Reservas r on p.ID_Poliza = r.ID_Poliza "
        //        sSql = sSql + "      inner join Usuarios u on u.ID_Usuario = e.ID_Usuario "
        //        sSql = sSql + " where 1=1 "
        //        sSql = sSql + "      and e.ID_Endoso = 27 "
        //        sSql = sSql + "      and e.Fecha_System between " & vParametros(1) & " and " & vParametros(2)
        //        sSql = sSql + "      and p.ID_InstitucionSS in (1,2) "
        //        sSql = sSql + "order by e.ID_HisEndoso "
        //    //FIN CEFB YA9A0E Sep2012

        //    End Select


        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic


        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iReportes = NoHayDatos
        //    Else
        //        iReportes = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iReportes = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        ////_______________________________________________________________
        ////Título: Consulta de Recibos Pendientes de Pago
        ////Clase: clsEndososCET.iRecibos_Pendientes
        ////Versión: 1.0
        ////Fecha:    08/05/07
        ////Autor:    Samuel Dueñas
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Consulta de recibos pendientes de pago para una póliza
        ////_______________________________________________________________
        //Public Function iRecibos_Pendientes(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror


        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    rsRecord.CursorLocation = adUseClient


        //    sSql = "select * from Recibos where ID_StaRecibo = 1 and ID_Poliza = " & vParametros(0)


        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic

        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iRecibos_Pendientes = NoHayDatos
        //    Else
        //        iRecibos_Pendientes = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iRecibos_Pendientes = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function



        ////_______________________________________________________________
        ////Título: Guarda Log de EndosoCET
        ////Versión: 1.0
        ////Fecha:    20-05-2010
        ////Autor:    Alexander Hdez
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Guarda el log de que se guardo un endosoCET
        ////_______________________________________________________________

        //Public Function bLog_EndosoCET(ByRef rsErrAdo As ADODB.Recordset, ByRef gsConexion As String, Optional vParametros As Variant) As Boolean
        //Dim sSql As String
        //On Error GoTo msgerror
        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gsConexion

        //    sSql = ""                                                                                                                                                        //Poliza                   //tipo endoso
        //    sSql = "Exec sp_BitacoraLog  //" & vParametros(0) & "//, //" & vParametros(1) & "//, //" & "ENDOSOS" & "//, //" & "ENDOSOS CON EFECTO TECNICO" & "//, //" & "" & "//, //" & vParametros(2) & "//, //" & vParametros(3) & "//, //" & "" & "//, //" & "" & "//, //" & "" & "//, //" & "" & "// "
        //    cnnConexion.Execute (sSql)

        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //    bLog_EndosoCET = True

        ////GetObjectContext.SetComplete
        //Exit Function

        //msgerror:
        //    bLog_EndosoCET = False
        //    Set cnnConexion = Nothing

        //End Function



        ////INI CEFB YA9A0E Sep2012
        ////_______________________________________________________________
        ////Título: Consulta la ROPC existente de la póliza
        ////Versión: 1.0
        ////Fecha:    29-Ago-2012
        ////Autor:    CEFB
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Consulta el ROPC de la póliza
        ////_______________________________________________________________
        //Public Function bCancelacionSinDevReserva(ByRef rsErrAdo As ADODB.Recordset, sConn As String, ByRef rsROPC As ADODB.Recordset, parametros As Variant) As Boolean
        //Dim sSql As String
        //Dim errVB As String
        //Dim cnnConexion As ADODB.Connection
        //On Error GoTo errGetPagos


        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open sConn
        //    Set rsErrAdo = Nothing

        //    //Traer información de ROPC
        //    sSql = "select  pg.ID_Pago, pg.ID_Grupo, ctp.ID_TpoROPC, pg.APagado, pg.PPagado, pg.Anio, pg.Periodo, "
        //    sSql = sSql + "     Tipo_Pago = ctc.Tipo_Corrida + // - //+  ctp.Abv_TipoPago, "
        //    sSql = sSql + "     Sta_Pago = convert(varchar(2), pg.ID_StaPago) + // - //+ convert(varchar(40), csp.Abv_StaPago), "
        //    sSql = sSql + "     pg.Imp_Pago, pg.Fecha_System, pg.ID_StaPago, c.ID_TipoCorrida "
        //    sSql = sSql + "from Pagos pg inner join Cat_TipoPagos ctp on pg.ID_TipoPago = ctp.ID_TipoPago and ctp.ID_TpoROPC in (2,3) "
        //    sSql = sSql + "     left join Cat_StaPagos csp on pg.ID_StaPago = csp.ID_StaPago "
        //    sSql = sSql + "     left join Corridas c on pg.ID_Corrida = c.ID_Corrida "
        //    sSql = sSql + "     left join Cat_TipoCorridas ctc on c.ID_TipoCorrida = ctc.ID_TipoCorrida "
        //    sSql = sSql + "where    pg.ID_Poliza = " & parametros(0) & " "
        //    sSql = sSql + "     and pg.ID_StaPago in (2,8) "
        //    sSql = sSql + "order by pg.APagado, pg.PPagado, pg.ID_Grupo, pg.ID_TipoPago, pg.Imp_Pago desc "


        //    Set rsROPC = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)


        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //    bCancelacionSinDevReserva = True
        //    Exit Function


        //errGetPagos:
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //    Set rsROPC = Nothing
        //    bCancelacionSinDevReserva = False
        //    errVB = Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(Nothing, errVB)
        //End Function
        ////FIN CEFB YA9A0E Sep2012


        ////INI CEFB YA9A0E Oct2012
        ////_______________________________________________________________
        ////Título: Consulta de Pagos a Banco para la Póliza
        ////Clase: clsEndososCET.iPagos
        ////Versión: 1.0
        ////Fecha:    10/10/2012
        ////Autor:    César Flores
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Consulta de Pagos a Banco para la Póliza
        ////_______________________________________________________________
        //Public Function iPagos(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror


        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    rsRecord.CursorLocation = adUseClient


        //    sSql = "select * from Pagos WHERE ID_StaPago = 1 and ID_TipoPago = 3 and ID_Poliza = " & vParametros(0)


        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic

        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iPagos = NoHayDatos
        //    Else
        //        iPagos = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iPagos = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function

        ////_______________________________________________________________
        ////Título: Consulta de Recibos Pagados
        ////Clase: clsEndososCET.iRecibos_Pagados
        ////Versión: 1.0
        ////Fecha:    11/10/2012
        ////Autor:    César Flores
        ////Modificación:
        ////Fecha de Modificacion:
        ////_______________________________________________________________
        ////Descripción:
        ////           Consulta de recibos pagados para una póliza
        ////_______________________________________________________________
        //Public Function iRecibos_Pagados(ByRef rsErrAdo As ADODB.Recordset, ByRef gDSN As String, Optional ByRef rsRecord As ADODB.Recordset, Optional vParametros As Variant) As Integer
        //On Error GoTo msgerror


        //    Set cnnConexion = New ADODB.Connection
        //    cnnConexion.Open gDSN
        //    Set rsRecord = New ADODB.Recordset
        //    rsRecord.CursorLocation = adUseClient


        //    sSql = "select * from Recibos where ID_StaRecibo = 2 and ID_Poliza = " & vParametros(0)


        //    rsRecord.Open sSql, cnnConexion, adOpenKeyset, adLockOptimistic

        //    if rsRecord.BOF And rsRecord.EOF Then
        //        iRecibos_Pagados = NoHayDatos
        //    Else
        //        iRecibos_Pagados = DatosOK
        //    End if

        //    Set rsRecord.ActiveConnection = Nothing
        //    cnnConexion.Close
        //    Set cnnConexion = Nothing
        //Exit Function
        //msgerror:
        //    iRecibos_Pagados = ExisteError
        //    sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //    Set rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //    Set cnnConexion = Nothing
        //End Function
        //FIN CEFB YA9A0E Oct2012

    }
}
