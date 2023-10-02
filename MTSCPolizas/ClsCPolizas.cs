using ADODB;

namespace MTSCPolizas
{
    public class ClsCPolizas
    {
        //_________________________________________________________________________
        //Titulo:   ClsCPolizas
        //Clase:    ClsCPolizas.cls
        //Versión:  1.0
        //Fecha:    05/Mayo/2003
        //Autor:    Claudia Torres Rodríguez
        //Modificación:
        //Fecha Modificación:
        //_________________________________________________________________________
        //Descripción:
        //     Crea el objeto ClsCPolizas con sus propiedades y métodos para la
        //     consulta general de Pólizas
        //_________________________________________________________________________
        private Connection cnnConexion;          //Variable de Conexión
        private string sErrVB;                                          //Variable Descripción de Error

        public int bGetDetPrestISS(ref Recordset rsData, string sConn, Recordset rsErrAdo, object[] parametros)
        {

            string sSql;
            string errVB;
            int bGetDetPrestISSRes = 0;

            // On Error GoTo bGetDetPrestISS_Err

            try
            {
                bGetDetPrestISSRes = Convert.ToInt32(MTSCPolizas.Modulos.ModRecordset.TipoResultado.NoHayDatos);

                sSql = "Select '1', pri.ID_Prestamo as ID_Prestamo, 'ISS' As Prestamo, pri.ID_Poliza, p.Fol_Poliza, 'Personal' As Tpo_Ptmo, cast(ID_PtmoISSSTE as bigint) As No_Prestamo,";
                sSql = sSql + " right(MesRpt, 2) + '/' + left(MesRpt, 4) As Fch_Aplicacion,";
                sSql = sSql + " pri.Importe, pri.Saldo_Inicial, pri.Saldo, pri.Plazo, pri.DescApli,pri.ID_StaPtmo, dbo.Cat_StaPtmo.Status_Prestamo as Status_Prestamo";
                sSql = sSql + " from Prestamos_ISSSTE pri inner join Polizas p on pri.ID_Poliza=p.ID_Poliza INNER JOIN dbo.Cat_StaPtmo ON pri.ID_StaPtmo = dbo.Cat_StaPtmo.ID_StaPtmo";
                //sSql = sSql & " Where ID_StaPtmo in (2, 4)";
                sSql = sSql + " Where ";
                sSql = sSql + " p.ID_Poliza=" + parametros[0].ToString();
                sSql = sSql + " Union";
                sSql = sSql + " Select '2', pri.ID_Prestamo, 'FVI' As Prestamo, pri.ID_Poliza, p.Fol_Poliza, (case Concepto when '64L' then 'Amortización'";
                sSql = sSql + "                                            when '65L' then 'Seguro de Daños'";
                sSql = sSql + "                                  end) As Tpo_Ptmo,";
                sSql = sSql + " cast(ID_Prestamo as bigint) As ID_Prestamo ,";
                sSql = sSql + " right(Fch_Aplicacion, 2) + '/' + left(Fch_Aplicacion, 4) As Fch_Aplicacion , pri.Importe, ";
                sSql = sSql + " pri.Saldo_Inicial, pri.Saldo, pri.Plazo, pri.DescApli,pri.ID_StaPtmo, dbo.Cat_StaPtmo.Status_Prestamo as Status_Prestamo";
                sSql = sSql + " from Prestamos_FOVISSSTE pri inner join Polizas p on pri.ID_Poliza=p.ID_Poliza INNER JOIN dbo.Cat_StaPtmo ON pri.ID_StaPtmo = dbo.Cat_StaPtmo.ID_StaPtmo";
                //' sSql = sSql & " Where ID_StaPtmo in (2, 4)";
                sSql = sSql + " Where ";
                sSql = sSql + "  p.ID_Poliza=" + parametros[0];
                sSql = sSql + " order by 1";

                bGetDetPrestISSRes = MTSCPolizas.Modulos.ModRecordset.EjecutaSql(ref rsData, sConn, sSql, rsErrAdo);

                //Exit Function
                return bGetDetPrestISSRes;
            }
            catch (Exception Err)
            {
                //bGetDetPrestISS_Err:
                errVB = Err.Source + "\t" + Err.Message;
                rsErrAdo = MTSCPolizas.Modulos.modErrores.ErroresDLL(null, errVB);
                return Convert.ToInt32(MTSCPolizas.Modulos.ModRecordset.TipoResultado.ExisteError);
            }

        }

        public int bGetPolizas(Recordset rsData, string sConn, Recordset rsErrAdo, object[] parametros)
        {
            string sSql;
            string errVB;
            MTSCPolizas.Modulos.ModRecordset.TipoResultado bGetPolizasRet;
            int result = 0;

            // On Error GoTo errGetPolizas

            try
            {

                bGetPolizasRet = MTSCPolizas.Modulos.ModRecordset.TipoResultado.NoHayDatos;

                //Validación, para saber si viene del Ajustes ROPC o de algun otro modulo de consulta
                //JCMN 09-06-2015 --Inicio
                if (Convert.ToInt32(parametros[7]) == 100)
                {

                    sSql = "Select Polizas.ID_Poliza, Polizas.AnioReval, Polizas.Fol_Poliza, Polizas.Num_SegSocial," +
                            " Cat_Ramos.Cve_Ramo Ramo," +
                            " Cat_Pensiones.Cve_Pension Pension," +
                            " Cat_Riesgos.Abv_Riesgo Riesgo," +
                            " Cat_StaPolizas.Sta_Poliza Sta_Poliza," +
                            " Empresas.Abv_Empresa Empresa, Polizas.ID_Empresa, InstitucionSS, Polizas.ID_InstitucionSS " +
                            " From Polizas, Cat_Ramos, Cat_Pensiones, Cat_Riesgos, Cat_StaPolizas, Empresas, Cat_InstitucionSS" +
                            " where Cat_Ramos.ID_Ramo = Polizas.ID_Ramo" +
                            " and Cat_Pensiones.ID_Pension = Polizas.ID_Pension" +
                            " and Cat_Riesgos.ID_Riesgo = Polizas.ID_Riesgo" +
                            " and Cat_StaPolizas.ID_StaPoliza = Polizas.ID_StaPoliza" +
                            " and Empresas.ID_Empresa = Polizas.ID_Empresa" +
                            " and Cat_InstitucionSS.ID_InstitucionSS = Polizas.ID_InstitucionSS" +
                            " and Polizas.ID_Empresa between " + parametros[0] + " and " + parametros[1];
                    sSql = sSql + " and Polizas.ID_Poliza = " + parametros[8] + " order by Polizas.Fol_Poliza";

                    result = MTSCPolizas.Modulos.ModRecordset.EjecutaSql(ref rsData, sConn, sSql, rsErrAdo);

                    return result;
                }
                else
                {

                    //SDC 2007-01-17 Añadir Riesgo y Año de Revaluación y quitar ID//s
                    sSql = "Select Polizas.ID_Poliza, Polizas.AnioReval, Polizas.Fol_Poliza, Polizas.Num_SegSocial," +
                        " Cat_Ramos.Cve_Ramo Ramo," +
                        " Cat_Pensiones.Cve_Pension Pension," +
                        " Cat_Riesgos.Abv_Riesgo Riesgo," +
                        " Cat_StaPolizas.Sta_Poliza Sta_Poliza," +
                        " Empresas.Abv_Empresa Empresa, Polizas.ID_Empresa, InstitucionSS, Polizas.ID_InstitucionSS " +
                        " From Polizas, Cat_Ramos, Cat_Pensiones, Cat_Riesgos, Cat_StaPolizas, Empresas, Cat_InstitucionSS" +
                        " where Cat_Ramos.ID_Ramo = Polizas.ID_Ramo" +
                        " and Cat_Pensiones.ID_Pension = Polizas.ID_Pension" +
                        " and Cat_Riesgos.ID_Riesgo = Polizas.ID_Riesgo" +
                        " and Cat_StaPolizas.ID_StaPoliza = Polizas.ID_StaPoliza" +
                        " and Empresas.ID_Empresa = Polizas.ID_Empresa" +
                        " and Cat_InstitucionSS.ID_InstitucionSS = Polizas.ID_InstitucionSS" +
                        " and Polizas.ID_Empresa between " + parametros[0] + " and " + parametros[1];

                    if (Convert.ToInt32(parametros[2]) != 0) sSql += " and Polizas.Fol_Poliza = " + parametros[2];
                    if (Convert.ToInt32(parametros[3]) != 0) sSql += " and Polizas.Num_SegSocial = //" + parametros[3] + "//";
                    if (parametros[4].ToString() != "" && Convert.ToInt32(parametros[5]) == 1) sSql += " and Polizas.Nom_Aseg+// //+Polizas.ApP_Aseg+// //+Polizas.ApM_Aseg like //%" + parametros[4].ToString() + "%//";
                    if (parametros[4].ToString() != "" && Convert.ToInt32(parametros[5]) == 2) sSql += " and Polizas.ID_Poliza in(select ID_Poliza from Titulares where Nom_Titular+// //+ApP_Titular+// //+ApM_Titular like //%" + parametros[4].ToString() + "%//)";
                    if (parametros[4].ToString() != "" && Convert.ToInt32(parametros[5]) == 3) sSql += sSql + " and Polizas.ID_Poliza in(select ID_Poliza from Pol_Benefs where Nom_Benef+// //+ApP_Benef+// //+ApM_Benef like //%" + parametros[4].ToString() + "%//)";
                    if (Convert.ToInt32(parametros[6]) != 0) sSql += " and Polizas.ID_Oferta in(select ID_Oferta from Emp_Ofertas where Fol_Oferta = " + parametros[6].ToString() + ")";
                    if (Convert.ToInt32(parametros[7]) != 0) sSql += " and Polizas.Num_Resolucion = //" + parametros[7].ToString() + "//";

                    sSql += " order by Polizas.Fol_Poliza";

                    result = MTSCPolizas.Modulos.ModRecordset.EjecutaSql(ref rsData, sConn, sSql, rsErrAdo);
                    return result;
                }
                //Exit Function
            }
            catch (Exception ex)
            {
                //errGetPolizas:
                errVB = ex.Source + "\t" + ex.Message;
                // rsErrAdo = modErrores.ErroresDLL(null, errVB);
                return result;
            }
        }


        public int bGetDetPtmoISS(ref Recordset rsData, string sConn, Recordset rsErrAdo, object[] parametros)
        {
            string sSql;
            string errVB;
            int bGetDetPtmoISSRes = 0;

            //On Error GoTo errGetDetPrest

            try
            {
                bGetDetPtmoISSRes = Convert.ToInt32(MTSCPolizas.Modulos.ModRecordset.TipoResultado.NoHayDatos);

                //sSql = "EXEC sp_ConsultaDescISS " & parametros(0) & ", " & parametros(1)  //Alexander Hdez 01/10/2012 Comente Linea Prestamos FOVISSSTE
                sSql = "EXEC sp_ConsultaDescISS " + parametros[0] + ", " + parametros[1] + ", " + parametros[2]; //Alexander Hdez 01/10/2012 Comente Linea Prestamos FOVISSSTE

                //bGetDetPtmoISS = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

                sSql = "select * from DetalleDescISS  "; //where ID_Prestamo=" & parametros(1)

                //sSql = "select pri.ID_Prestamo, pri.ID_PtmoISSSTE, pri.ID_Poliza, cat.Status_Prestamo, pri.Saldo_Inicial, pri.Saldo, pri.Plazo, " & _
                //        " pri.Importe , pri.DescApli, desci.Num_Desc, desci.Fch_Desc, desci.Observaciones " & _
                //        " from Prestamos_ISSSTE pri left join Descuentos_ISSSTE desci on pri.ID_Poliza = desci.ID_Poliza " & _
                //        "                               and pri.ID_Prestamo=desci.ID_Prestamo " & _
                //        "                           inner join Cat_StaPtmo cat on cat.ID_StaPtmo=pri.ID_StaPtmo " & _
                //        " Where pri.ID_Poliza = " & parametros(0) & " And pri.ID_Prestamo = " & parametros(1)

                bGetDetPtmoISSRes = MTSCPolizas.Modulos.ModRecordset.EjecutaSql(ref rsData, sConn, sSql, rsErrAdo);

                return bGetDetPtmoISSRes;
                //Exit Function
            }
            catch (Exception Err)
            {
                //errGetDetPrest:
                errVB = Err.Source + "\t" + Err.Message;
                rsErrAdo = MTSCPolizas.Modulos.modErrores.ErroresDLL(null, errVB);
                return bGetDetPtmoISSRes;
            }

        }


        //    public Function bGetPrestAut(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, sFecha As String) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetPolizas

        //        bGetPrestAut = TipoResultado.NoHayDatos

        //        sSql = " SP_CATEL //" & sFecha & "//"

        //        bGetPrestAut = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetPolizas:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    end function



        //    public Function bGetDetalle(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetalle

        //        bGetDetalle = TipoResultado.NoHayDatos

        //        //SDC 2007-01-17 Para mostrar el Pje de Inc, Pje de AA, MC y PV de artículo 14, %Comision y Comisión
        //        sSql = "select Polizas.ID_Poliza,Polizas.Fol_Poliza,Polizas.Nom_Aseg+// //+Polizas.ApP_Aseg+// //+Polizas.ApM_Aseg Asegurado," &
        //            " Polizas.ID_Oferta,Emp_Ofertas.Fol_Oferta,Polizas.ID_Empresa," &
        //            " Resolucion=isnull((select Num_Resolucion from Resoluciones where ID_Oferta = Polizas.ID_Oferta),Polizas.Num_Resolucion)," &
        //            " cast(Polizas.ID_StaPoliza as char(2))+// - //+Cat_StaPolizas.Sta_Poliza StaPoliza, Polizas.CURP," &
        //            " Polizas.RFC,Polizas.Num_SegSocial,Polizas.Fecha_Nacto,datediff(yy,Polizas.Fecha_Nacto,getdate()) Edad," &
        //            " cast(Polizas.ID_Sexo as char(2))+// - //+Cat_Sexos.Sexo Sexo,Cat_Sexos.Abv_Sexo," &
        //            " Polizas.Domicilio, Telefono=dbo.EliminaCerosIzquierda(REPLACE(Polizas.Telefono, //-//, ////)), Polizas.ID_CP," &
        //            " Cat_CodPostales.Cod_Postal,Cat_CodPostales.ID_Asento,Cat_Asentos.Asento," &
        //            " Cat_CodPostales.Nom_Asento," &
        //            " Cat_Municipios.Municipio,Cat_Estados.Estado, Cat_Ciudades.Ciudad," &
        //            " cast(Polizas.ID_Ramo as char(2))+// - //+cast(Cat_Ramos.Ramo as char(20)) Ramo," &
        //            " cast(isnull(Polizas.ID_EdoCivil,0) as char(2))+// - //+cast(Cat_EdoCivil.Edo_Civil as char(14)) EdoCivil," &
        //            " cast(Polizas.ID_Pension as char(2))+// - //+cast(Cat_Pensiones.Pension as char(20)) Pension," &
        //            " Polizas.Semanas_Cot,Polizas.Fecha_IniDer,Polizas.Fecha_IniVig,Polizas.Fecha_Emision,Polizas.Fecha_IniVigInc04, Polizas.Fecha_ABase," &
        //            " Polizas.Salario_RT,Polizas.Salario_IV,Polizas.Monto_ConstTotal,Polizas.Monto_ConstBasico,Polizas.Monto_ConstInc04,Polizas.Cuantia_BaseFC,isnull(Polizas.Pension_MensualFC,0) Pension_MensualFC," &
        //            " Polizas.Pagos_VTotal,Polizas.Pagos_VInc04,Polizas.ID_Ejecutivo," &
        //            " Ejecutivos.Fol_Ejecutivo,Ejecutivos.ApP_Ejec+// //+Ejecutivos.ApM_Ejec+// //+Ejecutivos.Nom_Ejec Nom_Ejec," &
        //            " Polizas.UMF, Polizas.Subdeleg, Polizas.Deleg," &
        //            " Polizas.Pje_Ayuda , Polizas.Pje_Valuacion, Ofertas.Porc_Comision, Comision = round(Polizas.Monto_ConstTotal * Ofertas.Porc_Comision / 100.00, 2), Polizas.Monto_ConstBAU, " &
        //            " Polizas.Tpo_Regimen, Ofertas.Modalidad_RP_RV ,Resoluciones.Email, Ofertas.Fecha_InicioPago " //Alexander Hdez 12/11/2013 se agrego ", Ofertas.Fecha_InicioPago" para el servicio R1100
        //        //Alexander Hdez 09/06/2010 se agrego Polizas.Tpo_Regimen que es donde lee el bono para mostrarlo en el detalle de la poliza   --- //CEFB 2011-06-21 YA96ST Se agrego Ofertas.Modalidad
        //        //Alexander Hdez 26/06/2012 Codigos Postales YA9A0F, para mostar email se agrego ,Resoluciones.Email en la linea de arriba
        //        //CEFB 27/Sep/2013 Se agregó Formato al campo telefono.
        //        sSql = sSql + ",Polizas.ID_InstitucionSS,MortalidadCV =ISNULL(Ofertas.MortalidadCV,0),Ofertas.Incremento_11" //GCR Normativa de Valuacion de Pasivos 2015-09-04

        //        // Proyecto: México - Migrar Obsolescencia de BD Oracle Auditoría KPMG
        //        // Juan Martínez Díaz
        //        // 2022-06-22
        //        // [Operadores *=, =*]
        //        sSql = sSql + " from Polizas "
        //        sSql = sSql + "inner join Emp_Ofertas on Emp_Ofertas.ID_Oferta = Polizas.ID_Oferta and Emp_Ofertas.ID_Empresa = Polizas.ID_Empresa  "
        //        sSql = sSql + "inner join Ofertas on Ofertas.ID_Oferta = Polizas.ID_Oferta  "
        //        sSql = sSql + "inner join Cat_StaPolizas on Cat_StaPolizas.ID_StaPoliza = Polizas.ID_StaPoliza  "
        //        sSql = sSql + "inner join Cat_Sexos on Cat_Sexos.ID_Sexo = Polizas.ID_Sexo  "
        //        sSql = sSql + "inner join Cat_CodPostales on Cat_CodPostales.ID_CP = Polizas.ID_CP  "
        //        sSql = sSql + "inner join Cat_Municipios on Cat_Municipios.ID_Municipio = Cat_CodPostales.ID_Municipio  "
        //        sSql = sSql + "inner join Cat_Ciudades on Cat_Ciudades.ID_Ciudad = Cat_CodPostales.ID_Ciudad  "
        //        sSql = sSql + "inner join Cat_Estados on Cat_Estados.ID_Estado = Cat_Municipios.ID_Estado  "
        //        sSql = sSql + "inner join Cat_Asentos on Cat_Asentos.ID_Asento = Cat_CodPostales.ID_Asento  "
        //        sSql = sSql + "inner join Cat_Ramos  on Cat_Ramos.ID_Ramo = Polizas.ID_Ramo  "
        //        sSql = sSql + "inner join Cat_Pensiones on Cat_Pensiones.ID_Pension = Polizas.ID_Pension  "
        //        sSql = sSql + "inner join Ejecutivos on Ejecutivos.ID_Ejecutivo = Polizas.ID_Ejecutivo  "
        //        sSql = sSql + "left join Cat_EdoCivil on Polizas.ID_EdoCivil = Cat_EdoCivil.ID_EdoCivil  "
        //        sSql = sSql + "left join Resoluciones on Polizas.Num_SegSocial = Resoluciones.Num_SegSocial or Polizas.ID_Oferta = Resoluciones.ID_Oferta "
        //        sSql = sSql + "where Polizas.ID_Poliza = " & parametros(0) & ";"

        //        // Reemplaza

        //        //   sSql = sSql + " from Polizas, Emp_Ofertas,Cat_StaPolizas, Cat_Sexos,Cat_CodPostales,Cat_Municipios,Cat_Estados," & _
        //        //           " Cat_Ramos , Cat_Pensiones, Ejecutivos, Cat_Ciudades,Cat_Asentos, Cat_EdoCivil, Ofertas ,Resoluciones " & _
        //        //           " Where Emp_Ofertas.ID_Oferta = Polizas.ID_Oferta" & _
        //        //           " and Ofertas.ID_Oferta = Polizas.ID_Oferta" & _
        //        //           " and Emp_Ofertas.ID_Empresa = Polizas.ID_Empresa" & _
        //        //           " and Cat_StaPolizas.ID_StaPoliza = Polizas.ID_StaPoliza" & _
        //        //           " and Cat_Sexos.ID_Sexo = Polizas.ID_Sexo" & _
        //        //           " and Cat_CodPostales.ID_CP = Polizas.ID_CP"
        //        //
        //        //           //Alexander Hdez 26/06/2012 Codigos Postales YA9A0F, para mostar email se agrego , en el inner la tabla de Resoluciones despues de la de Ofertas
        //        //
        //        //    sSql = sSql + " and Cat_Asentos.ID_Asento = Cat_CodPostales.ID_Asento" & _
        //        //                  " and Cat_Municipios.ID_Municipio = Cat_CodPostales.ID_Municipio" & _
        //        //                  " and Cat_Ciudades.ID_Ciudad = Cat_CodPostales.ID_Ciudad" & _
        //        //                  " and Cat_Estados.ID_Estado = Cat_Municipios.ID_Estado" & _
        //        //                  " and Cat_Ramos.ID_Ramo = Polizas.ID_Ramo" & _
        //        //                  " and Cat_Pensiones.ID_Pension = Polizas.ID_Pension" & _
        //        //                  " and Ejecutivos.ID_Ejecutivo = Polizas.ID_Ejecutivo" & _
        //        //                  " and Cat_EdoCivil.ID_EdoCivil =* Polizas.ID_EdoCivil" & _
        //        //                  " and Polizas.ID_Poliza = " & parametros(0) & " " & _
        //        //                  " and Resoluciones.Num_SegSocial=*Polizas.Num_SegSocial and Resoluciones.ID_Oferta=*Polizas.ID_Oferta "
        //        //
        //        //            //Alexander Hdez 26/06/2012 Codigos Postales YA9A0F, para mostar email se agrego lo siguiente al bloque de arriba
        //        //            //& " " & _
        //        //            //" and Resoluciones.Num_SegSocial=Polizas.Num_SegSocial and Resoluciones.ID_Oferta=Polizas.ID_Oferta "

        //        // [Operadores *=, =*]

        //        bGetDetalle = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetalle:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetTitulares(rsData As ADODB.Recordset, rsDataG As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetTit

        //        bGetTitulares = TipoResultado.NoHayDatos

        //        sSql = "select Polizas.Fol_Poliza, ID_Grupo, Fecha_UltimaNomina," &
        //            " cast(Grupos_Fam.ID_StaPago as char(2))+// - //+cast(Cat_StaPago.Sta_Pago as char(20)) Sta_Pago," &
        //            " cast(Grupos_Fam.ID_StaPVencido as char(2))+// - //+cast(Cat_StaPVencidos.Sta_PVencido as char(20)) Sta_PVencido, IncBAMI" &
        //            " From Polizas, Grupos_Fam,Cat_StaPago, Cat_StaPVencidos" &
        //            " Where Polizas.ID_Poliza = " & parametros(0) &
        //            " and Polizas.ID_Poliza = Grupos_Fam.ID_Poliza" &
        //            " and Grupos_Fam.ID_StaPago = Cat_StaPago.ID_StaPago" &
        //            " and Grupos_Fam.ID_StaPVencido = Cat_StaPVencidos.ID_StaPVencido" &
        //            " order by Grupos_Fam.ID_Grupo "
        //        bGetTitulares = EjecutaSql(rsDataG, sConn, sSql, rsErrAdo)

        //        sSql = "select Polizas.Fol_Poliza, ID_Grupo,Titulares.Num_SegSocial,Endoso = (select max(ID_HisEndoso) from Endosos where ID_Poliza = " & parametros(0) & " and ID_Endoso = 13)," &
        //            " Polizas.Nom_Aseg+// //+Polizas.ApP_Aseg+// //+Polizas.ApM_Aseg Asegurado," &
        //            " cast(Titulares.ID_Conducto as char(2))+// - //+cast(Cat_Conductos.Conducto as char(20)) Conducto," &
        //            " cast(Titulares.ID_StaTitular as char(2))+// - //+cast(Cat_StaTitular.Sta_Titular as char(20)) Sta_Titular," &
        //            " ID_Titular , Fol_Titular" &
        //            " From Titulares, Polizas, Cat_Conductos, Cat_StaTitular" &
        //            " Where Titulares.ID_Poliza = Polizas.ID_Poliza" &
        //            " and Polizas.ID_Poliza = " & parametros(0) &
        //            " and Titulares.ID_Conducto = Cat_Conductos.ID_Conducto" &
        //            " and Titulares.ID_StaTitular = Cat_StaTitular.ID_StaTitular" &
        //            " order by Titulares.ID_Grupo "
        //        bGetTitulares = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetTit:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetDetTit(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetTit

        //        //SDC 2007-04-17 Añadir campo Comentarios y CLABE
        //        //CEFB 20-03-2009 CAIRO-HERMES COMPORTAMENTAL Se agrego NumCliente
        //        bGetDetTit = TipoResultado.NoHayDatos

        //        //Inicio. Alexander Hdez 12/06/2012 Codigos Postales YA9A0F


        //        //    sSql = "select Polizas.Fol_Poliza, Grupos_Fam.ID_Grupo, CLABE = isnull(Titulares.CLABE,//SIN REGISTRAR//), NumCliente = isnull(Titulares.NumCliente,//SIN REGISTRAR//),Titulares.Comentarios, " & _
        //        //            " Polizas.Nom_Aseg+// //+Polizas.ApP_Aseg+// //+Polizas.ApM_Aseg Asegurado," & _
        //        //            " Endoso = isnull((select max(ID_HisEndoso) from Endosos where ID_Poliza = " & parametros(0) & " and ID_Endoso = 13),0)," & _
        //        //            " cast(Titulares.ID_StaTitular as char(2))+// - //+cast(Cat_StaTitular.Sta_Titular as char(20)) Sta_Titular," & _
        //        //            " ID_Titular,Fol_Titular," & _
        //        //            " Titulares.Nom_Titular+// //+Titulares.ApP_Titular+// //+Titulares.ApM_Titular Titular," & _
        //        //            " Domicilio= isnull(Titulares.Domicilio,//SIN REGISTRAR//), Colonia=isnull(Titulares.Colonia,//SIN REGISTRAR//)," & _
        //        //            " Titulares.ID_CP,Cat_CodPostales.Cod_Postal, Cat_CodPostales.ID_Asento,Cat_Asentos.Asento," & _
        //        //            " Cat_CodPostales.Nom_Asento, Cat_Municipios.Municipio,Cat_Estados.Estado, Cat_Ciudades.Ciudad," & _
        //        //            " Telefono=isnull(Titulares.Telefono,//SIN REGISTRAR//)," & _
        //        //            " cast(Titulares.ID_Identifica as char(2))+// - //+cast(Cat_Identifica.Identifica as char(20)) Identifica," & _
        //        //            " Num_Identifica=isnull(Titulares.Num_Identifica,//SIN REGISTRAR//)," & _
        //        //            " cast(Titulares.ID_Conducto as char(2))+// - //+cast(Cat_Conductos.Conducto as char(20)) Conducto," & _
        //        //            " cast(Titulares.ID_Banco as char(2))+// - //+cast(Cat_Bancos.Banco as char(20)) Banco," & _
        //        //            " Cuenta=isnull(Titulares.Cuenta,//SIN REGISTRAR//) , Plaza=isnull(Titulares.Plaza,//SIN REGISTRAR//), Sucursal=isnull(Titulares.Sucursal,//SIN REGISTRAR//)" & _
        //        //            " ,Titulares.Fecha_SystemADG,Titulares.Fecha_SystemADC," & _
        //        //            " Titulares.Fecha_SystemMDG,Titulares.Fecha_SystemMDC," & _
        //        //            " UADG=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioADG),"
        //        //
        //        //    sSql = sSql + " UADC=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioADC)," & _
        //        //            " UMDG=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioMDG)," & _
        //        //            " UMDC=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioMDC)," & _
        //        //            " UADCte=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioADCte)," & _
        //        //            " UMDCte=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioMDCte)," & _
        //        //            " Titulares.Fecha_SystemADCte,Titulares.Fecha_SystemMDCte" & _
        //        //            " from Titulares, Polizas, Grupos_Fam,Cat_Conductos,Cat_StaTitular,Cat_Identifica," & _
        //        //            " Cat_CodPostales , Cat_Municipios, Cat_Estados, Cat_Ciudades, Cat_Asentos, Cat_Bancos" & _
        //        //            " Where Polizas.ID_Poliza = Titulares.ID_Poliza" & _
        //        //            "    and Polizas.ID_Poliza = " & parametros(0) & _
        //        //            " and Titulares.ID_Titular = " & parametros(1) & _
        //        //            " and Grupos_Fam.ID_Poliza = Titulares.ID_Poliza" & _
        //        //            " and Grupos_Fam.ID_Grupo = Titulares.ID_Grupo" & _
        //        //            " and Cat_Conductos.ID_Conducto = Titulares.ID_Conducto" & _
        //        //            " and Cat_Bancos.ID_Banco = Titulares.ID_Banco" & _
        //        //            " and Cat_StaTitular.ID_StaTitular = Titulares.ID_StaTitular" & _
        //        //            " and Cat_CodPostales.ID_CP = Titulares.ID_CP" & _
        //        //            " and Cat_Asentos.ID_Asento = Cat_CodPostales.ID_Asento" & _
        //        //            " and Cat_Municipios.ID_Municipio = Cat_CodPostales.ID_Municipio" & _
        //        //            " and Cat_Ciudades.ID_Ciudad = Cat_CodPostales.ID_Ciudad" & _
        //        //            " and Cat_Estados.ID_Estado = Cat_Municipios.ID_Estado" & _
        //        //            " and Cat_Identifica.ID_Identifica = Titulares.ID_Identifica"


        //        //mafm ini 25082022 se agrega el telefono celular
        //        sSql = "select Polizas.Fol_Poliza, Grupos_Fam.ID_Grupo, CLABE = isnull(Titulares.CLABE,//SIN REGISTRAR//), " &
        //"NumCliente = isnull(Titulares.NumCliente,//SIN REGISTRAR//),Titulares.Comentarios, " &
        //"Polizas.Nom_Aseg+// //+Polizas.ApP_Aseg+// //+Polizas.ApM_Aseg Asegurado, " &
        //"Endoso = isnull((select max(ID_HisEndoso) from Endosos where ID_Poliza = 64769 and ID_Endoso = 13),0), " &
        //"cast(Titulares.ID_StaTitular as char(2))+// - //+cast(Cat_StaTitular.Sta_Titular as char(20)) Sta_Titular, " &
        //"Titulares.ID_Titular,Fol_Titular, Titulares.Nom_Titular+// //+Titulares.ApP_Titular+// //+Titulares.ApM_Titular Titular, " &
        //"Domicilio= isnull(Titulares.Domicilio,//SIN REGISTRAR//), Colonia=isnull(Titulares.Colonia,//SIN REGISTRAR//), " &
        //"Titulares.ID_CP,Cat_CodPostales.Cod_Postal, Cat_CodPostales.ID_Asento,Cat_Asentos.Asento, " &
        //"Cat_CodPostales.Nom_Asento, Cat_Municipios.Municipio,Cat_Estados.Estado, Cat_Ciudades.Ciudad, " &
        //"Telefono=isnull(Titulares.Telefono,//SIN REGISTRAR//), " &
        //"cast(Titulares.ID_Identifica as char(2))+// - //+cast(Cat_Identifica.Identifica as char(20)) Identifica, " &
        //"Num_Identifica=isnull(Titulares.Num_Identifica,//SIN REGISTRAR//), " &
        //"cast(Titulares.ID_Conducto as char(2))+// - //+cast(Cat_Conductos.Conducto as char(20)) Conducto, " &
        //"cast(Titulares.ID_Banco as char(2))+// - //+cast(Cat_Bancos.Banco as char(20)) Banco, " &
        //"Cuenta=isnull(Titulares.Cuenta,//SIN REGISTRAR//) , Plaza=isnull(Titulares.Plaza,//SIN REGISTRAR//), " &
        //"Sucursal=isnull(Titulares.Sucursal,//SIN REGISTRAR//) ,Titulares.Fecha_SystemADG,Titulares.Fecha_SystemADC, " &
        //"Titulares.Fecha_SystemMDG,Titulares.Fecha_SystemMDC, " &
        //"UADG=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioADG), "

        //        sSql = sSql + " UADC=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioADC)," &
        //"UMDG=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioMDG), " &
        //"UMDC=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioMDC), " &
        //"UADCte=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioADCte), " &
        //"UMDCte=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario=Titulares.ID_UsuarioMDCte), " &
        //"Titulares.Fecha_SystemADCte,Titulares.Fecha_SystemMDCte,Titulares.Email, Cat_Asentos.Asento, " &
        //"EstadoSaludTitulares.ID_EstadoSalud , Cat_EstadoSalud.Descripcion, " &
        //"TelefonoCelular=isnull(LTRIM(RTRIM(TT.Pre_Celular + // // + TT.Lada_Celular  + // // + TT.Celular)) ,//SIN REGISTRAR//), " &
        //"TelefonoRecados=isnull(LTRIM(RTRIM(TT.Lada_Recados + // // + TT.Recados )) ,//SIN REGISTRAR//) " &
        //"From Titulares inner join Polizas on  Polizas.ID_Poliza = Titulares.ID_Poliza " &
        //"inner join Grupos_Fam on Grupos_Fam.ID_Poliza = Titulares.ID_Poliza and Grupos_Fam.ID_Grupo = Titulares.ID_Grupo " &
        //"inner join Cat_Conductos on Cat_Conductos.ID_Conducto = Titulares.ID_Conducto " &
        //"inner join Cat_StaTitular on Cat_StaTitular.ID_StaTitular = Titulares.ID_StaTitular " &
        //"inner join Cat_Identifica on Cat_Identifica.ID_Identifica = Titulares.ID_Identifica " &
        //"inner join Cat_CodPostales on Cat_CodPostales.ID_CP = Titulares.ID_CP " &
        //"inner join Cat_Municipios on Cat_Municipios.ID_Municipio = Cat_CodPostales.ID_Municipio " &
        //"inner join Cat_Estados on Cat_Estados.ID_Estado = Cat_Municipios.ID_Estado " &
        //"inner join Cat_Ciudades on Cat_Ciudades.ID_Ciudad = Cat_CodPostales.ID_Ciudad " &
        //"inner join Cat_Asentos on Cat_Asentos.ID_Asento = Cat_CodPostales.ID_Asento " &
        //"inner join Cat_Bancos on Cat_Bancos.ID_Banco = Titulares.ID_Banco " &
        //"left join EstadoSaludTitulares on Titulares.ID_Poliza = EstadoSaludTitulares.ID_Poliza " &
        //"left join DBCLIENTE.dbo.Cat_EstadoSalud Cat_EstadoSalud on Cat_EstadoSalud.ID_EstadoSalud = EstadoSaludTitulares.ID_EstadoSalud " &
        //"left join (select TOP 1 ID_Titular, Pre_Celular ,Lada_Celular ,Celular, Lada_Recados, Recados from Telefonos_Titulares ORDER BY ID_Telefonos_Titulares DESC) TT on Titulares.ID_Titular = TT.ID_Titular " &
        //"Where Polizas.ID_Poliza = " & parametros(0) & " " &
        //"and Titulares.ID_Titular = " & parametros(1) & " "
        //        //mafm fin 25082022 se agrega el telefono celular
        //        //Fin. Alexander Hdez 12/06/2012 Codigos Postales YA9A0F

        //        bGetDetTit = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetTit:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    //mafm ini 23082022 se agrega el telefono celular
        //    public Function bGetDetTitularesTel(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer

        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetDetTitTel

        //        sSql = "execute SPCA_CONS_TitularesTel " & parametros(0) & "," & parametros(1) & ""
        //        bGetDetTitularesTel = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetTitTel:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    //mafm fin 23082022 se agrega el telefono celular


        //    public Function bGetBenefs(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetBen

        //        bGetBenefs = TipoResultado.NoHayDatos
        //        //    //GERMAN CARRANCO RIVERA 06/06/2011  INICIO
        //        //    //Para que muestre todos los Endosos se quito and ID_Endoso in(7,8,9,10)
        //        //    sSql = "select Polizas.Fol_Poliza, Pol_Benefs.ID_Grupo, Pol_Benefs.Num_Benef, Pol_Benefs.ID_PolBenef," & _
        //        //            " Endoso = (select max(ID_HisEndoso) from Endosos where ID_Poliza = " & parametros(0) & " and ID_PolBenef =Pol_Benefs.ID_PolBenef )," & _
        //        //            " Pol_Benefs.Nom_Benef+// //+Pol_Benefs.ApP_Benef+// //+Pol_Benefs.ApM_Benef Beneficiario," & _
        //        //            " cast(Pol_Benefs.ID_Parentesco as char(2))+// - //+Cat_Parentescos.Abv_Parentesco Parentesco," & _
        //        //            " cast(Pol_Benefs.ID_StaPolBenef as char(2))+// - //+cast(Cat_StaPolBenef.Sta_PolBenef as char(20)) Sta_PolBenef," & _
        //        //            " cast(Pol_Benefs.ID_StaPgoBenef as char(2))+// - //+cast(Cat_StaPgoBenef.Sta_PgoBenef as char(20)) Sta_PgoBenef," & _
        //        //            " Pol_Benefs.Fecha_System Fecha_Endoso" & _
        //        //            " ,FolioEnrolamiento = 100" & _
        //        //            " ,Fecha_Enrolamiento = //// " & _
        //        //            " From Polizas, Pol_Benefs, Cat_Parentescos, Cat_StaPolBenef, Cat_StaPgoBenef" & _
        //        //            " " & _
        //        //            " Where Polizas.ID_Poliza = Pol_Benefs.ID_Poliza" & _
        //        //            " and Polizas.ID_Poliza = " & parametros(0) & _
        //        //            " and Cat_Parentescos.ID_Parentesco = Pol_Benefs.ID_Parentesco" & _
        //        //            " and Cat_StaPolBenef.ID_StaPolBenef = Pol_Benefs.ID_StaPolBenef" & _
        //        //            " and Cat_StaPgoBenef.ID_StaPgoBenef = Pol_Benefs.ID_StaPgoBenef"
        //        //    //GERMAN CARRANCO RIVERA FIN

        //        ////Nuevo Query para mostrar el folio de verificacion de sobrevivencia por biometria de voz
        //        ////JCMN Inicio
        //        //sSql = "select po.Fol_Poliza, benef.ID_Grupo, benef.Num_Benef, benef.ID_PolBenef," & _
        //        //" Endoso = (select max(ID_HisEndoso) from Endosos where ID_Poliza =" & parametros(0) & " and ID_PolBenef =benef.ID_PolBenef )," & _
        //        //" benef.Nom_Benef+// //+benef.ApP_Benef+// //+benef.ApM_Benef Beneficiario," & _
        //        //" cast(benef.ID_Parentesco as char(2))+// - //+pa.Abv_Parentesco Parentesco," & _
        //        //" cast(benef.ID_StaPolBenef as char(2))+// - //+cast(sta.Sta_PolBenef as char(20)) Sta_PolBenef," & _
        //        //" cast(benef.ID_StaPgoBenef as char(2))+// - //+cast(pago.Sta_PgoBenef as char(20)) Sta_PgoBenef," & _
        //        //" benef.Fecha_System Fecha_Endoso" & _
        //        //" ,FolioEnrolamiento = COALESCE(enr.ID_Enrolamiento,0)" & _
        //        //" ,Fecha_Enrolamiento = COALESCE(enr.Fecha_Enrolamiento,//1900-01-01//)" & _
        //        //" From Polizas po inner join Pol_Benefs benef on po.ID_Poliza = benef.ID_Poliza" & _
        //        //" inner join Cat_Parentescos pa on pa.ID_Parentesco = benef.ID_Parentesco" & _
        //        //" inner join Cat_StaPolBenef sta on sta.ID_StaPolBenef = benef.ID_StaPolBenef" & _
        //        //" inner join Cat_StaPgoBenef pago on pago.ID_StaPgoBenef = benef.ID_StaPgoBenef" & _
        //        //" left join Enrolamientos enr on enr.ID_PolBenef = benef.ID_PolBenef AND enr.ID_StaEnrolamiento=1 AND enr.ID_TipoEnrolamiento=1" & _
        //        //" Where po.ID_Poliza = " & parametros(0)
        //        ////JCMN Fin

        //        // 2016-10-17 Alexander. Comento el bloque de arriba de Julio y agrego este para que ordene por benef.ID_PolBenef
        //        //Inicio
        //        sSql = "select po.Fol_Poliza, benef.ID_Grupo, benef.Num_Benef, benef.ID_PolBenef," &
        //" Endoso = (select max(ID_HisEndoso) from Endosos where ID_Poliza =" & parametros(0) & " and ID_PolBenef =benef.ID_PolBenef )," &
        //" benef.Nom_Benef+// //+benef.ApP_Benef+// //+benef.ApM_Benef Beneficiario," &
        //" cast(benef.ID_Parentesco as char(2))+// - //+pa.Abv_Parentesco Parentesco," &
        //" cast(benef.ID_StaPolBenef as char(2))+// - //+cast(sta.Sta_PolBenef as char(20)) Sta_PolBenef," &
        //" cast(benef.ID_StaPgoBenef as char(2))+// - //+cast(pago.Sta_PgoBenef as char(20)) Sta_PgoBenef," &
        //" benef.Fecha_System Fecha_Endoso" &
        //" ,FolioEnrolamiento = COALESCE(enr.ID_Enrolamiento,0)" &
        //" ,Fecha_Enrolamiento = COALESCE(enr.Fecha_Enrolamiento,//1900-01-01//)" &
        //"/*Inicio Alexander 20210510. Se agregar lineas para consulta de los enrolamientos en ASB*/" &
        //",FolioEnrolamientoASB = COALESCE(enrASB.ID_Enrolamiento,0) ," &
        //"Fecha_EnrolamientoASB = COALESCE(enrASB.Fecha_Enrolamiento,//1900-01-01//)" &
        //"/*Fin Alexander 20210510*/" &
        //" From Polizas po inner join Pol_Benefs benef on po.ID_Poliza = benef.ID_Poliza" &
        //" inner join Cat_Parentescos pa on pa.ID_Parentesco = benef.ID_Parentesco" &
        //" inner join Cat_StaPolBenef sta on sta.ID_StaPolBenef = benef.ID_StaPolBenef" &
        //" inner join Cat_StaPgoBenef pago on pago.ID_StaPgoBenef = benef.ID_StaPgoBenef" &
        //" left join Enrolamientos enr on enr.ID_PolBenef = benef.ID_PolBenef AND enr.ID_StaEnrolamiento=1 AND enr.ID_TipoEnrolamiento=1" &
        //"/*Inicio Alexander 20210510. Se agregar lineas para consulta de los enrolamientos en ASB*/" &
        //"left join Enrolamientos enrASB on enrASB.ID_PolBenef = benef.ID_PolBenef AND enrASB.ID_StaEnrolamiento=1 AND enrASB.ID_TipoEnrolamiento=3" &
        //"/*Fin Alexander 20210510*/" &
        //" Where po.ID_Poliza = " & parametros(0) & " " &
        //" order by benef.ID_PolBenef "
        //        //Fin
        //        // 2016-10-17 Alexander. Comento el bloque de arriba de Julio y agrego este para que ordene por benef.ID_PolBenef


        //        bGetBenefs = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetBen:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function


        //    public Function bGetDetBen(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetBen

        //        bGetDetBen = TipoResultado.NoHayDatos
        //        //SDC 2007-01-25 Para mostrar datos de artículo 14 2004
        //        //I***JGIMATE - 16/11/2017 - AGREGAR CURP BENEFICIARIO
        //        //    sSql = "select Polizas.Fol_Poliza, Pol_Benefs.ID_Grupo," & _
        //        //            " Pol_Benefs.Nom_Benef+// //+Pol_Benefs.ApP_Benef+// //+Pol_Benefs.ApM_Benef Beneficiario," & _
        //        //            " Pol_Benefs.Fecha_Nacto, datediff(yy,Pol_Benefs.Fecha_Nacto, getdate()) Edad," & _
        //        //            " cast(Pol_Benefs.ID_EdoCivil as char(2))+// - //+Cat_EdoCivil.Edo_Civil Edo_Civil," & _
        //        //            " cast(Pol_Benefs.ID_Sexo as char(2)),Cat_Sexos.Abv_Sexo Sexo," & _
        //        //            " Pol_Benefs.Fecha_Suspension,Pol_Benefs.Fecha_Reactivacion,Pol_Benefs.Fecha_VentoIni," & _
        //        //            " Pol_Benefs.Fecha_Vento,cast(Pol_Benefs.ID_Invalidez as char(2))+// - //+Cat_Invalidez.Invalidez Invalidez," & _
        //        //            " cast(Pol_Benefs.ID_StaPolBenef as char(2))+// - //+cast(Cat_StaPolBenef.Sta_PolBenef as char(20)) Sta_PolBenef," & _
        //        //            " cast(Pol_Benefs.ID_StaPgoBenef as char(2))+// - //+cast(Cat_StaPgoBenef.Sta_PgoBenef as char(20)) Sta_PgoBenef," & _
        //        //            " cast(Pol_Benefs.ID_StaFiniq as char(2))+// - //+cast(Cat_StaFiniq.Sta_Finiq as char(20)) Sta_Finiq," & _
        //        //            " cast(Pol_Benefs.ID_Parentesco as char(2))+// - //+Cat_Parentescos.Abv_Parentesco Parentesco," & _
        //        //            " cast(Pol_Benefs.ID_Orfandad as char(2))+// - //+Cat_Orfandades.Abv_Orfandad Orfandad, Pol_Benefs.Fecha_System," & _
        //        //            " Sta_Incremento = isnull(Sta_Incremento, //SIN DERECHO//), Modulo_Incremento, Pension = case when Pol_Benefs.ID_Parentesco = 1 then Incremento_Pensiones.Pension else null end, Aguinaldo = case when Pol_Benefs.ID_Parentesco = 1 then Incremento_Pensiones.Aguinaldo else null end, Polizas.Fecha_IniVigInc04, Incremento_Pensiones.Fecha_DerInc, Incremento_Pensiones.Fecha_Solicitud, Incremento_Pensiones.Fecha_Deposito, Cat_Curp_Benef.CURP " & _
        //        //            " from Polizas, Pol_Benefs left join Incremento_Pensiones on Pol_Benefs.ID_PolBenef = Incremento_Pensiones.ID_PolBenef left join Cat_StaIncremento on Incremento_Pensiones.ID_StaIncremento = Cat_StaIncremento.ID_StaIncremento left join Cat_ModuloIncremento on Incremento_Pensiones.ID_ModuloIncremento = Cat_ModuloIncremento.ID_ModuloIncremento left join Cat_Curp_Benef ON Cat_Curp_Benef.ID_PolBenef=Pol_Benefs.ID_PolBenef, Cat_Parentescos,Cat_StaPolBenef, Cat_StaPgoBenef," & _
        //        //            " Cat_EdoCivil , Cat_Sexos, Cat_Invalidez, Cat_StaFiniq, Cat_Orfandades" & _
        //        //            " Where Polizas.ID_Poliza = Pol_Benefs.ID_Poliza" & _
        //        //            " and Polizas.ID_Poliza = " & parametros(0) & _
        //        //            " and Pol_Benefs.ID_PolBenef = " & parametros(1) & _
        //        //            " and Cat_Parentescos.ID_Parentesco = Pol_Benefs.ID_Parentesco" & _
        //        //            " and Cat_StaPolBenef.ID_StaPolBenef = Pol_Benefs.ID_StaPolBenef" & _
        //        //            " and Cat_StaPgoBenef.ID_StaPgoBenef = Pol_Benefs.ID_StaPgoBenef" & _
        //        //            " and Cat_EdoCivil.ID_EdoCivil = Pol_Benefs.ID_EdoCivil" & _
        //        //            " and Cat_Sexos.ID_Sexo = Pol_Benefs.ID_Sexo" & _
        //        //            " and Cat_Invalidez.ID_Invalidez = Pol_Benefs.ID_Invalidez and Cat_Orfandades.ID_Orfandad = Pol_Benefs.ID_Orfandad" & _
        //        //            " and Cat_StaFiniq.ID_StaFiniq = Pol_Benefs.ID_StaFiniq"
        //        //----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //        //Se modificó consulta, para mostrar calificación de riesgo de Pensionados por parte de PLD 28/12/2021
        //        //linea 13 y 14 se agrega dato: st_Nivel_Riesgo y fh_Fecha_Calif de tabla TDC002_DictBP_PLD (DBCAIRO)
        //        //----------------------------------------------------------------------------------------------------
        //        sSql = "select Polizas.Fol_Poliza, Pol_Benefs.ID_Grupo," &
        //                " Pol_Benefs.Nom_Benef+// //+Pol_Benefs.ApP_Benef+// //+Pol_Benefs.ApM_Benef Beneficiario," &
        //                " Pol_Benefs.Fecha_Nacto, datediff(yy,Pol_Benefs.Fecha_Nacto, getdate()) Edad," &
        //                " cast(Pol_Benefs.ID_EdoCivil as char(2))+// - //+Cat_EdoCivil.Edo_Civil Edo_Civil," &
        //                " cast(Pol_Benefs.ID_Sexo as char(2)),Cat_Sexos.Abv_Sexo Sexo," &
        //                " Pol_Benefs.Fecha_Suspension,Pol_Benefs.Fecha_Reactivacion,Pol_Benefs.Fecha_VentoIni," &
        //                " Pol_Benefs.Fecha_Vento,cast(Pol_Benefs.ID_Invalidez as char(2))+// - //+Cat_Invalidez.Invalidez Invalidez," &
        //                " cast(Pol_Benefs.ID_StaPolBenef as char(2))+// - //+cast(Cat_StaPolBenef.Sta_PolBenef as char(20)) Sta_PolBenef," &
        //                " cast(Pol_Benefs.ID_StaPgoBenef as char(2))+// - //+cast(Cat_StaPgoBenef.Sta_PgoBenef as char(20)) Sta_PgoBenef," &
        //                " cast(Pol_Benefs.ID_StaFiniq as char(2))+// - //+cast(Cat_StaFiniq.Sta_Finiq as char(20)) Sta_Finiq," &
        //                " cast(Pol_Benefs.ID_Parentesco as char(2))+// - //+Cat_Parentescos.Abv_Parentesco Parentesco," &
        //                " cast(Pol_Benefs.ID_Orfandad as char(2))+// - //+Cat_Orfandades.Abv_Orfandad Orfandad, Pol_Benefs.Fecha_System," &
        //                " Sta_Incremento = isnull(Sta_Incremento, //SIN DERECHO//), Modulo_Incremento, Pension = case when Pol_Benefs.ID_Parentesco = 1 then Incremento_Pensiones.Pension else null end, Aguinaldo = case when Pol_Benefs.ID_Parentesco = 1 then Incremento_Pensiones.Aguinaldo else null end, Polizas.Fecha_IniVigInc04, Incremento_Pensiones.Fecha_DerInc, Incremento_Pensiones.Fecha_Solicitud, Incremento_Pensiones.Fecha_Deposito, Cat_Curp_Benef.CURP, TDC002_DictBP_PLD.st_Nivel_Riesgo, TDC002_DictBP_PLD.fh_Fecha_Calif " &
        //                " from Polizas, Pol_Benefs left join Incremento_Pensiones on Pol_Benefs.ID_PolBenef = Incremento_Pensiones.ID_PolBenef left join Cat_StaIncremento on Incremento_Pensiones.ID_StaIncremento = Cat_StaIncremento.ID_StaIncremento left join Cat_ModuloIncremento on Incremento_Pensiones.ID_ModuloIncremento = Cat_ModuloIncremento.ID_ModuloIncremento left join Cat_Curp_Benef ON Cat_Curp_Benef.ID_PolBenef=Pol_Benefs.ID_PolBenef left join TDC002_DictBP_PLD ON TDC002_DictBP_PLD.cd_PolBenef=Pol_Benefs.ID_PolBenef, Cat_Parentescos,Cat_StaPolBenef, Cat_StaPgoBenef," &
        //                " Cat_EdoCivil , Cat_Sexos, Cat_Invalidez, Cat_StaFiniq, Cat_Orfandades" &
        //                " Where Polizas.ID_Poliza = Pol_Benefs.ID_Poliza" &
        //                " and Polizas.ID_Poliza = " & parametros(0) &
        //                " and Pol_Benefs.ID_PolBenef = " & parametros(1) &
        //                " and Cat_Parentescos.ID_Parentesco = Pol_Benefs.ID_Parentesco" &
        //                " and Cat_StaPolBenef.ID_StaPolBenef = Pol_Benefs.ID_StaPolBenef" &
        //                " and Cat_StaPgoBenef.ID_StaPgoBenef = Pol_Benefs.ID_StaPgoBenef" &
        //                " and Cat_EdoCivil.ID_EdoCivil = Pol_Benefs.ID_EdoCivil" &
        //                " and Cat_Sexos.ID_Sexo = Pol_Benefs.ID_Sexo" &
        //                " and Cat_Invalidez.ID_Invalidez = Pol_Benefs.ID_Invalidez and Cat_Orfandades.ID_Orfandad = Pol_Benefs.ID_Orfandad" &
        //                " and Cat_StaFiniq.ID_StaFiniq = Pol_Benefs.ID_StaFiniq"
        //        //------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //        bGetDetBen = EjecutaSql(rsData, sConn, sSql, rsErrAdo)
        //        //F***JGIMATE - 16/11/2017 - AGREGAR CURP BENEFICIARIO
        //        Exit Function

        //errGetDetBen:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetBeneficios(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetBeneficios

        //        //SDC 2007-01-30 Nuevo esquema de consultas
        //        bGetBeneficios = TipoResultado.NoHayDatos
        //        rsErrAdo = Nothing
        //        sSql = "select cast(Cat_Beneficios.Beneficio as char(40)) Concepto," &
        //            " Pol_Beneficios.Imp_Beneficio, Fecha_Fin = case when Pol_Beneficios.ID_StaBeneficio in (2,4) or Pol_Beneficios.ID_Beneficio in (27,36,37) then Pol_Beneficios.Fecha_Fin else null end, " &
        //            " cast(Cat_StaBeneficios.Sta_Beneficio as char(3)) Sta_Beneficio," &
        //            " cast(Cat_Monedas.Descripcion as char(1)) Moneda" &
        //            " From Pol_Beneficios, Cat_Beneficios, Cat_StaBeneficios, Cat_Monedas, Cat_BenPgoNom " &
        //            " Where Pol_Beneficios.ID_PolBenef = " & parametros(1) &
        //            " and Cat_Beneficios.ID_Beneficio = Pol_Beneficios.ID_Beneficio" &
        //            " and Cat_StaBeneficios.ID_StaBeneficio = Pol_Beneficios.ID_StaBeneficio" &
        //            " and Cat_Monedas.ID_Moneda = Pol_Beneficios.ID_Moneda " &
        //            " and Cat_BenPgoNom.ID_Beneficio = Pol_Beneficios.ID_Beneficio" &
        //            " order by case when Pol_Beneficios.ID_StaBeneficio in (1,3,6) then Pol_Beneficios.ID_StaBeneficio else 10 end, Pol_Beneficios.ID_StaBeneficio, Cat_BenPgoNom.ID_BenPgoNom, Pol_Beneficios.Imp_Beneficio desc "
        //        bGetBeneficios = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetBeneficios:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetRecibos(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetRecibos

        //        bGetRecibos = TipoResultado.NoHayDatos
        //        sSql = "select Polizas.Fol_Poliza, Recibos.Fol_Recibo,Polizas.Num_SegSocial NSS, Polizas.Nom_Aseg+// //+Polizas.ApP_Aseg+// //+Polizas.ApM_Aseg Asegurado, Recibos.ID_Recibo" &
        //            " From Recibos, Polizas" &
        //            " Where Recibos.ID_Poliza = Polizas.ID_Poliza" &
        //            " and Polizas.ID_Poliza = " & parametros(0) &
        //            " order by Recibos.Fol_Recibo"
        //        bGetRecibos = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetRecibos:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetDetRec(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetRecibos

        //        bGetDetRec = TipoResultado.NoHayDatos
        //        sSql = "select Polizas.Fol_Poliza, Recibos.Fol_Recibo,Recibos.Fecha_ERecibo,Recibos.Monto_Total," &
        //            " Recibos.Fecha_Pago,Recibos.Fecha_Baja," &
        //            " cast(Recibos.ID_TipoRecibo as char(2))+// - //+cast(Cat_TpoRecibos.Tipo_Recibo as char(15)) Tipo_Recibo," &
        //            " cast(Recibos.ID_StaRecibo as char(2))+// - //+cast(Cat_StaRecibo.Sta_Recibo as char(15)) Sta_Recibo," &
        //            " Remesa=(select distinct Remesas.Fol_Remesa from Remesas, MRemesas where Remesas.ID_Remesa = MRemesas.ID_Remesa" &
        //            " and MRemesas.ID_TpoDocRem=2 and Remesas.ID_TpoRem=4 and cast(MRemesas.Folio as integer)=Recibos.ID_Recibo)," &
        //            " StaCont=case when Recibos.Sta_Contabilidad is null then //SIN CONTABILIZAR//" &
        //            " when Recibos.Sta_Contabilidad=//C// then //CONTABILIZADO// end" &
        //            " From Recibos, Polizas, Cat_TpoRecibos, Cat_StaRecibo" &
        //            " Where Recibos.ID_Poliza = Polizas.ID_Poliza" &
        //            " and Polizas.ID_Poliza = " & parametros(0) &
        //            " and Recibos.ID_Recibo = " & parametros(1) &
        //            " and Cat_TpoRecibos.ID_TipoRecibo = Recibos.ID_TipoRecibo" &
        //            " and Cat_StaRecibo.ID_StaRecibo = Recibos.ID_StaRecibo"
        //        bGetDetRec = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetRecibos:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetPrestamos(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetPrestamos

        //        bGetPrestamos = TipoResultado.NoHayDatos
        //        sSql = "SELECT Prestamos.ID_Grupo, Prestamos.ID_Prestamo,Prestamos.Saldo," &
        //            " cast(Prestamos.ID_StaPtmo as char(2))+// - //+Cat_StaPtmo.Status_Prestamo Sta_Ptmo," &
        //            " MesesFalt=(Prestamos.Num_Decto-Prestamos.Decto_Aplicados),Prestamos.Importe_Decto," &
        //            " Prestamos.Fecha_Deposito, Prestamos.Fecha_Captura" &
        //            " From Prestamos, Cat_StaPtmo" &
        //            " Where Prestamos.ID_StaPtmo = Cat_StaPtmo.ID_StaPtmo" &
        //            " and Prestamos.ID_Poliza = " & parametros(0)
        //        bGetPrestamos = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetPrestamos:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetDetPrest(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetPrest

        //        bGetDetPrest = TipoResultado.NoHayDatos
        //        sSql = "SELECT Prestamos.ID_Grupo, Prestamos.ID_Prestamo,Prestamos.Saldo, " &
        //            " cast(Prestamos.ID_StaPtmo as char(2))+// - //+Cat_StaPtmo.Status_Prestamo Sta_Ptmo," &
        //            " MesesFalt=(Prestamos.Num_Decto-Prestamos.Decto_Aplicados),Prestamos.Importe_Decto,Prestamos.Fecha_Deposito," &
        //            " Prestamos.Saldo_Inicial, Prestamos.Num_Decto,Prestamos.Decto_Aplicados," &
        //            " Prestamos.Periodo_Inicio, Prestamos.Anio_Inicio,Prestamos.Tipo," &
        //            " Prestamos.Fecha_Captura , Prestamos.Cobrar_UltPago" &
        //            " From Prestamos, Cat_StaPtmo" &
        //            " Where Prestamos.ID_StaPtmo = Cat_StaPtmo.ID_StaPtmo" &
        //            " and Prestamos.ID_Poliza = " & parametros(0) &
        //            " and Prestamos.ID_Prestamo = " & parametros(1)
        //        bGetDetPrest = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetPrest:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function


        //    public Function bGetDetPtmoFVI(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetPrest

        //        bGetDetPtmoFVI = TipoResultado.NoHayDatos
        //        sSql = "select pri.ID_Prestamo, pri.Concepto, pri.ID_Poliza, cat.Status_Prestamo, pri.Saldo_Inicial, pri.Saldo, pri.Plazo, " &
        //            " pri.Importe , pri.DescApli, desci.Num_Desc, desci.Fch_Desc, desci.Observaciones " &
        //            " from Prestamos_FOVISSSTE pri left join Descuentos_ISSSTE desci on pri.ID_Poliza = desci.ID_Poliza " &
        //            "                                   and pri.ID_Prestamo=desci.ID_Prestamo " &
        //            "                              inner join Cat_StaPtmo cat on cat.ID_StaPtmo=pri.ID_StaPtmo " &
        //            " Where pri.ID_Poliza = " & parametros(0) & " And pri.ID_Prestamo = " & parametros(1)

        //        bGetDetPtmoFVI = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetPrest:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function


        //    public Function bGetDetAbon(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetAbon

        //        bGetDetAbon = TipoResultado.NoHayDatos
        //        sSql = "select ID_Abono,Imp_Abono,Fecha_Captura,Fecha_System" &
        //            " From Abonos" &
        //            " where ID_Prestamo = " & parametros(0)
        //        bGetDetAbon = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetAbon:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetGrupo(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetGrupo

        //        bGetGrupo = TipoResultado.NoHayDatos

        //        sSql = "select p.Fol_Poliza, g.ID_Grupo, "
        //        sSql = sSql + "    Grupo_Familiar = convert(varchar, g.ID_Grupo) + // - // + Sta_Pago + // (// + Tpo_GpoFam "
        //        sSql = sSql + "        + //) - // + isnull(Nom_Titular + // // + ApP_Titular + // // + ApM_Titular, //SIN TITULAR DE COBRO ACTIVO//) "
        //        sSql = sSql + "from Polizas p left join Grupos_Fam g on p.ID_Poliza = g.ID_Poliza "
        //        sSql = sSql + "     left join Titulares t on g.ID_Poliza = t.ID_Poliza and g.ID_Grupo = t.ID_Grupo and ID_StaTitular = 1 "
        //        sSql = sSql + "     left join Cat_TpoGpoFam ct on g.ID_TpoGpoFam = ct.ID_TpoGpoFam "
        //        sSql = sSql + "     left join Cat_StaPago cs on g.ID_StaPago = cs.ID_StaPago "
        //        sSql = sSql + "where   p.ID_Poliza = " & parametros(0) & " "
        //        sSql = sSql + "order by g.ID_Grupo "

        //        bGetGrupo = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetGrupo:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetPagos(rsPagos As ADODB.Recordset, rsROPC As ADODB.Recordset, rsEspecial As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        Dim sSql As String
        //        Dim errVB As String
        //        Dim cnnConexion As ADODB.Connection
        //        On Error GoTo errGetPagos

        //        cnnConexion = New ADODB.Connection
        //        cnnConexion.Open(sConn)
        //        rsErrAdo = Nothing

        //        //Traer información de pagos normales de nómina
        //        sSql = "select  pg.ID_Pago, pg.ID_Grupo, pg.APagado, pg.PPagado, pg.Anio, pg.Periodo, "
        //        sSql = sSql + "     Tipo_Pago = ctc.Tipo_Corrida + // - //+  ctp.Abv_TipoPago, "
        //        sSql = sSql + "     Sta_Pago = convert(varchar(2), pg.ID_StaPago) + // - //+ convert(varchar(40), csp.Abv_StaPago), "
        //        sSql = sSql + "     pg.Imp_Pago, Fecha_StaPago = case when ctp.ID_TpoROPC in (2,3) and pg.ID_StaPago = 7 then pg.Fecha_StaPago else null end, pg.Fecha_System, pg.ID_StaPago, c.ID_TipoCorrida "
        //        sSql = sSql + "from Pagos pg inner join Cat_TipoPagos ctp on pg.ID_TipoPago = ctp.ID_TipoPago and ID_ConsultaPago = 0 "
        //        sSql = sSql + "     left join Cat_StaPagos csp on pg.ID_StaPago = csp.ID_StaPago "
        //        sSql = sSql + "     left join Corridas c on pg.ID_Corrida = c.ID_Corrida "
        //        sSql = sSql + "     left join Cat_TipoCorridas ctc on c.ID_TipoCorrida = ctc.ID_TipoCorrida "
        //        sSql = sSql + "where    pg.ID_Poliza = " & parametros(0) & " "
        //        sSql = sSql + "order by pg.APagado desc, pg.PPagado desc, pg.ID_Grupo asc, pg.Fecha_System, pg.Imp_Pago desc "

        //        rsPagos = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        //Traer información de ROPC pagada
        //        sSql = "select  pg.ID_Pago, pg.ID_Grupo, pg.APagado, pg.PPagado, pg.Anio, pg.Periodo, "
        //        sSql = sSql + "     Tipo_Pago = ctc.Tipo_Corrida + // - //+  ctp.Abv_TipoPago, "
        //        sSql = sSql + "     Sta_Pago = convert(varchar(2), pg.ID_StaPago) + // - //+ convert(varchar(40), csp.Abv_StaPago), "
        //        sSql = sSql + "     pg.Imp_Pago, pg.Fecha_System, pg.ID_StaPago, c.ID_TipoCorrida "
        //        sSql = sSql + "from Pagos pg inner join Cat_TipoPagos ctp on pg.ID_TipoPago = ctp.ID_TipoPago and ID_ConsultaPago = 1 "
        //        sSql = sSql + "     left join Cat_StaPagos csp on pg.ID_StaPago = csp.ID_StaPago "
        //        sSql = sSql + "     left join Corridas c on pg.ID_Corrida = c.ID_Corrida "
        //        sSql = sSql + "     left join Cat_TipoCorridas ctc on c.ID_TipoCorrida = ctc.ID_TipoCorrida "
        //        sSql = sSql + "where    pg.ID_Poliza = " & parametros(0) & " "
        //        sSql = sSql + "     and pg.ID_StaPago not in (12,13) "
        //        sSql = sSql + "order by pg.APagado desc, pg.PPagado desc, pg.ID_Grupo asc, pg.Imp_Pago desc "

        //        rsROPC = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        //Traer información de pagos especiales
        //        sSql = "select  pg.ID_Pago, pg.ID_Grupo, pg.APagado, pg.PPagado, pg.Anio, pg.Periodo, "
        //        sSql = sSql + "     Tipo_Pago = ctc.Tipo_Corrida + // - //+  ctp.Abv_TipoPago, "
        //        sSql = sSql + "     Sta_Pago = convert(varchar(2), pg.ID_StaPago) + // - //+ convert(varchar(40), csp.Abv_StaPago), "
        //        sSql = sSql + "     pg.Imp_Pago, pg.Fecha_System" //Se agrega el campo pg.Fecha_System a la linea de consulta ya que se comento su código abajo Francisco Arreola 15022023

        //        // rolando coellar serrano , gastos funerarios issste , 26 oct 2011
        //        //se comenta este codigo ya que npo tiene ningún uso y proboca perdida de información en la consulta Francisco Arreola 15022023
        //        //       sSql = sSql + "Fecha_System = case when pg.ID_TipoPago =(SELECT     ID_TipoPago FROM dbo.view_GastFunISSSTE_Configuracion) then "
        //        //
        //        //       sSql = sSql + "( SELECT top 1 dbo.Remesas.Fecha_Server "
        //        //       sSql = sSql + " From "
        //        //       sSql = sSql + " dbo.MRemesas INNER JOIN "
        //        //       sSql = sSql + " dbo.Remesas ON dbo.MRemesas.ID_Remesa = dbo.Remesas.ID_Remesa INNER JOIN "
        //        //       sSql = sSql + " dbo.Polizas ON dbo.MRemesas.Fol_Docto  = " & parametros(0) & " "
        //        //       sSql = sSql + " WHERE  Polizas.ID_Poliza = " & parametros(0) & "  And "
        //        //       sSql = sSql + " dbo.Remesas.ID_TpoRem  =(SELECT ID_TpoRem FROM dbo.view_GastFunISSSTE_Configuracion) and dbo.MRemesas.ID_MovConta =2  "
        //        //       sSql = sSql + "                     order by 1 desc "
        //        //       sSql = sSql + " )   "
        //        //
        //        //        sSql = sSql + "  else "
        //        //        sSql = sSql + "pg.Fecha_System"
        //        //        sSql = sSql + " end "
        //        // rolando coellar serrano , gastos funerarios issste , 26 oct 2011
        //        //Fin Francisco Arreola 15022023
        //        sSql = sSql + ",pg.ID_StaPago,pg.Fecha_System,c.ID_TipoCorrida "
        //        sSql = sSql + " from Pagos pg inner join Cat_TipoPagos ctp on pg.ID_TipoPago = ctp.ID_TipoPago and ID_ConsultaPago = 2 "
        //        sSql = sSql + "     left join Cat_StaPagos csp on pg.ID_StaPago = csp.ID_StaPago "
        //        sSql = sSql + "     left join Corridas c on pg.ID_Corrida = c.ID_Corrida "
        //        sSql = sSql + "     left join Cat_TipoCorridas ctc on c.ID_TipoCorrida = ctc.ID_TipoCorrida "
        //        sSql = sSql + "where    pg.ID_Poliza = " & parametros(0) & " "
        //        sSql = sSql + "     and pg.ID_StaPago not in (12,13) "
        //        sSql = sSql + "order by pg.APagado desc, pg.PPagado desc, pg.ID_Grupo asc, pg.Imp_Pago desc "

        //        rsEspecial = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        bGetPagos = True
        //        Exit Function

        //errGetPagos:
        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        rsPagos = Nothing
        //        rsROPC = Nothing
        //        rsEspecial = Nothing
        //        bGetPagos = False
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetDetPagos(rsData As ADODB.Recordset, rsTotales As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        Dim sSql As String
        //        Dim errVB As String
        //        Dim cnnConexion As ADODB.Connection
        //        On Error GoTo errGetDetPagos

        //        cnnConexion = New ADODB.Connection
        //        cnnConexion.Open(sConn)
        //        rsErrAdo = Nothing

        //        //Buscar datos de cuenta
        //        sSql = "select     Detalle = isnull(case when ID_CatDPago = 3 then // - // + d.Detalle  "
        //        sSql = sSql + "         else // - // + substring(d.Detalle, charindex(//,//, d.Detalle, charindex(//,//, d.Detalle)+1), charindex(//,//, d.Detalle, charindex(//,//, d.Detalle, charindex(//,//, d.Detalle)+1)+1) - charindex(//,//, d.Detalle, charindex(//,//, d.Detalle)+1)) end, ////), "
        //        sSql = sSql + "    Conducto "
        //        sSql = sSql + "from    Pagos pg left join DPagos d on pg.Folio_DPago = d.Folio_DPago and pg.Fecha_System = d.Fecha_System and ID_CatDPago in (3,5) "
        //        sSql = sSql + "        left join Cat_Conductos cc on pg.ID_Conducto = cc.ID_Conducto "
        //        sSql = sSql + "where   pg.ID_Pago = " & parametros(1)
        //        rsData = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)
        //        parametros(2) = rsData.Fields("Conducto").Value & Replace(Replace(rsData.Fields("Detalle").Value, ",", ""), "/", " / ")

        //        //Crear tabla con nombres de beneficiarios
        //        sSql = "drop table #benefs"
        //        On Error Resume Next
        //        cnnConexion.Execute(sSql)

        //        //Crear tabla con nombres de beneficiarios
        //        sSql = "select distinct ID_PolBenef into #benefs from MPagos where ID_Pago = " & parametros(1)
        //        cnnConexion.Execute(sSql)

        //        //Alexander Hdez 09/09/2010 se omenta esta linea y se agrega la misma pero agregandole un "IN" en lugar de solo el ID 4
        //        //sSql = sSql & vbCrLf & "     Aguinaldo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 4), "

        //        //Traer detalle de pagos de forma Pgo_Nom
        //        sSql = "select      "
        //        sSql = sSql & vbCrLf & "     Prestamos_ISS =  (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 23), " //SGRR 2010-01-13 Descuentos de Ptmos ISSSTE
        //        sSql = sSql & vbCrLf & "     Prestamos_FVIA = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 24), " //SGRR 2010-01-13 Descuentos de Ptmos FOVISSSTE Amortización
        //        sSql = sSql & vbCrLf & "     Prestamos_FVIS = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 25), " //SGRR 2010-01-13 Descuentos de Ptmos FOVISSSTE Seguro de Daños
        //        sSql = sSql & vbCrLf & "     Nombre = convert(varchar(2), Num_Benef) + // - // + isnull(Nom_Benef, ////) +  // // + isnull(ApP_Benef, ////) +  // // + isnull(ApM_Benef, ////), "
        //        sSql = sSql & vbCrLf & "     Pagos_Vencidos = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 0), "
        //        sSql = sSql & vbCrLf & "     Pension_Basica = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 1), "
        //        sSql = sSql & vbCrLf & "     BAU = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom in (21,30)), " //ASB - 2010-11-23 - SE SUMA BAU POR ENDOSO
        //        sSql = sSql & vbCrLf & "     Aguinaldo_BAU = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 22), "
        //        sSql = sSql & vbCrLf & "     Art14_2002 = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 2), "
        //        sSql = sSql & vbCrLf & "     Art14_2004 = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 3), "
        //        sSql = sSql & vbCrLf & "     Aguinaldo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom in (4,26,27,28,29)), " //Alexander 09/09/2010 para que muestre aguinaldo o gratificacion
        //        sSql = sSql & vbCrLf & "     Ag_Art14_2004 = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 5), "
        //        sSql = sSql & vbCrLf & "     Finiquito = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 6), "
        //        sSql = sSql & vbCrLf & "     Finiquito_Art14_2004 = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 7), "
        //        sSql = sSql & vbCrLf & "     Retroactivo_ROPC_HS = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 8), "
        //        sSql = sSql & vbCrLf & "     BAMI = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 9), "
        //        sSql = sSql & vbCrLf & "     BAMI_Aguinaldo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 10), "
        //        sSql = sSql & vbCrLf & "     BAMI_Finiquito = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 11), "
        //        sSql = sSql & vbCrLf & "     Pension_Adicional = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 12), "
        //        sSql = sSql & vbCrLf & "     Aguinaldo_Adicional = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 13), "
        //        sSql = sSql & vbCrLf & "     Ayuda_Escolar = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 14), "
        //        sSql = sSql & vbCrLf & "     Abono_Grupo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 15), "

        //        //<inicio> rcs 9  de junio 2011
        //        sSql = sSql & vbCrLf & "     Descuentos_Grupo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 16), "
        //        //<inicio> rcs 9  de junio 2011
        //        sSql = sSql & vbCrLf & "     Descuentos_Otros = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 17), "
        //        sSql = sSql & vbCrLf & "     Prestamos_ATM = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 18), "
        //        sSql = sSql & vbCrLf & "     Prestamos_Seguros = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 19), "
        //        sSql = sSql & vbCrLf & "     Total = (select sum(Imp_PBenef) from MPagos m where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef), "
        //        sSql = sSql & vbCrLf & "     Prestamos_PB = (select sum(Imp_PBenef) from MPagos m where m.ID_PolBenef = pb.ID_PolBenef AND pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = 129), " //CEFB 2009-08-03 Descuento del primer Prestamo
        //        sSql = sSql & vbCrLf & "     Prestamos_PB2 = (select sum(Imp_PBenef) from MPagos m where m.ID_PolBenef = pb.ID_PolBenef AND pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = 139), " //CEFB 2009-08-03 Descuento del segundo Prestamo
        //        sSql = sSql & vbCrLf & "     GtosFun_Endoso = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 31) " //ASB - 2010-11-23 - Gtos fun por Endoso
        //        sSql = sSql & vbCrLf & "from Pagos pg inner join #benefs m on pg.ID_Pago = " & parametros(1) & " "
        //        sSql = sSql & vbCrLf & "     left join Pol_Benefs pb on pb.ID_PolBenef = m.ID_PolBenef "
        //        sSql = sSql & vbCrLf & "order by Num_Benef "

        //        rsData = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        //Alexander Hdez 09/09/2010 se omenta esta linea y se agrega la misma pero agregandole un "IN"
        //        //sSql = sSql & vbCrLf & "     Aguinaldo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 4), "
        //        //Traer totales por columna
        //        sSql = "select   "
        //        sSql = sSql & vbCrLf & "     Prestamos_ISS =  (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 23), " // SGRR 2010-01-13 Descuentos de Ptmos ISSSTE
        //        sSql = sSql & vbCrLf & "     Prestamos_FVIA = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 24), " // SGRR 2010-01-13 Descuentos de Ptmos FOVISSSTE Amortización
        //        sSql = sSql & vbCrLf & "     Prestamos_FVIS = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 25), " // SGRR 2010-01-13 Descuentos de Ptmos FOVISSSTE Seguro de Daños
        //        sSql = sSql & vbCrLf & "     Nombre = //TOTAL//, "
        //        sSql = sSql & vbCrLf & "     Pagos_Vencidos = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 0), "
        //        sSql = sSql & vbCrLf & "     Pension_Basica = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 1), "
        //        sSql = sSql & vbCrLf & "     BAU = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom IN (21,30)), " //asb - 2010-11-23 - se suma con bau por endoso
        //        sSql = sSql & vbCrLf & "     Aguinaldo_BAU = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 22), "
        //        sSql = sSql & vbCrLf & "     Art14_2002 = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 2), "
        //        sSql = sSql & vbCrLf & "     Art14_2004 = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 3), "
        //        sSql = sSql & vbCrLf & "     Aguinaldo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom in (4,26,27,28,29)), " // Alexander 09/09/2010 para que muestre aguinaldo o gratificacion
        //        sSql = sSql & vbCrLf & "     Ag_Art14_2004 = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 5), "
        //        sSql = sSql & vbCrLf & "     Finiquito = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 6), "
        //        sSql = sSql & vbCrLf & "     Finiquito_Art14_2004 = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 7), "
        //        sSql = sSql & vbCrLf & "     Retroactivo_ROPC_HS = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 8), "
        //        sSql = sSql & vbCrLf & "     BAMI = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 9), "
        //        sSql = sSql & vbCrLf & "     BAMI_Aguinaldo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 10), "
        //        sSql = sSql & vbCrLf & "     BAMI_Finiquito = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 11), "
        //        sSql = sSql & vbCrLf & "     Pension_Adicional = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 12), "
        //        sSql = sSql & vbCrLf & "     Aguinaldo_Adicional = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 13), "
        //        sSql = sSql & vbCrLf & "     Ayuda_Escolar = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 14), "
        //        sSql = sSql & vbCrLf & "     Abono_Grupo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 15), "
        //        //<inicio> rcs 9  de junio 2011
        //        sSql = sSql & vbCrLf & "     Descuentos_Grupo = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 16), "
        //        //<inicio> rcs 9  de junio 2011

        //        sSql = sSql & vbCrLf & "     Descuentos_Otros = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 17), "
        //        sSql = sSql & vbCrLf & "     Prestamos_ATM = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 18), "
        //        sSql = sSql & vbCrLf & "     Prestamos_Seguros = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 19), "
        //        sSql = sSql & vbCrLf & "     Total = (select sum(Imp_PBenef) from MPagos m where pg.ID_Pago = m.ID_Pago), "
        //        sSql = sSql & vbCrLf & "     Prestamos_PB = (select sum(Imp_PBenef) from MPagos m where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = 129), " //CEFB 2009-08-03 Descuentos de primer Prestamo
        //        sSql = sSql & vbCrLf & "     Prestamos_PB2 = (select sum(Imp_PBenef) from MPagos m where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = 139), " //CEFB 2009-08-03 Descuentos de segundo Prestamo
        //        sSql = sSql & vbCrLf & "     GtosFun_Endoso = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 31) " //ASB - 2010-11-23 - GTOS FUN POR ENDOSO
        //        sSql = sSql & vbCrLf & " from Pagos pg "
        //        sSql = sSql & vbCrLf & "where pg.ID_Pago = " & parametros(1) & " "

        //        rsTotales = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        bGetDetPagos = True
        //        Exit Function

        //errGetDetPagos:
        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        bGetDetPagos = False
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function



        //    public Function bGetBeneficiosA(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetBeneficios

        //        bGetBeneficiosA = TipoResultado.NoHayDatos
        //        sSql = " Select p.Fol_Poliza, p.Num_SegSocial, p.Nom_Aseg  + // // + p.ApP_Aseg + // // + p.ApM_Aseg as Asegurado, cp.Pension, cr.Ramo, p.Fecha_Emision " &
        //            " from Polizas p, Cat_Ramos cr, Cat_Pensiones cp " &
        //            " Where p.ID_Ramo = cr.ID_Ramo " &
        //            " and p.ID_Pension = cp.ID_Pension " &
        //            " and p.ID_Poliza = " & parametros(0)

        //        bGetBeneficiosA = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetBeneficios:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetDetBeneficiosA(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetBeneficios

        //        bGetDetBeneficiosA = TipoResultado.NoHayDatos
        //        if parametros(2) != 1 Then   //IMSS 97. GRCC: 20/01/2009
        //            //SDC 2007-01-29 Para mostrar correctamente la edad limite escogida en AE
        //            sSql = "select  "
        //            sSql = sSql + " CASE "
        //            sSql = sSql + " WHEN cb.ID_Beneficio  IN (55,56,27,36,37,57) "
        //            sSql = sSql + "     THEN cast(cb.ID_Beneficio as char(3))+// - //+ cb.Beneficio + // (// + convert(varchar(2), cop.Limite) + //)// "
        //            sSql = sSql + "     ELSE cast(cb.ID_Beneficio as char(3))+// - //+ cb.Beneficio "
        //            sSql = sSql + " END AS Beneficio, 1,1, "
        //            sSql = sSql + " isnull(p.Pension_MensualBAU,0) as Pension_MensualBAU,pb.Fecha_System as Fecha  "
        //            sSql = sSql + " from     Polizas p inner join Pol_BASelec pb on p.ID_Poliza = pb.ID_Poliza "
        //            sSql = sSql + "         inner join Cat_Beneficios cb on pb.ID_Beneficio = cb.ID_Beneficio "
        //            sSql = sSql + "         left join Cat_OpEscolar cop on pb.Op_Escolar = cop.Op_Escolar and p.Fecha_ABase between cop.Fecha_Inicial and cop.Fecha_Final "
        //            sSql = sSql + " where    cb.ID_TpoBenef = 2  "
        //            sSql = sSql + "         and pb.ID_Poliza = " & parametros(0)
        //        Else
        //            //SDC 2007-01-29 Para mostrar correctamente la edad limite escogida en AE
        //            sSql = "select  "
        //            sSql = sSql + " CASE "
        //            sSql = sSql + " WHEN cb.ID_Beneficio  IN (55,56,27,36,37,57) "
        //            sSql = sSql + "     THEN cast(cb.ID_Beneficio as char(3))+// - //+ cb.Beneficio + // (// + convert(varchar(2), cop.Limite) + //)// "
        //            sSql = sSql + "     ELSE cast(cb.ID_Beneficio as char(3))+// - //+ cb.Beneficio "
        //            sSql = sSql + " END AS Beneficio, "
        //            if parametros(1) = 1 Then   //Bancomer
        //                sSql = sSql + " isnull(pb.Pje_Escogido/100,0) * isnull(pb.Pje_BenAdic,0) as Porcentaje,  "
        //                sSql = sSql + " isnull(pb.Pje_Escogido/100,0) * isnull(pb.Dias_Beneficio,0) as Dias , pb.Fecha_System as Fecha, cb.Comentario,isnull(p.Pension_MensualBAU,0)  "
        //            Elseif parametros(1) = 2 Then //BBV
        //                sSql = sSql + " isnull(pb.Pje_BenAdic,0) as Porcentaje,  "
        //                sSql = sSql + " isnull(pb.Dias_Beneficio,0) as Dias , pb.Fecha_System as Fecha, cb.Comentario "
        //            End if
        //            sSql = sSql + "from     Polizas p inner join Pol_BASelec pb on p.ID_Poliza = pb.ID_Poliza "
        //            sSql = sSql + "         inner join Cat_Beneficios cb on pb.ID_Beneficio = cb.ID_Beneficio "
        //            sSql = sSql + "         left join Cat_OpEscolar cop on pb.Op_Escolar = cop.Op_Escolar and p.Fecha_ABase between cop.Fecha_Inicial and cop.Fecha_Final "
        //            sSql = sSql + "where    cb.ID_TpoBenef = 2  "
        //            sSql = sSql + "         and pb.ID_Poliza = " & parametros(0)
        //        End if
        //        bGetDetBeneficiosA = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetBeneficios:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    //SDC 2007-01-29 Nuevo esquema de consulta, Endosos separados de Suspensiones
        //    public Function bGetEndosos(rsEndosos As ADODB.Recordset, rsSuspAct As ADODB.Recordset, rsSuspHist As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        Dim sSql As String
        //        Dim errVB As String
        //        Dim cnnConexion As ADODB.Connection
        //        On Error GoTo errGetEndosos

        //        cnnConexion = New ADODB.Connection
        //        cnnConexion.Open(sConn)
        //        rsErrAdo = Nothing

        //        sSql = " select p.ID_Poliza, "
        //        sSql = sSql + "     e.ID_HisEndoso, "
        //        sSql = sSql + "     c.Endoso, "
        //        sSql = sSql + "     e.Fecha_System, "
        //        sSql = sSql + "     u.Usuario "
        //        sSql = sSql + "from Endosos e inner join Polizas p on e.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + "     inner join Cat_Endosos c on e.ID_Endoso = c.ID_Endoso "
        //        sSql = sSql + "     inner join Usuarios u on e.ID_Usuario = u.ID_Usuario "
        //        sSql = sSql + "where e.ID_Poliza = " & parametros(0)
        //        sSql = sSql + "order by e.Fecha_Server desc "
        //        rsEndosos = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        sSql = " select p.ID_Poliza, "
        //        sSql = sSql + "     s.ID_Grupo, s.ID_Suspension, "
        //        sSql = sSql + "     Motivo_Suspension = upper(cm.Motivo_Suspension), "
        //        sSql = sSql + "     s.Fecha_System, "
        //        sSql = sSql + "     u.Usuario "
        //        sSql = sSql + "from Suspension_PolGpo s inner join Polizas p on s.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + "     inner join Cat_MotivoSus cm on s.ID_MotivoSus = cm.ID_MotivoSus "
        //        sSql = sSql + "     inner join Cat_StaSuspension cs on s.ID_StaSuspension = cs.ID_StaSuspension "
        //        sSql = sSql + "     inner join Usuarios u on s.ID_Usuario = u.ID_Usuario "
        //        sSql = sSql + "where s.ID_Poliza = " & parametros(0)
        //        sSql = sSql + "     and s.ID_StaSuspension = 1 "
        //        sSql = sSql + "order by s.Fecha_Server desc "
        //        rsSuspAct = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        sSql = " select p.ID_Poliza, "
        //        sSql = sSql + "     s.ID_Grupo, s.ID_Suspension, "
        //        sSql = sSql + "     Motivo_Suspension = upper(cm.Motivo_Suspension), "
        //        sSql = sSql + "     Sta_Suspension = upper(cs.Sta_Suspension), "
        //        sSql = sSql + "     s.Fecha_System, "
        //        sSql = sSql + "     u.Usuario "
        //        sSql = sSql + "from Suspension_PolGpo s inner join Polizas p on s.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + "     inner join Cat_MotivoSus cm on s.ID_MotivoSus = cm.ID_MotivoSus "
        //        sSql = sSql + "     inner join Cat_StaSuspension cs on s.ID_StaSuspension = cs.ID_StaSuspension "
        //        sSql = sSql + "     inner join Usuarios u on s.ID_Usuario = u.ID_Usuario "
        //        sSql = sSql + "where s.ID_Poliza = " & parametros(0)
        //        sSql = sSql + "     and s.ID_StaSuspension != 1 "
        //        sSql = sSql + "order by s.Fecha_Server desc "
        //        rsSuspHist = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        bGetEndosos = True
        //        cnnConexion = Nothing
        //        Exit Function

        //errGetEndosos:
        //        bGetEndosos = False
        //        cnnConexion = Nothing
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetDetSuspension(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetEndosos

        //        rsErrAdo = Nothing
        //        bGetDetSuspension = TipoResultado.NoHayDatos
        //        sSql = " select s.ID_Suspension, "
        //        sSql = sSql + "     Motivo_Suspension = upper(cm.Motivo_Suspension), Fecha_Susp = s.Fecha_System, Descripcion_Susp = s.Descripcion, "
        //        sSql = sSql + "     Motivo_Reactivacion = upper(cr.Motivo_Reactivacion), Fecha_Reac = r.Fecha_System, Descripcion_Reac = r.Descripcion, "
        //        sSql = sSql + "     Usuario = cast(u.Usuario as char(12)) " //CEFB PaqIns Abr2012 Usuario Responsable de la Reactivacion
        //        sSql = sSql + "from Suspension_PolGpo s inner join Cat_MotivoSus cm on s.ID_MotivoSus = cm.ID_MotivoSus "
        //        sSql = sSql + "     left join Reactivacion_PolGpo r on r.ID_Suspension = s.ID_Suspension "
        //        sSql = sSql + "     left join Cat_MotivoReac cr on r.ID_MotivoReac = cr.ID_MotivoReac "
        //        sSql = sSql + "     left join Usuarios u on u.ID_Usuario = r.ID_Usuario " //CEFB PaqIns Abr2012 Usuario Responsable de la Reactivacion
        //        sSql = sSql + "where s.ID_Suspension = " & parametros(1)
        //        bGetDetSuspension = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetEndosos:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetDetEndosos(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetEndosos2

        //        rsErrAdo = Nothing
        //        bGetDetEndosos = TipoResultado.NoHayDatos

        //        sSql = " select e.ID_HisEndoso, e.Fecha_Aplicacion, e.Fecha_Valuacion, "
        //        sSql = sSql + "     Endoso = convert(varchar(10), e.ID_HisEndoso) + // - // + ce.Endoso, e.ID_Endoso, "
        //        sSql = sSql + "     Beneficiario = convert(varchar(2), Num_Benef) + // - // + Abv_Parentesco + // - // + isnull(Nom_Benef, ////) + // // + isnull(ApP_Benef, ////) + // // + isnull(ApM_Benef, ////), "
        //        sSql = sSql + "     ct.Tpo_Valor, e.Valor_Ant, e.Valor_Nvo, e.Descripcion, "
        //        sSql = sSql + "     cb.Causa_Baja,  ca.Causa_Alta, "
        //        sSql = sSql + "     d.Fecha_Resol, d.Num_Resolucion, d.Num_ActaDef, d.Liquidacion_Prestamo, "
        //        sSql = sSql + "     d.Finiq_Total, d.Finiq_Neto, d.Dif_Prima, d.Pagos_Venc, d.Rentas_Mensuales, "
        //        sSql = sSql + "     d.Aguinaldo, d.Pagos_Indebidos, d.Tot_TransAlAseg, d.Devolucion "
        //        sSql = sSql + "from Endosos e inner join Cat_Endosos ce on e.ID_Endoso = ce.ID_Endoso "
        //        sSql = sSql + "     left join Cat_TpoEndosos cte on ce.ID_TpoEndoso = cte.ID_TpoEndoso "
        //        sSql = sSql + "     left join Pol_Benefs pb on e.ID_PolBenef = pb.ID_PolBenef "
        //        sSql = sSql + "     left join Cat_Parentescos cp on pb.ID_Parentesco = cp.ID_Parentesco "
        //        sSql = sSql + "     left join Cat_TpoValor ct on e.ID_TpoValor = ct.ID_TpoValor "
        //        sSql = sSql + "     left join Cat_CausaBajaEA cb on e.ID_CausaBaja = cb.ID_CausaBaja "
        //        sSql = sSql + "     left join Cat_CausaAltaEA ca on e.ID_CausaAlta = ca.ID_CausaAlta "
        //        sSql = sSql + "     left join DEndosos d on e.ID_HisEndoso = d.ID_HisEndoso "
        //        sSql = sSql + "where e.ID_HisEndoso = " & parametros(1)
        //        bGetDetEndosos = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetEndosos2:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetDetReac(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetReac

        //        bGetDetReac = TipoResultado.NoHayDatos
        //        sSql = "select " &
        //            " ID_Reactivacion," &
        //            " Motivo=(select cast(ID_MotivoReac as char(2))+// - Reactivación por //+cast(Motivo_Reactivacion as char(25)) from Cat_MotivoReac where ID_MotivoReac=Reactivacion_PolGpo.ID_MotivoReac)," &
        //            " Fecha_System," &
        //            " Usuario=(select cast(Usuario as char(12)) from Usuarios where ID_Usuario= Reactivacion_PolGpo.ID_Usuario)," &
        //            " Descripcion" &
        //            " From Reactivacion_PolGpo" &
        //            " where ID_Suspension = " & parametros(1)
        //        bGetDetReac = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetReac:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetImagenPolizas(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errorvb

        //        sSql = "select  Ramo, Pension, Riesgo, i.Fecha_IniDer, i.Fecha_Calculo, i.Fecha_CalculoInc, i.Fecha_ABase, i.Pje_Valuacion, i.Pje_AyudaAsist, i.Salario_RT, i.Salario_IV, Fecha_NactoAseg = i.Fecha_Nacto, Sexo = Abv_Sexo, "
        //        sSql = sSql + "    Cuantia_FIV = round(Cuantia_BaseFC / ((select Incremento from Cat_INPC where Periodo = 12 and Anio = AnioReval-1) / (select Incremento from Cat_INPC where Periodo = 12 and Anio = year(p.Fecha_IniVig)-1)), 2)"
        //        sSql = sSql + "from  Img_Polizas i, Polizas p, Cat_Sexos cs, Cat_Ramos cr, Cat_Pensiones cp, Cat_Riesgos cri "
        //        sSql = sSql + "where i.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + "  and i.ID_Sexo = cs.ID_Sexo "
        //        sSql = sSql + "  and i.ID_Ramo = cr.ID_Ramo "
        //        sSql = sSql + "  and i.ID_Pension = cp.ID_Pension "
        //        sSql = sSql + "  and i.ID_Riesgo = cri.ID_Riesgo "
        //        sSql = sSql + "  and p.ID_Poliza = " & parametros(0)

        //        bGetImagenPolizas = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function
        //errorvb:
        //        bGetImagenPolizas = TipoResultado.ExisteError
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetReservas(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errorvb

        //        sSql = "select * from Reservas where ID_Poliza = " & parametros(0)

        //        bGetReservas = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function
        //errorvb:
        //        bGetReservas = TipoResultado.ExisteError
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetImagenBenefs(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errorvb

        //        sSql = "select  Num_Benef, BenefIni, Abv_Parentesco, Cve_Orfandad, Abv_Sexo, Invalidez, Sta_Incremento = case when i.ID_StaIncremento = 0 then //NO// else //SI// end, i.Fecha_Nacto, "
        //        sSql = sSql + "    Edad = dbo.Age(i.Fecha_Nacto, Fecha_Calculo),  Edad_Inc04 = dbo.Age(i.Fecha_Nacto, Fecha_CalculoInc)"
        //        sSql = sSql + "from     Img_PolBenefs i,  Img_Polizas ip, "
        //        sSql = sSql + " Polizas p, Cat_BenefIni cb, Cat_Parentescos cp, Cat_Orfandades co, Cat_Sexos cs, Cat_Invalidez ci "
        //        sSql = sSql + "where    i.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + " and i.ID_Poliza = ip.ID_Poliza "
        //        sSql = sSql + " and i.ID_BenefIni = cb.ID_BenefIni "
        //        sSql = sSql + " and i.ID_Parentesco = cp.ID_Parentesco "
        //        sSql = sSql + " and i.ID_Orfandad = co.ID_Orfandad "
        //        sSql = sSql + " and i.ID_Sexo = cs.ID_Sexo "
        //        sSql = sSql + " and i.ID_Invalidez = ci.ID_Invalidez "
        //        sSql = sSql + " and p.ID_Poliza = " & parametros(0)
        //        sSql = sSql + " order by Num_Benef "

        //        bGetImagenBenefs = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function
        //errorvb:
        //        bGetImagenBenefs = TipoResultado.ExisteError
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetImagenBA(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errorvb

        //        sSql = "select  Abv_Beneficio, Proporcional, Opcion, Cuantia_AyEscolar, Pago_Art14, Num_Benef, "
        //        sSql = sSql + " Reserva = case when i.ID_Beneficio in (25,32,86,51) then Res_MatematicaPenAdic "
        //        sSql = sSql + "                when i.ID_Beneficio in (26,33,87,52,53) then Res_MatematicaAgAdic "
        //        sSql = sSql + "                when i.ID_Beneficio in (27,37,55,56,57) then Res_MatematicaAyEsc "
        //        sSql = sSql + "                else 0 end "
        //        sSql = sSql + "from  Img_PolBASelec i, "
        //        sSql = sSql + " Polizas p, "
        //        sSql = sSql + " Cat_Beneficios cb, "
        //        sSql = sSql + " Reservas r "
        //        sSql = sSql + "where    i.ID_Poliza = p.ID_Poliza "
        //        sSql = sSql + " and i.ID_Poliza = r.ID_Poliza "
        //        sSql = sSql + " and i.ID_Beneficio = cb.ID_Beneficio "
        //        sSql = sSql + " and i.ID_Beneficio in (25,32,86,51,26,33,87,52,53,27,37,55,56,57,28,54) "
        //        sSql = sSql + " and p.ID_Poliza = " & parametros(0)

        //        bGetImagenBA = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function
        //errorvb:
        //        bGetImagenBA = TipoResultado.ExisteError
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function


        //    public Function bGetPrestamosSegBan(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetPrestamosSegBan

        //        bGetPrestamosSegBan = TipoResultado.NoHayDatos
        //        ////////////////    sSql = "SELECT PB.ID_FolPrestamo, PB.ID_Poliza , P.Fol_Poliza, " & _
        //        ////////////////            " PB.Monto, PB.Monto+PB.PrimaSeguro as Monto_Total, " & _
        //        ////////////////            " CS.Sta_PtoSeg, PB.Plazo, " & _
        //        ////////////////            " PB.Fecha_Cotizado, PB.Fecha_Autorizado, PB.Fecha_Pagado " & _
        //        ////////////////            " From Prestamos_Bancomer as PB, Polizas as P, Cat_StaPtoSeg as CS" & _
        //        ////////////////            " Where PB.ID_StaPtoSeg = CS.ID_StaPrestamoSeg " & _
        //        ////////////////            " and PB.ID_Poliza = P.ID_Poliza " & _
        //        ////////////////            " and PB.ID_Poliza = " & parametros(0) & "order by ID_FolPrestamo DESC"

        //        sSql = "SELECT PB.ID_FolPrestamo as Folio, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, PB.Digito_Verificador, "
        //        ////sSql = sSql + " PB.Monto, PB.Monto+PB.PrimaSeguro as Monto_Total, CS.Sta_PtoSeg, PB.Fecha_Pagado, PB.Plazo, "
        //        //// Cambio a Monto_Total AGV 18/oct/2007
        //        sSql = sSql + " PB.Monto, isnull(PB.Monto_Total,0) as Monto_Total, CS.Sta_PtoSeg, PB.Fecha_Pagado, PB.Plazo, "
        //        sSql = sSql + " CASE when isnull(PB.Flag_MontoTotalSinSeguro, 0) = 0 then //NO// else //SI// END as SegIncMonto, "
        //        sSql = sSql + " SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE 1 END) As Pagos, "
        //        sSql = sSql + " SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE Imp_Capital END) As Pagado, "
        //        ////sSql = sSql + " Plazo-SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE 1 END) as Meses_Faltantes, "
        //        sSql = sSql + "SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 1 ELSE 0 END) as Meses_Faltantes, "
        //        ////sSql = sSql + " PB.Monto+PB.PrimaSeguro-SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE Imp_Capital END) as Saldo, PrimaSeguro, "
        //        //// Cambio a Monto_Total AGV 18/oct/2007
        //        sSql = sSql + " isnull(PB.Monto_Total,0) - SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE Imp_Capital END) as Saldo, PrimaSeguro, "
        //        sSql = sSql + " PB.ID_FolPrestamo, Fecha_Deposito, ID_StaPtoSeg, Importe_PagoFijo, Fecha_Liquidado, PB.Monto_Pension "
        //        sSql = sSql + " FROM Prestamos_Bancomer AS PB "
        //        sSql = sSql + " INNER JOIN Polizas as P ON PB.ID_Poliza = P.ID_Poliza "
        //        sSql = sSql + " INNER JOIN Cat_StaPtoSeg AS CS ON PB.ID_StaPtoSeg = CS.ID_StaPrestamoSeg "
        //        sSql = sSql + " INNER JOIN Tabla_AmortizacionPrestamos as TAP ON PB.ID_FolPrestamo = TAP.ID_FolPrestamo "
        //        sSql = sSql + " WHERE PB.ID_Poliza = " & parametros(0) & " "
        //        if parametros(1) = 1 Then
        //            sSql = sSql + " AND PB.ID_Grupo = " & parametros(2) & " "
        //        End if
        //        sSql = sSql + " GROUP BY  PB.ID_FolPrestamo, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, Digito_Verificador, "
        //        sSql = sSql + " Monto, Monto_Total, PrimaSeguro, Sta_PtoSeg, PB.Fecha_Pagado, Plazo, PrimaSeguro, Fecha_Deposito, ID_StaPtoSeg, Importe_PagoFijo, Fecha_Liquidado, PB.Monto_Pension, PB.Flag_MontoTotalSinSeguro "
        //        sSql = sSql + " ORDER BY PB.ID_FolPrestamo DESC "



        //        bGetPrestamosSegBan = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetPrestamosSegBan:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    ////bGetDetAmortiz
        //    public Function bGetDetAmortiz(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        On Error GoTo errGetDetAmortizaciones

        //        bGetDetAmortiz = TipoResultado.NoHayDatos
        //        sSql = "select * " &
        //            " From Tabla_AmortizacionPrestamos " &
        //            " where ID_FolPrestamo = " & parametros(0) & " " &
        //            " order by Num_Pago "
        //        bGetDetAmortiz = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetAmortizaciones:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function


        //    ////bGetDetPrestSeg
        //    public Function bGetDetPrestSeg(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer

        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetDetPrestSeg

        //        bGetDetPrestSeg = TipoResultado.NoHayDatos

        //        ////////////////    sSql = "SELECT PB.ID_FolPrestamo, PB.ID_Poliza , P.Fol_Poliza, PB.ID_Grupo, "
        //        ////////////////    sSql = sSql + " PB.Monto, PB.Monto+PB.PrimaSeguro as Monto_Total, PB.Importe_PagoFijo, "
        //        ////////////////    sSql = sSql + " CS.Sta_PtoSeg, PB.Plazo, "
        //        ////////////////    sSql = sSql + " PB.Fecha_Cotizado, PB.Fecha_Autorizado, PB.Fecha_Pagado, PB.Digito_Verificador "
        //        ////////////////    sSql = sSql + " From Prestamos_Bancomer as PB, Polizas as P, Cat_StaPtoSeg as CS"
        //        ////////////////    sSql = sSql + " Where PB.ID_StaPtoSeg = CS.ID_StaPrestamoSeg "
        //        ////////////////    sSql = sSql + " and PB.ID_Poliza = P.ID_Poliza "
        //        ////////////////    ////sSql = sSql + " and PB.ID_Poliza = " & parametros(0)
        //        ////////////////    sSql = sSql + " and PB.ID_FolPrestamo = " & parametros(1)

        //        sSql = "SELECT PB.ID_FolPrestamo as Folio, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, PB.Digito_Verificador, "
        //        ////sSql = sSql + " PB.Monto, PB.Monto+PB.PrimaSeguro as Monto_Total, CS.Sta_PtoSeg, PB.Fecha_Pagado, PB.Plazo, "
        //        //// Cambio por columna Monto_Total AGV 18/oct/2007
        //        sSql = sSql + " PB.Monto, isnull(PB.Monto_Total, 0) as Monto_Total, CS.Sta_PtoSeg, PB.Fecha_Pagado, PB.Plazo, "
        //        sSql = sSql + " SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE 1 END) As Pagos, "
        //        sSql = sSql + " SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE Imp_Capital END) As Pagado, "
        //        ////sSql = sSql + " Plazo-SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE 1 END) as Meses_Faltantes, "
        //        sSql = sSql + "SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 1 ELSE 0 END) as Meses_Faltantes, "
        //        ////sSql = sSql + " PB.Monto+PB.PrimaSeguro-SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE Imp_Capital END) as Saldo, PrimaSeguro, "
        //        //// Cambio por columna Monto_Total AGV 18/oct/2007
        //        sSql = sSql + " isnull(PB.Monto_Total, 0) - SUM(CASE WHEN Fecha_PagoSeguros IS NULL THEN 0 ELSE Imp_Capital END) as Saldo, PrimaSeguro, "
        //        sSql = sSql + " PB.ID_FolPrestamo, Fecha_Deposito, ID_StaPtoSeg, Importe_PagoFijo, Fecha_Liquidado, PB.Monto_Pension "
        //        sSql = sSql + " FROM Prestamos_Bancomer AS PB "
        //        sSql = sSql + " INNER JOIN Polizas as P ON PB.ID_Poliza = P.ID_Poliza "
        //        sSql = sSql + " INNER JOIN Cat_StaPtoSeg AS CS ON PB.ID_StaPtoSeg = CS.ID_StaPrestamoSeg "
        //        sSql = sSql + " INNER JOIN Tabla_AmortizacionPrestamos as TAP ON PB.ID_FolPrestamo = TAP.ID_FolPrestamo "
        //        sSql = sSql + " WHERE PB.ID_FolPrestamo = " & parametros(1) & " "
        //        sSql = sSql + " GROUP BY  PB.ID_FolPrestamo, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, Digito_Verificador, "
        //        sSql = sSql + " Monto, Monto_Total, PrimaSeguro, Sta_PtoSeg, PB.Fecha_Pagado, Plazo, PrimaSeguro, Fecha_Deposito, ID_StaPtoSeg, Importe_PagoFijo, Fecha_Liquidado, PB.Monto_Pension "
        //        sSql = sSql + " ORDER BY PB.ID_FolPrestamo DESC "


        //        bGetDetPrestSeg = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetPrestSeg:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function


        //    public Function bGetDescuentos(rsData As ADODB.Recordset, rsHistorico As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        Dim sSql As String
        //        Dim errVB As String
        //        Dim cnnConexion As ADODB.Connection
        //        On Error GoTo msgerror

        //        cnnConexion = New ADODB.Connection
        //        cnnConexion.Open(sConn)
        //        rsErrAdo = Nothing

        //        sSql = "select Fol_Poliza, d.*, Tpo_Descuento, Abv_Descuento, Sta_Descuento, "
        //        sSql = sSql + "     GrupoD = convert(varchar(3), d.ID_GrupoD) + // - // + isnull(rtrim(t1.Nom_Titular), ////) + // // + isnull(rtrim(t1.ApP_Titular), ////) + // // + isnull(rtrim(t1.ApM_Titular), ////), "
        //        sSql = sSql + "     GrupoA = convert(varchar(3), d.ID_GrupoA) + // - // + isnull(rtrim(t2.Nom_Titular), ////) + // // + isnull(rtrim(t2.ApP_Titular), ////) + // // + isnull(rtrim(t2.ApM_Titular), ////) "
        //        sSql = sSql + "from Descuentos d left join Polizas p on p.ID_Poliza = d.ID_Poliza "
        //        sSql = sSql + "     left join Titulares t1 on p.ID_Poliza = t1.ID_Poliza and d.ID_GrupoD = t1.ID_Grupo and t1.ID_StaTitular = 1 "
        //        sSql = sSql + "     left join Titulares t2 on p.ID_Poliza = t2.ID_Poliza and d.ID_GrupoA = t2.ID_Grupo and t2.ID_StaTitular = 1 "
        //        sSql = sSql + "     left join Cat_TpoDescuento cd on d.ID_TpoDescuento = cd.ID_TpoDescuento "
        //        sSql = sSql + "     left join Cat_StaDescuento cs on d.ID_StaDescuento = cs.ID_StaDescuento "
        //        sSql = sSql + "where p.ID_Empresa = " & parametros(0) & " "
        //        sSql = sSql + "     and d.ID_Poliza = " & parametros(1) & " "
        //        sSql = sSql + "order by GrupoD, Sta_Descuento, Fecha_StaDescuento, ID_Descuento "

        //        rsData = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        sSql = "select d.ID_Descuento, Mes = convert(datetime, convert(varchar(4), c.Anio) + //-// + convert(varchar(2), c.Periodo) + //-01//), pg.Importe_Descuento "
        //        sSql = sSql + "from Pagos_Descuentos pg inner join Corridas c on pg.ID_Corrida = c.ID_Corrida "
        //        sSql = sSql + "     inner join Descuentos d on pg.ID_Descuento = d.ID_Descuento "
        //        sSql = sSql + "where d.ID_Poliza = " & parametros(1) & " "
        //        sSql = sSql + "order by pg.ID_Descuento, Mes desc "

        //        rsHistorico = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        bGetDescuentos = True
        //        Exit Function

        //msgerror:
        //        bGetDescuentos = False
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetROPC(rsROPC As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        Dim sSql As String
        //        Dim errVB As String
        //        Dim cnnConexion As ADODB.Connection
        //        On Error GoTo errGetPagos

        //        cnnConexion = New ADODB.Connection
        //        cnnConexion.Open(sConn)
        //        rsErrAdo = Nothing

        //        //Traer información de ROPC
        //        sSql = "select  pg.ID_Pago, pg.ID_Grupo, ctp.ID_TpoROPC, pg.APagado, pg.PPagado, pg.Anio, pg.Periodo, "
        //        sSql = sSql + "     Tipo_Pago = ctc.Tipo_Corrida + // - //+  ctp.Abv_TipoPago, "
        //        sSql = sSql + "     Sta_Pago = convert(varchar(2), pg.ID_StaPago) + // - //+ convert(varchar(40), csp.Abv_StaPago), "
        //        sSql = sSql + "     pg.Imp_Pago, pg.Fecha_System, pg.ID_StaPago, c.ID_TipoCorrida "
        //        sSql = sSql + "from Pagos pg inner join Cat_TipoPagos ctp on pg.ID_TipoPago = ctp.ID_TipoPago and ctp.ID_TpoROPC in (2,3) "
        //        sSql = sSql + "     left join Cat_StaPagos csp on pg.ID_StaPago = csp.ID_StaPago "
        //        sSql = sSql + "     left join Corridas c on pg.ID_Corrida = c.ID_Corrida "
        //        sSql = sSql + "     left join Cat_TipoCorridas ctc on c.ID_TipoCorrida = ctc.ID_TipoCorrida "
        //        sSql = sSql + "where    pg.ID_Poliza = " & parametros(0) & " "
        //        sSql = sSql + "     and pg.ID_StaPago in (2,8) "
        //        sSql = sSql + "order by pg.APagado, pg.PPagado, pg.ID_Grupo, pg.ID_TipoPago, pg.Imp_Pago desc "

        //        rsROPC = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        bGetROPC = True
        //        Exit Function

        //errGetPagos:
        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        rsROPC = Nothing
        //        bGetROPC = False
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetPrestamosPB(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer

        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetPrestamosSegBan

        //        bGetPrestamosPB = TipoResultado.NoHayDatos
        //        //GCR  Nueva Regulación Prestamos Pensionados 2015-02-11 Inicio
        //        //sSql = "SELECT PB.ID_FolPrestamo as Folio, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, PB.Digito_Verificador, "
        //        //GCR  Nueva Regulación Prestamos Pensionados 2015-06-16 Inicio
        //        //    sSql = "SELECT PB.ID_FolPrestamo as Folio,convert(varchar(10),PB.ID_FolPrestamo)+CONVERT(varchar(1),Digito_Verificador) as Referencia, PB.ID_FolioOrigenReestructuracion as FolOri, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, Digito_Verificador=ROW_NUMBER() OVER(PARTITION BY PB.ID_FolioOrigenReestructuracion ORDER BY PB.ID_FolPrestamo)-1, "
        //        sSql = "SELECT PB.ID_FolPrestamo as Folio,convert(varchar(10),PB.ID_FolPrestamo)+CONVERT(varchar(1),Digito_Verificador) as Referencia,"
        //        sSql = sSql + "case when PFF.ID_FolPrestamo IS NULL then PB.ID_FolioOrigenReestructuracion else PFF.ID_FolPrestamo end as FolOri,"
        //        //GCR Fin
        //        //GCR Estatus de prestamos 2020/11/23 Inicio
        //        //sSql = sSql + "PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, Digito_Verificador=ROW_NUMBER() OVER(PARTITION BY PB.ID_FolioOrigenReestructuracion ORDER BY PB.ID_FolPrestamo)-1,"
        //        sSql = sSql + "PB.ID_Poliza, P.Fol_Poliza, PB.ID_Grupo, Digito_Verificador=ROW_NUMBER() OVER(PARTITION BY PB.ID_FolioOrigenReestructuracion ORDER BY PB.ID_FolPrestamo)-1,"
        //        //sSql = sSql + " PB.Monto, isnull(PB.Monto_DepReal,0)as //Monto Depositado Real//,isnull(PB.Monto_Total,0) as Monto_Total, CS.Sta_PtoPB as Sta_PtoPB  , PB.Fecha_Pagado, PB.Plazo, "
        //        sSql = sSql + " PB.Monto, isnull(PB.Monto_DepReal,0)as //Monto Depositado Real//,isnull(PB.Monto_Total,0) as Monto_Total, "
        //        sSql = sSql + "case when PB.ID_StaPtoPB=11 then case when E.ID_HisEndoso is null then //LIQUIDADO POR PÉRDIDA// else //LIQUIDADO POR SEGUNDAS NUPCIAS// end else CS.Sta_PtoPB  end  as Sta_PtoPB"
        //        sSql = sSql + ", PB.Fecha_Pagado, PB.Plazo,"
        //        //GCR Fin

        //        ////sSql = sSql + " CASE when isnull(PB.Flag_MontoTotalSinSeguro, 0) = 0 then //NO// else //SI// END as SegIncMonto, "
        //        sSql = sSql + " CASE when isnull(PB.Flag_MontoTotalSinSeguro, 0) = 0 then //SI// else //NO// END as SegIncMonto, "
        //        sSql = sSql + " SUM(CASE WHEN Fecha_PagoPensiones IS NULL THEN 0 ELSE 1 END) As Pagos, "
        //        sSql = sSql + " SUM(CASE WHEN Fecha_PagoPensiones IS NULL THEN 0 ELSE Imp_Capital END) As Pagado, "
        //        sSql = sSql + "SUM(CASE WHEN Fecha_PagoPensiones IS NULL THEN 1 ELSE 0 END) as Meses_Faltantes, "
        //        sSql = sSql + " isnull(PB.Monto_Total,0) - SUM(CASE WHEN Fecha_PagoPensiones IS NULL THEN 0 ELSE Imp_Capital END) as Saldo, PrimaSeguro, "
        //        sSql = sSql + " PB.ID_FolPrestamo, Fecha_Deposito, PB.ID_StaPtoPB, Importe_PagoFijo, Fecha_Liquidado, PB.Monto_Pension "
        //        sSql = sSql + " FROM Prestamos_PB AS PB "
        //        sSql = sSql + " INNER JOIN Polizas as P ON PB.ID_Poliza = P.ID_Poliza "
        //        sSql = sSql + " INNER JOIN Cat_StaPtoPB AS CS ON PB.ID_StaPtoPB = CS.ID_StaPtoPB "
        //        sSql = sSql + " INNER JOIN Prestamos_PBTablaAmortizacion as TAP ON PB.ID_FolPrestamo = TAP.ID_FolPrestamo "
        //        //GCR  Nueva Regulación Prestamos Pensionados 2015-06-16 Inicio
        //        sSql = sSql + " LEFT JOIN Prest_FolFijo PFF ON PB.ID_FolioOrigenReestructuracion=PFF.ID_FolioOrigenReestructuracion  "
        //        //GCR Fin

        //        //GCR Estatus de prestamos 2020/11/23 Inicio
        //        sSql = sSql + " LEFT JOIN Pol_Benefs B ON PB.ID_Poliza=B.ID_Poliza and PB.Num_Benef=B.Num_Benef"
        //        sSql = sSql + " LEFT JOIN Endosos E on PB.ID_Poliza=E.ID_Poliza and E.ID_PolBenef=B.ID_PolBenef and E.ID_Endoso=9"
        //        //GCR Fin

        //        sSql = sSql + " WHERE PB.ID_Poliza = " & parametros(0) & " "
        //        //GCR  Nueva Regulación Prestamos Pensionados 2015-02-11 Inicio
        //        //sSql = sSql + " GROUP BY  PB.ID_FolPrestamo, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, Digito_Verificador, "
        //        //GCR  Nueva Regulación Prestamos Pensionados 2015-06-16 Inicio
        //        //sSql = sSql + " GROUP BY  PB.ID_FolPrestamo,convert(varchar(10),PB.ID_FolPrestamo)+CONVERT(varchar(1),Digito_Verificador),PB.ID_FolioOrigenReestructuracion , PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, "

        //        //GCR Estatus de prestamos 2020/11/23 Inicio
        //        //sSql = sSql + " GROUP BY  PB.ID_FolPrestamo,convert(varchar(10),PB.ID_FolPrestamo)+CONVERT(varchar(1),Digito_Verificador),PFF.ID_FolPrestamo,PB.ID_FolioOrigenReestructuracion, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, "
        //        sSql = sSql + " GROUP BY E.ID_HisEndoso,PB.ID_FolPrestamo,convert(varchar(10),PB.ID_FolPrestamo)+CONVERT(varchar(1),Digito_Verificador),PFF.ID_FolPrestamo,PB.ID_FolioOrigenReestructuracion, PB.ID_Poliza, P.Fol_Poliza, PB.ID_Grupo, "
        //        //GCR Fin

        //        //GCR Fin
        //        //GCR Fin
        //        sSql = sSql + " Monto, Monto_Total, PrimaSeguro, Sta_PtoPB, PB.Fecha_Pagado, Plazo, PrimaSeguro, Fecha_Deposito, PB.ID_StaPtoPB, Importe_PagoFijo, Fecha_Liquidado, PB.Monto_Pension, PB.Flag_MontoTotalSinSeguro,PB.Monto_DepReal "
        //        sSql = sSql + " ORDER BY PB.ID_FolPrestamo DESC "

        //        bGetPrestamosPB = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetPrestamosSegBan:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)

        //    End Function


        //    public Function bGetDetPrestPB(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer

        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetDetPrestSeg

        //        bGetDetPrestPB = TipoResultado.NoHayDatos

        //        sSql = "SELECT PB.ID_FolPrestamo as Folio, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, PB.Digito_Verificador, "
        //        sSql = sSql + " PB.Monto, isnull(PB.Monto_Total, 0) as Monto_Total, CS.Sta_PtoPB, PB.Fecha_Pagado, PB.Plazo, "
        //        sSql = sSql + " SUM(CASE WHEN Fecha_PagoPensiones IS NULL THEN 0 ELSE 1 END) As Pagos, "
        //        sSql = sSql + " SUM(CASE WHEN Fecha_PagoPensiones IS NULL THEN 0 ELSE Imp_Capital END) As Pagado, "
        //        sSql = sSql + "SUM(CASE WHEN Fecha_PagoPensiones IS NULL THEN 1 ELSE 0 END) as Meses_Faltantes, "
        //        sSql = sSql + " isnull(PB.Monto_Total, 0) - SUM(CASE WHEN Fecha_PagoPensiones IS NULL THEN 0 ELSE Imp_Capital END) as Saldo, PrimaSeguro, "
        //        sSql = sSql + " PB.ID_FolPrestamo, Fecha_Deposito, PB.ID_StaPtoPB, Importe_PagoFijo, Fecha_Liquidado, PB.Monto_Pension, "
        //        sSql = sSql + " Tasa = convert(varchar(15), (PB.TIIE + PB.Spread)), " //CEFB PaqIns Abr2012 Tasa de Interes del Prestamo CATEL
        //        sSql = sSql + " Cuota = PB.CuotaSeguro " //CEFB MAPFRE Oct2012 Se agrega Cuota de Seguro
        //        sSql = sSql + " FROM Prestamos_PB AS PB "
        //        sSql = sSql + " INNER JOIN Polizas as P ON PB.ID_Poliza = P.ID_Poliza "
        //        sSql = sSql + " INNER JOIN Cat_StaPtoPB AS CS ON PB.ID_StaPtoPB = CS.ID_StaPtoPB "
        //        sSql = sSql + " INNER JOIN Prestamos_PBTablaAmortizacion as TAP ON PB.ID_FolPrestamo = TAP.ID_FolPrestamo "
        //        sSql = sSql + " WHERE PB.ID_FolPrestamo = " & parametros(1) & " "
        //        sSql = sSql + " GROUP BY  PB.ID_FolPrestamo, PB.ID_Poliza, P.Fol_Poliza, ID_Grupo, Digito_Verificador, "
        //        sSql = sSql + " Monto, Monto_Total, PrimaSeguro, Sta_PtoPB, PB.Fecha_Pagado, Plazo, PrimaSeguro, Fecha_Deposito, PB.ID_StaPtoPB, Importe_PagoFijo, Fecha_Liquidado, PB.Monto_Pension, "
        //        sSql = sSql + " convert(varchar(15), (PB.TIIE + PB.Spread)), PB.CuotaSeguro" //CEFB PaqIns Abr2012 Tasa de Interes del Prestamo CATEL //CEFB MAPFRE Oct2012 Se agrega Cuota de Seguro
        //        sSql = sSql + " ORDER BY PB.ID_FolPrestamo DESC "


        //        bGetDetPrestPB = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetPrestSeg:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

       

        //    public Function bGetDetAbonPB(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer

        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetDetAbon

        //        bGetDetAbonPB = TipoResultado.NoHayDatos
        //        sSql = "select ID_Abono,Imp_Abono,Fecha_Captura,Fecha_System" &
        //            " From Abonos" &
        //            " where ID_Prestamo = " & parametros(0)
        //        bGetDetAbonPB = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetAbon:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function


        //    public Function bGetDetAmortizPB(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer

        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetDetAmortizaciones

        //        bGetDetAmortizPB = TipoResultado.NoHayDatos
        //        sSql = "select * " &
        //            " From Prestamos_PBTablaAmortizacion " &
        //            " where ID_FolPrestamo = " & parametros(0) & " " &
        //            " order by Num_Pago "
        //        bGetDetAmortizPB = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetDetAmortizaciones:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)

        //    End Function


        //    public Function bGetDetPtmosISS(rsData As ADODB.Recordset, rsTotales As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        Dim sSql As String
        //        Dim errVB As String
        //        Dim cnnConexion As ADODB.Connection
        //        On Error GoTo errGetDetPagos

        //        cnnConexion = New ADODB.Connection
        //        cnnConexion.Open(sConn)
        //        rsErrAdo = Nothing

        //        //Buscar datos de cuenta
        //        sSql = "select     Detalle = isnull(case when ID_CatDPago = 3 then // - // + d.Detalle  "
        //        sSql = sSql + "         else // - // + substring(d.Detalle, charindex(//,//, d.Detalle, charindex(//,//, d.Detalle)+1), charindex(//,//, d.Detalle, charindex(//,//, d.Detalle, charindex(//,//, d.Detalle)+1)+1) - charindex(//,//, d.Detalle, charindex(//,//, d.Detalle)+1)) end, ////), "
        //        sSql = sSql + "    Conducto "
        //        sSql = sSql + "from    Pagos pg left join DPagos d on pg.Folio_DPago = d.Folio_DPago and pg.Fecha_System = d.Fecha_System and ID_CatDPago in (3,5) "
        //        sSql = sSql + "        left join Cat_Conductos cc on pg.ID_Conducto = cc.ID_Conducto "
        //        sSql = sSql + "where   pg.ID_Pago = " & parametros(1)
        //        rsData = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)
        //        parametros(2) = rsData.Fields("Conducto").Value & Replace(Replace(rsData.Fields("Detalle").Value, ",", ""), "/", " / ")

        //        //Crear tabla con nombres de beneficiarios
        //        sSql = "drop table #benefs"
        //        On Error Resume Next
        //        cnnConexion.Execute(sSql)

        //        //Crear tabla con nombres de beneficiarios
        //        sSql = "select distinct ID_PolBenef into #benefs from MPagos where ID_Pago = " & parametros(1)
        //        cnnConexion.Execute(sSql)

        //        //Traer detalle de pagos de forma Pgo_Nom
        //        sSql = "select      "
        //        sSql = sSql & vbCrLf & "     Prestamos_ISS = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 23), " //SGRR 2010-01-13 Descuentos de Ptmos ISSSTE
        //        sSql = sSql & vbCrLf & "     Prestamos_FVIA = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 24), " //SGRR 2010-01-13 Descuentos de Ptmos FOVISSSTE Amortización
        //        sSql = sSql & vbCrLf & "     Prestamos_FVIS = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_PolBenef = pb.ID_PolBenef and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 25) " //SGRR 2010-01-13 Descuentos de Ptmos FOVISSSTE Seguro de Daños
        //        sSql = sSql & vbCrLf & "from Pagos pg inner join #benefs m on pg.ID_Pago = " & parametros(1) & " "
        //        sSql = sSql & vbCrLf & "     left join Pol_Benefs pb on pb.ID_PolBenef = m.ID_PolBenef "
        //        sSql = sSql & vbCrLf & "order by Num_Benef "

        //        rsData = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        //Traer totales por columna
        //        sSql = "select  "
        //        sSql = sSql & vbCrLf & "     Prestamos_ISS = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 23), " // SGRR 2010-01-13 Descuentos de Ptmos ISSSTE
        //        sSql = sSql & vbCrLf & "     Prestamos_FVIA = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 24), " // SGRR 2010-01-13 Descuentos de Ptmos FOVISSSTE Amortización
        //        sSql = sSql & vbCrLf & "     Prestamos_FVIS = (select sum(Imp_PBenef) from MPagos m, Cat_BenPgoNom c where pg.ID_Pago = m.ID_Pago and m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom = 25) " // SGRR 2010-01-13 Descuentos de Ptmos FOVISSSTE Seguro de Daños
        //        sSql = sSql & vbCrLf & "from Pagos pg "
        //        sSql = sSql & vbCrLf & "where pg.ID_Pago = " & parametros(1) & " "

        //        rsTotales = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenDynamic, adUseClient, sSql)

        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        bGetDetPtmosISS = True
        //        Exit Function

        //errGetDetPagos:
        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        bGetDetPtmosISS = False
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    //ASB - 06/06/2010
        //    //FUNCION QUE SE UTILIZA EN PARA CONSULTAR HISTORIAL DE ENCUESTAS
        //    public Function bGetHistorialEncuestas(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetHistorialEncuestas

        //        bGetHistorialEncuestas = TipoResultado.NoHayDatos
        //        sSql = "select  top 3 p.Fol_Poliza,es.ID_Poliza," &
        //        "es.ID_Encuesta," &
        //        "bei.ID_Folio," &
        //        "es.Fecha_Captura," &
        //        "es.Fecha_ProxEncuesta," &
        //        "coe.Descripcion OrigenEncuesta," &
        //        "dbes.Descripcion EstadoSalud," &
        //        "es.Num_Promotor," &
        //        "aceptada= case es.Encuesta_Aceptada when 1 then //SI// else //NO// end," &
        //        "es.Comentarios " &
        //        "from Encuesta_DeServicio es " &
        //        "left join Bit_EncuestaImpresion bei " &
        //        "on bei.ID_Encuesta=es.ID_Encuesta " &
        //        "inner join Polizas p " &
        //        "on es.ID_Poliza=p.ID_Poliza " &
        //        "inner join Cat_OrigenEncuesta coe " &
        //        "on es.ID_OrigenEncuesta=coe.ID_OrigenEncuesta " &
        //        "inner join DBCLIENTE.dbo.Cat_EstadoSalud dbes " &
        //        "on es.ID_EstadoSalud=dbes.ID_EstadoSalud " &
        //        "Where es.ID_Poliza =" & parametros(0) &
        //        " and ID_InstitucionSS=" & parametros(1) &
        //        " order by es.Fecha_Captura desc"
        //        bGetHistorialEncuestas = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetHistorialEncuestas:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    //ASB - 06/06/2010
        //    //FUNCION QUE SE UTILIZA EN PARA CONSULTAR LIQUIDACIONES DE SEGURO BANCOMER
        //    public Function bGetConsultaLiquidacionesSB(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetLiquidacionesSB

        //        //bGetHistorialEncuestas = TipoResultado.NoHayDatos
        //        sSql = "execute sp_EstadoActualPrestamo " & parametros(0) & "," & parametros(1) & ",//" & parametros(2) & "//"
        //        bGetConsultaLiquidacionesSB = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetLiquidacionesSB:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    //ASB - 06/06/2010
        //    //FUNCION QUE SE UTILIZA EN PARA CONSULTAR LIQUIDACIONES DE PENSIONES BANCOMER
        //    public Function bGetConsultaLiquidacionesPB(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetLiquidacionesPB

        //        //bGetHistorialEncuestas = TipoResultado.NoHayDatos
        //        sSql = "execute sp_EstadoActualPrestamoPB " & parametros(0) & "," & parametros(1) & ",//" & parametros(2) & "//"
        //        bGetConsultaLiquidacionesPB = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetLiquidacionesPB:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    //ASB - 06/06/2010
        //    //FUNCION QUE SE UTILIZA EN PARA CONSULTAR LIQUIDACIONES DE PENSIONES BANCOMER
        //    public Function bGetConvenioCIE(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetConvenioCIE

        //        bGetConvenioCIE = TipoResultado.NoHayDatos
        //        sSql = "SELECT  Descripcion descripcionCIE " &
        //          "From DBCLIENTE.dbo.Prestamos_PBParametros " &
        //          "Where ID_Parametro =16" //  Parametro Convenio CIE  en base de datos DBCLIENTE
        //        bGetConvenioCIE = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetConvenioCIE:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    //FUNCION QUE SE UTILIZA EN PARA CONSULTAR LIQUIDACIONES DE SEGUROS BANCOMER
        //    public Function bGetConvenioCIE_SB(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetConvenioCIE

        //        bGetConvenioCIE_SB = TipoResultado.NoHayDatos
        //        sSql = "SELECT  Descripcion descripcionCIE " &
        //          "From DBCLIENTE.dbo.Cat_Parametro " &
        //          "Where ID_Parametro =20" //  Parametro Convenio CIE  en base de datos DBCLIENTE
        //        bGetConvenioCIE_SB = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetConvenioCIE:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    //ASB - 06/06/2010
        //    //FUNCION QUE SE UTILIZA EN PARA CONSULTAR LIQUIDACIONES DE SEGURO BANCOMER
        //    public Function bGetReferencia(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetReferencia

        //        bGetReferencia = TipoResultado.NoHayDatos
        //        sSql = "select  referencia= convert(varchar,ID_FolPrestamo) +////+ convert(varchar,Digito_Verificador) " &
        //            "From Prestamos_Bancomer " &
        //           "Where ID_FolPrestamo = " & parametros(0)
        //        bGetReferencia = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetReferencia:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    //ASB - 22/06/2010
        //    //FUNCION QUE SE UTILIZA EN PARA CONSULTAR LIQUIDACIONES DE PENSIONES BANCOMER
        //    public Function bGetReferenciaPB(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGetReferencia

        //        bGetReferenciaPB = TipoResultado.NoHayDatos
        //        sSql = "select  referencia= convert(varchar,ID_FolPrestamo) +////+ convert(varchar,Digito_Verificador) " &
        //            "From Prestamos_PB " &
        //           "Where ID_FolPrestamo = " & parametros(0)
        //        bGetReferenciaPB = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetReferencia:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    //ASB - 22/06/2010
        //    //FUNCION QUE SE UTILIZA EN PARA CONSULTAR LIQUIDACIONES DE SEGURO BANCOMER Y PENSIONES
        //    public Function bGetTitularActivo(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String

        //        On Error GoTo errGettitularActivo

        //        bGetTitularActivo = TipoResultado.NoHayDatos
        //        sSql = "select Titular= ApP_Titular +// //+ ApM_Titular +// //+ Nom_Titular " &
        //           "From Titulares " &
        //           "Where ID_Poliza =" & parametros(3) & " and ID_StaTitular=1 and ID_Grupo=" & parametros(4)
        //        bGetTitularActivo = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGettitularActivo:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    public Function bGetPrestamosActualAntiguo(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        // 26/08/2010 -ASB
        //        // SE AGREGA A CONSULTA EL ID_Grupo y ID_StaPtoPB a 2 o 4 segun sea el caso
        //        On Error GoTo errGetPrestamosActualAntiguo

        //        bGetPrestamosActualAntiguo = TipoResultado.NoHayDatos
        //        sSql = "select " &
        //    "prestamoAntiguo=isnull((select Fecha_Deposito from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " and ID_StaPtoPB=4),0)," &
        //    "folioAntiguo=isnull((select ID_FolPrestamo from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " and ID_StaPtoPB=4),0)," &
        //    "prestamoNuevo=isnull((select Fecha_Deposito from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " and ID_StaPtoPB=2),0)," &
        //    "folioNuevo=isnull((select ID_FolPrestamo from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " and ID_StaPtoPB=2),0)," &
        //    "MontoTotalSegundoPrestamo=isnull((select Monto_Total from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " and ID_StaPtoPB=2),0)"

        //        bGetPrestamosActualAntiguo = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetPrestamosActualAntiguo:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    public Function bGetPrestamosActualAntiguoC(rsData As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Integer
        //        Dim sSql As String
        //        Dim errVB As String
        //        // 26/08/2010 -ASB
        //        // SE AGREGA A CONSULTA EL ID_Grupo y ID_StaPtoPB
        //        // 28-02-2011 - ASB
        //        // SE Agrega status de prestamos para mejor filtrado al obtener prestamo actual (status 4) y antiguo (status 5)
        //        On Error GoTo errGetPrestamosActualAntiguoc

        //        bGetPrestamosActualAntiguoC = TipoResultado.NoHayDatos
        //        sSql = "select " &
        //    "prestamoAntiguo=(select top 1 Fecha_Pagado from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " and ID_StaPtoPB=4 order by Fecha_Pagado asc)," &
        //    "folioAntiguo=(select top 1 ID_FolPrestamo from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " and ID_StaPtoPB=4 order by Fecha_Pagado asc)," &
        //    "prestamoNuevo=(select top 1 Fecha_Pagado from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " order by Fecha_Pagado desc)," &
        //    "folioNuevo=(select top 1 ID_FolPrestamo from Prestamos_PB where ID_Poliza=" & parametros(0) & " and ID_Grupo= " & parametros(2) & " order by Fecha_Pagado desc)"



        //        bGetPrestamosActualAntiguoC = EjecutaSql(rsData, sConn, sSql, rsErrAdo)

        //        Exit Function

        //errGetPrestamosActualAntiguoc:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function




        //    //Alexander Hdez 04/10/2012, Prestamos FOVISSSTE, Muestra el total de descuentos hechos al momento de la consulta.
        public int bGetDescAplifOVI(Recordset rsData, string sConn, Recordset rsErrAdo, object[] parametros)
        {
            string sSql;
            string errVB;
            int bGetDescAplifOVIRes = 0;

            //On Error GoTo errGetDescAplifOVI

            try
            {

            bGetDescAplifOVIRes = Convert.ToInt32(MTSCPolizas.Modulos.ModRecordset.TipoResultado.NoHayDatos);

            sSql = "SELECT Count(m.ID_Beneficio)ID_Beneficio   FROM Pagos p ";
            sSql = sSql + " INNER JOIN MPagos m ON m.ID_Pago=p.ID_Pago ";
            sSql = sSql + " INNER JOIN Prestamos_FOVISSSTE pf ON p.ID_Poliza=pf.ID_Poliza ";
            sSql = sSql + " Where p.ID_Poliza = //" + parametros[1] + "// And m.ID_Beneficio = //" + parametros[2] + "// And pf.ID_Prestamo = //" + parametros[0] + "// ";
            sSql = sSql + " AND p.Fecha_Server>(SELECT pf2.Fch_Captura from Prestamos_FOVISSSTE pf2 WHERE pf2.ID_Poliza=//" + parametros[1] + "// ";
            sSql = sSql + " AND pf2.ID_Prestamo=//" + parametros[0] + "//) ";


            bGetDescAplifOVIRes = MTSCPolizas.Modulos.ModRecordset.EjecutaSql(ref rsData, sConn, sSql, rsErrAdo);

            return bGetDescAplifOVIRes;

                //Exit Function
            }
            catch (Exception Err)
            {
                //errGetDescAplifOVI:
                errVB = Err.Source + "\t" + Err.Message;
                rsErrAdo = Modulos.modErrores.ErroresDLL(null, errVB);
                return Convert.ToInt32(MTSCPolizas.Modulos.ModRecordset.TipoResultado.ExisteError);
            }

        }


            //    //-----------------------------------------------------------------------------------------------
            //    //------------------------Nueva funcionalidad de ROPC Modificaciones-----------------------------
            //    //-------------------------------------JCMN 23092014---------------------------------------------
            //    //-----------------------------------------Inicio------------------------------------------------

        //    public Function bUpdatePagoROPC(sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object, rsDatosUser As ADODB.Recordset) As Boolean
        //        Dim sSqlTable As String
        //        Dim sSqlConta As String
        //        Dim errVB As String
        //        Dim ID_BenPgoNom As String
        //        Dim cnnConexion As ADODB.Connection
        //        Dim rsPolBenefsROPC As ADODB.Recordset
        //        Dim TpoMovConta As Integer
        //        Dim RegPolBenefs As Integer
        //        Dim sSqlPolBenefs As String
        //        Dim No_PolBenef As Long
        //        Dim i As Integer
        //        On Error GoTo errdatePagoROPC
        //        //--Inicio---Definición de parametros----------
        //        //        vParametrosROPC(0) = rsDetalleROPC!ID_Poliza //ID_Poliza
        //        //        vParametrosROPC(1) = rsDetalleROPC!Fol_Poliza //Fol_Poliza
        //        //        vParametrosROPC(2) = rsDetalleROPC!ID_Empresa //Empresa
        //        //        vParametrosROPC(3) = rsDetalleROPC!ID_Conducto //ID_Conducto
        //        //        vParametrosROPC(4) = rsDetalleROPC!ID_Beneficio //ID_Beneficio
        //        //        vParametrosROPC(5) = rsDetalleROPC!CD_Concepto //CD_Concepto
        //        //        vParametrosROPC(6) = TextEdit.Text
        //        //        vParametrosROPC(7) = IdPagoPolizaROPC
        //        //        vParametrosROPC(8) = Id_GrupoROPC
        //        //        vParametrosROPC(9) = gsFechaSistema
        //        //        vParametrosROPC(10) = columnaROPC
        //        //--Fin------Definición de parametros----------




        //        //Obtiene el ID_PolBenef para el beneficiario correcto
        //        sSqlPolBenefs = "select distinct ID_PolBenef from MPagos where ID_Pago = " + Str(parametros(7))
        //        bUpdatePagoROPC = EjecutaSql(rsPolBenefsROPC, sConn, sSqlPolBenefs, rsErrAdo)

        //        For i = 0 To rsPolBenefsROPC.RecordCount - 1
        //            if parametros(13) = i Then
        //                No_PolBenef = rsPolBenefsROPC(0).Value
        //                Exit For
        //            Else
        //                rsPolBenefsROPC.MoveNext()
        //            End if
        //        Next i


        //        if CDec(parametros(6)) > CDec(parametros(11)) Then
        //            TpoMovConta = 1
        //        Else
        //            TpoMovConta = 2
        //        End if

        //        sSqlTable = "Update tmp tmp.Imp_PBenef =" + Str(parametros(11))
        //        sSqlTable = sSqlTable + " from DBCAIRO.dbo.MPagos as tmp inner join DBCAIRO.dbo.Pagos p on tmp.ID_Pago = p.ID_Pago"
        //        sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_BenPgoNom c on  tmp.ID_Beneficio = c.ID_Beneficio"
        //        sSqlTable = sSqlTable + " and tmp.ID_Beneficio=" + Str(parametros(4)) + " and p.ID_Pago =" + Str(parametros(7))
        //        sSqlTable = sSqlTable + " and ID_PolBenef= " + Str(No_PolBenef)

        //        sSqlConta = "EXEC BCNBD001.dbo.SPCN_UpdateMontosROPC " + Str(parametros(0)) + "," + Str(parametros(8)) + "," + Str(parametros(2))
        //        sSqlConta = sSqlConta + "," + Str(parametros(3)) + "," + Str(parametros(5)) + "," // + "//" + "Ajuste ROPC" + "//"


        //        if parametros(10) = 4 Then // Columna de Pagos vencidos JCMN 12/01/2015
        //            sSqlConta = sSqlConta + "//" + "Ajuste ROPC-PV" + "//"
        //        Else
        //            sSqlConta = sSqlConta + "//" + "Ajuste ROPC" + "//"
        //        End if
        //        sSqlConta = sSqlConta + "," + Str(Math.Abs(parametros(6) - parametros(11))) + "," + "//" + parametros(9) + "//" + "," + Str(TpoMovConta)

        //        Dim grsErrADO As ADODB.Recordset
        //        grsErrADO = New Recordset
        //        rsDatosUser("IDAccion").Value = 1130

        //        if parametros(12) = 1 Then
        //            if ExecSql(sConn, sSqlTable, rsErrAdo) Then ////Afecta SQL MPagos
        //                bUpdatePagoROPC = ExecSqlROPC(sConn, sSqlConta, rsErrAdo) ////Afecta SOL Asientos Contables

        //                if InsertaBitacora(sConn, "Actualiza Monto: " + sSqlTable + "   ,Genera contabilidad: " + sSqlConta + " --:-- Monto Anterior: $" + parametros(6) + " --:-- Monto Nuevo: $" + parametros(11), rsDatosUser, grsErrADO) Then

        //                End if
        //            End if
        //        Else
        //            bUpdatePagoROPC = ExecSql(sConn, sSqlTable, rsErrAdo)  ////Afecta SOL MPagos
        //            if InsertaBitacora(sConn, "Primer Mov: " + sSqlTable + "   ,Genera contabilidad: NA" + " --:-- Monto Anterior: $" + Str(parametros(6)) + " --:-- Monto Nuevo: $" + Str(parametros(11)), rsDatosUser, grsErrADO) Then

        //            End if
        //        End if
        //        rsDatosUser("IDAccion").Value = 0

        //        Exit Function

        //errdatePagoROPC:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function
        //    public Function bGetDetPagoROPC(rsDetalleROPC As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        //public Function bGetDetPagoROPC(rsDetalleROPC As ADODB.Recordset, rsPolBenefsROPC As ADODB.Recordset, sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        Dim sConnRopc As String
        //        Dim IdBeneficio As String
        //        Dim sSqlTable As String
        //        Dim errVB As String
        //        Dim icont As Integer
        //        Dim cnnConexion As ADODB.Connection
        //        Dim RecordPrueba As ADODB.Recordset
        //        On Error GoTo errDetPagoROPC

        //        sConnRopc = sConn

        //        Select Case parametros(2)
        //            Case 0 //Prestamos_ISS = 23
        //                IdBeneficio = "= 23"
        //            Case 1 //Prestamos_FVIA = 24
        //                IdBeneficio = "= 24"
        //            Case 2 //Prestamos_FVIS = 25
        //                IdBeneficio = "= 25"
        //            Case 3 //Nombre = //TOTAL//
        //        //IdBeneficio =
        //            Case 4 //Pagos_Vencidos = 0
        //                IdBeneficio = "= 0"
        //            Case 5 //Pension_Basica 1
        //                IdBeneficio = "= 1"
        //            Case 6 //BAU = IN (21,30)
        //                IdBeneficio = "in(21,30)"
        //            Case 7 //Aguinaldo_BAU = 22
        //                IdBeneficio = "= 22"
        //            Case 8 //Art14_2002 = 2
        //                IdBeneficio = "= 2"
        //            Case 9 //Art14_2004 = 3
        //                IdBeneficio = "= 3"
        //            Case 10 //Aguinaldo = in (4,26,27,28,29)
        //                IdBeneficio = "in(4,26,27,28,29)"
        //            Case 11 //Ag_Art14_2004 = 5
        //                IdBeneficio = "= 5"
        //            Case 12 //Finiquito = 6
        //                IdBeneficio = "= 6"
        //            Case 13 //Finiquito_Art14_2004 = 7
        //                IdBeneficio = "= 7"
        //            Case 14 //Retroactivo_ROPC_HS = 8
        //                IdBeneficio = "= 8"
        //            Case 15 //BAMI = 9
        //                IdBeneficio = "= 9"
        //            Case 16 //BAMI_Aguinaldo = 10
        //                IdBeneficio = " = 10"
        //            Case 17 //BAMI_Finiquito = 11
        //                IdBeneficio = "= 11"
        //            Case 18 //Pension_Adicional = 12
        //                IdBeneficio = "= 12"
        //            Case 19 //Aguinaldo_Adicional = 13
        //                IdBeneficio = "= 13"
        //            Case 20 //Ayuda_Escolar = 14
        //                IdBeneficio = "= 14"
        //            Case 21 //Abono_Grupo = 15
        //                IdBeneficio = "= 15"
        //            Case 22 //Descuentos_Grupo = 16
        //                IdBeneficio = "= 16"
        //            Case 23 //Descuentos_Otros = 17
        //                IdBeneficio = "= 17"
        //            Case 24 //Prestamos_ATM = 18
        //                IdBeneficio = "= 18"
        //            Case 25 //Prestamos_Seguros = 19
        //                IdBeneficio = "= 19"
        //    //Case 26 //Total = (select sum(Imp_PBenef) from MPagos m where pg.ID_Pago = m.ID_Pago),
        //     //   IdBeneficio = 26
        //            Case 27 //Prestamos_PB = 129
        //                IdBeneficio = "= 129"
        //            Case 28 //Prestamos_PB2 = 139
        //                IdBeneficio = "= 139"
        //            Case 29 //GtosFun_Endoso = 31
        //                IdBeneficio = "= 31"
        //        End Select

        //        sSqlTable = "select Imp_PBenef,m.ID_Beneficio from Pagos pg inner join MPagos m on pg.ID_Pago=m.ID_Pago"
        //        sSqlTable = sSqlTable + " inner join Cat_BenPgoNom c on m.ID_Beneficio = c.ID_Beneficio and ID_BenPgoNom " + IdBeneficio
        //        sSqlTable = sSqlTable + " and pg.ID_Pago =" + Str(parametros(1))

        //        bGetDetPagoROPC = EjecutaSql(rsDetalleROPC, sConn, sSqlTable, rsErrAdo)
        //        //rsDetalleROPC.MoveFirst

        //        //sSqlTable = "select A.ID_Poliza,A.Fol_Poliza,A.ID_Empresa,B.ID_Conducto,B.ID_Pago,D.Imp_PBenef,F.Beneficio,D.ID_Beneficio,F.ID_Concepto,F.ID_ConceptoContable,H.CD_MovRem,H.CD_NivMovRem"
        //        //sSqlTable = sSqlTable + ",O.Cta_Mayor +// //+O.Cta_N1+// //+O.Cta_N2+// //+O.Cta_N3+// //+O.Cta_N4 as Cuenta_Contable,H.ID_CtaMayor,H.CD_Concepto"
        //        //sSqlTable = sSqlTable + " from DBCAIRO.dbo.Polizas A inner join DBCAIRO.dbo.Pagos B on A.ID_Poliza=B.ID_Poliza"
        //        //sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_Conductos J on J.ID_Conducto=B.ID_Conducto"
        //        //sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_RamoContable G on G.ID_RamoContable=B.ID_RamoContable"
        //        //sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.MPagos D on D.ID_Pago = B.ID_Pago"
        //        //sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_BenPgoNom E on E.ID_Beneficio = D.ID_Beneficio"
        //        //sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_Beneficios F on F.ID_Beneficio=E.ID_Beneficio"
        //        //sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_ConceptoContable M on M.ID_ConceptoContable=F.ID_ConceptoContable"
        //        //sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_StaPago C on C.ID_StaPago = B.ID_StaPago"
        //        //sSqlTable = sSqlTable + " inner join BCNBD001.dbo.TCN002_MovRemConCtas H on H.CD_Concepto=E.ID_BenPgoNom"
        //        //sSqlTable = sSqlTable + " inner join BCNBD001.dbo.TCN013_Cat_Cuentas_Mayor I on I.ID_CtaMayor=H.ID_CtaMayor and I.Cuenta in (2121,5401,5403,6506)"
        //        //sSqlTable = sSqlTable + " inner join BCNBD001.dbo.TCN014_Cat_Ctas_Contables O on O.Cta_Mayor=I.Cuenta and O.ID_CtaCont = I.ID_CtaMayor"
        //        //sSqlTable = sSqlTable + " Where B.ID_Pago =" + Str(parametros(1)) + " And A.ID_Poliza =" + Str(parametros(0))
        //        //sSqlTable = sSqlTable + " and D.ID_Beneficio= " + Str(rsDetalleROPC(1))

        //        ////Nueva logica para beneficios compuestos JCMN --Inicio
        //        //if rsDetalleROPC.RecordCount > 1 Then
        //        //    rsDetalleROPC.MoveFirst
        //        //    For i = 0 To rsDetalleROPC.RecordCount - 1
        //        //      sSqlTable = "EXEC DBCAIRO.dbo.SP_Busca_Beneficio " & parametros(0) & ", " & parametros(1) & ", " & rsDetalleROPC(1)
        //        //      bGetDetPagoROPC = EjecutaSql(rsDetalleROPC, sConnRopc, sSqlTable, rsErrAdo)
        //        //    Next i
        //        //End if
        //        ////Nueva logica para beneficios compuestos JCMN --Fin


        //        //    sSqlTable = "EXEC DBCAIRO.dbo.SP_Busca_Beneficio " & parametros(0) & ", " & parametros(1) & ", " & rsDetalleROPC(1)
        //        //    bGetDetPagoROPC = EjecutaSql(rsDetalleROPC, sConnRopc, sSqlTable, rsErrAdo)
        //        //
        //        //    if rsDetalleROPC.RecordCount = 0 And parametros(2) = 4 Then // Columna de Pagos vencidos JCMN
        //        //        sSqlTable = "select A.ID_Poliza,A.Fol_Poliza,A.ID_Empresa,B.ID_Conducto,B.ID_Pago,D.Imp_PBenef,F.Beneficio,D.ID_Beneficio,F.ID_Concepto,F.ID_ConceptoContable,CD_MovRem = 10,CD_NivMovRem= 1"
        //        //        sSqlTable = sSqlTable + ",Cuenta_Contable = //21210102107500//,ID_CtaMayor = 2121,CD_Concepto=1"
        //        //        sSqlTable = sSqlTable + " from DBCAIRO.dbo.Polizas A inner join DBCAIRO.dbo.Pagos B on A.ID_Poliza=B.ID_Poliza"
        //        //        sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_Conductos J on J.ID_Conducto=B.ID_Conducto"
        //        //        sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_RamoContable G on G.ID_RamoContable=B.ID_RamoContable"
        //        //        sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.MPagos D on D.ID_Pago = B.ID_Pago"
        //        //        sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_BenPgoNom E on E.ID_Beneficio = D.ID_Beneficio"
        //        //        sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_Beneficios F on F.ID_Beneficio=E.ID_Beneficio"
        //        //        sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_ConceptoContable M on M.ID_ConceptoContable=F.ID_ConceptoContable"
        //        //        sSqlTable = sSqlTable + " inner join DBCAIRO.dbo.Cat_StaPago C on C.ID_StaPago = B.ID_StaPago"
        //        //        sSqlTable = sSqlTable + " Where B.ID_Pago =" + Str(parametros(1)) + " And A.ID_Poliza =" + Str(parametros(0))
        //        //        bGetDetPagoROPC = EjecutaSql(rsDetalleROPC, sConnRopc, sSqlTable, rsErrAdo)
        //        //    End if



        //        Exit Function

        //errDetPagoROPC:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        bGetDetPagoROPC = False
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function




        //    //---JCMN 01-12-2014  Actualización de status en Pagos (DBCAIRO)
        //    public Function bUpdateStaPagoROPC(sConn As String, rsErrAdo As ADODB.Recordset, parametros As Object) As Boolean
        //        Dim sSql As String
        //        Dim sSqlPolBenefs As String
        //        Dim errVB As String
        //        Dim ID_BenPgoNom As String
        //        Dim cnnConexion As ADODB.Connection
        //        Dim TpoMovConta As Integer
        //        On Error GoTo errdatePagoROPC

        //        //--Inicio---Definición de parametros----------
        //        //   vParametrosROPC(1) --> //ID_Pago
        //        //   vParametrosROPC(0) --> //ID_StaPago

        //        sSql = "Update tmp tmp.ID_StaPago =" + Str(parametros(0))
        //        sSql = sSql + " from DBCAIRO.dbo.Pagos as tmp where "
        //        sSql = sSql + " tmp.ID_Pago=" + Str(parametros(1))

        //        bUpdateStaPagoROPC = ExecSql(sConn, sSql, rsErrAdo)  //Afecta SOL Pagos

        //        Exit Function

        //errdatePagoROPC:
        //        errVB = String.Format(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(Nothing, errVB)
        //    End Function

        //    //------------------------------------------Fin--------------------------------------------------
        //    //------------------------Nueva funcionalidad de ROPC Modificaciones-----------------------------
        //    //-------------------------------------JCMN 23092014---------------------------------------------
        //    //-----------------------------------------------------------------------------------------------


        //    //------------------------------------------Inicio-----------------------------------------------
        //    //------------------------Nueva funcionalidad de Consulta NSS Beneficiarios ISSSTE---------------
        //    //-------------------------------------JCMN 06/04/2016-------------------------------------------
        //    //-----------------------------------------------------------------------------------------------

        //    public Function bGetNSS(ByRef rsErrAdo As ADODB.Recordset, ByRef rsRecord As ADODB.Recordset, ByRef gDsn As String, parametros As Object) As Boolean
        //        On Error GoTo msgerror

        //        cnnConexion = New ADODB.Connection
        //        cnnConexion.Open(gDsn)

        //        cnnConexion.CommandTimeout = 0
        //        Dim sSql As String


        //        sSql = " Select p.Fol_Poliza as poliza,"
        //        sSql = sSql + " p.ApP_Aseg+// //+p.ApM_Aseg+// //+p.Nom_Aseg as pensionado,"
        //        sSql = sSql + " p.Num_SegSocial as nssPensionado,"
        //        sSql = sSql + " ob.ApPBenef+// //+ob.ApMBenef+// //+ob.NomBenef as beneficiario,"
        //        sSql = sSql + " ob.Num_SegSocialBenef as nssBeneficiario,"
        //        sSql = sSql + " cp.Parentesco as parentesco,"
        //        sSql = sSql + " p.Fecha_Emision"
        //        sSql = sSql + " From DBCAIRO.dbo.Of_Benefs ob"
        //        sSql = sSql + " Inner Join DBCAIRO.dbo.Polizas p"
        //        sSql = sSql + " on ob.ID_Oferta = p.ID_Oferta"
        //        sSql = sSql + " Inner Join DBCAIRO.dbo.Cat_Parentescos cp"
        //        sSql = sSql + " on ob.ID_Parentesco = cp.ID_Parentesco"
        //        sSql = sSql + "  WHERE p.ID_StaPoliza = 2 And p.ID_InstitucionSS = 3 And ob.Num_SegSocialBenef IS NOT NULL And"
        //        sSql = sSql + " Fecha_Emision Between //" + parametros(0) + "// and //" + parametros(1) + "//"

        //        rsRecord = ADORecordset(cnnConexion, rsErrAdo, adLockOptimistic, adOpenKeyset, adUseClient, sSql)

        //        cnnConexion.Close()
        //        cnnConexion = Nothing
        //        bGetNSS = True
        //        Exit Function

        //msgerror:
        //        bGetNSS = False
        //        sErrVB = String.Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        rsErrAdo = ErroresDLL(cnnConexion, sErrVB)
        //        cnnConexion = Nothing
        //    End Function
        //    //------------------------------------------Fin--------------------------------------------------
        //    //------------------------Nueva funcionalidad de Consulta NSS Beneficiarios ISSSTE---------------
        //    //-------------------------------------JCMN 06/04/2016-------------------------------------------
        //    //-----------------------------------------------------------------------------------------------

        //}


    }
}
