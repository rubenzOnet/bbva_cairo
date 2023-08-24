using System.Configuration;
using MTSCPolizas;
using static bbva_cairo.Modulos.ModGeneral;
using System.Data;
//using System.Data.SqlClient;
using System.Data.OleDb;
using ADODB;

namespace bbva_cairo.Formularios
{
    public partial class frmCISSSTE : Form
    {

        public long iPolizaPrest;
        public bool bRe;
        public object gObjCPolizas;

        private Recordset RsPrestamos;
        private Recordset rsAbonos;
        private int iRes;

        public frmCISSSTE()
        {
            InitializeComponent();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCISSSTE_Load(object sender, EventArgs e)
        {
            //SDC 2010-01-15 Esto es para que cuando switchen entre pantalla y pantalla no se vuelva
            //a consultar, solo cuando venga de un click de frmCPolizas
            gsConexion = ConfigurationManager.AppSettings["ConString"];
            iPolizaPrest = 1002639723;

            GetPrestamos();
            grdPtmosActivos_Click();

            //if bRe Then
            //    GetPrestamos()
            //    //grdPtmosActivos_Click()
            //End if
            //bRe = False
        }

        private void GetPrestamos()
        {

            object[] vParametros = new object[3];

            this.Cursor = Cursors.WaitCursor;
            RsPrestamos = null;
            ClsCPolizas gObjCPolizas = new ClsCPolizas();


            vParametros[0] = iPolizaPrest;

            iRes = gObjCPolizas.bGetDetPrestISS(ref RsPrestamos, gsConexion, grsErrADO, vParametros);

            if (iRes == Convert.ToInt32(TipoResultado.DatosOK))
            {
                //double ID_Prestamo = Convert.ToDouble(RsPrestamos.Fields["ID_Prestamo"].Value);
                //string Prestamo = RsPrestamos.Fields["Prestamo"].Value;
                //string Tpo_Ptmo = RsPrestamos.Fields["Tpo_Ptmo"].Value;

                OleDbDataAdapter DataAd = new OleDbDataAdapter();
                DataTable dt = new DataTable();
                DataAd.Fill(dt, RsPrestamos);


                //SqlDataAdapter DataAd = new SqlDataAdapter();
                //DataTable dt = new DataTable();
                //dt = ModRecordset.ConvertRecordsetToDataTable(RsPrestamos);
                //DataAd.Fill(dt, RsPrestamos);

                grdPtmosActivos.DataSource = dt.DefaultView;
                grdPtmosActivos.Columns[0].Visible = false;

                //grdPtmosActivos.Caption = "Prestamos ISSSTE y FOVISSSTE (" & dt.DefaultView.Count & ")"
                GetDetalleISS(iPolizaPrest, Convert.ToDouble(RsPrestamos.Fields["ID_Prestamo"].Value), RsPrestamos.Fields["Prestamo"].Value, RsPrestamos.Fields["Tpo_Ptmo"].Value); //Alexander Hdez 01/10/2012 Agregue Prestamos FOVISSSTE

                //GetDetalleISS(iPolizaPrest, ID_Prestamo, Prestamo, Tpo_Ptmo); //Alexander Hdez 01/10/2012 Agregue Prestamos FOVISSSTE
                // grdDescXPtmos.Caption = "Descuentos Aplicados al Préstamo: " & RsPrestamos.Fields("No_Prestamo").Value
            }
            else
            {
                grdPtmosActivos.DataSource = null;
                //grdPtmosActivos.Caption = "Prestamos  (0)"
                LimpiaCampos();
            }

            gObjCPolizas = null;

            this.Cursor = Cursors.Default;

        }

        private void LimpiaCampos()
        {
            // Campos FmeISSSTE
            lblNoPtmo.Text = "";
            lblEstado.Text = "";
            lblNoDesc.Text = "0";
            lblSaldoActual.Text = "0.00";
            lblSaldoIni.Text = "0.00";
            lblDescApli.Text = "0";
            lblImporte.Text = "0.00";
        }

        private bool PrestamoConStatusBaja(long ID_Poliza, ref double ID_PrestamoActivo, ref int Plazo, ref double Saldo, ref double Saldo_Inicial, ref double Importe, ref int DescApli)
        {
            bool PrestamoConStatusBajaRes = false;
            Connection con;
            con = new Connection();
            Recordset rs;
            rs = new Recordset();
            con.Open(gsConexion);

            rs.Open(" select ID_Poliza,ID_Prestamo,Tipo_Orden, Plazo , Saldo, Saldo_Inicial, Importe, DescApli      from Prestamos_ISSSTE where Prestamos_ISSSTE.ID_Poliza = " + ID_Poliza, con, CursorTypeEnum.adOpenDynamic, LockTypeEnum.adLockOptimistic);


            while (!rs.EOF)
            {
                if (Convert.ToInt32(rs.Fields["Tipo_Orden"].Value) == 2)
                {
                    PrestamoConStatusBajaRes = true;
                }
                else
                {
                    ID_PrestamoActivo = Convert.ToDouble(rs.Fields["ID_Prestamo"].Value);
                    Plazo = Convert.ToInt32(rs.Fields["Plazo"].Value);
                    Saldo = Convert.ToDouble(rs.Fields["Saldo"].Value);
                    Saldo_Inicial = Convert.ToDouble(rs.Fields["Saldo_Inicial"].Value);
                    Importe = Convert.ToDouble(rs.Fields["Importe"].Value);
                    DescApli = Convert.ToInt32(rs.Fields["DescApli"].Value);
                }
                rs.MoveNext();
            }

            rs.Close();
            con.Close();

            return PrestamoConStatusBajaRes;
        }

        //private void GetDetalleISS(long lPoliza, double dIDPrestamo, string sPtmo, string Tpo_Ptmo) //Alexander Hdez 01/10/2012 Agregue Prestamos FOVISSSTE
        //{
        //    object[] vParametros = new object[2];
        //    this.Cursor = Cursors.WaitCursor;


        //    ClsCPolizas gObjCPolizas = new ClsCPolizas();

        //    vParametros[0] = lPoliza;
        //    vParametros[1] = dIDPrestamo;
        //    vParametros[2] = Tpo_Ptmo; //Alexander Hdez 01 / 10 / 2012 Agregue Prestamos FOVISSSTE

        //    fmeISSSTE.Visible = true;

        //    iRes = gObjCPolizas.bGetDetPtmoISS(ref rsAbonos, gsConexion, grsErrADO, vParametros);

        //    if (iRes == Convert.ToInt32(TipoResultado.DatosOK))
        //    {
        //        SqlDataAdapter DataAd = new SqlDataAdapter();
        //        DataSet ds = new DataSet();
        //        DataAd.Fill(ds, "rsAbonos");

        //        grdDescXPtmos.DataSource = ds;
        //    }

        //    gObjCPolizas = null;
        //    this.Cursor = Cursors.Default;
        //}

        private void grdPtmosActivos_Click()
        {
            string strEstatus;
            string TipoDescuento; //Alexander Hdez 01/10/2012 Prestamos FOVISSSTE


            if (RsPrestamos.RecordCount > 0)
            {
                //grdDescXPtmos.Caption = "Descuentos Aplicados al Préstamo: " & RsPrestamos.Fields("No_Prestamo").Value
                lblNoPtmo.Text = (RsPrestamos.Fields["No_Prestamo"].Value == null) ? "" : RsPrestamos.Fields["No_Prestamo"].Value.ToString();


                //<i> RCS, 29/DIC/2010 ,  10-0154 Prestamos ISSSTE

                //SI EL RsPrestamos!No_Prestamo , sehace una consulta y se encuentra que esta dado de baja
                //con el tipo_Orden =2en la tabla de Prestamos_ISSSTE , entonces regresar el prestamo del primero en estatus de vigente.

                double UltimoPrestamoActivo = 0;
                int Plazo = 0;
                double Saldo = 0;
                double Saldo_Inicial = 0;
                double Importe = 0;
                int DescApli = 0;


                if (PrestamoConStatusBaja(Convert.ToInt64(RsPrestamos.Fields["ID_Poliza"].Value), ref UltimoPrestamoActivo, ref Plazo, ref Saldo, ref Saldo_Inicial, ref Importe, ref DescApli))
                {
                    //GetDetalleISS iPolizaPrest, UltimoPrestamoActivo, RsPrestamos!Prestamo //Alexander Hdez 01 / 10 / 2012 Comente, Prestamos FOVISSSTE
                    GetDetalleISS(iPolizaPrest, UltimoPrestamoActivo, RsPrestamos.Fields["Prestamo"].Value, RsPrestamos.Fields["Tpo_Ptmo"].Value); //Alexander Hdez 01/10/2012 Agregue, Prestamos FOVISSSTE
                    lblNoDesc.Text = Plazo.ToString();
                    lblSaldoActual.Text = string.Format("{0:0.2}", Saldo);
                    lblSaldoIni.Text = string.Format("{0:0.2}", Saldo_Inicial);
                    //lblImporte.Text = String.Format(Importe, "#,##0.00")
                    lblImporte.Text = string.Format("{0:0.2}", RsPrestamos.Fields["Importe"].Value);
                    // lblDescApli.Caption = DescApli //Alexander Hdez 04 / 10 / 2012 Prestamos FOVISSSTE comente linea

                    //Alexander Hdez 04/10/2012, Prestamos FOVISSSTE, Muestra Descuentos Aplicados
                    //Inicio
                    if (RsPrestamos.Fields["Tpo_Ptmo"].Value == "Amortización" || RsPrestamos.Fields["Tpo_Ptmo"].Value == "Seguro de Daños")
                        ObtenDescApliFOVI(RsPrestamos.Fields["No_Prestamo"].Value, iPolizaPrest, RsPrestamos.Fields["Tpo_Ptmo"].Value);
                    else
                        lblDescApli.Text = string.Format("{0:0.2}", RsPrestamos.Fields["DescApli"].Value);
                    // End if
                    //Fin
                }
                else
                {
                    //GetDetalleISS iPolizaPrest, RsPrestamos!ID_Prestamo, RsPrestamos!Prestamo //Alexander Hdez 01 / 10 / 2012 Comente, Prestamos FOVISSSTE
                    GetDetalleISS(iPolizaPrest, Convert.ToDouble(RsPrestamos.Fields["ID_Prestamo"].Value), RsPrestamos.Fields["Prestamo"].Value, RsPrestamos.Fields["Tpo_Ptmo"].Value);  //Alexander Hdez 01/10/2012 Agregue, Prestamos FOVISSSTE
                    if (!rsAbonos.EOF)
                    {
                        //            Select Case rsAbonos!ID_StaPtmo
                        //                Case 1: strEstatus = "TRAMITE"
                        //                Case 2: strEstatus = "VIGENTE"
                        //                Case 3: strEstatus = "CANCELADO"
                        //                Case 4: strEstatus = "LIQUIDADO"
                        //            End Select

                        lblNoDesc.Text = string.Format("{0:0.2}", RsPrestamos.Fields["Plazo"].Value);
                        lblSaldoActual.Text = string.Format("{0:0.2}", RsPrestamos.Fields["Saldo"].Value);
                        lblSaldoIni.Text = string.Format("{0:0.2}", RsPrestamos.Fields["Saldo_Inicial"].Value);
                        lblImporte.Text = string.Format("{0:0.2}", RsPrestamos.Fields["Importe"].Value);
                        lblDescApli.Text = string.Format("{0:0.2}", RsPrestamos.Fields["DescApli"].Value); //Alexander Hdez 04/10/2012 Prestamos FOVISSSTE comente linea

                        //Alexander Hdez 04/10/2012, Prestamos FOVISSSTE, Muestra Descuentos Aplicados
                        //Inicio
                        if (RsPrestamos.Fields["Tpo_Ptmo"].Value == "Amortización" || RsPrestamos.Fields["Tpo_Ptmo"].Value == "Seguro de Daños")
                            ObtenDescApliFOVI(RsPrestamos.Fields["No_Prestamo"].Value, iPolizaPrest, RsPrestamos.Fields["Tpo_Ptmo"].Value);
                        else
                            lblDescApli.Text = string.Format("{0:0.2}", RsPrestamos.Fields["DescApli"].Value);

                        //Fin
                    }
                }

                //<f>  RCS, 29/DIC/2010 ,  10-0154 , Prestamos ISSSTE


                strEstatus = RsPrestamos.Fields["Status_Prestamo"].Value;
                lblEstado.Text = strEstatus;

            }


        }

        private void miFuncion()
        { }

        private void ObtenDescApliFOVI(int No_Prestamo, long iPolizaPrest, string Tpo_Ptmo)
        {
            Recordset RsDescApliFOVI;
            RsDescApliFOVI = new Recordset();
            object[] vParametros = new object[3];

            this.Cursor = Cursors.WaitCursor;
            RsDescApliFOVI = null;
            // gObjCPolizas = CreateObject("MTSCPolizas.ClsCPolizas")
            ClsCPolizas gObjCPolizas = new ClsCPolizas();

            vParametros[0] = No_Prestamo;
            vParametros[1] = iPolizaPrest;

            if (Tpo_Ptmo == "Amortización")
                vParametros[2] = "137";

            if (Tpo_Ptmo == "Seguro de Daños")
                vParametros[2] = "138";

            iRes = gObjCPolizas.bGetDescAplifOVI(RsDescApliFOVI, gsConexion, grsErrADO, vParametros); //.bGetDescApliFOVI(RsDescApliFOVI, gsConexion, grsErrADO, vParametros);

            if (iRes == Convert.ToInt32(TipoResultado.DatosOK))
                lblDescApli.Text = string.Format("{0:2}", RsDescApliFOVI.Fields["ID_Beneficio"].Value);


            gObjCPolizas = null;

            this.Cursor = Cursors.Default;
        }

        private void GetDetalleISS(long lPoliza, double dIDPrestamo, string sPtmo, string Tpo_Ptmo) // 'Alexander Hdez 01/10/2012 Agregue Prestamos FOVISSSTE
        {
            object[] vParametros = new object[3];

            this.Cursor = Cursors.WaitCursor;
            rsAbonos = null;

            ClsCPolizas gObjCPolizas = new ClsCPolizas();

            vParametros[0] = lPoliza;
            vParametros[1] = dIDPrestamo;
            vParametros[2] = Tpo_Ptmo; //''Alexander Hdez 01 / 10 / 2012 Agregue Prestamos FOVISSSTE

            fmeISSSTE.Visible = true;
            iRes = gObjCPolizas.bGetDetPtmoISS(ref rsAbonos, gsConexion, grsErrADO, vParametros);

            if (iRes == Convert.ToInt32(TipoResultado.DatosOK))
            {
                OleDbDataAdapter DataAd = new OleDbDataAdapter();
                DataTable dt = new DataTable();
                DataAd.Fill(dt, rsAbonos);

                grdDescXPtmos.DataSource = dt.DefaultView;
            }

            gObjCPolizas = null;
            this.Cursor = Cursors.Default;
        }




    }
}
