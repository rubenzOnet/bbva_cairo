//using System;
//using System.Collections.Generic;
//using System.Drawing.Printing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

namespace bbva_cairo.Classes
{
    public class ClsGeneral
    {

        //        //***************************************************************
        //        //Título:    Clase General del Sistema Cairo
        //        //Forma:     clsGeneral.cls
        //        //Versión:   1.0
        //        //Fecha:     15 / Junio / 2000
        //        //Autor:
        //        //---------------------------------------------------------------
        //        //Descripción:  Clase General del Sistema Cairo, contiene funciones de uso
        //        //              General para todo el Sistema
        //        //---------------------------------------------------------------

        //        //******Variables para Centrar una Forma a la pantalla origen*******************************************************************
        //        //5/07/2000 RReyes
        //        private bool bAjuste;
        //        private float sFactorX;
        //        private float sFactorY;

        //        //********API´S para la ejecución del SUC, sincroniza el tiempo y espera en un ciclo infinito la respuesta al sistema**************************
        //        //5/07/2000 RReyes
        //        private const SYNCHRONIZE = &H100000;
        //        private const INFINITE = &HFFFFFFFF       //  Infinite timeout

        //        private declare function OpenProcess Lib "kernel32" (dwDesiredAccess As Long, bInheritHandle As Long, dwProcessId As Long) As Long
        //        private Declare Function WaitForSingleObject Lib "kernel32" (hHandle As Long, dwMilliseconds As Long) As Long
        //        private Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Long

        //        //********API//S para el Nombre del Usuario y el de la Máquina*******************************************************************
        //        //5/07/2000 RReyes
        //        private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (lpBuffer As String, nSize As Long) As Long
        //        private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (lpBuffer As String, nSize As Long) As Long

        //        //********API//S para Movimiento del Mouse**********************************************************************************
        //        //5/07/2000 RReyes
        //        private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (lpFileName As String) As Long
        //        private Declare Function GlobalLock Lib "kernel32" (hMem As Long) As Long
        //        private Declare Function SetClassWord Lib "user32" (hwnd As Long, nIndex As Long, wNewWord As Long) As Long
        //        private Declare Function DestroyCursor Lib "user32" (hCursor As Long) As Long
        //        private Declare Function GlobalUnlock Lib "kernel32" (hMem As Long) As Long

        //        private long lCursorAni;

        //        //******Variables para El Módulo de Error VBError*******************************************************************
        //        //5/07/2000 RReyes
        //        private string sErrVB;

        //        //Activa el Movimiento del Mouse en True y desactiva el Movimiento en False
        //        //5/07/2000 RReyes
        //        public bool MouseActivo(ref object FForma, bool bValor )
        //        {
        //            if (bValor)
        //            { 
        //                lCursorAni = LoadCursorFromFile(Application.StartupPath & "\" & "Mouse.ani");
        //                if (lCursorAni != 0)
        //                {
        //                    GlobalLock(lCursorAni);
        //                }
        //                //SelectCursor = SetClassWord(FForma.hwnd, -12, lCursorAni)
        //                else
        //                {
        //                    GlobalUnlock(lCursorAni);
        //                    DestroyCursor(lCursorAni);
        //                }
        //            }
        //            return true;
        //        }

        //        //Rutina extrae el Nombre del Usuario
        //        //5/07/2000 RReyes
        //        public string Usuario()
        //        {
        //            string sBuffer = "";
        //            int lTamañoBuffer;
        //            bool bRes;

        //            //Crea Buffer
        //            // sBuffer = Space(255);
        //            lTamañoBuffer = sBuffer.Length;
        //            bRes = GetUserName(sBuffer, lTamañoBuffer);
        //            if (bRes)
        //            {
        //                return sBuffer.Substring(0, lTamañoBuffer - 1); //  left(sBuffer, lTamañoBuffer - 1);
        //                //Usuario = "MB22563"
        //            }
        //            else
        //            {
        //                return "No encontrado";
        //            }
        //        }

        //        //Rutina que extrae el Nombre de la Máquina
        //        //5/07/2000 RReyes
        //        public string Computadora()
        //        {
        //            string sBuffer;
        //            int lTamañoBuffer;
        //            bool bRes;
        //            //Crea Buffer
        //            sBuffer = "";
        //            lTamañoBuffer = sBuffer.Length;

        //            bRes = GetComputerName(sBuffer, lTamañoBuffer);

        //            if (bRes)
        //            {
        //                return sBuffer.Substring(0, lTamañoBuffer); //Left$(sBuffer, lTamañoBuffer)
        //                //Computadora = "MB22563"
        //            }
        //            else
        //            {
        //                return "No encontrado";
        //            }
        //        }

        //        //Se llama desde la forma pra la ejecución del SUC
        //        //5/07/2000 RReyes
        //        private bool EjecutaApp(string sarchivo)
        //        {
        //            long lpid;
        //            lpid = Shell(sarchivo, vbNormalFocus);
        //            Application.DoEvents();

        //            if (lpid != 0)
        //                EsperaATerminar(lpid, sarchivo);
        //            return true;
        //        }

        //        //Espera hasta que se regresa el el control de la aplicación del SUC
        //        //5/07/2000 RReyes
        //        private void EsperaATerminar(long lpid, string sarchivo)
        //        {
        //            long lphnd;
        //            lphnd = OpenProcess(SYNCHRONIZE, 0, lpid);
        //            if (lphnd != 0)
        //            {
        //                WaitForSingleObject(lphnd, INFINITE);
        //                CloseHandle(lphnd);
        //            }
        //        }

        //        //Inicializa el directorio del SUC para la ruta de su ejecutable.
        //        //5/07/2000 RReyes
        //        public void InicializaSUC(OpenFileDialog ObjCdgSUC, object vRuta)
        //        {
        //            // On Error Resume Next
        //            bbva_print bbva_Print = new bbva_print();
        //            object vExe;
        //            string sPath;


        //            if (new DirectoryInfo(Application.StartupPath + "\\" + vRuta + ".dat").Name == string.Empty )
        //            {
        //                ObjCdgSUC.ShowDialog();    

        //                if (ObjCdgSUC.FileName == "*.exe") return;

        //                sPath = ObjCdgSUC.FileName;

        //                bbva_Print.imprimir(sPath);

        //                //File.Open(Application.StartupPath + "\\" + vRuta + ".dat", FileMode.Open, FileAccess.Read);
        //                //PrintDocument pd = new PrintDocument();
        //                //pd.PrintPage += new PrintPageEventHandler(PrintPageCallback);

        //                //pd.Print();
        //                //Print(1, sPath);
        //                //FileClose(1);
        //                EjecutaApp(sPath);
        //            }
        //            else
        //                FileOpen(1, Application.StartupPath + "\" + vRuta + ".dat", OpenMode.Input)
        //                //Line(Input(1), sPath)
        //                FileClose(1)
        //                if (Dir(sPath) == String.Empty)
        //                    ObjCdgSUC.ShowOpen
        //                    sPath = ObjCdgSUC.FileName
        //                    if sPath = String.Empty Then Exit Sub
        //                    FileOpen(1, Application.StartupPath & "\" & vRuta & ".dat", OpenMode.Output)
        //                    Print(1, sPath)
        //                    FileClose(1)
        //                    EjecutaApp(sPath)
        //                else
        //                    if sPath = String.Empty Or sPath = "*.exe" Then
        //                        Kill(Application.StartupPath & "\" & vRuta & ".dat")
        //                        ObjCdgSUC.ShowOpen
        //                        sPath = ObjCdgSUC.FileName
        //                        if sPath = String.Empty Then Exit Sub
        //                        FileOpen(1, Application.StartupPath & "\" & vRuta & ".dat", OpenMode.Output)
        //                        Print(1, sPath)
        //                        FileClose(1)
        //                    End If
        //                    EjecutaApp(sPath)
        //                End If
        //            End If

        //        }

        //        //Ajusta la pantalla a una Resolución
        //        //5/07/2000 RReyes
        //        public Sub AjustarAResolucion(Optional f As Object = Nothing, Optional iXRes As Integer = 0)
        //            Dim iYRes As Integer

        //            On Error Resume Next

        //            if IsNothing(iXRes) Then iXRes = 800

        //            if iXRes = 800 Then
        //                iYRes = 600
        //            elseif iXRes = 640 Then
        //                iYRes = 480
        //            elseif iXRes = 1024 Then
        //                iYRes = 768
        //            elseif iXRes = 1280 Then
        //                iYRes = 1024
        //            else
        //                iXRes = 800
        //                iYRes = 600
        //            End If
        //            sFactorX = iXRes * Microsoft.VisualBasic.Compatibility.VB6.Support.PixelsToTwipsX(1) / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width
        //            sFactorY = iYRes * Microsoft.VisualBasic.Compatibility.VB6.Support.PixelsToTwipsY(1) / System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height
        //            If(sFactorX = 1 And sFactorY = 1) Or bAjuste Then
        //                bAjuste = True
        //                Exit Sub
        //            End If
        //            f.Visible = False
        //            Dim C As Object
        //            if f.WindowState = vbNormal Then
        //                AjusteNormal(f)
        //            End If
        //            For Each C In f.Controls
        //                Select Case LCase(TypeName(C.Container))
        //                    Case LCase(f.Name)
        //                        Select Case LCase(TypeName(C))
        //                            Case "label"
        //                                AjusteNormal(C)
        //                                C.AutoSize = C.AutoSize
        //                            Case "line"
        //                                C.X1 = C.X1 / sFactorX
        //                                C.X2 = C.X2 / sFactorX
        //                                C.Y1 = C.Y1 / sFactorY
        //                                C.Y2 = C.Y2 / sFactorY
        //                            Case "picturebox"
        //                                AjusteNormal(C)
        //              //C.Align = C.Align
        //                            Case "shape"
        //                                AjusteNormal(C)
        //           //No se ha detectado nada
        //                            Case "textbox"
        //                                AjusteNormal(C)
        //           //No se ha detectado nada excepto la escalabilidad de la fuente
        //                            Case "activebar2"
        //                                AjusteNormal(C)
        //                            Case else
        //                                //Shape
        //                                AjusteNormal(C)
        //                        End Select
        //                    Case "sstab"
        //                        Dim T As Integer
        //                        T = C.Container.Tab
        //                        C.Container.Tab = 0
        //                        Do
        //                            if Left$(Str(C.Left), 1) = "-" Then
        //                                C.Container.Tab = C.Container.Tab + 1
        //                            else
        //                                Exit Do
        //                            End If
        //                        Loop //While C.Container.Tab <= C.Container.Tabs
        //                        AjusteNormal(C)
        //                        C.Container.Tab = T
        //                    Case else
        //                        AjusteNormal(C)
        //                End Select
        //            Next
        //            bAjuste = True
        //            f.Visible = True
        //        End Sub

        //        //Función Privada para Ajustar la Pantalla Normal, esta función se coloca en el Load de las formas
        //        //5/07/2000 RReyes
        //        private Sub AjusteNormal(C2 As Object)
        //            On Error Resume Next
        //            //C2.Font.Size = C2.FontSize / sFactorX
        //            //C2.Height = C2.Height / sFactorY
        //            //C2.Width = C2.Width / sFactorX
        //            //C2.Left = C2.Left / sFactorX
        //            //C2.Top = C2.Top / sFactorY
        //        End Sub

        //        //Centra la Forma a la Pantalla, esta función se coloca en el Load de las formas
        //        //5/07/2000 RReyes
        //        public Sub Centrar(Optional FForma As Form = Nothing, Optional iXRes As Integer = 0)
        //            On Error Resume Next
        //            if IsNothing(iXRes) Then iXRes = 800
        //            if TypeName(FForma) = "Nothing" Then
        //                FForma = Form.ActiveForm
        //            End If
        //            AjustarAResolucion(FForma, iXRes)

        //            if FForma.IsMdiChild = True Then
        //                FForma.Top = ((frmRentas.Height - FForma.Height) / 2) - 850
        //                FForma.Left = ((frmRentas.Width - FForma.Width) / 2) - 100
        //            else
        //                //FForma.Move(frmRentas.Width - FForma.Width) / 2, (frmRentas.Height - FForma.Height) / 2
        //            End If
        //        End Sub

        //        //Selecciona el Texto, esta función se coloca en el GotFocus
        //        //5/07/2000 RReyes
        //        public Sub TextSelect()
        //            Dim iLenStr As Integer
        //            Dim ObjMyTextBox As Object

        //            //ObjMyTextBox = Screen.ActiveControl
        //            ObjMyTextBox = Form.ActiveForm.ActiveControl
        //            if TypeName(ObjMyTextBox) = "TextBox" Then
        //                iLenStr = Len(ObjMyTextBox.Text)
        //                ObjMyTextBox.SelStart = 0
        //                ObjMyTextBox.SelLength = iLenStr
        //            End If
        //        End Sub

        //        //Selecciona el Texto, esta función se coloca en el GotFocus
        //        //5/07/2000 RReyes
        //        public Function Valida09(ikeyascii As Integer) As Integer
        //            Dim sChar As String
        //            //Permitiendo unicamente la captura de mayusculas
        //            if Not IsNumeric(Chr(ikeyascii)) And ikeyascii!= 8 And ikeyascii!= 42 Then ikeyascii = 0
        //            Valida09 = ikeyascii
        //        End Function

        //        //Funcion SinApostrofe Busca la Comilla y la Quita del String
        //        public Function sSinApostrofe(sCad As String) As String
        //            On Error GoTo msgerror
        //            Dim iLong As Integer

        //            iLong = InStr(sCad, "//")
        //            if iLong!= 0 Then
        //                Mid(sCad, iLong) = " "
        //                sSinApostrofe = sSinApostrofe(sCad)
        //            else
        //                sSinApostrofe = sCad
        //            End If

        //            Exit Function
        //    msgerror:
        //            //Screen.MousePointer = vbDefault
        //            sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //            //cMsgbox.Mensaje(" Error del Sistema Prospect ", vbOKDetails + vbCritical, "C A I R O", , sErrVB)

        //        End Function

        //        //Ordena los datos del TDBGrid Ascendente y Descendentemente, se manda el recordset, el datafield del grid y el orden deseado
        //        //2/10/2000 CTorres
        //        public Function OrdenaGrid(rsdatos As ADODB.Recordset, sDataField As String, sOrden As String) As Integer
        //            On Error GoTo msgerror
        //            OrdenaGrid = 1
        //            rsdatos.Sort = sDataField & " " & sOrden
        //            Exit Function

        //    msgerror:
        //            sErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //            gfMsgbox.Mensaje(" Error del Sistema Prospect ", frmMensaje.ETipos.vbOKDetails + vbCritical, "C A I R O", , sErrVB)
        //        End Function

        //        //Ejecuta un paquete DTS en el servidor
        //        //22/10/2003 Samuel Dueñas
        //        //    public Function ExecuteDTS(sDTSName As String, sPkgPwd As String, Optional sFullPath As String = Empty, Optional sName As String = Empty, Optional bSolicitaNomArchivo As Boolean = False, Optional ParameterName As String, Optional ParameterValue As String, Optional EliminaArchivoPrevio As Boolean) As Boolean
        //        //        Dim fs As FileSystemObject
        //        //        Dim f As File
        //        //        Dim oPkg As DTS.Package
        //        //        Dim sarchivo As String
        //        //        Dim sExtension As String
        //        //        Dim sMsg As String
        //        //        On Error GoTo msgerror

        //        //        //SDC 2007 - 06 - 15 Modificaciones para ejecutar un DTS que genera archivos de Excel
        //        //        //SDC 2007 - 08 - 23 Corregir lógica para archivos xls cuando no se pide solicitar nombre
        //        //        fs = CreateObject("Scripting.FileSystemObject")

        //        //        //Revisar si mandan el parámetro de path de archivo de salida
        //        //        if sFullPath != String.Empty Then
        //        //            sExtension = Mid(sFullPath, InStr(sFullPath, ".") + 1)

        //        //            // * ********************************OJO * ************************************/
        //        //            //En caso de que sea de excel debe haber un archivo template para cada DTS
        //        //            //en el application path con el nombre del DTS con los nombres de las columnas
        //        //            //del reporte y el nombre del a hoja
        //        //            if LCase(sExtension) = "xls" Then
        //        //                fs.CopyFile app.path & "\" & sDTSName & ".xls", sFullPath, True //CEFB Oct2010 Se puede sobreescribir el archivo
        //        //            End If
        //        //        End If

        //        //        //SDC 2007 - 07 - 06 Solicitar el nombre antes de ejecutar el DTS
        //        //        if bSolicitaNomArchivo Then
        //        //            sarchivo = ObtenNombreArchivo(sExtension, "Proporcione el nombre del archivo")

        //        //            if sarchivo = String.Empty Then
        //        //                fs = Nothing
        //        //                Exit Function
        //        //            End If
        //        //        End If

        //        //        oPkg = New DTS.Package
        //        //        oPkg.LoadFromSQLServer gsServidor, "", "", DTSSQLStgFlag_UseTrustedConnection, "", "", "", sDTSName, 0

        //        //    // Rolando Coellar, Agregue el uso de parametros a los DTS
        //        //        if ParameterName != "" Then
        //        //            oPkg.GlobalVariables(ParameterName).Value = ParameterValue
        //        //        End If

        //        //        oPkg.FailOnError = True

        //        //        //SDC 2007 - 07 - 13 Para cambiar la base de datos depeniendo del connection string
        //        //        For X = 1 To oPkg.Connections.Count
        //        //            if oPkg.Connections(X).ProviderID = "SQLOLEDB" Then
        //        //                oPkg.Connections(X).Catalog = gsBaseDatos
        //        //            End If
        //        //        Next

        //        //        oPkg.Execute
        //        //        oPkg.UnInitialize
        //        //        oPkg = Nothing


        //        //        //RCS 2010 nov 24, para eliminar el archivo de excel de la ruta para evitar que se vayan acumulando registros y aparezcan duplicados.

        //        //        //if EliminaArchivoPrevio Then
        //        //        //    if fs.FileExists(sFullPath) Then
        //        //        //        fs.GetFile(sFullPath).Delete True
        //        //        //    End If
        //        //        //End If


        //        //        //Nueva funcionalidad para preguntar al usuario donde y con que nombre quiere su archivo.
        //        //        if bSolicitaNomArchivo And sFullPath != String.Empty Then
        //        //            if fs.FileExists(sarchivo) Then
        //        //                fs.GetFile(sarchivo).Delete True
        //        //        End If

        //        //            fs.MoveFile sFullPath, sarchivo
        //        //    else
        //        //            //Para cambiarle el nombre al archivo
        //        //            if sFullPath != String.Empty And sName != String.Empty Then
        //        //                f = fs.GetFile(sFullPath)

        //        //                sarchivo = f.ParentFolder & "\" & sName

        //        //                //Si ya existe el archivo, borrarlo
        //        //                if fs.FileExists(sarchivo) Then
        //        //                    fs.GetFile(sarchivo).Delete True
        //        //            End If
        //        //                f.Name = sName
        //        //                f = Nothing
        //        //            End If
        //        //        End If

        //        //        fs = Nothing
        //        //        ExecuteDTS = True
        //        //        Exit Function
        //        //msgerror:
        //        //        fs = Nothing
        //        //        ExecuteDTS = False
        //        //        if Not oPkg Is Nothing Then
        //        //            sMsg = "Error en el paquete " & oPkg.Name & ": " & vbCrLf & sAccumStepErrors(oPkg)
        //        //        End If
        //        //        gsErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        //        gfMsgbox.Mensaje "Error de sistema" & vbCrLf & sMsg, vbOKDetails + vbCritical, "CAIRO", , gsErrVB
        //        //End Function

        //        //private Function sAccumStepErrors(objPackage As DTS.Package) As String
        //        //    //Accumulate the step error info into the error message.
        //        //    Dim oStep As DTS.Step
        //        //    Dim sMessage As String
        //        //    Dim lErrNum As Long
        //        //    Dim sDescr As String
        //        //    Dim sSource As String

        //        //    //Look for steps that completed and failed.
        //        //    For Each oStep In objPackage.Steps
        //        //        if oStep.ExecutionStatus = DTSStepExecStat_Completed Then
        //        //            if oStep.ExecutionResult = DTSStepExecResult_Failure Then

        //        //                //Get the step error information and append it to the message.
        //        //                oStep.GetExecutionErrorInfo lErrNum, sSource, sDescr
        //        //            sMessage = sMessage & vbCrLf &
        //        //                    "Step: " & oStep.Description & " failed, error: " &
        //        //                    vbCrLf & sDescr & vbCrLf
        //        //            End If
        //        //        End If
        //        //    Next
        //        //    sAccumStepErrors = sMessage
        //        //End Function

        //        //private Function sErrorNumConv(lErrNum As Long) As String
        //        //    //Convert the error number into readable forms, both hexadecimal and decimal for the low - order word.

        //        //    if lErrNum < 65536 And lErrNum > -65536 Then
        //        //        sErrorNumConv = "x" & Hex(lErrNum) & ",  " & CStr(lErrNum)
        //        //    else
        //        //        sErrorNumConv = "x" & Hex(lErrNum) & ",  x" &
        //        //            Hex(lErrNum And -65536) & " + " & CStr(lErrNum And 65535)
        //        //    End If
        //        //End Function

        //        //Incia construccion del archivo DTS_PgoNom.xlsx Erick Bejarano 06/11/2017 ------------------------------------------
        //        //public Sub DTS_PgoNom(sFileName As String, iHoja As Integer, rs As ADODB.Recordset, iCorrida As Integer) // 20200123 Julio
        //        //    public Sub DTS_PgoNom(sFileName As String, iHoja As Integer, rs As ADODB.Recordset, iCorrida As String) // 20200123 Julio

        //        //        Dim oExcel As Excel.Application
        //        //        Dim oSheet1 As Excel.Worksheet
        //        //        Dim oSheet2 As Excel.Worksheet

        //        //        Dim sEtiqueta As String
        //        //        Dim Cadena As String
        //        //        Dim sarchivo As String

        //        //        On Error GoTo msgerror

        //        //        SetStatusBar("Generando reporte, espere un momento ...")
        //        //        Screen.MousePointer = vbHourglass

        //        //        oExcel = CreateObject("Excel.Application")
        //        //        DoEvents

        //        //        //FJMH 05 / 02 / 2021 Inicio Generando Cuadro de Dialogo para Guardar archivo en ruta
        //        //        //sarchivo = "C:\Dispersion\" & sFileName & iCorrida & ".xls"
        //        //        sarchivo = sFileName
        //        //        //FJMH 05 / 02 / 2021 Fin Generando Cuadro de Dialogo para Guardar archivo en ruta
        //        //        if Len(Trim$(Dir$(sarchivo))) = 0 Then

        //        //            //Crea el archivo
        //        //            oExcel.Workbooks.Add()

        //        //            While oExcel.Sheets.Count < 2  //ASG correcion para nuevas versiones de excel que solo abren una hoja
        //        //                oExcel.Sheets.Add , oExcel.ActiveSheet
        //        //        Wend

        //        //        oSheet1 = oExcel.Worksheets("Hoja1")
        //        //            oSheet1.Name = "CONDUCTO"

        //        //            oSheet2 = oExcel.Worksheets("Hoja2")
        //        //            oSheet2.Name = "ROPC"

        //        //            //else: oExcel.Workbooks.Open FileName:= "C:\Dispersion\" & sFileName & iCorrida & ".xls"
        //        //        else : oExcel.Workbooks.Open Filename:=sarchivo
        //        //    //FJMH 05 / 02 / 2021 Fin Generando Cuadro de Dialogo para Guardar archivo en ruta
        //        //        End If

        //        //        //construye rotulos------------------------------------------------------------
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 1).Formula = "Corrida"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 2).Formula = "Conducto"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 3).Formula = "Empresa"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 4).Formula = "InstitucionSS"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 5).Formula = "Fol_Poliza"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 6).Formula = "ID_Grupo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 7).Formula = "Ramo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 8).Formula = "Mes"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 9).Formula = "Cambios"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 10).Formula = "Pagos Vencidos"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 11).Formula = "Pension Basica"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 12).Formula = "Viudas"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 13).Formula = "Incremento 11%"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 14).Formula = "Aguinaldo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 15).Formula = "Aguinaldo Inc 11%"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 16).Formula = "Finiquito"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 17).Formula = "Finiquito Inc 11%"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 18).Formula = "Retroactivo HS"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 19).Formula = "BAMI"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 20).Formula = "BAMI Aguinaldo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 21).Formula = "BAMI Finiquito"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 22).Formula = "Pension Adicional"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 23).Formula = "Aguinaldo Adicional"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 24).Formula = "Ayuda Escolar"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 25).Formula = "Abono Grupo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 26).Formula = "Descuentos Grupo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 27).Formula = "Descuentos Otros"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 28).Formula = "Prestamo ATM"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 29).Formula = "Prestamo Seguros"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 30).Formula = "Prestamo Pensiones"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 31).Formula = "BAU"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 32).Formula = "Aguinaldo BAU"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 33).Formula = "Prestamo Personal ISSSTE"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 34).Formula = "Prestamo Hipotecario"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 35).Formula = "Descuento de Seguro"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 36).Formula = "GtosFun_Endoso"
        //        //        if (oExcel.Workbooks(1).Sheets(iHoja).Name = "CONDUCTO") Then
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 37).Formula = "Retension_ISR"
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 38).Formula = "Total"
        //        //        else
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 37).Formula = "Total"
        //        //        End If


        //        //        //   //construye detalle --------------------------------------------------------------
        //        //        //    X = 2
        //        //        //    Do While Not rs.EOF = True
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 1).Formula = rs.Fields(0)                           //--< corrida
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 2).Formula = rs.Fields(1)                           //--< Conducto
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 3).Formula = rs.Fields(2)                           //--< Empresa
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 4).Formula = rs.Fields(3)                           //--< InstitucionSS
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 5).Formula = rs.Fields(4)                           //--< Fol_Poliza
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 6).Formula = rs.Fields(5)                           //--< ID_Grupo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 7).Formula = rs.Fields(6)                           //--< Ramo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 8).Formula = Format(rs.Fields(7), "yyyy-mm-dd")  //--< Mes
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 9).Formula = rs.Fields(8)                               //--< Cambios
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 10).Formula = FormatCurrency(rs.Fields(9), 2, vbFalse, vbFalse, vbFalse) //--< Pagos_Vencidos
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 11).Formula = FormatCurrency(rs.Fields(10), 2, vbFalse, vbFalse, vbFalse) //--< Pension_Basica
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 12).Formula = FormatCurrency(rs.Fields(11), 2, vbFalse, vbFalse, vbFalse) //--< Viudas
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 13).Formula = FormatCurrency(rs.Fields(12), 2, vbFalse, vbFalse, vbFalse) //--< Incremento_11
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 14).Formula = FormatCurrency(rs.Fields(13), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 15).Formula = FormatCurrency(rs.Fields(14), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo_Inc_11
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 16).Formula = FormatCurrency(rs.Fields(15), 2, vbFalse, vbFalse, vbFalse) //--< Finiquito
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 17).Formula = FormatCurrency(rs.Fields(16), 2, vbFalse, vbFalse, vbFalse) //--< Finiquito_Inc_11
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 18).Formula = FormatCurrency(rs.Fields(17), 2, vbFalse, vbFalse, vbFalse) //--< Retroactivo_HS
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 19).Formula = FormatCurrency(rs.Fields(18), 2, vbFalse, vbFalse, vbFalse) //--< BAMI
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 20).Formula = FormatCurrency(rs.Fields(19), 2, vbFalse, vbFalse, vbFalse) //--< BAMI_Aguinaldo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 21).Formula = FormatCurrency(rs.Fields(20), 2, vbFalse, vbFalse, vbFalse) //--< BAMI_Finiquito
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 22).Formula = FormatCurrency(rs.Fields(21), 2, vbFalse, vbFalse, vbFalse) //--< Pension_Adicional
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 23).Formula = FormatCurrency(rs.Fields(22), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo_Adicional
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 24).Formula = FormatCurrency(rs.Fields(23), 2, vbFalse, vbFalse, vbFalse) //--< Ayuda_Escolar
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 25).Formula = FormatCurrency(rs.Fields(24), 2, vbFalse, vbFalse, vbFalse) //--< Abono_Grupo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 26).Formula = FormatCurrency(rs.Fields(25), 2, vbFalse, vbFalse, vbFalse) //--< Descuentos_Grupo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 27).Formula = FormatCurrency(rs.Fields(26), 2, vbFalse, vbFalse, vbFalse) //--< Descuentos_Otros
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 28).Formula = FormatCurrency(rs.Fields(27), 2, vbFalse, vbFalse, vbFalse) //--< Prestamo_ATM
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 29).Formula = FormatCurrency(rs.Fields(28), 2, vbFalse, vbFalse, vbFalse) //--< Prestamo_Seguros
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 30).Formula = FormatCurrency(rs.Fields(29), 2, vbFalse, vbFalse, vbFalse) //--< Prestamo_Pensiones
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 31).Formula = FormatCurrency(rs.Fields(30), 2, vbFalse, vbFalse, vbFalse) //--< BAU
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 32).Formula = FormatCurrency(rs.Fields(31), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo_BAU
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 33).Formula = FormatCurrency(rs.Fields(32), 2, vbFalse, vbFalse, vbFalse) //--< Prest_Perso_ISSSTE
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 34).Formula = FormatCurrency(rs.Fields(33), 2, vbFalse, vbFalse, vbFalse) //--< Prest_Hipotecario
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 35).Formula = FormatCurrency(rs.Fields(34), 2, vbFalse, vbFalse, vbFalse) //--< Descuento_Seguro
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 36).Formula = FormatCurrency(rs.Fields(35), 2, vbFalse, vbFalse, vbFalse) //--< GtosFun_Endoso
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 37).Formula = FormatCurrency(rs.Fields(36), 2, vbFalse, vbFalse, vbFalse) //--< Total
        //        //        //
        //        //        //        X = X + 1
        //        //        //        rs.MoveNext
        //        //        //    Loop

        //        //        oExcel.Workbooks(1).Sheets(iHoja).Range("A2").CopyFromRecordrs

        //        //        if Not rs.EOF Then
        //        //            //Formato a las celdas de importes
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Activate
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Range(oExcel.Workbooks(1).Sheets(iHoja).Cells(2, 10), oExcel.Workbooks(1).Sheets(iHoja).Cells(rs.RecordCount + 1, 37)).Select
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Range(oExcel.Workbooks(1).Sheets(iHoja).Cells(2, 10), oExcel.Workbooks(1).Sheets(iHoja).Cells(rs.RecordCount + 1, 37)).NumberFormat = "$#,##0.00"
        //        //        End If

        //        //        if Len(Trim$(Dir$(sarchivo))) = 0 Then
        //        //            //Guarda el libro
        //        //            //FJMH Inicio
        //        //            //oExcel.Workbooks(1).SaveAs(sFileName & iCorrida & ".xls")
        //        //            oExcel.Workbooks(1).SaveAs(sFileName)
        //        //            //FJMH Fin
        //        //        else : oExcel.Workbooks(1).Save()
        //        //        End If

        //        //        oExcel.Workbooks.Close()
        //        //        oExcel.Quit()
        //        //        oSheet = Nothing
        //        //        oExcel = Nothing

        //        //        rs.Close()
        //        //        rs = Nothing

        //        //        if iHoja = 2 Then
        //        //            //FJMH Inicio
        //        //            //gfMsgbox.Mensaje "Se generó exitosamente en: " & sFileName & iCorrida & ".xls", vbInformation, "CAIRO"
        //        //            gfMsgbox.Mensaje "Se generó exitosamente en: " & sFileName, vbInformation, "CAIRO"
        //        //        //FJMH Fin
        //        //            SetStatusBar("Listo ...")
        //        //            Screen.MousePointer = vbDefault
        //        //        End If

        //        //        Exit Sub

        //        //msgerror:
        //        //        Screen.MousePointer = vbDefault
        //        //        gsErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        //        gfMsgbox.Mensaje " Error del Sistema Cairo ", vbOKDetails + vbCritical, "C A I R O", , gsErrVB
        //        //End Sub
        //        //    //Fin construccion del archivo DTS_PgoNom.xlsx Erick Bejarano 26 / 07 / 2017------------------------------------------

        //        //    //Incia construccion del archivo DTS_PgoNomComp.xlsx Erick Bejarano 06 / 11 / 2017------------------------------------------
        //        //    //public Sub DTS_PgoNomComp(iHoja As Integer, iTipo As Integer, rs As ADODB.Recordset, iCorrida As Integer)28 / 01 / 2020
        //        //    public Sub DTS_PgoNomComp(iHoja As Integer, iTipo As Integer, rs As ADODB.Recordset, iCorrida As String)

        //        //        Dim oExcel As Excel.Application
        //        //        Dim oSheet1 As Excel.Worksheet
        //        //        Dim oSheet2 As Excel.Worksheet

        //        //        Dim sEtiqueta As String
        //        //        Dim Cadena As String
        //        //        Dim sarchivo As String
        //        //        Dim sFileName As String

        //        //        On Error GoTo msgerror

        //        //        SetStatusBar("Generando reporte, espere un momento ...")
        //        //        Screen.MousePointer = vbHourglass

        //        //        oExcel = CreateObject("Excel.Application")
        //        //        DoEvents

        //        //        sarchivo = "C:\Dispersion\" & "Pgo_NomC" & IIf(iTipo = 1, "F", "") & iCorrida & ".xls"
        //        //        if Len(Trim$(Dir$(sarchivo))) = 0 Then

        //        //            //Crea el archivo
        //        //            oExcel.Workbooks.Add()

        //        //            oSheet1 = oExcel.Worksheets("Hoja1")
        //        //            oSheet1.Name = "Pgo_Nom1"

        //        //            oSheet2 = oExcel.Worksheets("Hoja2")
        //        //            oSheet2.Name = "Pgo_Nom2"

        //        //        else : oExcel.Workbooks.Open Filename:="C:\Dispersion\" & "Pgo_NomC" & IIf(iTipo = 1, "F", "") & iCorrida & ".xls"
        //        //    End If

        //        //        //construye rotulos------------------------------------------------------------
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 1).Formula = "ID_CorridaAnt"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 2).Formula = "ID_CorridaNva"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 3).Formula = "Cambios"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 4).Formula = "Empresa"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 5).Formula = "InstitucionSS"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 6).Formula = "Fol_Poliza"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 7).Formula = "Pension_Basica"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 8).Formula = "Art14_2002"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 9).Formula = "Art14_2004"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 10).Formula = "Aguinaldo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 11).Formula = "Ag_Art14_2004"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 12).Formula = "Finiquito"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 13).Formula = "Finiquito_Art14_2004"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 14).Formula = "BAMI"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 15).Formula = "BAMI_Aguinaldo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 16).Formula = "BAMI_Finiquito"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 17).Formula = "Pension_Adicional"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 18).Formula = "Aguinaldo_Adicional"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 19).Formula = "Ayuda_Escolar"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 20).Formula = "Abono_Grupo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 21).Formula = "Descuentos_Grupo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 22).Formula = "Descuentos_Otros"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 23).Formula = "Prestamos_ATM"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 24).Formula = "Prestamos_Seguros"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 25).Formula = "Prestamos_Pensiones"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 26).Formula = "BAU"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 27).Formula = "Aguinaldo_BAU"

        //        //        if iHoja = 2 Then
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 28).Formula = "Gratif_AnualPagMensual"
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 29).Formula = "Gratif_AnualPagMensualND"
        //        //        End If

        //        //        //construye detalle--------------------------------------------------------------
        //        //        //    X = 2
        //        //        //    Do While Not rs.EOF = True
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 1).Formula = rs.Fields(0)                           //--< ID_CorridaAnt
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 2).Formula = rs.Fields(1)                           //--< ID_CorridaNva
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 3).Formula = rs.Fields(2)                           //--< Cambios
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 4).Formula = rs.Fields(3)                           //--< Empresa
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 5).Formula = rs.Fields(26)                          //--< InstitucionSS
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 6).Formula = rs.Fields(4)                           //--< Fol_Poliza
        //        //        //
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 7).Formula = FormatCurrency(rs.Fields(5), 2, vbFalse, vbFalse, vbFalse)  //--< Pension_Basica
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 8).Formula = FormatCurrency(rs.Fields(6), 2, vbFalse, vbFalse, vbFalse)  //--< Art14_2002
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 9).Formula = FormatCurrency(rs.Fields(7), 2, vbFalse, vbFalse, vbFalse)  //--< Art14_2004
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 10).Formula = FormatCurrency(rs.Fields(8), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 11).Formula = FormatCurrency(rs.Fields(9), 2, vbFalse, vbFalse, vbFalse) //--< Ag_Art14_2004
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 12).Formula = FormatCurrency(rs.Fields(10), 2, vbFalse, vbFalse, vbFalse) //--< Finiquito             -- CORRECTO
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 13).Formula = FormatCurrency(rs.Fields(11), 2, vbFalse, vbFalse, vbFalse) //--< Finiquito_Art14_2004
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 14).Formula = FormatCurrency(rs.Fields(12), 2, vbFalse, vbFalse, vbFalse) //--< BAMI
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 15).Formula = FormatCurrency(rs.Fields(13), 2, vbFalse, vbFalse, vbFalse) //--< BAMI_Aguinaldo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 16).Formula = FormatCurrency(rs.Fields(14), 2, vbFalse, vbFalse, vbFalse) //--< BAMI_Finiquito
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 17).Formula = FormatCurrency(rs.Fields(15), 2, vbFalse, vbFalse, vbFalse) //--< Pension_Adicional
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 18).Formula = FormatCurrency(rs.Fields(16), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo_Adicional
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 19).Formula = FormatCurrency(rs.Fields(17), 2, vbFalse, vbFalse, vbFalse) //--< Ayuda_Escolar
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 20).Formula = FormatCurrency(rs.Fields(18), 2, vbFalse, vbFalse, vbFalse) //--< Abono_Grupo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 21).Formula = FormatCurrency(rs.Fields(19), 2, vbFalse, vbFalse, vbFalse) //--< Descuentos_Grupo
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 22).Formula = FormatCurrency(rs.Fields(20), 2, vbFalse, vbFalse, vbFalse) //--< Descuentos_Otros      -- CORRECTO
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 23).Formula = FormatCurrency(rs.Fields(21), 2, vbFalse, vbFalse, vbFalse) //--< Prestamos_ATM
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 24).Formula = FormatCurrency(rs.Fields(22), 2, vbFalse, vbFalse, vbFalse) //--< Prestamos_Seguros
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 25).Formula = FormatCurrency(rs.Fields(23), 2, vbFalse, vbFalse, vbFalse) //--< Prestamos_Pensiones
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 26).Formula = FormatCurrency(rs.Fields(24), 2, vbFalse, vbFalse, vbFalse) //--< BAU
        //        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 27).Formula = FormatCurrency(rs.Fields(25), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo_BAU
        //        //        //
        //        //        //        if iHoja = 2 Then
        //        //        //            oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 28).Formula = FormatCurrency(rs.Fields(27), 2, vbFalse, vbFalse, vbFalse) //--< Gratif_AnualPagMensual
        //        //        //            oExcel.Workbooks(1).Sheets(iHoja).Cells(X, 29).Formula = FormatCurrency(rs.Fields(28), 2, vbFalse, vbFalse, vbFalse) //--< Gratif_AnualPagMensualND
        //        //        //        End If
        //        //        //
        //        //        //        X = X + 1
        //        //        //        rs.MoveNext
        //        //        //    Loop

        //        //        oExcel.Workbooks(1).Sheets(iHoja).Range("A2").CopyFromRecordrs

        //        //        if Not rs.EOF Then
        //        //            //Formato a las celdas de importes
        //        //            oExcel.Workbooks(1).Sheets(iHoja).Activate

        //        //            if iHoja = 1 Then
        //        //                oExcel.Workbooks(1).Sheets(iHoja).Range(oExcel.Workbooks(1).Sheets(iHoja).Cells(2, 7), oExcel.Workbooks(1).Sheets(iHoja).Cells(rs.RecordCount + 1, 27)).Select
        //        //                oExcel.Workbooks(1).Sheets(iHoja).Range(oExcel.Workbooks(1).Sheets(iHoja).Cells(2, 7), oExcel.Workbooks(1).Sheets(iHoja).Cells(rs.RecordCount + 1, 27)).NumberFormat = "$#,##0.00"
        //        //            else
        //        //                oExcel.Workbooks(1).Sheets(iHoja).Range(oExcel.Workbooks(1).Sheets(iHoja).Cells(2, 7), oExcel.Workbooks(1).Sheets(iHoja).Cells(rs.RecordCount + 1, 29)).Select
        //        //                oExcel.Workbooks(1).Sheets(iHoja).Range(oExcel.Workbooks(1).Sheets(iHoja).Cells(2, 7), oExcel.Workbooks(1).Sheets(iHoja).Cells(rs.RecordCount + 1, 29)).NumberFormat = "$#,##0.00"
        //        //            End If
        //        //        End If

        //        //        if Len(Trim$(Dir$(sarchivo))) = 0 Then
        //        //            //Guarda el libro
        //        //            oExcel.Workbooks(1).SaveAs("C:\Dispersion\" & "Pgo_NomC" & IIf(iTipo = 1, "F", "") & iCorrida & ".xls")
        //        //        else : oExcel.Workbooks(1).Save()
        //        //        End If

        //        //        oExcel.Workbooks.Close()
        //        //        oExcel.Quit()
        //        //        oSheet = Nothing
        //        //        oExcel = Nothing

        //        //        rs.Close()
        //        //        rs = Nothing

        //        //        if iHoja = 2 Then
        //        //            gfMsgbox.Mensaje "Se generó exitosamente en C:\Dispersion\" & "Pgo_NomC" & IIf(iTipo = 1, "F", "") & iCorrida & ".xls", vbInformation, "CAIRO"
        //        //        SetStatusBar("Listo ...")
        //        //            Screen.MousePointer = vbDefault
        //        //        End If

        //        //        Exit Sub

        //        //msgerror:
        //        //        Screen.MousePointer = vbDefault
        //        //        gsErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        //        gfMsgbox.Mensaje " Error del Sistema Cairo ", vbOKDetails + vbCritical, "C A I R O", , gsErrVB
        //        //End Sub
        //        //    //Fin construccion del archivo DTS_PgoNomComp.xlsx Erick Bejarano 26 / 07 / 2017------------------------------------------

        //        //    //Incia construccion del archivo PgoNomEmision.xlsx Erick Bejarano 06 / 11 / 2017------------------------------------------
        //        //    //public Sub DTS_PgoNomEmision(sFileName As String, rs As ADODB.Recordset, iCorrida As Integer)28 / 01 / 2020
        //        //    public Sub DTS_PgoNomEmision(sFileName As String, rs As ADODB.Recordset, iCorrida As String)

        //        //        Dim oExcel As Excel.Application
        //        //        Dim oSheet1 As Excel.Worksheet
        //        //        Dim oSheet2 As Excel.Worksheet

        //        //        Dim sEtiqueta As String
        //        //        Dim Cadena As String
        //        //        Dim sarchivo As String

        //        //        On Error GoTo msgerror

        //        //        SetStatusBar("Generando reporte, espere un momento ...")
        //        //        Screen.MousePointer = vbHourglass

        //        //        oExcel = CreateObject("Excel.Application")
        //        //        DoEvents
        //        //        //FJMH 04 / 02 / 2021 Inicio Guardar archivo
        //        //        //sarchivo = "C:\Dispersion\" & sFileName & iCorrida & ".xls"
        //        //        sarchivo = sFileName
        //        //        //FJMH 04 / 02 / 2021 Fin
        //        //        if Len(Trim$(Dir$(sarchivo))) = 0 Then

        //        //            //Crea el archivo
        //        //            oExcel.Workbooks.Add()

        //        //            oSheet1 = oExcel.Worksheets("Hoja1")
        //        //            //FJMH 04 / 02 / 2021 Inicio
        //        //            //else: oExcel.Workbooks.Open FileName:= "C:\Dispersion\" & sFileName & iCorrida & ".xls"
        //        //        else : oExcel.Workbooks.Open Filename:=sFileName
        //        //    //FJMH 04 / 02 / 2021 Fin
        //        //        End If

        //        //        //construye rotulos------------------------------------------------------------
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 1).Formula = "Corrida"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 2).Formula = "Conducto"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 3).Formula = "Empresa"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 4).Formula = "Fol_Poliza"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 5).Formula = "ID_Grupo"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 6).Formula = "Ramo"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 7).Formula = "Mes"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 8).Formula = "Cambios"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 9).Formula = "Pagos Vencidos"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 10).Formula = "Pension Basica"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 11).Formula = "Viudas"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 12).Formula = "Incremento 11%"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 13).Formula = "Aguinaldo"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 14).Formula = "Aguinaldo Inc 11%"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 15).Formula = "Finiquito"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 16).Formula = "Finiquito Inc 11%"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 17).Formula = "Retroactivo HS"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 18).Formula = "BAMI"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 19).Formula = "BAMI Aguinaldo"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 20).Formula = "BAMI Finiquito"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 21).Formula = "Pension Adicional"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 22).Formula = "Aguinaldo Adicional"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 23).Formula = "Ayuda Escolar"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 24).Formula = "Abono Grupo"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 25).Formula = "Descuentos Grupo"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 26).Formula = "Descuentos Otros"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 27).Formula = "Prestamo ATM"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 28).Formula = "Prestamo Seguros"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 29).Formula = "Prestamo Pensiones"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 30).Formula = "BAU"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 31).Formula = "Aguinaldo BAU"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 32).Formula = "Ptmo.Hipotecario"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 33).Formula = "Descto.Seguro"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 34).Formula = "InstitucionSS"
        //        //        oExcel.Workbooks(1).Sheets(1).Cells(1, 35).Formula = "Total"

        //        //        //construye detalle--------------------------------------------------------------
        //        //        //    X = 2
        //        //        //    Do While Not rs.EOF = True
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 1).Formula = rs.Fields(0)                           //--< corrida
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 2).Formula = rs.Fields(1)                           //--< Conducto
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 3).Formula = rs.Fields(2)                           //--< Empresa
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 4).Formula = rs.Fields(3)                           //--< Fol_Poliza
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 5).Formula = rs.Fields(4)                           //--< ID_Grupo
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 6).Formula = rs.Fields(5)                           //--< Ramo
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 7).Formula = Format(rs.Fields(6), "yyyy-mm-dd")      //--< Mes
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 8).Formula = rs.Fields(7)                               //--< Cambios
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 9).Formula = FormatCurrency(rs.Fields(8), 2, vbFalse, vbFalse, vbFalse) //--< Pagos_Vencidos
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 10).Formula = FormatCurrency(rs.Fields(9), 2, vbFalse, vbFalse, vbFalse) //--< Pension_Basica
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 11).Formula = FormatCurrency(rs.Fields(10), 2, vbFalse, vbFalse, vbFalse) //--< Viudas
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 12).Formula = FormatCurrency(rs.Fields(11), 2, vbFalse, vbFalse, vbFalse) //--< Incremento_11
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 13).Formula = FormatCurrency(rs.Fields(12), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 14).Formula = FormatCurrency(rs.Fields(13), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo_Inc_11
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 15).Formula = FormatCurrency(rs.Fields(14), 2, vbFalse, vbFalse, vbFalse) //--< Finiquito
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 16).Formula = FormatCurrency(rs.Fields(15), 2, vbFalse, vbFalse, vbFalse) //--< Finiquito_Inc_11
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 17).Formula = FormatCurrency(rs.Fields(16), 2, vbFalse, vbFalse, vbFalse) //--< Retroactivo_HS
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 18).Formula = FormatCurrency(rs.Fields(17), 2, vbFalse, vbFalse, vbFalse) //--< BAMI
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 19).Formula = FormatCurrency(rs.Fields(18), 2, vbFalse, vbFalse, vbFalse) //--< BAMI_Aguinaldo
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 20).Formula = FormatCurrency(rs.Fields(19), 2, vbFalse, vbFalse, vbFalse) //--< BAMI_Finiquito
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 21).Formula = FormatCurrency(rs.Fields(20), 2, vbFalse, vbFalse, vbFalse) //--< Pension_Adicional
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 22).Formula = FormatCurrency(rs.Fields(21), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo_Adicional
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 23).Formula = FormatCurrency(rs.Fields(22), 2, vbFalse, vbFalse, vbFalse) //--< Ayuda_Escolar
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 24).Formula = FormatCurrency(rs.Fields(23), 2, vbFalse, vbFalse, vbFalse) //--< Abono_Grupo
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 25).Formula = FormatCurrency(rs.Fields(24), 2, vbFalse, vbFalse, vbFalse) //--< Descuentos_Grupo
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 26).Formula = FormatCurrency(rs.Fields(25), 2, vbFalse, vbFalse, vbFalse) //--< Descuentos_Otros
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 27).Formula = FormatCurrency(rs.Fields(26), 2, vbFalse, vbFalse, vbFalse) //--< Prestamo_ATM
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 28).Formula = FormatCurrency(rs.Fields(27), 2, vbFalse, vbFalse, vbFalse) //--< Prestamo_Seguros
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 29).Formula = FormatCurrency(rs.Fields(28), 2, vbFalse, vbFalse, vbFalse) //--< Prestamo_Pensiones
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 30).Formula = FormatCurrency(rs.Fields(29), 2, vbFalse, vbFalse, vbFalse) //--< BAU
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 31).Formula = FormatCurrency(rs.Fields(30), 2, vbFalse, vbFalse, vbFalse) //--< Aguinaldo_BAU
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 32).Formula = FormatCurrency(rs.Fields(31), 2, vbFalse, vbFalse, vbFalse) //--< Prest_Hipotecario
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 33).Formula = FormatCurrency(rs.Fields(32), 2, vbFalse, vbFalse, vbFalse) //--< Descuento_Seguro
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 34).Formula = rs.Fields(33)                                                //--< InstitucionSS
        //        //        //        oExcel.Workbooks(1).Sheets(1).Cells(X, 35).Formula = FormatCurrency(rs.Fields(34), 2, vbFalse, vbFalse, vbFalse) //--< Total
        //        //        //
        //        //        //        X = X + 1
        //        //        //        rs.MoveNext
        //        //        //    Loop

        //        //        oExcel.Workbooks(1).Sheets(1).Range("A2").CopyFromRecordrs

        //        //        if Not rs.EOF Then
        //        //            //Formato a las celdas de importes
        //        //            oExcel.Workbooks(1).Sheets(1).Activate
        //        //            oExcel.Workbooks(1).Sheets(1).Range(oExcel.Workbooks(1).Sheets(1).Cells(2, 9), oExcel.Workbooks(1).Sheets(1).Cells(rs.RecordCount + 1, 35)).Select
        //        //            oExcel.Workbooks(1).Sheets(1).Range(oExcel.Workbooks(1).Sheets(1).Cells(2, 9), oExcel.Workbooks(1).Sheets(1).Cells(rs.RecordCount + 1, 35)).NumberFormat = "$#,##0.00"
        //        //        End If

        //        //        if Len(Trim$(Dir$(sarchivo))) = 0 Then
        //        //            //Guarda el libro
        //        //            //FJMH 04 / 02 / 2021 Inicio
        //        //            //oExcel.Workbooks(1).SaveAs("C:\Dispersion\" & sFileName & iCorrida & ".xls")
        //        //            oExcel.Workbooks(1).SaveAs(sFileName)
        //        //            //FJMH 04 / 02 / 2021 Fin
        //        //        else : oExcel.Workbooks(1).Save()
        //        //        End If

        //        //        oExcel.Workbooks.Close()
        //        //        oExcel.Quit()
        //        //        oSheet = Nothing
        //        //        oExcel = Nothing

        //        //        rs.Close()
        //        //        rs = Nothing
        //        //        //FJMH 04 / 02 / 2021 Inicio
        //        //        //gfMsgbox.Mensaje "Se generó exitosamente en C:\Dispersion\" & sFileName & iCorrida & ".xls", vbInformation, "CAIRO"
        //        //        gfMsgbox.Mensaje "Se generó exitosamente en:" & sFileName, vbInformation, "CAIRO"
        //        //    //FJMH 04 / 02 / 2021 Fin
        //        //        SetStatusBar("Listo ...")
        //        //        Screen.MousePointer = vbDefault

        //        //        Exit Sub

        //        //msgerror:
        //        //        Screen.MousePointer = vbDefault
        //        //        gsErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        //        gfMsgbox.Mensaje " Error del Sistema Cairo ", vbOKDetails + vbCritical, "C A I R O", , gsErrVB
        //        //End Sub
        //        //    //Fin construccion del archivo PgoNomEmision.xlsx Erick Bejarano 26 / 07 / 2017------------------------------------------

        //        //    //Incia construccion del archivo DTS_EstadoSalud.xls Erick Bejarano 06 / 11 / 2017------------------------------------------
        //        //    public Sub DTS_EstadoSalud(sFileName As String, iHoja As Integer, rs As ADODB.Recordset)

        //        //        Dim oExcel As Excel.Application
        //        //        Dim oSheet1 As Excel.Worksheet
        //        //        Dim oSheet2 As Excel.Worksheet

        //        //        Dim sEtiqueta As String
        //        //        Dim Cadena As String
        //        //        Dim sarchivo As String

        //        //        On Error GoTo msgerror

        //        //        SetStatusBar("Generando reporte, espere un momento ...")
        //        //        Screen.MousePointer = vbHourglass

        //        //        oExcel = CreateObject("Excel.Application")
        //        //        DoEvents
        //        //        //FJMH Inicio 09 / 03 / 2021
        //        //        //sarchivo = sFileName & ".xls"
        //        //        sarchivo = sFileName
        //        //        //FJMH Fin
        //        //        if Len(Trim$(Dir$(sarchivo))) = 0 Then

        //        //            //Crea el archivo
        //        //            oExcel.Workbooks.Add()

        //        //            oSheet1 = oExcel.Worksheets("Hoja1")
        //        //            oSheet1.Name = "Estado_de_Salud"

        //        //            oSheet2 = oExcel.Worksheets("Hoja2")
        //        //            oSheet2.Name = "No_Acudio"
        //        //            //FJMH 09 / 03 / 2021 Inicio
        //        //            //else: oExcel.Workbooks.Open FileName:= "C:\Dispersion\" & sFileName & ".xls"
        //        //        else : oExcel.Workbooks.Open Filename:=sFileName
        //        //    //FJMH 09 / 03 / 2021 Fin
        //        //        End If

        //        //        //construye rotulos------------------------------------------------------------
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 1).Formula = "Empresa"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 2).Formula = "Instituto"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 3).Formula = "Poliza"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 4).Formula = "Grupo"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 5).Formula = "NSS"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 6).Formula = "Nombre_Beneficiario"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 7).Formula = "Pensionado"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 8).Formula = "Nivel_Encuesta"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 9).Formula = "Tipo_Pension"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 10).Formula = "Direccion"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 11).Formula = "Colonia"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 12).Formula = "CP"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 13).Formula = "Municipio"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 14).Formula = "Estado"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 15).Formula = "Telefono"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 16).Formula = "Gerencia"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 17).Formula = "Fecha_Encuesta"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 18).Formula = "Estado_Salud"
        //        //        oExcel.Workbooks(1).Sheets(iHoja).Cells(1, 19).Formula = "Estatus"

        //        //        oExcel.Workbooks(1).Sheets(iHoja).Range("A2").CopyFromRecordrs

        //        //        if Len(Trim$(Dir$(sarchivo))) = 0 Then
        //        //            //Guarda el libro
        //        //            //FJMH 09 / 03 / 2021 Inicio
        //        //            //oExcel.Workbooks(1).SaveAs("C:\Dispersion\" & sFileName & ".xls")
        //        //            oExcel.Workbooks(1).SaveAs(sFileName)
        //        //            //FJMH 09 / 03 / 2021 Fin
        //        //        else : oExcel.Workbooks(1).Save()
        //        //        End If

        //        //        oExcel.Workbooks.Close()
        //        //        oExcel.Quit()
        //        //        oSheet = Nothing
        //        //        oExcel = Nothing

        //        //        rs.Close()
        //        //        rs = Nothing

        //        //        if iHoja = 2 Then
        //        //            //FJMH 09 / 03 / 2021 Inicio
        //        //            //gfMsgbox.Mensaje "Se generó exitosamente en C:\Dispersion\" & sFileName & ".xls", vbInformation, "CAIRO"
        //        //            gfMsgbox.Mensaje "Se generó exitosamente en" & sFileName, vbInformation, "CAIRO"
        //        //        //FJMH 09 / 03 / 2021 Fin
        //        //            SetStatusBar("Listo ...")
        //        //            Screen.MousePointer = vbDefault
        //        //        End If

        //        //        Exit Sub

        //        //msgerror:
        //        //        Screen.MousePointer = vbDefault
        //        //        gsErrVB = Format$(Err.Number) & vbTab & Err.Source & vbTab & Err.Description
        //        //        gfMsgbox.Mensaje " Error del Sistema Cairo ", vbOKDetails + vbCritical, "C A I R O", , gsErrVB
        //        //End Sub
        //        //    //Fin construccion del archivo DTS_EstadoSalud.xls Erick Bejarano 26 / 07 / 2017------------------------------------------


        }
    }
