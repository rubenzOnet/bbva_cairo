using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ADODB;
using Microsoft.VisualBasic.FileIO;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MTSEndososCET.Modulos
{
    public static  class modErrores
    {
        

        //public static ErroresDLL(Connection cnConn As ADODB.Connection, err1 As String) As ADODB.Recordset
        //Dim ipos As Long
        //Dim i As Integer
        //Dim sCadena As String
        //Dim sError As String
        //Dim sDescripcion As String
        //Dim sOrigen As String
        //Dim Errs1 As ADODB.Errors
        //Dim errLoop As ADODB.Error
        //Dim rsdatos As ADODB.Recordset

        //'Se crea el recordset
        //Set rsdatos = New ADODB.Recordset

        //With rsdatos
        //    .CursorLocation = adUseClient
    
        //     'Agreamos los campos
        //     .Fields.Append "Tipo", adVarChar, 10, adFldIsNullable
        //     .Fields.Append "Error", adVarChar, 50, adFldIsNullable
        //     .Fields.Append "Descripcion", adVarChar, 255, adFldIsNullable
        //     .Fields.Append "Fuente", adVarChar, 255, adFldIsNullable
        //     'Creamos el recorset
        //      .Open , , adOpenStatic, adLockBatchOptimistic
        //      'Agreamos los renglones
      
        //      'Se introducen primero los Errores de VB
        //      If err1 <> "" Then
        //             sCadena = err1 & vbTab
        //            For i = 1 To 3
        //                ipos = InStr(1, sCadena, vbTab)
        //                Select Case i
        //                    Case 1
        //                         sError = Mid$(sCadena, 1, ipos - 1)
        //                     Case 2
        //                         sOrigen = Mid$(sCadena, 1, ipos - 1)
        //                    Case 3
        //                         sDescripcion = Mid$(sCadena, 1, ipos - 1)
        //                End Select
        //                sCadena = Mid$(sCadena, ipos + 1)
        //            Next i

        //           .AddNew
        //           !Tipo = "VB"
        //           !Error = sError
        //           !Descripcion = sDescripcion
        //           !Fuente = sOrigen
        //           .Update


        //      End If
      
        //      'Se colocan los errores de ADO
        //      If cnConn Is Nothing Then
        //      Else
        //           Set Errs1 = cnConn.Errors
        //           For Each errLoop In Errs1
        //                 With errLoop
        //                     sError = Format$(.NativeError)
        //                     sDescripcion = .Description
        //                     sOrigen = .Source
        //                 End With
        //                 .AddNew Array("Tipo", "Error", "Descripcion", "Fuente"), Array("ADO", sError, sDescripcion, sOrigen)
        //           Next
        //      End If

        //End With

        //Set ErroresDLL = rsdatos
        //End Function


    }
}
