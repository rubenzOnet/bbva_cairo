﻿using System.Security.Cryptography;
using ADODB;

namespace MTSEndososCET.Modulos
{
    public static  class modErrores
    {


        public static Recordset ErroresDLL(Connection cnConn, string err1)
        {
            int ipos;
            int i;
            string sCadena;
            string sError = "";
            string sDescripcion = "";
            string sOrigen = "";
            ADODB.Errors Errs1;
            ADODB.Error errLoop;
            Recordset rsdatos = null; 

            try
            {


            //Se crea el recordset
            rsdatos = new Recordset();

            // With rsdatos
            rsdatos.CursorLocation = CursorLocationEnum.adUseClient;

            //Agreamos los campos
            rsdatos.Fields.Append("Tipo", DataTypeEnum.adVarChar, 10, FieldAttributeEnum.adFldIsNullable);
            rsdatos.Fields.Append("Error", DataTypeEnum.adVarChar, 50, FieldAttributeEnum.adFldIsNullable);
            rsdatos.Fields.Append("Descripcion", DataTypeEnum.adVarChar, 255, FieldAttributeEnum.adFldIsNullable);
            rsdatos.Fields.Append("Fuente", DataTypeEnum.adVarChar, 255, FieldAttributeEnum.adFldIsNullable);
            //Creamos el recorset
            rsdatos.Open();
            //Agreamos los renglones

            //Se introducen primero los Errores de VB
            if (err1 != "")
            {
                sCadena = err1 + "\t";
                for (int y = 0; y < 3; y++)
                {
                    ipos = sCadena.IndexOf("\t", 1); // InStr(1, sCadena, vbTab)
                    switch (y)
                    {
                        case 1:
                            sError = sCadena.Substring(1, ipos - 1); //Mid$(sCadena, 1, ipos - 1)
                            break;
                        case 2:
                            sOrigen = sCadena.Substring(1, ipos - 1); // Mid$(sCadena, 1, ipos - 1)
                            break;
                        case 3:
                            sDescripcion = sCadena.Substring(1, ipos - 1); // Mid$(sCadena, 1, ipos - 1)
                            break;
                    }

                    sCadena = sCadena.Substring(ipos + 1); // Mid$(sCadena, ipos + 1)
                }

                rsdatos.AddNew();

                rsdatos.Fields["Tipo"].Value = "VB";
                rsdatos.Fields["Error"].Value = sError;
                rsdatos.Fields["Descripcion"].Value = sDescripcion;
                rsdatos.Fields["Fuente"].Value = sOrigen;
                rsdatos.Update();

            }
            }
            catch (Exception ex)
            {
                //Se colocan los errores de ADO
                if (cnConn != null)
                {
                    Errors adoErrors = cnConn.Errors;
                    Errs1 = cnConn.Errors;
                    foreach (Error adoErrLoop in adoErrors)
                    {
                        sError = adoErrLoop.NativeError.ToString();
                        sDescripcion = adoErrLoop.Description;
                        sOrigen = adoErrLoop.Source;

                        // adoErrLoop.AddNew(Array("Tipo", "Error", "Descripcion", "Fuente"), Array("ADO", sError, sDescripcion, sOrigen));

                    }
                }
            }

            return rsdatos;
        }


    }
}