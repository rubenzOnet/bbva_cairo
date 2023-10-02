
using ADODB;

namespace MTSCPolizas.Modulos
{
    public static class modErrores
    {
        public static Recordset ErroresDLL(Connection cnConn, string err1 )
        {
            int ipos;
            int i;
            string sCadena;
            string sError = "";
            string sDescripcion = "";
            string sOrigen = "";
            Errors Errs1;
            Errors errLoop;
            Recordset rsdatos;

            //Se crea el recordset
            rsdatos = new Recordset();

            try
            {


            rsdatos.CursorLocation = CursorLocationEnum.adUseClient;

            //Agreamos los campos
            rsdatos.Fields.Append("Tipo", DataTypeEnum.adVarChar, 10, FieldAttributeEnum.adFldIsNullable);
            rsdatos.Fields.Append("Error", DataTypeEnum.adVarChar, 50, FieldAttributeEnum.adFldIsNullable);
            rsdatos.Fields.Append("Descripcion", DataTypeEnum.adVarChar, 255, FieldAttributeEnum.adFldIsNullable);
            rsdatos.Fields.Append("Fuente", DataTypeEnum.adVarChar, 255, FieldAttributeEnum.adFldIsNullable);
            //Creamos el recorset
            //rsdatos.Open(, , CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockBatchOptimistic)
            rsdatos.Open();
            //Agreamos los renglones

            //Se introducen primero los Errores de VB
            if (err1 != "")
            {
                sCadena = err1 + "\t";
                for (int y = 0; y < 3; y++)
                {
                    ipos = sCadena.IndexOf("\t", 1);

                    switch (y)
                    {
                        case 1:
                            sError = sCadena.Substring(1, ipos - 1);
                            break;
                        case 2:
                            sOrigen = sCadena.Substring(1, ipos - 1);
                            break;
                        case 3:
                            sDescripcion = sCadena.Substring(1, ipos - 1);
                            break;
                    }

                    sCadena = sCadena.Substring(ipos + 1);
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
                    string[] valor1 = new string[] { "Tipo", "Error", "Descripcion", "Fuente" };
                    string[] valor2 = new string[] { "ADO", sError, sDescripcion, sOrigen };


                    Errs1 = cnConn.Errors;

                    foreach (Error adoError in Errs1)
                    {
                        Console.WriteLine("ADO Error Number: " + adoError.Number);
                        Console.WriteLine("ADO Error Description: " + adoError.Description);
                        sError = ex.Message;
                        sDescripcion = adoError.Description;
                        sOrigen = rsdatos.Source.ToString();

                        rsdatos.AddNew(valor1, valor2);

                    }
                }
                
            }

            return rsdatos;
        }


    }
}
