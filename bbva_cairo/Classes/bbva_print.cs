using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace bbva_cairo.Classes
{
    public class bbva_print
    {

        string rutaArchivo = "";

        public void imprimir(string file)
        {

            rutaArchivo = file;

            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(ManejadorPaginaImpresion);

            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = pd;

            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                pd.Print();
            }
        }

        private void ManejadorPaginaImpresion(object sender, PrintPageEventArgs e)
        {
            using (var streamReader = new System.IO.StreamReader(rutaArchivo))
            {
                Font fuente = new Font("Arial", 12);
                float yPos = 0;
                int contador = 0;
                string linea = null;

                while (contador < e.MarginBounds.Height / fuente.GetHeight(e.Graphics) && ((linea = streamReader.ReadLine()) != null))
                {
                    yPos = e.MarginBounds.Top + contador * fuente.GetHeight(e.Graphics);
                    e.Graphics.DrawString(linea, fuente, Brushes.Black, e.MarginBounds.Left, yPos, new StringFormat());
                    contador++;
                }
            }

            e.HasMorePages = false;
        }

    }
}
