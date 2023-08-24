﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace bbva_cairo.Formularios
{
    public partial class frmRentas : Form
    {
        public static frmCISSSTE CurrentForm1Instance;


        public frmRentas()
        {
            InitializeComponent();
        }

        private void prestamoISSSTEToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (CurrentForm1Instance == null || CurrentForm1Instance.IsDisposed)
            {
                CurrentForm1Instance = new frmCISSSTE();
                CurrentForm1Instance.MdiParent = this;
                CurrentForm1Instance.Show();
            }
            else
            {
                CurrentForm1Instance.Focus();
            }

        }

        private void frmRentas_Load(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (CurrentForm1Instance == null || CurrentForm1Instance.IsDisposed)
            {
                CurrentForm1Instance = new frmCISSSTE();
                CurrentForm1Instance.MdiParent = this;
                CurrentForm1Instance.Show();
            }
            else
            {
                CurrentForm1Instance.Focus();
            }
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("¿Estás seguro que deseas salir?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Application.Exit();
            }


        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("¿Estás seguro que deseas salir?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                Application.Exit();
            }

        }
    }
}



