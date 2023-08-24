namespace bbva_cairo.Formularios
{
    partial class frmCISSSTE
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCISSSTE));
            toolStrip1 = new ToolStrip();
            toolStripButton1 = new ToolStripButton();
            groupBox1 = new GroupBox();
            grdDescXPtmos = new DataGridView();
            fmeISSSTE = new GroupBox();
            label7 = new Label();
            label6 = new Label();
            label5 = new Label();
            label4 = new Label();
            label3 = new Label();
            label2 = new Label();
            label1 = new Label();
            lblImporte = new TextBox();
            lblDescApli = new TextBox();
            lblSaldoActual = new TextBox();
            lblSaldoIni = new TextBox();
            lblNoDesc = new TextBox();
            lblEstado = new TextBox();
            lblNoPtmo = new TextBox();
            grdPtmosActivos = new DataGridView();
            toolStrip1.SuspendLayout();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)grdDescXPtmos).BeginInit();
            fmeISSSTE.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)grdPtmosActivos).BeginInit();
            SuspendLayout();
            // 
            // toolStrip1
            // 
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripButton1 });
            toolStrip1.Location = new Point(0, 0);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(800, 25);
            toolStrip1.TabIndex = 0;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            toolStripButton1.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripButton1.Image = Properties.Resources.Salir;
            toolStripButton1.ImageTransparentColor = Color.Magenta;
            toolStripButton1.Name = "toolStripButton1";
            toolStripButton1.Size = new Size(23, 22);
            toolStripButton1.Text = "toolStripButton1";
            toolStripButton1.Click += toolStripButton1_Click;
            // 
            // groupBox1
            // 
            groupBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBox1.Controls.Add(grdDescXPtmos);
            groupBox1.Controls.Add(fmeISSSTE);
            groupBox1.Controls.Add(grdPtmosActivos);
            groupBox1.Location = new Point(0, 28);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(788, 626);
            groupBox1.TabIndex = 1;
            groupBox1.TabStop = false;
            // 
            // grdDescXPtmos
            // 
            grdDescXPtmos.AllowUserToAddRows = false;
            grdDescXPtmos.BorderStyle = BorderStyle.Fixed3D;
            grdDescXPtmos.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            grdDescXPtmos.Location = new Point(6, 348);
            grdDescXPtmos.Name = "grdDescXPtmos";
            grdDescXPtmos.ReadOnly = true;
            grdDescXPtmos.RowTemplate.Height = 25;
            grdDescXPtmos.Size = new Size(776, 272);
            grdDescXPtmos.TabIndex = 2;
            // 
            // fmeISSSTE
            // 
            fmeISSSTE.Controls.Add(label7);
            fmeISSSTE.Controls.Add(label6);
            fmeISSSTE.Controls.Add(label5);
            fmeISSSTE.Controls.Add(label4);
            fmeISSSTE.Controls.Add(label3);
            fmeISSSTE.Controls.Add(label2);
            fmeISSSTE.Controls.Add(label1);
            fmeISSSTE.Controls.Add(lblImporte);
            fmeISSSTE.Controls.Add(lblDescApli);
            fmeISSSTE.Controls.Add(lblSaldoActual);
            fmeISSSTE.Controls.Add(lblSaldoIni);
            fmeISSSTE.Controls.Add(lblNoDesc);
            fmeISSSTE.Controls.Add(lblEstado);
            fmeISSSTE.Controls.Add(lblNoPtmo);
            fmeISSSTE.Location = new Point(6, 200);
            fmeISSSTE.Name = "fmeISSSTE";
            fmeISSSTE.Size = new Size(776, 142);
            fmeISSSTE.TabIndex = 1;
            fmeISSSTE.TabStop = false;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(626, 76);
            label7.Name = "label7";
            label7.Size = new Size(49, 15);
            label7.TabIndex = 13;
            label7.Text = "Importe";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(626, 19);
            label6.Name = "label6";
            label6.Size = new Size(90, 15);
            label6.TabIndex = 12;
            label6.Text = "Desc. Aplicados";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(502, 19);
            label5.Name = "label5";
            label5.Size = new Size(73, 15);
            label5.TabIndex = 11;
            label5.Text = "Saldo Actual";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(378, 19);
            label4.Name = "label4";
            label4.Size = new Size(70, 15);
            label4.TabIndex = 10;
            label4.Text = "Saldo Inicial";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(254, 19);
            label3.Name = "label3";
            label3.Size = new Size(35, 15);
            label3.TabIndex = 9;
            label3.Text = "Plazo";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(130, 19);
            label2.Name = "label2";
            label2.Size = new Size(42, 15);
            label2.TabIndex = 8;
            label2.Text = "Estado";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(6, 19);
            label1.Name = "label1";
            label1.Size = new Size(90, 15);
            label1.TabIndex = 7;
            label1.Text = "Num. Préstamo";
            // 
            // lblImporte
            // 
            lblImporte.Location = new Point(626, 94);
            lblImporte.Name = "lblImporte";
            lblImporte.Size = new Size(118, 23);
            lblImporte.TabIndex = 6;
            // 
            // lblDescApli
            // 
            lblDescApli.Location = new Point(626, 37);
            lblDescApli.Name = "lblDescApli";
            lblDescApli.Size = new Size(118, 23);
            lblDescApli.TabIndex = 5;
            // 
            // lblSaldoActual
            // 
            lblSaldoActual.Location = new Point(502, 37);
            lblSaldoActual.Name = "lblSaldoActual";
            lblSaldoActual.Size = new Size(118, 23);
            lblSaldoActual.TabIndex = 4;
            // 
            // lblSaldoIni
            // 
            lblSaldoIni.Location = new Point(378, 37);
            lblSaldoIni.Name = "lblSaldoIni";
            lblSaldoIni.Size = new Size(118, 23);
            lblSaldoIni.TabIndex = 3;
            // 
            // lblNoDesc
            // 
            lblNoDesc.Location = new Point(254, 37);
            lblNoDesc.Name = "lblNoDesc";
            lblNoDesc.Size = new Size(118, 23);
            lblNoDesc.TabIndex = 2;
            // 
            // lblEstado
            // 
            lblEstado.Location = new Point(130, 37);
            lblEstado.Name = "lblEstado";
            lblEstado.Size = new Size(118, 23);
            lblEstado.TabIndex = 1;
            // 
            // lblNoPtmo
            // 
            lblNoPtmo.Location = new Point(6, 37);
            lblNoPtmo.Name = "lblNoPtmo";
            lblNoPtmo.Size = new Size(118, 23);
            lblNoPtmo.TabIndex = 0;
            // 
            // grdPtmosActivos
            // 
            grdPtmosActivos.AllowUserToAddRows = false;
            grdPtmosActivos.BorderStyle = BorderStyle.Fixed3D;
            grdPtmosActivos.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            grdPtmosActivos.Location = new Point(6, 12);
            grdPtmosActivos.Name = "grdPtmosActivos";
            grdPtmosActivos.ReadOnly = true;
            grdPtmosActivos.RowTemplate.Height = 25;
            grdPtmosActivos.Size = new Size(776, 182);
            grdPtmosActivos.TabIndex = 0;
            // 
            // frmCISSSTE
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 666);
            Controls.Add(groupBox1);
            Controls.Add(toolStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "frmCISSSTE";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "CAIRO - PRÉSTAMOS ISSSTE";
            Load += frmCISSSTE_Load;
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)grdDescXPtmos).EndInit();
            fmeISSSTE.ResumeLayout(false);
            fmeISSSTE.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)grdPtmosActivos).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButton1;
        private GroupBox groupBox1;
        private DataGridView grdPtmosActivos;
        private GroupBox fmeISSSTE;
        private TextBox lblNoPtmo;
        private Label label7;
        private Label label6;
        private Label label5;
        private Label label4;
        private Label label3;
        private Label label2;
        private Label label1;
        private TextBox lblImporte;
        private TextBox lblDescApli;
        private TextBox lblSaldoActual;
        private TextBox lblSaldoIni;
        private TextBox lblNoDesc;
        private TextBox lblEstado;
        private DataGridView grdDescXPtmos;
    }
}