namespace bbva_cairo.Formularios
{
    partial class frmRentas
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
            menuStrip1 = new MenuStrip();
            prestamoISSSTEToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem1 = new ToolStripMenuItem();
            toolStrip1 = new ToolStrip();
            toolStripButton1 = new ToolStripButton();
            statusStrip1 = new StatusStrip();
            statusLabel = new ToolStripStatusLabel();
            menuStrip1.SuspendLayout();
            toolStrip1.SuspendLayout();
            statusStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { prestamoISSSTEToolStripMenuItem, toolStripMenuItem1 });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(878, 24);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // prestamoISSSTEToolStripMenuItem
            // 
            prestamoISSSTEToolStripMenuItem.Name = "prestamoISSSTEToolStripMenuItem";
            prestamoISSSTEToolStripMenuItem.Size = new Size(105, 20);
            prestamoISSSTEToolStripMenuItem.Text = "Prestamo ISSSTE";
            prestamoISSSTEToolStripMenuItem.Click += prestamoISSSTEToolStripMenuItem_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(72, 20);
            toolStripMenuItem1.Text = "Ca&tálogos";
            // 
            // toolStrip1
            // 
            toolStrip1.Items.AddRange(new ToolStripItem[] { toolStripButton1 });
            toolStrip1.Location = new Point(0, 24);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Size = new Size(878, 39);
            toolStrip1.TabIndex = 2;
            toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButton1
            // 
            toolStripButton1.AutoSize = false;
            toolStripButton1.DisplayStyle = ToolStripItemDisplayStyle.Image;
            toolStripButton1.Image = Properties.Resources.icons8_insurance_64;
            toolStripButton1.ImageTransparentColor = Color.Magenta;
            toolStripButton1.Name = "toolStripButton1";
            toolStripButton1.Size = new Size(36, 36);
            toolStripButton1.Text = "toolStripButton1";
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new ToolStripItem[] { statusLabel });
            statusStrip1.Location = new Point(0, 487);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(878, 22);
            statusStrip1.TabIndex = 3;
            statusStrip1.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            statusLabel.AutoSize = false;
            statusLabel.ImageScaling = ToolStripItemImageScaling.None;
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(400, 17);
            statusLabel.Text = "Listo...";
            statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // frmRentas
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(878, 509);
            Controls.Add(statusStrip1);
            Controls.Add(toolStrip1);
            Controls.Add(menuStrip1);
            IsMdiContainer = true;
            MainMenuStrip = menuStrip1;
            Name = "frmRentas";
            Text = "frmRentas";
            WindowState = FormWindowState.Maximized;
            Load += frmRentas_Load;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem prestamoISSSTEToolStripMenuItem;
        private ToolStripMenuItem toolStripMenuItem1;
        private ToolStrip toolStrip1;
        private ToolStripButton toolStripButton1;
        private StatusStrip statusStrip1;
        protected ToolStripStatusLabel statusLabel;
    }
}