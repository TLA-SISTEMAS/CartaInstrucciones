namespace CartaInstrucciones
{
    partial class CartadeInstruccionesMenu
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CartadeInstruccionesMenu));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnAltasCarta = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.NumCarta = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Importador = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnBajas = new System.Windows.Forms.Button();
            this.btnCambios = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.flowLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnAltasCarta
            // 
            this.btnAltasCarta.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnAltasCarta.BackgroundImage")));
            this.btnAltasCarta.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnAltasCarta.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAltasCarta.Location = new System.Drawing.Point(3, 3);
            this.btnAltasCarta.Name = "btnAltasCarta";
            this.btnAltasCarta.Size = new System.Drawing.Size(89, 65);
            this.btnAltasCarta.TabIndex = 0;
            this.btnAltasCarta.TabStop = false;
            this.btnAltasCarta.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnAltasCarta.UseVisualStyleBackColor = true;
            this.btnAltasCarta.Click += new System.EventHandler(this.btnAltasCarta_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.AppWorkspace;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.NumCarta,
            this.Importador});
            this.dataGridView1.Location = new System.Drawing.Point(3, 7);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(888, 306);
            this.dataGridView1.TabIndex = 1;
            // 
            // NumCarta
            // 
            this.NumCarta.HeaderText = "Num. Carta";
            this.NumCarta.Name = "NumCarta";
            this.NumCarta.ReadOnly = true;
            // 
            // Importador
            // 
            this.Importador.HeaderText = "Importador";
            this.Importador.Name = "Importador";
            this.Importador.ReadOnly = true;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.btnAltasCarta);
            this.flowLayoutPanel1.Controls.Add(this.btnBajas);
            this.flowLayoutPanel1.Controls.Add(this.btnCambios);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(289, 71);
            this.flowLayoutPanel1.TabIndex = 2;
            // 
            // btnBajas
            // 
            this.btnBajas.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnBajas.BackgroundImage")));
            this.btnBajas.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnBajas.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBajas.Location = new System.Drawing.Point(98, 3);
            this.btnBajas.Name = "btnBajas";
            this.btnBajas.Size = new System.Drawing.Size(89, 65);
            this.btnBajas.TabIndex = 1;
            this.btnBajas.TabStop = false;
            this.btnBajas.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnBajas.UseVisualStyleBackColor = true;
            // 
            // btnCambios
            // 
            this.btnCambios.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnCambios.BackgroundImage")));
            this.btnCambios.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnCambios.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCambios.Location = new System.Drawing.Point(193, 3);
            this.btnCambios.Name = "btnCambios";
            this.btnCambios.Size = new System.Drawing.Size(89, 65);
            this.btnCambios.TabIndex = 2;
            this.btnCambios.TabStop = false;
            this.btnCambios.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnCambios.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.AutoSize = true;
            this.panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Location = new System.Drawing.Point(3, 74);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(894, 316);
            this.panel1.TabIndex = 3;
            // 
            // CartadeInstruccionesMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1026, 504);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CartadeInstruccionesMenu";
            this.Text = "Carta de Instrucciones";
            this.Load += new System.EventHandler(this.CartadeInstruccionesMenu_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAltasCarta;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button btnBajas;
        private System.Windows.Forms.Button btnCambios;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridViewTextBoxColumn NumCarta;
        private System.Windows.Forms.DataGridViewTextBoxColumn Importador;
    }
}