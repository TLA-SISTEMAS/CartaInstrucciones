namespace CartaInstrucciones
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.performanceCounter1 = new System.Diagnostics.PerformanceCounter();
            this.msPrincipal = new System.Windows.Forms.MenuStrip();
            this.tsmCatalogo = new System.Windows.Forms.ToolStripMenuItem();
            this.importadoresToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.proveedoresToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmCartadeInstrucciones = new System.Windows.Forms.ToolStripMenuItem();
            this.catálogoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.performanceCounter1)).BeginInit();
            this.msPrincipal.SuspendLayout();
            this.SuspendLayout();
            // 
            // msPrincipal
            // 
            this.msPrincipal.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.msPrincipal.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmCatalogo,
            this.tsmCartadeInstrucciones});
            this.msPrincipal.Location = new System.Drawing.Point(0, 0);
            this.msPrincipal.Name = "msPrincipal";
            this.msPrincipal.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.msPrincipal.Size = new System.Drawing.Size(1311, 25);
            this.msPrincipal.TabIndex = 273;
            this.msPrincipal.Text = "Menu";
            // 
            // tsmCatalogo
            // 
            this.tsmCatalogo.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importadoresToolStripMenuItem,
            this.proveedoresToolStripMenuItem});
            this.tsmCatalogo.Name = "tsmCatalogo";
            this.tsmCatalogo.Size = new System.Drawing.Size(79, 21);
            this.tsmCatalogo.Text = "Catálogos";
            this.tsmCatalogo.Click += new System.EventHandler(this.tsmCatalogo_Click);
            // 
            // importadoresToolStripMenuItem
            // 
            this.importadoresToolStripMenuItem.Name = "importadoresToolStripMenuItem";
            this.importadoresToolStripMenuItem.Size = new System.Drawing.Size(156, 22);
            this.importadoresToolStripMenuItem.Text = "Importadores";
            this.importadoresToolStripMenuItem.Click += new System.EventHandler(this.importadoresToolStripMenuItem_Click);
            // 
            // proveedoresToolStripMenuItem
            // 
            this.proveedoresToolStripMenuItem.Name = "proveedoresToolStripMenuItem";
            this.proveedoresToolStripMenuItem.Size = new System.Drawing.Size(156, 22);
            this.proveedoresToolStripMenuItem.Text = "Proveedores";
            // 
            // tsmCartadeInstrucciones
            // 
            this.tsmCartadeInstrucciones.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.catálogoToolStripMenuItem});
            this.tsmCartadeInstrucciones.Name = "tsmCartadeInstrucciones";
            this.tsmCartadeInstrucciones.Size = new System.Drawing.Size(149, 21);
            this.tsmCartadeInstrucciones.Text = "Carta de Instrucciones";
            this.tsmCartadeInstrucciones.Click += new System.EventHandler(this.tsmCartadeInstrucciones_Click);
            // 
            // catálogoToolStripMenuItem
            // 
            this.catálogoToolStripMenuItem.Name = "catálogoToolStripMenuItem";
            this.catálogoToolStripMenuItem.Size = new System.Drawing.Size(129, 22);
            this.catálogoToolStripMenuItem.Text = "Catálogo";
            this.catálogoToolStripMenuItem.Click += new System.EventHandler(this.catálogoToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(1311, 613);
            this.Controls.Add(this.msPrincipal);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.DoubleBuffered = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.msPrincipal;
            this.Name = "Form1";
            this.Text = "Carta de Instrucciones Tecno Logistica Aduanal";
            this.Load += new System.EventHandler(this.Form1_Load);
            //((System.ComponentModel.ISupportInitialize)(this.performanceCounter1)).EndInit();
            this.msPrincipal.ResumeLayout(false);
            this.msPrincipal.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Diagnostics.PerformanceCounter performanceCounter1;
        private System.Windows.Forms.MenuStrip msPrincipal;
        private System.Windows.Forms.ToolStripMenuItem tsmCatalogo;
        private System.Windows.Forms.ToolStripMenuItem importadoresToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem proveedoresToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tsmCartadeInstrucciones;
        private System.Windows.Forms.ToolStripMenuItem catálogoToolStripMenuItem;
    }
}

