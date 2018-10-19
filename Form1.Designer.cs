namespace GeneradorConstancias
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
            this.btnGenerar = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtRuta = new System.Windows.Forms.TextBox();
            this.btnAbrirExcel = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btn_abrirWord = new System.Windows.Forms.Button();
            this.txtRutaWord = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgreso = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnGenerar
            // 
            this.btnGenerar.Enabled = false;
            this.btnGenerar.Location = new System.Drawing.Point(658, 21);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Size = new System.Drawing.Size(200, 57);
            this.btnGenerar.TabIndex = 0;
            this.btnGenerar.Text = "Generar";
            this.btnGenerar.UseVisualStyleBackColor = true;
            this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(125, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(188, 31);
            this.label1.TabIndex = 1;
            this.label1.Text = "Archivo excel";
            // 
            // txtRuta
            // 
            this.txtRuta.Enabled = false;
            this.txtRuta.Location = new System.Drawing.Point(319, 21);
            this.txtRuta.Name = "txtRuta";
            this.txtRuta.Size = new System.Drawing.Size(230, 20);
            this.txtRuta.TabIndex = 2;
            // 
            // btnAbrirExcel
            // 
            this.btnAbrirExcel.Location = new System.Drawing.Point(567, 21);
            this.btnAbrirExcel.Name = "btnAbrirExcel";
            this.btnAbrirExcel.Size = new System.Drawing.Size(62, 23);
            this.btnAbrirExcel.TabIndex = 3;
            this.btnAbrirExcel.Text = "...";
            this.btnAbrirExcel.UseVisualStyleBackColor = true;
            this.btnAbrirExcel.Click += new System.EventHandler(this.btnAbrirExcel_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btn_abrirWord
            // 
            this.btn_abrirWord.Location = new System.Drawing.Point(567, 70);
            this.btn_abrirWord.Name = "btn_abrirWord";
            this.btn_abrirWord.Size = new System.Drawing.Size(62, 23);
            this.btn_abrirWord.TabIndex = 6;
            this.btn_abrirWord.Text = "...";
            this.btn_abrirWord.UseVisualStyleBackColor = true;
            this.btn_abrirWord.Click += new System.EventHandler(this.btn_abrirWord_Click);
            // 
            // txtRutaWord
            // 
            this.txtRutaWord.Enabled = false;
            this.txtRutaWord.Location = new System.Drawing.Point(319, 70);
            this.txtRutaWord.Name = "txtRutaWord";
            this.txtRutaWord.Size = new System.Drawing.Size(230, 20);
            this.txtRutaWord.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(125, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(183, 31);
            this.label2.TabIndex = 4;
            this.label2.Text = "Archivo word";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(1, 1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(100, 100);
            this.pictureBox1.TabIndex = 7;
            this.pictureBox1.TabStop = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(658, 109);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(200, 23);
            this.progressBar1.TabIndex = 8;
            // 
            // lblProgreso
            // 
            this.lblProgreso.AutoSize = true;
            this.lblProgreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProgreso.Location = new System.Drawing.Point(655, 91);
            this.lblProgreso.Name = "lblProgreso";
            this.lblProgreso.Size = new System.Drawing.Size(57, 15);
            this.lblProgreso.TabIndex = 9;
            this.lblProgreso.Text = "Progreso";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(12, 144);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(435, 15);
            this.label3.TabIndex = 10;
            this.label3.Text = "Desarrollado por: Sánchez Mundo Martín Ignacio y Acosta Rocha Jesús Carlos";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(875, 168);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblProgreso);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btn_abrirWord);
            this.Controls.Add(this.txtRutaWord);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnAbrirExcel);
            this.Controls.Add(this.txtRuta);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnGenerar);
            this.Name = "Form1";
            this.Text = "Generador de constancias";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenerar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtRuta;
        private System.Windows.Forms.Button btnAbrirExcel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btn_abrirWord;
        private System.Windows.Forms.TextBox txtRutaWord;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProgreso;
        private System.Windows.Forms.Label label3;
    }
}

