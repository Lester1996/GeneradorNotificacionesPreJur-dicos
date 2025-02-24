namespace GeneradorNotificacionesPreJurídicos
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            panel1 = new Panel();
            pictureBox1 = new PictureBox();
            label1 = new Label();
            groupBox1 = new GroupBox();
            label4 = new Label();
            dtpFecha = new DateTimePicker();
            rbtnGrupoClaves = new RadioButton();
            grpBoxArchivo = new GroupBox();
            btnProcesarArchivo = new Button();
            btnSeleccionarArchivo = new Button();
            label3 = new Label();
            txtArchivo = new TextBox();
            panel2 = new Panel();
            panel3 = new Panel();
            panel4 = new Panel();
            panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            groupBox1.SuspendLayout();
            grpBoxArchivo.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.BackColor = SystemColors.ActiveCaptionText;
            panel1.Controls.Add(pictureBox1);
            panel1.Controls.Add(label1);
            panel1.Location = new Point(-1, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(632, 65);
            panel1.TabIndex = 0;
            // 
            // pictureBox1
            // 
            pictureBox1.Image = Properties.Resources.EEH;
            pictureBox1.Location = new Point(401, 5);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(125, 58);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 1;
            pictureBox1.TabStop = false;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.BackColor = SystemColors.ActiveCaptionText;
            label1.Font = new Font("Verdana", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            label1.ForeColor = Color.White;
            label1.Location = new Point(3, 25);
            label1.Name = "label1";
            label1.Size = new Size(392, 20);
            label1.TabIndex = 1;
            label1.Text = "Generador de notificaciones pre jurídico";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(label4);
            groupBox1.Controls.Add(dtpFecha);
            groupBox1.Controls.Add(rbtnGrupoClaves);
            groupBox1.Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox1.Location = new Point(12, 89);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(513, 100);
            groupBox1.TabIndex = 1;
            groupBox1.TabStop = false;
            groupBox1.Text = "Generar para:";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold, GraphicsUnit.Point);
            label4.Location = new Point(231, 43);
            label4.Name = "label4";
            label4.Size = new Size(52, 17);
            label4.TabIndex = 3;
            label4.Text = "Fecha";
            // 
            // dtpFecha
            // 
            dtpFecha.Location = new Point(289, 41);
            dtpFecha.Name = "dtpFecha";
            dtpFecha.Size = new Size(170, 23);
            dtpFecha.TabIndex = 2;
            // 
            // rbtnGrupoClaves
            // 
            rbtnGrupoClaves.AutoSize = true;
            rbtnGrupoClaves.Location = new Point(5, 42);
            rbtnGrupoClaves.Name = "rbtnGrupoClaves";
            rbtnGrupoClaves.Size = new Size(133, 21);
            rbtnGrupoClaves.TabIndex = 1;
            rbtnGrupoClaves.TabStop = true;
            rbtnGrupoClaves.Text = "Grupo de claves";
            rbtnGrupoClaves.UseVisualStyleBackColor = true;
            rbtnGrupoClaves.CheckedChanged += rbtnGrupoClaves_CheckedChanged;
            // 
            // grpBoxArchivo
            // 
            grpBoxArchivo.Controls.Add(btnProcesarArchivo);
            grpBoxArchivo.Controls.Add(btnSeleccionarArchivo);
            grpBoxArchivo.Controls.Add(label3);
            grpBoxArchivo.Controls.Add(txtArchivo);
            grpBoxArchivo.Location = new Point(12, 193);
            grpBoxArchivo.Name = "grpBoxArchivo";
            grpBoxArchivo.Size = new Size(513, 125);
            grpBoxArchivo.TabIndex = 6;
            grpBoxArchivo.TabStop = false;
            // 
            // btnProcesarArchivo
            // 
            btnProcesarArchivo.Enabled = false;
            btnProcesarArchivo.Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Regular, GraphicsUnit.Point);
            btnProcesarArchivo.Location = new Point(410, 73);
            btnProcesarArchivo.Name = "btnProcesarArchivo";
            btnProcesarArchivo.Size = new Size(94, 29);
            btnProcesarArchivo.TabIndex = 3;
            btnProcesarArchivo.Text = "Procesar";
            btnProcesarArchivo.UseVisualStyleBackColor = true;
            btnProcesarArchivo.Click += btnProcesarArchivo_Click;
            // 
            // btnSeleccionarArchivo
            // 
            btnSeleccionarArchivo.Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Regular, GraphicsUnit.Point);
            btnSeleccionarArchivo.Location = new Point(410, 26);
            btnSeleccionarArchivo.Name = "btnSeleccionarArchivo";
            btnSeleccionarArchivo.Size = new Size(94, 29);
            btnSeleccionarArchivo.TabIndex = 2;
            btnSeleccionarArchivo.Text = "Seleccionar";
            btnSeleccionarArchivo.UseVisualStyleBackColor = true;
            btnSeleccionarArchivo.Click += btnSeleccionarArchivo_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold, GraphicsUnit.Point);
            label3.Location = new Point(9, 55);
            label3.Name = "label3";
            label3.Size = new Size(62, 17);
            label3.TabIndex = 1;
            label3.Text = "Archivo";
            // 
            // txtArchivo
            // 
            txtArchivo.Enabled = false;
            txtArchivo.Location = new Point(77, 49);
            txtArchivo.Name = "txtArchivo";
            txtArchivo.Size = new Size(320, 27);
            txtArchivo.TabIndex = 0;
            txtArchivo.TextChanged += txtArchivo_TextChanged;
            // 
            // panel2
            // 
            panel2.BackColor = Color.FromArgb(44, 200, 220);
            panel2.Location = new Point(-3, 65);
            panel2.Name = "panel2";
            panel2.Size = new Size(165, 9);
            panel2.TabIndex = 3;
            // 
            // panel3
            // 
            panel3.BackColor = Color.FromArgb(249, 119, 120);
            panel3.Location = new Point(357, 65);
            panel3.Name = "panel3";
            panel3.Size = new Size(199, 9);
            panel3.TabIndex = 4;
            // 
            // panel4
            // 
            panel4.BackColor = Color.FromArgb(250, 193, 82);
            panel4.Location = new Point(161, 65);
            panel4.Name = "panel4";
            panel4.Size = new Size(199, 9);
            panel4.TabIndex = 5;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(535, 330);
            Controls.Add(grpBoxArchivo);
            Controls.Add(panel3);
            Controls.Add(panel4);
            Controls.Add(panel2);
            Controls.Add(groupBox1);
            Controls.Add(panel1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MaximumSize = new Size(553, 377);
            MinimizeBox = false;
            MinimumSize = new Size(553, 377);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Inicio";
            Load += Form1_Load;
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            grpBoxArchivo.ResumeLayout(false);
            grpBoxArchivo.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private Label label1;
        private PictureBox pictureBox1;
        private GroupBox groupBox1;
        private RadioButton rbtnGrupoClaves;
        private Panel panel2;
        private Panel panel3;
        private Panel panel4;
        private GroupBox grpBoxArchivo;
        private Label label3;
        private TextBox txtArchivo;
        private Button btnProcesarArchivo;
        private Button btnSeleccionarArchivo;
        private DateTimePicker dtpFecha;
        private Label label4;
    }
}
