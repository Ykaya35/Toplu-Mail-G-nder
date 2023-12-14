namespace TopluMailSed
{
    partial class Form1
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.txtDetay = new System.Windows.Forms.RichTextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.GondericiMail = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.GondericiPass = new System.Windows.Forms.TextBox();
            this.GondericiAdSoyad = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.SmtpHost = new System.Windows.Forms.TextBox();
            this.SmtpPt = new System.Windows.Forms.TextBox();
            this.Baslik = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.excelSendMailAdd = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(616, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(273, 184);
            this.dataGridView1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Chocolate;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.button1.ForeColor = System.Drawing.SystemColors.Control;
            this.button1.Location = new System.Drawing.Point(640, 202);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(249, 59);
            this.button1.TabIndex = 1;
            this.button1.Text = "Excel Veri Ekle";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.btnOpenExcel_Click);
            // 
            // txtDetay
            // 
            this.txtDetay.Location = new System.Drawing.Point(12, 72);
            this.txtDetay.Name = "txtDetay";
            this.txtDetay.Size = new System.Drawing.Size(598, 222);
            this.txtDetay.TabIndex = 2;
            this.txtDetay.Text = "";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.groupBox1.Location = new System.Drawing.Point(12, 316);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(877, 232);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Mail Gönderim Bilgileri ";
            // 
            // GondericiMail
            // 
            this.GondericiMail.Location = new System.Drawing.Point(6, 126);
            this.GondericiMail.Name = "GondericiMail";
            this.GondericiMail.Size = new System.Drawing.Size(214, 22);
            this.GondericiMail.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 189);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 16);
            this.label5.TabIndex = 8;
            this.label5.Text = "Mail Şifre";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(26, 153);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(36, 16);
            this.label4.TabIndex = 8;
            this.label4.Text = "Mail";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 105);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 16);
            this.label3.TabIndex = 8;
            this.label3.Text = "Ssl";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(74, 16);
            this.label2.TabIndex = 7;
            this.label2.Text = "SmtpHost";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "SmtpPort :";
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Green;
            this.button2.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.button2.Location = new System.Drawing.Point(735, 130);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(120, 86);
            this.button2.TabIndex = 5;
            this.button2.Text = "Gönder";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // GondericiPass
            // 
            this.GondericiPass.Location = new System.Drawing.Point(6, 162);
            this.GondericiPass.Name = "GondericiPass";
            this.GondericiPass.Size = new System.Drawing.Size(214, 22);
            this.GondericiPass.TabIndex = 4;
            // 
            // GondericiAdSoyad
            // 
            this.GondericiAdSoyad.Location = new System.Drawing.Point(411, 24);
            this.GondericiAdSoyad.Name = "GondericiAdSoyad";
            this.GondericiAdSoyad.Size = new System.Drawing.Size(189, 22);
            this.GondericiAdSoyad.TabIndex = 3;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Evet",
            "Hayır"});
            this.comboBox1.Location = new System.Drawing.Point(6, 84);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(214, 24);
            this.comboBox1.TabIndex = 2;
            // 
            // SmtpHost
            // 
            this.SmtpHost.Location = new System.Drawing.Point(6, 49);
            this.SmtpHost.Name = "SmtpHost";
            this.SmtpHost.Size = new System.Drawing.Size(214, 22);
            this.SmtpHost.TabIndex = 1;
            // 
            // SmtpPt
            // 
            this.SmtpPt.Location = new System.Drawing.Point(6, 14);
            this.SmtpPt.Name = "SmtpPt";
            this.SmtpPt.Size = new System.Drawing.Size(214, 22);
            this.SmtpPt.TabIndex = 0;
            // 
            // Baslik
            // 
            this.Baslik.Location = new System.Drawing.Point(103, 10);
            this.Baslik.Multiline = true;
            this.Baslik.Name = "Baslik";
            this.Baslik.Size = new System.Drawing.Size(507, 29);
            this.Baslik.TabIndex = 4;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.excelSendMailAdd);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.SmtpPt);
            this.groupBox2.Controls.Add(this.GondericiMail);
            this.groupBox2.Controls.Add(this.SmtpHost);
            this.groupBox2.Controls.Add(this.comboBox1);
            this.groupBox2.Controls.Add(this.GondericiPass);
            this.groupBox2.Controls.Add(this.GondericiAdSoyad);
            this.groupBox2.Location = new System.Drawing.Point(107, 21);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(606, 205);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(283, 27);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(122, 16);
            this.label6.TabIndex = 10;
            this.label6.Text = "Gönderenin Adı :";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(28, 20);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 13);
            this.label7.TabIndex = 7;
            this.label7.Text = "Mail Başlığı :";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(15, 56);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(85, 13);
            this.label8.TabIndex = 8;
            this.label8.Text = "Mail Açıklaması :";
            // 
            // excelSendMailAdd
            // 
            this.excelSendMailAdd.Location = new System.Drawing.Point(322, 64);
            this.excelSendMailAdd.Name = "excelSendMailAdd";
            this.excelSendMailAdd.Size = new System.Drawing.Size(278, 44);
            this.excelSendMailAdd.TabIndex = 11;
            this.excelSendMailAdd.Text = "Excelden Gönderim Mail\'i Ekle";
            this.excelSendMailAdd.UseVisualStyleBackColor = true;
            this.excelSendMailAdd.Click += new System.EventHandler(this.excelSendMailAdd_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 560);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.Baslik);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.txtDetay);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.RichTextBox txtDetay;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TextBox SmtpHost;
        private System.Windows.Forms.TextBox Baslik;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox GondericiPass;
        private System.Windows.Forms.TextBox GondericiAdSoyad;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox GondericiMail;
        private System.Windows.Forms.TextBox SmtpPt;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button excelSendMailAdd;
    }
}

