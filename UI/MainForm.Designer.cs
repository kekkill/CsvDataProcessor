namespace WinFormsApp1
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.btnImport = new System.Windows.Forms.Button();
            this.btnExportExcel = new System.Windows.Forms.Button();
            this.btnExportXml = new System.Windows.Forms.Button();
            this.btnClearData = new System.Windows.Forms.Button();
            this.lblTotalRecords = new System.Windows.Forms.Label();
            this.dtpStartDate = new System.Windows.Forms.DateTimePicker();
            this.dtpEndDate = new System.Windows.Forms.DateTimePicker();
            this.chkStartDate = new System.Windows.Forms.CheckBox();
            this.chkEndDate = new System.Windows.Forms.CheckBox();
            this.txtFirstName = new System.Windows.Forms.TextBox();
            this.txtLastName = new System.Windows.Forms.TextBox();
            this.txtSurName = new System.Windows.Forms.TextBox();
            this.txtCity = new System.Windows.Forms.TextBox();
            this.txtCountry = new System.Windows.Forms.TextBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label(); 
            this.SuspendLayout();
  
            this.btnImport.Location = new System.Drawing.Point(12, 12);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(120, 30);
            this.btnImport.TabIndex = 0;
            this.btnImport.Text = "Импорт CSV";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
 
            this.btnExportExcel.Location = new System.Drawing.Point(138, 12);
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.Size = new System.Drawing.Size(120, 30);
            this.btnExportExcel.TabIndex = 1;
            this.btnExportExcel.Text = "Экспорт в Excel";
            this.btnExportExcel.UseVisualStyleBackColor = true;
            this.btnExportExcel.Click += new System.EventHandler(this.btnExportExcel_Click);

            this.btnExportXml.Location = new System.Drawing.Point(264, 12);
            this.btnExportXml.Name = "btnExportXml";
            this.btnExportXml.Size = new System.Drawing.Size(120, 30);
            this.btnExportXml.TabIndex = 2;
            this.btnExportXml.Text = "Экспорт в XML";
            this.btnExportXml.UseVisualStyleBackColor = true;
            this.btnExportXml.Click += new System.EventHandler(this.btnExportXml_Click);

            this.btnClearData.Location = new System.Drawing.Point(390, 12);
            this.btnClearData.Name = "btnClearData";
            this.btnClearData.Size = new System.Drawing.Size(120, 30);
            this.btnClearData.TabIndex = 3;
            this.btnClearData.Text = "Очистить данные";
            this.btnClearData.UseVisualStyleBackColor = true;
            this.btnClearData.Click += new System.EventHandler(this.btnClearData_Click);
  
            this.lblTotalRecords.AutoSize = true;
            this.lblTotalRecords.Location = new System.Drawing.Point(12, 55);
            this.lblTotalRecords.Name = "lblTotalRecords";
            this.lblTotalRecords.Size = new System.Drawing.Size(132, 15);
            this.lblTotalRecords.TabIndex = 4;
            this.lblTotalRecords.Text = "Всего записей в базе: 0";
   
            this.dtpStartDate.Enabled = false;
            this.dtpStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpStartDate.Location = new System.Drawing.Point(100, 90);
            this.dtpStartDate.Name = "dtpStartDate";
            this.dtpStartDate.Size = new System.Drawing.Size(120, 23);
            this.dtpStartDate.TabIndex = 5;

            this.dtpEndDate.Enabled = false;
            this.dtpEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpEndDate.Location = new System.Drawing.Point(100, 120);
            this.dtpEndDate.Name = "dtpEndDate";
            this.dtpEndDate.Size = new System.Drawing.Size(120, 23);
            this.dtpEndDate.TabIndex = 6;

            this.chkStartDate.AutoSize = true;
            this.chkStartDate.Location = new System.Drawing.Point(20, 92);
            this.chkStartDate.Name = "chkStartDate";
            this.chkStartDate.Size = new System.Drawing.Size(74, 19);
            this.chkStartDate.TabIndex = 7;
            this.chkStartDate.Text = "Дата с:";
            this.chkStartDate.UseVisualStyleBackColor = true;
            this.chkStartDate.CheckedChanged += new System.EventHandler(this.chkStartDate_CheckedChanged);

            this.chkEndDate.AutoSize = true;
            this.chkEndDate.Location = new System.Drawing.Point(20, 122);
            this.chkEndDate.Name = "chkEndDate";
            this.chkEndDate.Size = new System.Drawing.Size(79, 19);
            this.chkEndDate.TabIndex = 8;
            this.chkEndDate.Text = "Дата по:";
            this.chkEndDate.UseVisualStyleBackColor = true;
            this.chkEndDate.CheckedChanged += new System.EventHandler(this.chkEndDate_CheckedChanged);

            this.txtFirstName.Location = new System.Drawing.Point(300, 90);
            this.txtFirstName.Name = "txtFirstName";
            this.txtFirstName.PlaceholderText = "Имя";
            this.txtFirstName.Size = new System.Drawing.Size(150, 23);
            this.txtFirstName.TabIndex = 9;
 
            this.txtLastName.Location = new System.Drawing.Point(300, 120);
            this.txtLastName.Name = "txtLastName";
            this.txtLastName.PlaceholderText = "Фамилия";
            this.txtLastName.Size = new System.Drawing.Size(150, 23);
            this.txtLastName.TabIndex = 10;

            this.txtSurName.Location = new System.Drawing.Point(300, 150);
            this.txtSurName.Name = "txtSurName";
            this.txtSurName.PlaceholderText = "Отчество";
            this.txtSurName.Size = new System.Drawing.Size(150, 23);
            this.txtSurName.TabIndex = 11;

            this.txtCity.Location = new System.Drawing.Point(300, 180);
            this.txtCity.Name = "txtCity";
            this.txtCity.PlaceholderText = "Город";
            this.txtCity.Size = new System.Drawing.Size(150, 23);
            this.txtCity.TabIndex = 12;

            this.txtCountry.Location = new System.Drawing.Point(300, 210);
            this.txtCountry.Name = "txtCountry";
            this.txtCountry.PlaceholderText = "Страна";
            this.txtCountry.Size = new System.Drawing.Size(150, 23);
            this.txtCountry.TabIndex = 13;

            this.progressBar.Location = new System.Drawing.Point(12, 250);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(498, 23);
            this.progressBar.TabIndex = 14;
            this.progressBar.Visible = false;

            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(240, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 15);
            this.label1.TabIndex = 15;
            this.label1.Text = "Имя:";
  
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(240, 123);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 15);
            this.label2.TabIndex = 16;
            this.label2.Text = "Фамилия:";

            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(240, 153);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 15);
            this.label3.TabIndex = 17;
            this.label3.Text = "Отчество:";

            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(240, 183);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(43, 15);
            this.label4.TabIndex = 18;
            this.label4.Text = "Город:";

            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(240, 213);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 15);
            this.label5.TabIndex = 19;
            this.label5.Text = "Страна:";

            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label6.Location = new System.Drawing.Point(20, 70);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(120, 15);
            this.label6.TabIndex = 20;
            this.label6.Text = "ФИЛЬТРЫ ДАТЫ:";

            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label7.Location = new System.Drawing.Point(240, 70);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(149, 15);
            this.label7.TabIndex = 21;
            this.label7.Text = "ФИЛЬТРЫ ДАННЫХ:";
      
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 280);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 15);
            this.lblStatus.TabIndex = 22;

            this.ClientSize = new System.Drawing.Size(522, 300); 
            this.Controls.Add(this.lblStatus); 
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.txtCountry);
            this.Controls.Add(this.txtCity);
            this.Controls.Add(this.txtSurName);
            this.Controls.Add(this.txtLastName);
            this.Controls.Add(this.txtFirstName);
            this.Controls.Add(this.chkEndDate);
            this.Controls.Add(this.chkStartDate);
            this.Controls.Add(this.dtpEndDate);
            this.Controls.Add(this.dtpStartDate);
            this.Controls.Add(this.lblTotalRecords);
            this.Controls.Add(this.btnClearData);
            this.Controls.Add(this.btnExportXml);
            this.Controls.Add(this.btnExportExcel);
            this.Controls.Add(this.btnImport);
            this.Name = "MainForm"; 
            this.Text = "CSV Import/Export Application";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private Button btnImport;
        private Button btnExportExcel;
        private Button btnExportXml;
        private Button btnClearData;
        private Label lblTotalRecords;
        private DateTimePicker dtpStartDate;
        private DateTimePicker dtpEndDate;
        private CheckBox chkStartDate;
        private CheckBox chkEndDate;
        private TextBox txtFirstName;
        private TextBox txtLastName;
        private TextBox txtSurName;
        private TextBox txtCity;
        private TextBox txtCountry;
        private ProgressBar progressBar;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private Label label7;
        private Label lblStatus; 
    }
}