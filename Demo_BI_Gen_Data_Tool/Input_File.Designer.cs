
namespace Demo_BI_Gen_Data_Tool
{
    partial class Input_File
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
            this.label1 = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.btnSelect = new System.Windows.Forms.Button();
            this.cbxSheet = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cbxColumn = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cbxData = new System.Windows.Forms.ComboBox();
            this.dataGrid = new System.Windows.Forms.DataGridView();
            this.btnExport = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.lblCount = new System.Windows.Forms.Label();
            this.dateFrom = new System.Windows.Forms.DateTimePicker();
            this.dateTo = new System.Windows.Forms.DateTimePicker();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "File Name:";
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(87, 12);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(503, 22);
            this.txtFileName.TabIndex = 2;
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(596, 12);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(38, 23);
            this.btnSelect.TabIndex = 3;
            this.btnSelect.Text = ". . .";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // cbxSheet
            // 
            this.cbxSheet.FormattingEnabled = true;
            this.cbxSheet.Location = new System.Drawing.Point(87, 53);
            this.cbxSheet.Name = "cbxSheet";
            this.cbxSheet.Size = new System.Drawing.Size(121, 24);
            this.cbxSheet.TabIndex = 4;
            this.cbxSheet.SelectionChangeCommitted += new System.EventHandler(this.cbxSheet_SelectionChangeCommitted);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(32, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 17);
            this.label2.TabIndex = 5;
            this.label2.Text = "Sheet:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(234, 56);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 17);
            this.label3.TabIndex = 7;
            this.label3.Text = "Column:";
            // 
            // cbxColumn
            // 
            this.cbxColumn.FormattingEnabled = true;
            this.cbxColumn.Location = new System.Drawing.Point(299, 53);
            this.cbxColumn.Name = "cbxColumn";
            this.cbxColumn.Size = new System.Drawing.Size(121, 24);
            this.cbxColumn.TabIndex = 6;
            this.cbxColumn.SelectionChangeCommitted += new System.EventHandler(this.cbxColumn_SelectionChangeCommitted);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(465, 56);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(42, 17);
            this.label4.TabIndex = 9;
            this.label4.Text = "Data:";
            // 
            // cbxData
            // 
            this.cbxData.FormattingEnabled = true;
            this.cbxData.Location = new System.Drawing.Point(513, 53);
            this.cbxData.Name = "cbxData";
            this.cbxData.Size = new System.Drawing.Size(121, 24);
            this.cbxData.TabIndex = 8;
            this.cbxData.SelectionChangeCommitted += new System.EventHandler(this.cbxData_SelectionChangeCommitted);
            // 
            // dataGrid
            // 
            this.dataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid.Location = new System.Drawing.Point(12, 143);
            this.dataGrid.Name = "dataGrid";
            this.dataGrid.RowHeadersWidth = 51;
            this.dataGrid.RowTemplate.Height = 24;
            this.dataGrid.Size = new System.Drawing.Size(926, 387);
            this.dataGrid.TabIndex = 10;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(863, 536);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 11;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 542);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 17);
            this.label5.TabIndex = 12;
            this.label5.Text = "Count:";
            // 
            // lblCount
            // 
            this.lblCount.AutoSize = true;
            this.lblCount.Location = new System.Drawing.Point(67, 542);
            this.lblCount.Name = "lblCount";
            this.lblCount.Size = new System.Drawing.Size(16, 17);
            this.lblCount.TabIndex = 13;
            this.lblCount.Text = "0";
            // 
            // dateFrom
            // 
            this.dateFrom.CustomFormat = "dd/MM/yyyy";
            this.dateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateFrom.Location = new System.Drawing.Point(87, 93);
            this.dateFrom.Name = "dateFrom";
            this.dateFrom.Size = new System.Drawing.Size(121, 22);
            this.dateFrom.TabIndex = 14;
            // 
            // dateTo
            // 
            this.dateTo.CustomFormat = "dd/MM/yyyy";
            this.dateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTo.Location = new System.Drawing.Point(299, 93);
            this.dateTo.Name = "dateTo";
            this.dateTo.Size = new System.Drawing.Size(121, 22);
            this.dateTo.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(32, 98);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(44, 17);
            this.label6.TabIndex = 16;
            this.label6.Text = "From:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(249, 98);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 17);
            this.label7.TabIndex = 17;
            this.label7.Text = "To:";
            // 
            // Input_File
            // 
            this.ClientSize = new System.Drawing.Size(950, 570);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.dateTo);
            this.Controls.Add(this.dateFrom);
            this.Controls.Add(this.lblCount);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.dataGrid);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cbxData);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbxColumn);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cbxSheet);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.label1);
            this.Name = "Input_File";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "BI Tool";
            this.Load += new System.EventHandler(this.Input_File_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.ComboBox cbxSheet;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbxColumn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbxData;
        private System.Windows.Forms.DataGridView dataGrid;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblCount;
        private System.Windows.Forms.DateTimePicker dateFrom;
        private System.Windows.Forms.DateTimePicker dateTo;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
    }
}