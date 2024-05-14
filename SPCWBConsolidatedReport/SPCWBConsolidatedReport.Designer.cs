
namespace SPCWBConsolidatedReport
{
    partial class SPCWBConsolidatedReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SPCWBConsolidatedReport));
            this.pnlFile = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnClearFiles = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnExportToExcel = new System.Windows.Forms.Button();
            this.pnlSearch = new System.Windows.Forms.Panel();
            this.btnSearch = new System.Windows.Forms.Button();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.pnlsfdatagrid = new System.Windows.Forms.Panel();
            this.sfdgSPCWBConsolidatedReport = new Syncfusion.WinForms.DataGrid.SfDataGrid();
            this.pnlFile.SuspendLayout();
            this.panel1.SuspendLayout();
            this.pnlSearch.SuspendLayout();
            this.pnlsfdatagrid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sfdgSPCWBConsolidatedReport)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlFile
            // 
            this.pnlFile.AutoScroll = true;
            this.pnlFile.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.pnlFile.Controls.Add(this.pnlsfdatagrid);
            this.pnlFile.Controls.Add(this.panel1);
            this.pnlFile.Controls.Add(this.pnlSearch);
            this.pnlFile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFile.Location = new System.Drawing.Point(0, 0);
            this.pnlFile.Name = "pnlFile";
            this.pnlFile.Size = new System.Drawing.Size(397, 512);
            this.pnlFile.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel1.Controls.Add(this.btnClearFiles);
            this.panel1.Controls.Add(this.btnClose);
            this.panel1.Controls.Add(this.btnAdd);
            this.panel1.Controls.Add(this.btnExportToExcel);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 472);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(397, 40);
            this.panel1.TabIndex = 2;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // btnClearFiles
            // 
            this.btnClearFiles.Location = new System.Drawing.Point(299, 5);
            this.btnClearFiles.Name = "btnClearFiles";
            this.btnClearFiles.Size = new System.Drawing.Size(75, 23);
            this.btnClearFiles.TabIndex = 4;
            this.btnClearFiles.Text = "Clear Files";
            this.btnClearFiles.UseVisualStyleBackColor = true;
            this.btnClearFiles.Click += new System.EventHandler(this.btnClearFiles_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(218, 5);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(3, 5);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 3;
            this.btnAdd.Text = "Add Files";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnExportToExcel
            // 
            this.btnExportToExcel.Location = new System.Drawing.Point(84, 5);
            this.btnExportToExcel.Name = "btnExportToExcel";
            this.btnExportToExcel.Size = new System.Drawing.Size(128, 23);
            this.btnExportToExcel.TabIndex = 0;
            this.btnExportToExcel.Text = "Export To Excel";
            this.btnExportToExcel.UseVisualStyleBackColor = true;
            this.btnExportToExcel.Click += new System.EventHandler(this.btnExportToExcel_Click);
            // 
            // pnlSearch
            // 
            this.pnlSearch.Controls.Add(this.btnSearch);
            this.pnlSearch.Controls.Add(this.txtSearch);
            this.pnlSearch.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlSearch.Location = new System.Drawing.Point(0, 0);
            this.pnlSearch.Name = "pnlSearch";
            this.pnlSearch.Size = new System.Drawing.Size(397, 40);
            this.pnlSearch.TabIndex = 0;
            // 
            // btnSearch
            // 
            this.btnSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSearch.Location = new System.Drawing.Point(310, 7);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(75, 23);
            this.btnSearch.TabIndex = 1;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // txtSearch
            // 
            this.txtSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSearch.Location = new System.Drawing.Point(204, 7);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(100, 20);
            this.txtSearch.TabIndex = 0;
            this.txtSearch.TextChanged += new System.EventHandler(this.txtSearch_TextChanged);
            // 
            // pnlsfdatagrid
            // 
            this.pnlsfdatagrid.Controls.Add(this.sfdgSPCWBConsolidatedReport);
            this.pnlsfdatagrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlsfdatagrid.Location = new System.Drawing.Point(0, 40);
            this.pnlsfdatagrid.Name = "pnlsfdatagrid";
            this.pnlsfdatagrid.Size = new System.Drawing.Size(397, 432);
            this.pnlsfdatagrid.TabIndex = 3;
            // 
            // sfdgSPCWBConsolidatedReport
            // 
            this.sfdgSPCWBConsolidatedReport.AccessibleName = "Table";
            this.sfdgSPCWBConsolidatedReport.AllowDeleting = true;
            this.sfdgSPCWBConsolidatedReport.AllowEditing = false;
            this.sfdgSPCWBConsolidatedReport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sfdgSPCWBConsolidatedReport.Location = new System.Drawing.Point(0, 0);
            this.sfdgSPCWBConsolidatedReport.Name = "sfdgSPCWBConsolidatedReport";
            this.sfdgSPCWBConsolidatedReport.SelectionMode = Syncfusion.WinForms.DataGrid.Enums.GridSelectionMode.Extended;
            this.sfdgSPCWBConsolidatedReport.Size = new System.Drawing.Size(397, 432);
            this.sfdgSPCWBConsolidatedReport.TabIndex = 2;
            this.sfdgSPCWBConsolidatedReport.Text = "sfDataGrid1";
            // 
            // SPCWBConsolidatedReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(397, 512);
            this.Controls.Add(this.pnlFile);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SPCWBConsolidatedReport";
            this.Text = "SPCWB Consolidated Report";
            this.Load += new System.EventHandler(this.SPCWBConsolidatedReport_Load);
            this.pnlFile.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.pnlSearch.ResumeLayout(false);
            this.pnlSearch.PerformLayout();
            this.pnlsfdatagrid.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.sfdgSPCWBConsolidatedReport)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlFile;
        private System.Windows.Forms.Panel pnlSearch;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnExportToExcel;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnClearFiles;
        private System.Windows.Forms.Panel pnlsfdatagrid;
        private Syncfusion.WinForms.DataGrid.SfDataGrid sfdgSPCWBConsolidatedReport;
    }
}

