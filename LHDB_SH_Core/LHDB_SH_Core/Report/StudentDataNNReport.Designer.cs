namespace LHDB_SH_Core.Report
{
    partial class StudentDataNNReport
    {
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
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.iptSemester = new DevComponents.Editors.IntegerInput();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.iptSchoolYear = new DevComponents.Editors.IntegerInput();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnExport = new DevComponents.DotNetBar.ButtonX();
            this.iptClassDefault = new DevComponents.Editors.IntegerInput();
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.iptDepDefault = new DevComponents.Editors.IntegerInput();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.lnkDepSetup = new System.Windows.Forms.LinkLabel();
            this.txtDesc = new DevComponents.DotNetBar.Controls.TextBoxX();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptClassDefault)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptDepDefault)).BeginInit();
            this.SuspendLayout();
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(323, 12);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(34, 21);
            this.labelX2.TabIndex = 7;
            this.labelX2.Text = "學期";
            // 
            // iptSemester
            // 
            this.iptSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSemester.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSemester.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSemester.Location = new System.Drawing.Point(361, 10);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 1;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.ShowUpDown = true;
            this.iptSemester.Size = new System.Drawing.Size(53, 25);
            this.iptSemester.TabIndex = 6;
            this.iptSemester.Value = 1;
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(130, 14);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(47, 21);
            this.labelX1.TabIndex = 5;
            this.labelX1.Text = "學年度";
            // 
            // iptSchoolYear
            // 
            this.iptSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSchoolYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSchoolYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSchoolYear.Location = new System.Drawing.Point(183, 12);
            this.iptSchoolYear.MaxValue = 999;
            this.iptSchoolYear.MinValue = 1;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.ShowUpDown = true;
            this.iptSchoolYear.Size = new System.Drawing.Size(68, 25);
            this.iptSchoolYear.TabIndex = 4;
            this.iptSchoolYear.Value = 1;
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(437, 391);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 10;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnExport
            // 
            this.btnExport.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExport.AutoSize = true;
            this.btnExport.BackColor = System.Drawing.Color.Transparent;
            this.btnExport.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExport.Location = new System.Drawing.Point(347, 391);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 25);
            this.btnExport.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExport.TabIndex = 9;
            this.btnExport.Text = "產生";
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // iptClassDefault
            // 
            this.iptClassDefault.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptClassDefault.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptClassDefault.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptClassDefault.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptClassDefault.Location = new System.Drawing.Point(361, 57);
            this.iptClassDefault.MaxValue = 6;
            this.iptClassDefault.MinValue = 1;
            this.iptClassDefault.Name = "iptClassDefault";
            this.iptClassDefault.ShowUpDown = true;
            this.iptClassDefault.Size = new System.Drawing.Size(53, 25);
            this.iptClassDefault.TabIndex = 14;
            this.iptClassDefault.Value = 1;
            // 
            // labelX4
            // 
            this.labelX4.AutoSize = true;
            this.labelX4.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX4.BackgroundStyle.Class = "";
            this.labelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX4.Location = new System.Drawing.Point(261, 59);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(101, 21);
            this.labelX4.TabIndex = 13;
            this.labelX4.Text = "班別代碼預設值";
            // 
            // iptDepDefault
            // 
            this.iptDepDefault.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptDepDefault.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptDepDefault.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptDepDefault.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptDepDefault.Location = new System.Drawing.Point(183, 57);
            this.iptDepDefault.MaxValue = 6;
            this.iptDepDefault.MinValue = 1;
            this.iptDepDefault.Name = "iptDepDefault";
            this.iptDepDefault.ShowUpDown = true;
            this.iptDepDefault.Size = new System.Drawing.Size(68, 25);
            this.iptDepDefault.TabIndex = 12;
            this.iptDepDefault.Value = 1;
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(76, 59);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(101, 21);
            this.labelX3.TabIndex = 11;
            this.labelX3.Text = "部別代碼預設值";
            // 
            // lnkDepSetup
            // 
            this.lnkDepSetup.AutoSize = true;
            this.lnkDepSetup.BackColor = System.Drawing.Color.Transparent;
            this.lnkDepSetup.Location = new System.Drawing.Point(16, 112);
            this.lnkDepSetup.Name = "lnkDepSetup";
            this.lnkDepSetup.Size = new System.Drawing.Size(86, 17);
            this.lnkDepSetup.TabIndex = 15;
            this.lnkDepSetup.TabStop = true;
            this.lnkDepSetup.Text = "特殊類別對照";
            this.lnkDepSetup.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDepSetup_LinkClicked);
            // 
            // txtDesc
            // 
            this.txtDesc.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // 
            // 
            this.txtDesc.Border.Class = "TextBoxBorder";
            this.txtDesc.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtDesc.Location = new System.Drawing.Point(16, 147);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtDesc.Size = new System.Drawing.Size(501, 230);
            this.txtDesc.TabIndex = 23;
            // 
            // StudentDataNNReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(532, 427);
            this.Controls.Add(this.txtDesc);
            this.Controls.Add(this.lnkDepSetup);
            this.Controls.Add(this.iptClassDefault);
            this.Controls.Add(this.labelX4);
            this.Controls.Add(this.iptDepDefault);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.iptSemester);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.iptSchoolYear);
            this.DoubleBuffered = true;
            this.Name = "StudentDataNNReport";
            this.Text = "學生資料名冊(非國教署主管學校)";
            this.Load += new System.EventHandler(this.StudentDataNNReport_Load);
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptClassDefault)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptDepDefault)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.Editors.IntegerInput iptSemester;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.Editors.IntegerInput iptSchoolYear;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.ButtonX btnExport;
        private DevComponents.Editors.IntegerInput iptClassDefault;
        private DevComponents.DotNetBar.LabelX labelX4;
        private DevComponents.Editors.IntegerInput iptDepDefault;
        private DevComponents.DotNetBar.LabelX labelX3;
        private System.Windows.Forms.LinkLabel lnkDepSetup;
        private DevComponents.DotNetBar.Controls.TextBoxX txtDesc;
    }
}