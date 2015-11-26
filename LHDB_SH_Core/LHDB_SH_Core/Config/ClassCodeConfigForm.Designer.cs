namespace LHDB_SH_Core.Config
{
    partial class ClassCodeConfigForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.dgClassCode = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.lnStData = new System.Windows.Forms.LinkLabel();
            this.colClassName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStClassCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colLHClassCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgClassCode)).BeginInit();
            this.SuspendLayout();
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
            this.labelX1.Location = new System.Drawing.Point(17, 14);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(60, 21);
            this.labelX1.TabIndex = 6;
            this.labelX1.Text = "班級代碼";
            // 
            // dgClassCode
            // 
            this.dgClassCode.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgClassCode.BackgroundColor = System.Drawing.Color.White;
            this.dgClassCode.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgClassCode.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colClassName,
            this.colStClassCode,
            this.colLHClassCode});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgClassCode.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgClassCode.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgClassCode.Location = new System.Drawing.Point(12, 43);
            this.dgClassCode.Name = "dgClassCode";
            this.dgClassCode.RowTemplate.Height = 24;
            this.dgClassCode.Size = new System.Drawing.Size(466, 317);
            this.dgClassCode.TabIndex = 5;
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(397, 376);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 8;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.AutoSize = true;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(316, 376);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // lnStData
            // 
            this.lnStData.AutoSize = true;
            this.lnStData.BackColor = System.Drawing.Color.Transparent;
            this.lnStData.Location = new System.Drawing.Point(17, 380);
            this.lnStData.Name = "lnStData";
            this.lnStData.Size = new System.Drawing.Size(112, 17);
            this.lnStData.TabIndex = 9;
            this.lnStData.TabStop = true;
            this.lnStData.Text = "取得繁星班級代碼";
            this.lnStData.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnStData_LinkClicked);
            // 
            // colClassName
            // 
            this.colClassName.HeaderText = "班級名稱";
            this.colClassName.Name = "colClassName";
            // 
            // colStClassCode
            // 
            this.colStClassCode.HeaderText = "繁星班級代碼";
            this.colStClassCode.Name = "colStClassCode";
            this.colStClassCode.ReadOnly = true;
            this.colStClassCode.Width = 120;
            // 
            // colLHClassCode
            // 
            this.colLHClassCode.HeaderText = "學習歷程班級代碼";
            this.colLHClassCode.Name = "colLHClassCode";
            this.colLHClassCode.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colLHClassCode.Width = 150;
            // 
            // ClassCodeConfigForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 414);
            this.Controls.Add(this.lnStData);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.dgClassCode);
            this.DoubleBuffered = true;
            this.Name = "ClassCodeConfigForm";
            this.Text = "設定班級代碼";
            this.Load += new System.EventHandler(this.ClassCodeConfigForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgClassCode)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgClassCode;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private System.Windows.Forms.LinkLabel lnStData;
        private System.Windows.Forms.DataGridViewTextBoxColumn colClassName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStClassCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn colLHClassCode;
    }
}