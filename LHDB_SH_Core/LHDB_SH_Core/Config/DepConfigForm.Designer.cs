namespace LHDB_SH_Core.Config
{
    partial class DepConfigForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgDepData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.dgClsData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.colDepName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDepValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDepTagName = new DevComponents.DotNetBar.Controls.DataGridViewComboBoxExColumn();
            this.colClsName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colClsValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colClsTagName = new DevComponents.DotNetBar.Controls.DataGridViewComboBoxExColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgDepData)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgClsData)).BeginInit();
            this.SuspendLayout();
            // 
            // dgDepData
            // 
            this.dgDepData.AllowUserToAddRows = false;
            this.dgDepData.AllowUserToDeleteRows = false;
            this.dgDepData.BackgroundColor = System.Drawing.Color.White;
            this.dgDepData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgDepData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colDepName,
            this.colDepValue,
            this.colDepTagName});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgDepData.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgDepData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgDepData.Location = new System.Drawing.Point(20, 47);
            this.dgDepData.Name = "dgDepData";
            this.dgDepData.RowTemplate.Height = 24;
            this.dgDepData.Size = new System.Drawing.Size(633, 169);
            this.dgDepData.TabIndex = 0;
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.AutoSize = true;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(491, 419);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(575, 419);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // dgClsData
            // 
            this.dgClsData.AllowUserToAddRows = false;
            this.dgClsData.AllowUserToDeleteRows = false;
            this.dgClsData.BackgroundColor = System.Drawing.Color.White;
            this.dgClsData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgClsData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colClsName,
            this.colClsValue,
            this.colClsTagName});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgClsData.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgClsData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgClsData.Location = new System.Drawing.Point(20, 249);
            this.dgClsData.Name = "dgClsData";
            this.dgClsData.RowTemplate.Height = 24;
            this.dgClsData.Size = new System.Drawing.Size(633, 155);
            this.dgClsData.TabIndex = 3;
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
            this.labelX1.Location = new System.Drawing.Point(20, 19);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(60, 21);
            this.labelX1.TabIndex = 4;
            this.labelX1.Text = "部別設定";
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
            this.labelX2.Location = new System.Drawing.Point(20, 222);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(60, 21);
            this.labelX2.TabIndex = 5;
            this.labelX2.Text = "班別設定";
            // 
            // colDepName
            // 
            this.colDepName.HeaderText = "部別名稱";
            this.colDepName.Name = "colDepName";
            this.colDepName.ReadOnly = true;
            this.colDepName.Width = 200;
            // 
            // colDepValue
            // 
            this.colDepValue.HeaderText = "值";
            this.colDepValue.Name = "colDepValue";
            this.colDepValue.ReadOnly = true;
            this.colDepValue.Width = 60;
            // 
            // colDepTagName
            // 
            this.colDepTagName.DisplayMember = "Text";
            this.colDepTagName.DropDownHeight = 106;
            this.colDepTagName.DropDownWidth = 121;
            this.colDepTagName.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.colDepTagName.HeaderText = "學生類別";
            this.colDepTagName.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.colDepTagName.IntegralHeight = false;
            this.colDepTagName.ItemHeight = 17;
            this.colDepTagName.Name = "colDepTagName";
            this.colDepTagName.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colDepTagName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.colDepTagName.Width = 200;
            // 
            // colClsName
            // 
            this.colClsName.HeaderText = "班別名稱";
            this.colClsName.Name = "colClsName";
            this.colClsName.ReadOnly = true;
            this.colClsName.Width = 200;
            // 
            // colClsValue
            // 
            this.colClsValue.HeaderText = "值";
            this.colClsValue.Name = "colClsValue";
            this.colClsValue.ReadOnly = true;
            this.colClsValue.Width = 60;
            // 
            // colClsTagName
            // 
            this.colClsTagName.DisplayMember = "Text";
            this.colClsTagName.DropDownHeight = 106;
            this.colClsTagName.DropDownWidth = 121;
            this.colClsTagName.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.colClsTagName.HeaderText = "學生類別";
            this.colClsTagName.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.colClsTagName.IntegralHeight = false;
            this.colClsTagName.ItemHeight = 17;
            this.colClsTagName.Name = "colClsTagName";
            this.colClsTagName.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colClsTagName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.colClsTagName.Width = 200;
            // 
            // DepConfigForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(674, 455);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.dgClsData);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.dgDepData);
            this.DoubleBuffered = true;
            this.Name = "DepConfigForm";
            this.Text = "部別與班別對照設定";
            this.Load += new System.EventHandler(this.DepConfigForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgDepData)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgClsData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgDepData;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgClsData;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDepName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDepValue;
        private DevComponents.DotNetBar.Controls.DataGridViewComboBoxExColumn colDepTagName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colClsName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colClsValue;
        private DevComponents.DotNetBar.Controls.DataGridViewComboBoxExColumn colClsTagName;
    }
}