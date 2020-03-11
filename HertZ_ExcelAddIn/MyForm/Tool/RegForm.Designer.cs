namespace HertZ_ExcelAddIn.MyForm.Tool
{
    partial class RegForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RegForm));
            this.comboBox = new System.Windows.Forms.ComboBox();
            this.ConfirmBtn = new System.Windows.Forms.Button();
            this.ChangeBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // comboBox
            // 
            this.comboBox.DropDownWidth = 700;
            this.comboBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.comboBox.FormattingEnabled = true;
            this.comboBox.Items.AddRange(new object[] {
            "去中文: [\\u4e00-\\u9fa5]",
            "留中文: [^\\u4e00-\\u9fa5]",
            "去字母: [A-Za-z]",
            "留字母: [^A-Za-z]",
            "去数字: \\d+(\\.\\d)?",
            "留数字: ^\\d+(\\.\\d)?",
            "去数字字符: \\d",
            "去非数字字符: \\D",
            "去换页字符: \\f",
            "去换行字符: \\n",
            "去回车符字符: \\r",
            "去任何空白: \\s",
            "去任何非空白字符: \\S",
            "去任何非空白字符: [^\\f\\n\\r\\t\\v]",
            "去制表字符: \\t",
            "去垂直制表符: \\v",
            "去包括下划线在内的任何字字符: \\w",
            "去包括下划线在内的任何字字符: [A-Za-z0-9_]",
            "去任何非字字符: \\W",
            "去任何非字字符: [^A-Za-z0-9_]",
            "去双字节字符(包括汉字在内): [^\\x00-\\xff]",
            "去HTML标记: <(\\S*?)[^>]*>.*?</\\1>|<.*?/>",
            "去首尾空白字符: ^\\s*|\\s*$",
            "去Email地址: \\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*"});
            this.comboBox.Location = new System.Drawing.Point(170, 65);
            this.comboBox.Name = "comboBox";
            this.comboBox.Size = new System.Drawing.Size(560, 44);
            this.comboBox.TabIndex = 0;
            // 
            // ConfirmBtn
            // 
            this.ConfirmBtn.Location = new System.Drawing.Point(605, 150);
            this.ConfirmBtn.Name = "ConfirmBtn";
            this.ConfirmBtn.Size = new System.Drawing.Size(130, 50);
            this.ConfirmBtn.TabIndex = 1;
            this.ConfirmBtn.Text = "确 认";
            this.ConfirmBtn.UseVisualStyleBackColor = true;
            this.ConfirmBtn.Click += new System.EventHandler(this.ConfirmBtn_Click);
            // 
            // ChangeBtn
            // 
            this.ChangeBtn.BackColor = System.Drawing.SystemColors.Window;
            this.ChangeBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ChangeBtn.Location = new System.Drawing.Point(50, 65);
            this.ChangeBtn.Name = "ChangeBtn";
            this.ChangeBtn.Size = new System.Drawing.Size(120, 44);
            this.ChangeBtn.TabIndex = 2;
            this.ChangeBtn.Text = "Replace";
            this.ChangeBtn.UseVisualStyleBackColor = false;
            this.ChangeBtn.Click += new System.EventHandler(this.ChangeBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.label1.Location = new System.Drawing.Point(201, 160);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(398, 31);
            this.label1.TabIndex = 3;
            this.label1.Text = "注意：此操作无法撤销，请手动备份";
            // 
            // RegForm
            // 
            this.AcceptButton = this.ConfirmBtn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(774, 229);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ChangeBtn);
            this.Controls.Add(this.ConfirmBtn);
            this.Controls.Add(this.comboBox);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RegForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "正则表达式";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox;
        private System.Windows.Forms.Button ConfirmBtn;
        private System.Windows.Forms.Button ChangeBtn;
        private System.Windows.Forms.Label label1;
    }
}