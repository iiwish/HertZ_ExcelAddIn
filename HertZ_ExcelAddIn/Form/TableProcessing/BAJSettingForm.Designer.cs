namespace HertZ_ExcelAddIn
{
    partial class BAJSettingForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BAJSettingForm));
            this.SubjectCodeButton1 = new System.Windows.Forms.RadioButton();
            this.SubjectCodeButton2 = new System.Windows.Forms.RadioButton();
            this.SubjectCodeGroupBox = new System.Windows.Forms.GroupBox();
            this.SubjectCodeLength6 = new System.Windows.Forms.TextBox();
            this.SubjectCodeLength5 = new System.Windows.Forms.TextBox();
            this.SubjectCodeLength4 = new System.Windows.Forms.TextBox();
            this.SubjectCodeLength3 = new System.Windows.Forms.TextBox();
            this.SubjectCodeLength2 = new System.Windows.Forms.TextBox();
            this.SubjectCodeLength1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SubjectCodeSign = new System.Windows.Forms.TextBox();
            this.SubjectCodeButton3 = new System.Windows.Forms.RadioButton();
            this.ConfirmBtn = new System.Windows.Forms.Button();
            this.CancelBtn = new System.Windows.Forms.Button();
            this.OrderCheckBox = new System.Windows.Forms.CheckBox();
            this.SubjectCodeGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // SubjectCodeButton1
            // 
            this.SubjectCodeButton1.AutoSize = true;
            this.SubjectCodeButton1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.SubjectCodeButton1.Location = new System.Drawing.Point(6, 50);
            this.SubjectCodeButton1.Name = "SubjectCodeButton1";
            this.SubjectCodeButton1.Size = new System.Drawing.Size(189, 35);
            this.SubjectCodeButton1.TabIndex = 0;
            this.SubjectCodeButton1.TabStop = true;
            this.SubjectCodeButton1.Text = "自动识别编码";
            this.SubjectCodeButton1.UseVisualStyleBackColor = true;
            this.SubjectCodeButton1.Click += new System.EventHandler(this.ChangeState);
            // 
            // SubjectCodeButton2
            // 
            this.SubjectCodeButton2.AutoSize = true;
            this.SubjectCodeButton2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.SubjectCodeButton2.Location = new System.Drawing.Point(6, 100);
            this.SubjectCodeButton2.Name = "SubjectCodeButton2";
            this.SubjectCodeButton2.Size = new System.Drawing.Size(262, 35);
            this.SubjectCodeButton2.TabIndex = 1;
            this.SubjectCodeButton2.Text = "使用分隔符       拆分";
            this.SubjectCodeButton2.UseVisualStyleBackColor = true;
            this.SubjectCodeButton2.Click += new System.EventHandler(this.ChangeState);
            // 
            // SubjectCodeGroupBox
            // 
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeLength6);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeLength5);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeLength4);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeLength3);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeLength2);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeLength1);
            this.SubjectCodeGroupBox.Controls.Add(this.label2);
            this.SubjectCodeGroupBox.Controls.Add(this.label1);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeSign);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeButton3);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeButton1);
            this.SubjectCodeGroupBox.Controls.Add(this.SubjectCodeButton2);
            this.SubjectCodeGroupBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SubjectCodeGroupBox.Location = new System.Drawing.Point(137, 39);
            this.SubjectCodeGroupBox.Name = "SubjectCodeGroupBox";
            this.SubjectCodeGroupBox.Size = new System.Drawing.Size(502, 304);
            this.SubjectCodeGroupBox.TabIndex = 2;
            this.SubjectCodeGroupBox.TabStop = false;
            this.SubjectCodeGroupBox.Text = "请选择编码格式";
            // 
            // SubjectCodeLength6
            // 
            this.SubjectCodeLength6.Location = new System.Drawing.Point(436, 244);
            this.SubjectCodeLength6.Name = "SubjectCodeLength6";
            this.SubjectCodeLength6.Size = new System.Drawing.Size(38, 39);
            this.SubjectCodeLength6.TabIndex = 12;
            // 
            // SubjectCodeLength5
            // 
            this.SubjectCodeLength5.Location = new System.Drawing.Point(291, 244);
            this.SubjectCodeLength5.Name = "SubjectCodeLength5";
            this.SubjectCodeLength5.Size = new System.Drawing.Size(38, 39);
            this.SubjectCodeLength5.TabIndex = 11;
            // 
            // SubjectCodeLength4
            // 
            this.SubjectCodeLength4.Location = new System.Drawing.Point(146, 244);
            this.SubjectCodeLength4.Name = "SubjectCodeLength4";
            this.SubjectCodeLength4.Size = new System.Drawing.Size(38, 39);
            this.SubjectCodeLength4.TabIndex = 10;
            // 
            // SubjectCodeLength3
            // 
            this.SubjectCodeLength3.Location = new System.Drawing.Point(436, 195);
            this.SubjectCodeLength3.Name = "SubjectCodeLength3";
            this.SubjectCodeLength3.Size = new System.Drawing.Size(38, 39);
            this.SubjectCodeLength3.TabIndex = 9;
            // 
            // SubjectCodeLength2
            // 
            this.SubjectCodeLength2.Location = new System.Drawing.Point(291, 195);
            this.SubjectCodeLength2.Name = "SubjectCodeLength2";
            this.SubjectCodeLength2.Size = new System.Drawing.Size(38, 39);
            this.SubjectCodeLength2.TabIndex = 8;
            // 
            // SubjectCodeLength1
            // 
            this.SubjectCodeLength1.Location = new System.Drawing.Point(146, 195);
            this.SubjectCodeLength1.Name = "SubjectCodeLength1";
            this.SubjectCodeLength1.Size = new System.Drawing.Size(38, 39);
            this.SubjectCodeLength1.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(40, 247);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(400, 31);
            this.label2.TabIndex = 6;
            this.label2.Text = "四级科目       五级科目       六级科目";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(40, 200);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(400, 31);
            this.label1.TabIndex = 5;
            this.label1.Text = "一级科目       二级科目       三级科目";
            // 
            // SubjectCodeSign
            // 
            this.SubjectCodeSign.Location = new System.Drawing.Point(165, 98);
            this.SubjectCodeSign.Name = "SubjectCodeSign";
            this.SubjectCodeSign.Size = new System.Drawing.Size(38, 39);
            this.SubjectCodeSign.TabIndex = 4;
            // 
            // SubjectCodeButton3
            // 
            this.SubjectCodeButton3.AutoSize = true;
            this.SubjectCodeButton3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.SubjectCodeButton3.Location = new System.Drawing.Point(6, 150);
            this.SubjectCodeButton3.Name = "SubjectCodeButton3";
            this.SubjectCodeButton3.Size = new System.Drawing.Size(285, 35);
            this.SubjectCodeButton3.TabIndex = 2;
            this.SubjectCodeButton3.TabStop = true;
            this.SubjectCodeButton3.Text = "按照以下编码长度拆分";
            this.SubjectCodeButton3.UseVisualStyleBackColor = true;
            this.SubjectCodeButton3.Click += new System.EventHandler(this.ChangeState);
            // 
            // ConfirmBtn
            // 
            this.ConfirmBtn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ConfirmBtn.Location = new System.Drawing.Point(137, 469);
            this.ConfirmBtn.Name = "ConfirmBtn";
            this.ConfirmBtn.Size = new System.Drawing.Size(150, 50);
            this.ConfirmBtn.TabIndex = 4;
            this.ConfirmBtn.Text = "保存";
            this.ConfirmBtn.UseVisualStyleBackColor = true;
            this.ConfirmBtn.Click += new System.EventHandler(this.ConfirmBtn_Click);
            // 
            // CancelBtn
            // 
            this.CancelBtn.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CancelBtn.Location = new System.Drawing.Point(489, 469);
            this.CancelBtn.Name = "CancelBtn";
            this.CancelBtn.Size = new System.Drawing.Size(150, 50);
            this.CancelBtn.TabIndex = 5;
            this.CancelBtn.Text = "取消";
            this.CancelBtn.UseVisualStyleBackColor = true;
            this.CancelBtn.Click += new System.EventHandler(this.CancelBtn_Click);
            // 
            // OrderCheckBox
            // 
            this.OrderCheckBox.AutoSize = true;
            this.OrderCheckBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OrderCheckBox.Location = new System.Drawing.Point(143, 373);
            this.OrderCheckBox.Name = "OrderCheckBox";
            this.OrderCheckBox.Size = new System.Drawing.Size(454, 35);
            this.OrderCheckBox.TabIndex = 6;
            this.OrderCheckBox.Text = "加工时将余额表按照科目编码重新排序";
            this.OrderCheckBox.UseVisualStyleBackColor = true;
            // 
            // BAJSettingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(774, 565);
            this.Controls.Add(this.OrderCheckBox);
            this.Controls.Add(this.CancelBtn);
            this.Controls.Add(this.ConfirmBtn);
            this.Controls.Add(this.SubjectCodeGroupBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "BAJSettingForm";
            this.Text = "账表加工设置";
            this.Load += new System.EventHandler(this.BAJSettingForm_Load);
            this.SubjectCodeGroupBox.ResumeLayout(false);
            this.SubjectCodeGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton SubjectCodeButton1;
        private System.Windows.Forms.RadioButton SubjectCodeButton2;
        private System.Windows.Forms.GroupBox SubjectCodeGroupBox;
        private System.Windows.Forms.RadioButton SubjectCodeButton3;
        private System.Windows.Forms.TextBox SubjectCodeLength6;
        private System.Windows.Forms.TextBox SubjectCodeLength5;
        private System.Windows.Forms.TextBox SubjectCodeLength4;
        private System.Windows.Forms.TextBox SubjectCodeLength3;
        private System.Windows.Forms.TextBox SubjectCodeLength2;
        private System.Windows.Forms.TextBox SubjectCodeLength1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox SubjectCodeSign;
        private System.Windows.Forms.Button ConfirmBtn;
        private System.Windows.Forms.Button CancelBtn;
        private System.Windows.Forms.CheckBox OrderCheckBox;
    }
}