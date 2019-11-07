namespace HertZ_ExcelAddIn
{
    partial class CASetting
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CASetting));
            this.ConfirmationSetting = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.AccountingFirmName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Auditee = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.ReplyAddress = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.PostalCode = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.AuditDeadline = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.Contact = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.Telephone = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.Department = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.Leading = new System.Windows.Forms.TextBox();
            this.ConfirmationSetting.SuspendLayout();
            this.SuspendLayout();
            // 
            // ConfirmationSetting
            // 
            this.ConfirmationSetting.Controls.Add(this.Leading);
            this.ConfirmationSetting.Controls.Add(this.label9);
            this.ConfirmationSetting.Controls.Add(this.Department);
            this.ConfirmationSetting.Controls.Add(this.label8);
            this.ConfirmationSetting.Controls.Add(this.Telephone);
            this.ConfirmationSetting.Controls.Add(this.label7);
            this.ConfirmationSetting.Controls.Add(this.Contact);
            this.ConfirmationSetting.Controls.Add(this.label6);
            this.ConfirmationSetting.Controls.Add(this.AuditDeadline);
            this.ConfirmationSetting.Controls.Add(this.label5);
            this.ConfirmationSetting.Controls.Add(this.PostalCode);
            this.ConfirmationSetting.Controls.Add(this.label4);
            this.ConfirmationSetting.Controls.Add(this.ReplyAddress);
            this.ConfirmationSetting.Controls.Add(this.label3);
            this.ConfirmationSetting.Controls.Add(this.Auditee);
            this.ConfirmationSetting.Controls.Add(this.label2);
            this.ConfirmationSetting.Controls.Add(this.AccountingFirmName);
            this.ConfirmationSetting.Controls.Add(this.label1);
            this.ConfirmationSetting.Font = new System.Drawing.Font("微软雅黑 Light", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ConfirmationSetting.Location = new System.Drawing.Point(137, 45);
            this.ConfirmationSetting.Name = "ConfirmationSetting";
            this.ConfirmationSetting.Size = new System.Drawing.Size(524, 390);
            this.ConfirmationSetting.TabIndex = 3;
            this.ConfirmationSetting.TabStop = false;
            this.ConfirmationSetting.Text = "函证信息设置";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label1.Location = new System.Drawing.Point(7, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 31);
            this.label1.TabIndex = 0;
            this.label1.Text = "事务所：";
            // 
            // AccountingFirmName
            // 
            this.AccountingFirmName.Location = new System.Drawing.Point(101, 38);
            this.AccountingFirmName.MaxLength = 100;
            this.AccountingFirmName.Name = "AccountingFirmName";
            this.AccountingFirmName.Size = new System.Drawing.Size(398, 39);
            this.AccountingFirmName.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(7, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(158, 31);
            this.label2.TabIndex = 2;
            this.label2.Text = "被审计单位：";
            // 
            // Auditee
            // 
            this.Auditee.Location = new System.Drawing.Point(149, 99);
            this.Auditee.MaxLength = 100;
            this.Auditee.Name = "Auditee";
            this.Auditee.Size = new System.Drawing.Size(350, 39);
            this.Auditee.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label3.Location = new System.Drawing.Point(7, 161);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(134, 31);
            this.label3.TabIndex = 6;
            this.label3.Text = "回函地址：";
            // 
            // ReplyAddress
            // 
            this.ReplyAddress.Location = new System.Drawing.Point(127, 158);
            this.ReplyAddress.Name = "ReplyAddress";
            this.ReplyAddress.Size = new System.Drawing.Size(372, 39);
            this.ReplyAddress.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(7, 222);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(86, 31);
            this.label4.TabIndex = 8;
            this.label4.Text = "邮编：";
            // 
            // PostalCode
            // 
            this.PostalCode.Location = new System.Drawing.Point(82, 219);
            this.PostalCode.Name = "PostalCode";
            this.PostalCode.Size = new System.Drawing.Size(124, 39);
            this.PostalCode.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label5.Location = new System.Drawing.Point(212, 222);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(158, 31);
            this.label5.TabIndex = 10;
            this.label5.Text = "审计截止日：";
            // 
            // AuditDeadline
            // 
            this.AuditDeadline.Location = new System.Drawing.Point(354, 219);
            this.AuditDeadline.Name = "AuditDeadline";
            this.AuditDeadline.Size = new System.Drawing.Size(145, 39);
            this.AuditDeadline.TabIndex = 11;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label6.Location = new System.Drawing.Point(7, 276);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(110, 31);
            this.label6.TabIndex = 12;
            this.label6.Text = "联系人：";
            // 
            // Contact
            // 
            this.Contact.Location = new System.Drawing.Point(106, 273);
            this.Contact.Name = "Contact";
            this.Contact.Size = new System.Drawing.Size(124, 39);
            this.Contact.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label7.Location = new System.Drawing.Point(236, 276);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(134, 31);
            this.label7.TabIndex = 14;
            this.label7.Text = "联系电话：";
            // 
            // Telephone
            // 
            this.Telephone.Location = new System.Drawing.Point(354, 273);
            this.Telephone.Name = "Telephone";
            this.Telephone.Size = new System.Drawing.Size(145, 39);
            this.Telephone.TabIndex = 15;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label8.Location = new System.Drawing.Point(7, 332);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(86, 31);
            this.label8.TabIndex = 16;
            this.label8.Text = "部门：";
            // 
            // Department
            // 
            this.Department.Location = new System.Drawing.Point(82, 329);
            this.Department.Name = "Department";
            this.Department.Size = new System.Drawing.Size(124, 39);
            this.Department.TabIndex = 17;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.label9.Location = new System.Drawing.Point(212, 332);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(158, 31);
            this.label9.TabIndex = 18;
            this.label9.Text = "项目负责人：";
            // 
            // Leading
            // 
            this.Leading.Location = new System.Drawing.Point(354, 329);
            this.Leading.Name = "Leading";
            this.Leading.Size = new System.Drawing.Size(145, 39);
            this.Leading.TabIndex = 19;
            // 
            // CASetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 564);
            this.Controls.Add(this.ConfirmationSetting);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CASetting";
            this.Text = "往来款项加工设置";
            this.Load += new System.EventHandler(this.CASetting_Load);
            this.ConfirmationSetting.ResumeLayout(false);
            this.ConfirmationSetting.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.GroupBox ConfirmationSetting;
        private System.Windows.Forms.TextBox Telephone;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox Contact;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox AuditDeadline;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox PostalCode;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox ReplyAddress;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Auditee;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox AccountingFirmName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Leading;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox Department;
        private System.Windows.Forms.Label label8;
    }
}