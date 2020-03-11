namespace HertZ_ExcelAddIn.MyForm.WorkSheet
{
    partial class UnionSheetForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UnionSheetForm));
            this.ListBox = new System.Windows.Forms.CheckedListBox();
            this.SelectAll = new System.Windows.Forms.Button();
            this.Confirm = new System.Windows.Forms.Button();
            this.SelectNone = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.NumUpDown = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.NumUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // ListBox
            // 
            this.ListBox.FormattingEnabled = true;
            this.ListBox.Location = new System.Drawing.Point(70, 120);
            this.ListBox.Name = "ListBox";
            this.ListBox.Size = new System.Drawing.Size(420, 328);
            this.ListBox.TabIndex = 0;
            // 
            // SelectAll
            // 
            this.SelectAll.Location = new System.Drawing.Point(70, 50);
            this.SelectAll.Name = "SelectAll";
            this.SelectAll.Size = new System.Drawing.Size(150, 50);
            this.SelectAll.TabIndex = 1;
            this.SelectAll.Text = "全 选";
            this.SelectAll.UseVisualStyleBackColor = true;
            this.SelectAll.Click += new System.EventHandler(this.SelectAll_Click);
            // 
            // Confirm
            // 
            this.Confirm.Location = new System.Drawing.Point(200, 550);
            this.Confirm.Name = "Confirm";
            this.Confirm.Size = new System.Drawing.Size(150, 50);
            this.Confirm.TabIndex = 2;
            this.Confirm.Text = "确 定";
            this.Confirm.UseVisualStyleBackColor = true;
            this.Confirm.Click += new System.EventHandler(this.Confirm_Click);
            // 
            // SelectNone
            // 
            this.SelectNone.Location = new System.Drawing.Point(340, 50);
            this.SelectNone.Name = "SelectNone";
            this.SelectNone.Size = new System.Drawing.Size(150, 50);
            this.SelectNone.TabIndex = 3;
            this.SelectNone.Text = "取消全选";
            this.SelectNone.UseVisualStyleBackColor = true;
            this.SelectNone.Click += new System.EventHandler(this.SelectNone_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(164, 476);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(239, 31);
            this.label1.TabIndex = 4;
            this.label1.Text = "跳过标题               行";
            // 
            // NumUpDown
            // 
            this.NumUpDown.Location = new System.Drawing.Point(279, 474);
            this.NumUpDown.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.NumUpDown.Name = "NumUpDown";
            this.NumUpDown.Size = new System.Drawing.Size(82, 39);
            this.NumUpDown.TabIndex = 5;
            this.NumUpDown.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // UnionSheetForm
            // 
            this.AcceptButton = this.Confirm;
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(574, 629);
            this.Controls.Add(this.NumUpDown);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.SelectNone);
            this.Controls.Add(this.Confirm);
            this.Controls.Add(this.SelectAll);
            this.Controls.Add(this.ListBox);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UnionSheetForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "选择表单";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.UnionSheetForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.NumUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox ListBox;
        private System.Windows.Forms.Button SelectAll;
        private System.Windows.Forms.Button Confirm;
        private System.Windows.Forms.Button SelectNone;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown NumUpDown;
    }
}