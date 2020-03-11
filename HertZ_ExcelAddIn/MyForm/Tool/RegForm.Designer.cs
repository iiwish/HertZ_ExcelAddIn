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
            "去Email地址: \\w+([-+.]\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*",
            "去网址URL: [a-zA-z]+: //[^\\s]*",
            "匹配帐号是否合法(字母开头,允许5-16字节,允许字母数字下划线): ^[a-zA-Z][a-zA-Z0-9_]{4,15}$",
            "匹配国内电话号码(如0511-4405222或021-87888822): \\d{3}-\\d{8}|\\d{4}-\\d{7}",
            "匹配腾讯QQ号: [1-9][0-9]{4,}",
            "匹配中国邮政编码: [1-9]\\d{5}(?!\\d)",
            "匹配身份证: \\d{15}|\\d{18}",
            "匹配ip地址: \\d+\\.\\d+\\.\\d+\\.\\d+",
            "匹配正整数: [1-9]\\d*$",
            "匹配负整数: -[1-9]\\d*$",
            "匹配整数: -?[1-9]\\d*$",
            "匹配非负整数（正整数+0）: [1-9]\\d*|0$",
            "匹配非正整数（负整数+0）: -[1-9]\\d*|0$",
            "匹配正浮点数: [1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*$",
            "匹配负浮点数: -([1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*)$",
            "匹配浮点数: -?([1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*|0?\\.0+|0)$",
            "匹配非负浮点数（正浮点数+0）: [1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*|0?\\.0+|0$",
            "匹配非正浮点数（负浮点数+0）: (-([1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*))|0?\\.0+|0$",
            "匹配由26个英文字母组成的字符串: [A-Za-z]+$",
            "匹配由26个英文字母的大写组成的字符串: [A-Z]+$",
            "匹配由26个英文字母的小写组成的字符串: [a-z]+$",
            "匹配由数字和26个英文字母组成的字符串: [A-Za-z0-9]+$",
            "匹配由数字26个英文字母或者下划线组成的字符串: \\w+$"});
            this.comboBox.Location = new System.Drawing.Point(35, 65);
            this.comboBox.Name = "comboBox";
            this.comboBox.Size = new System.Drawing.Size(700, 44);
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
            // 
            // RegForm
            // 
            this.AcceptButton = this.ConfirmBtn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(774, 229);
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

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox;
        private System.Windows.Forms.Button ConfirmBtn;
    }
}