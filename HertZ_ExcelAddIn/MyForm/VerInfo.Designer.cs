namespace HertZ_ExcelAddIn
{
    partial class VerInfo
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VerInfo));
            this.Manual = new System.Windows.Forms.Button();
            this.OnlineVideo = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Manual
            // 
            this.Manual.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Manual.Location = new System.Drawing.Point(150, 300);
            this.Manual.Name = "Manual";
            this.Manual.Size = new System.Drawing.Size(150, 50);
            this.Manual.TabIndex = 0;
            this.Manual.Text = "使用说明";
            this.Manual.UseVisualStyleBackColor = true;
            this.Manual.Click += new System.EventHandler(this.Manual_Click);
            // 
            // OnlineVideo
            // 
            this.OnlineVideo.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OnlineVideo.Location = new System.Drawing.Point(400, 300);
            this.OnlineVideo.Name = "OnlineVideo";
            this.OnlineVideo.Size = new System.Drawing.Size(150, 50);
            this.OnlineVideo.TabIndex = 1;
            this.OnlineVideo.Text = "视频教程";
            this.OnlineVideo.UseVisualStyleBackColor = true;
            this.OnlineVideo.Click += new System.EventHandler(this.OnlineVideo_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(239, 87);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(222, 31);
            this.label1.TabIndex = 2;
            this.label1.Text = "当前版本：0.0.0.01";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(241, 150);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(218, 93);
            this.label2.TabIndex = 3;
            this.label2.Text = "何未生\r\n\r\nQQ：1215678765";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // VerInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(674, 429);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OnlineVideo);
            this.Controls.Add(this.Manual);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "VerInfo";
            this.Text = "版本信息";
            this.Load += new System.EventHandler(this.VerInfo_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Manual;
        private System.Windows.Forms.Button OnlineVideo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}