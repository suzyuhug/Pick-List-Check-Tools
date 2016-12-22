namespace Pick_List_Check_Tools
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.loglabel = new System.Windows.Forms.Label();
            this.PONumLabel = new System.Windows.Forms.Label();
            this.PB = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(590, 117);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(97, 31);
            this.button1.TabIndex = 0;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // loglabel
            // 
            this.loglabel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.loglabel.Location = new System.Drawing.Point(2, 52);
            this.loglabel.Name = "loglabel";
            this.loglabel.Size = new System.Drawing.Size(260, 16);
            this.loglabel.TabIndex = 2;
            this.loglabel.Text = "log";
            this.loglabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // PONumLabel
            // 
            this.PONumLabel.AutoSize = true;
            this.PONumLabel.Location = new System.Drawing.Point(9, 7);
            this.PONumLabel.Name = "PONumLabel";
            this.PONumLabel.Size = new System.Drawing.Size(29, 13);
            this.PONumLabel.TabIndex = 3;
            this.PONumLabel.Text = "PO#";
            // 
            // PB
            // 
            this.PB.Location = new System.Drawing.Point(26, 36);
            this.PB.Name = "PB";
            this.PB.Size = new System.Drawing.Size(214, 8);
            this.PB.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(266, 79);
            this.Controls.Add(this.PB);
            this.Controls.Add(this.PONumLabel);
            this.Controls.Add(this.loglabel);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PO校验工具";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label loglabel;
        private System.Windows.Forms.Label PONumLabel;
        private System.Windows.Forms.ProgressBar PB;
    }
}

