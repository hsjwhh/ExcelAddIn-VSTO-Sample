namespace ExcelAddIn_VSTO_Sample
{
    partial class MyUserControl
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

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btSubmit = new System.Windows.Forms.Button();
            this.cbAutoR = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btSubmit
            // 
            this.btSubmit.Location = new System.Drawing.Point(18, 44);
            this.btSubmit.Name = "btSubmit";
            this.btSubmit.Size = new System.Drawing.Size(75, 23);
            this.btSubmit.TabIndex = 0;
            this.btSubmit.Text = "提交";
            this.btSubmit.UseVisualStyleBackColor = true;
            this.btSubmit.Click += new System.EventHandler(this.btSubmit_Click);
            // 
            // cbAutoR
            // 
            this.cbAutoR.AutoSize = true;
            this.cbAutoR.Location = new System.Drawing.Point(18, 22);
            this.cbAutoR.Name = "cbAutoR";
            this.cbAutoR.Size = new System.Drawing.Size(72, 16);
            this.cbAutoR.TabIndex = 1;
            this.cbAutoR.Text = "自动刷新";
            this.cbAutoR.UseVisualStyleBackColor = true;
            // 
            // MyUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbAutoR);
            this.Controls.Add(this.btSubmit);
            this.Name = "MyUserControl";
            this.Size = new System.Drawing.Size(150, 95);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btSubmit;
        private System.Windows.Forms.CheckBox cbAutoR;
    }
}
