namespace convertPointToPoint
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.excelLocation = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.convertstate = new System.Windows.Forms.Label();
            this.conn = new System.Windows.Forms.Button();
            this.connAccess = new System.Windows.Forms.Button();
            this.localsearch = new System.Windows.Forms.Button();
            this.runstate = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // excelLocation
            // 
            this.excelLocation.AutoSize = true;
            this.excelLocation.Location = new System.Drawing.Point(120, 31);
            this.excelLocation.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.excelLocation.Name = "excelLocation";
            this.excelLocation.Size = new System.Drawing.Size(0, 12);
            this.excelLocation.TabIndex = 1;
            // 
            // convertstate
            // 
            this.convertstate.AutoSize = true;
            this.convertstate.Location = new System.Drawing.Point(30, 218);
            this.convertstate.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.convertstate.Name = "convertstate";
            this.convertstate.Size = new System.Drawing.Size(0, 12);
            this.convertstate.TabIndex = 6;
            // 
            // conn
            // 
            this.conn.Location = new System.Drawing.Point(58, 22);
            this.conn.Margin = new System.Windows.Forms.Padding(2);
            this.conn.Name = "conn";
            this.conn.Size = new System.Drawing.Size(86, 30);
            this.conn.TabIndex = 7;
            this.conn.Text = "连接oracle";
            this.conn.UseVisualStyleBackColor = true;
            this.conn.Click += new System.EventHandler(this.conn_Click);
            // 
            // connAccess
            // 
            this.connAccess.Location = new System.Drawing.Point(58, 56);
            this.connAccess.Margin = new System.Windows.Forms.Padding(2);
            this.connAccess.Name = "connAccess";
            this.connAccess.Size = new System.Drawing.Size(86, 38);
            this.connAccess.TabIndex = 8;
            this.connAccess.Text = "连接access";
            this.connAccess.UseVisualStyleBackColor = true;
            this.connAccess.Click += new System.EventHandler(this.connAccess_Click);
            // 
            // localsearch
            // 
            this.localsearch.Location = new System.Drawing.Point(58, 138);
            this.localsearch.Margin = new System.Windows.Forms.Padding(2);
            this.localsearch.Name = "localsearch";
            this.localsearch.Size = new System.Drawing.Size(86, 36);
            this.localsearch.TabIndex = 11;
            this.localsearch.Text = "内网匹配";
            this.localsearch.UseVisualStyleBackColor = true;
            this.localsearch.Click += new System.EventHandler(this.localsearch_Click);
            // 
            // runstate
            // 
            this.runstate.AutoSize = true;
            this.runstate.Location = new System.Drawing.Point(61, 110);
            this.runstate.Name = "runstate";
            this.runstate.Size = new System.Drawing.Size(0, 12);
            this.runstate.TabIndex = 12;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(211, 185);
            this.Controls.Add(this.runstate);
            this.Controls.Add(this.localsearch);
            this.Controls.Add(this.connAccess);
            this.Controls.Add(this.conn);
            this.Controls.Add(this.convertstate);
            this.Controls.Add(this.excelLocation);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "地名地址匹配";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label excelLocation;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label convertstate;
        private System.Windows.Forms.Button conn;
        private System.Windows.Forms.Button connAccess;
        private System.Windows.Forms.Button localsearch;
        private System.Windows.Forms.Label runstate;
    }
}

