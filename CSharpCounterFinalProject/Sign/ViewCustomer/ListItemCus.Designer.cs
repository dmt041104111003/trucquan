namespace CSharpCounterFinalProject.ViewCustomer
{
    partial class ListItemCus
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ListItemCus));
            this.btbPaymentCus = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.labelName = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.labelCus_IDCus = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btbPaymentCus
            // 
            this.btbPaymentCus.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btbPaymentCus.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btbPaymentCus.ForeColor = System.Drawing.Color.White;
            this.btbPaymentCus.Location = new System.Drawing.Point(78, 387);
            this.btbPaymentCus.Name = "btbPaymentCus";
            this.btbPaymentCus.Size = new System.Drawing.Size(226, 48);
            this.btbPaymentCus.TabIndex = 0;
            this.btbPaymentCus.Text = "Tiến hành thanh toán";
            this.btbPaymentCus.UseVisualStyleBackColor = false;
            this.btbPaymentCus.Click += new System.EventHandler(this.btbPaymentCus_Click_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(181, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 33);
            this.label1.TabIndex = 1;
            this.label1.Text = "Giỏ hàng";
            // 
            // tableLayoutPanel
            // 
            this.tableLayoutPanel.AutoScroll = true;
            this.tableLayoutPanel.AutoSize = true;
            this.tableLayoutPanel.ColumnCount = 1;
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel.Location = new System.Drawing.Point(438, 80);
            this.tableLayoutPanel.Name = "tableLayoutPanel";
            this.tableLayoutPanel.RowCount = 1;
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel.Size = new System.Drawing.Size(387, 355);
            this.tableLayoutPanel.TabIndex = 2;
            this.tableLayoutPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.tableLayoutPanel1_Paint);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(37, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 21);
            this.label2.TabIndex = 3;
            this.label2.Text = "Xin chào:";
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Location = new System.Drawing.Point(131, 80);
            this.labelName.Name = "labelName";
            this.labelName.Size = new System.Drawing.Size(53, 21);
            this.labelName.TabIndex = 4;
            this.labelName.Text = "label3";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(37, 112);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(126, 21);
            this.label3.TabIndex = 5;
            this.label3.Text = "Mã KH của bạn:";
            // 
            // labelCus_IDCus
            // 
            this.labelCus_IDCus.AutoSize = true;
            this.labelCus_IDCus.Location = new System.Drawing.Point(186, 112);
            this.labelCus_IDCus.Name = "labelCus_IDCus";
            this.labelCus_IDCus.Size = new System.Drawing.Size(54, 21);
            this.labelCus_IDCus.TabIndex = 6;
            this.labelCus_IDCus.Text = "label4";
            // 
            // ListItemCus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.ClientSize = new System.Drawing.Size(852, 475);
            this.Controls.Add(this.labelCus_IDCus);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.labelName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tableLayoutPanel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btbPaymentCus);
            this.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ListItemCus";
            this.Text = "Giỏ hàng";
            this.Load += new System.EventHandler(this.ListItemCus_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btbPaymentCus;
        private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label labelName;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label labelCus_IDCus;
        public System.Windows.Forms.TableLayoutPanel tableLayoutPanel;
    }
}