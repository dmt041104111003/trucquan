namespace CSharpCounterFinalProject.ViewCustomer
{
    partial class DetailItemCus
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DetailItemCus));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtItemIDDetail = new System.Windows.Forms.TextBox();
            this.txtItemNameDetail = new System.Windows.Forms.TextBox();
            this.txtPriceDetail = new System.Windows.Forms.TextBox();
            this.nudQuantityDetail = new System.Windows.Forms.NumericUpDown();
            this.btnAddDetail = new System.Windows.Forms.Button();
            this.btnCancelDetail = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.nudQuantityDetail)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(258, 92);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 21);
            this.label1.TabIndex = 1;
            this.label1.Text = "Tên sản phẩm: ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(258, 138);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 21);
            this.label2.TabIndex = 2;
            this.label2.Text = "Giá bán:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(258, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(113, 21);
            this.label3.TabIndex = 3;
            this.label3.Text = "Mã sản phẩm:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(258, 180);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(119, 21);
            this.label4.TabIndex = 4;
            this.label4.Text = "Số lượng mua:";
            // 
            // txtItemIDDetail
            // 
            this.txtItemIDDetail.Enabled = false;
            this.txtItemIDDetail.Location = new System.Drawing.Point(427, 42);
            this.txtItemIDDetail.Name = "txtItemIDDetail";
            this.txtItemIDDetail.Size = new System.Drawing.Size(227, 28);
            this.txtItemIDDetail.TabIndex = 5;
            // 
            // txtItemNameDetail
            // 
            this.txtItemNameDetail.Enabled = false;
            this.txtItemNameDetail.Location = new System.Drawing.Point(427, 85);
            this.txtItemNameDetail.Name = "txtItemNameDetail";
            this.txtItemNameDetail.Size = new System.Drawing.Size(227, 28);
            this.txtItemNameDetail.TabIndex = 6;
            // 
            // txtPriceDetail
            // 
            this.txtPriceDetail.Enabled = false;
            this.txtPriceDetail.Location = new System.Drawing.Point(427, 131);
            this.txtPriceDetail.Name = "txtPriceDetail";
            this.txtPriceDetail.Size = new System.Drawing.Size(227, 28);
            this.txtPriceDetail.TabIndex = 7;
            // 
            // nudQuantityDetail
            // 
            this.nudQuantityDetail.Location = new System.Drawing.Point(427, 173);
            this.nudQuantityDetail.Name = "nudQuantityDetail";
            this.nudQuantityDetail.Size = new System.Drawing.Size(228, 28);
            this.nudQuantityDetail.TabIndex = 8;
            this.nudQuantityDetail.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudQuantityDetail.Validating += new System.ComponentModel.CancelEventHandler(this.nudQuantityDetail_Validating);
            // 
            // btnAddDetail
            // 
            this.btnAddDetail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnAddDetail.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAddDetail.FlatAppearance.BorderSize = 2;
            this.btnAddDetail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAddDetail.ForeColor = System.Drawing.Color.White;
            this.btnAddDetail.Location = new System.Drawing.Point(148, 280);
            this.btnAddDetail.Name = "btnAddDetail";
            this.btnAddDetail.Size = new System.Drawing.Size(159, 47);
            this.btnAddDetail.TabIndex = 9;
            this.btnAddDetail.Text = "Thêm giỏ hàng";
            this.btnAddDetail.UseVisualStyleBackColor = false;
            this.btnAddDetail.Click += new System.EventHandler(this.btnAddDetail_Click);
            // 
            // btnCancelDetail
            // 
            this.btnCancelDetail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnCancelDetail.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCancelDetail.FlatAppearance.BorderSize = 2;
            this.btnCancelDetail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelDetail.ForeColor = System.Drawing.Color.White;
            this.btnCancelDetail.Location = new System.Drawing.Point(384, 280);
            this.btnCancelDetail.Name = "btnCancelDetail";
            this.btnCancelDetail.Size = new System.Drawing.Size(159, 47);
            this.btnCancelDetail.TabIndex = 10;
            this.btnCancelDetail.Text = "Thoát";
            this.btnCancelDetail.UseVisualStyleBackColor = false;
            this.btnCancelDetail.Click += new System.EventHandler(this.btnCancelDetail_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(51, 49);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(174, 192);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // DetailItemCus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.ClientSize = new System.Drawing.Size(708, 359);
            this.Controls.Add(this.btnCancelDetail);
            this.Controls.Add(this.btnAddDetail);
            this.Controls.Add(this.nudQuantityDetail);
            this.Controls.Add(this.txtPriceDetail);
            this.Controls.Add(this.txtItemNameDetail);
            this.Controls.Add(this.txtItemIDDetail);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "DetailItemCus";
            this.Text = "Thông tin chi tiết mặt hàng";
            this.Load += new System.EventHandler(this.DetailItemCus_Load);
            ((System.ComponentModel.ISupportInitialize)(this.nudQuantityDetail)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtItemIDDetail;
        private System.Windows.Forms.TextBox txtItemNameDetail;
        private System.Windows.Forms.TextBox txtPriceDetail;
        private System.Windows.Forms.NumericUpDown nudQuantityDetail;
        private System.Windows.Forms.Button btnAddDetail;
        private System.Windows.Forms.Button btnCancelDetail;
    }
}