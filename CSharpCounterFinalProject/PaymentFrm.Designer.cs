namespace CSharpCounterFinalProject
{
    partial class PaymentFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PaymentFrm));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtCustomerName = new System.Windows.Forms.TextBox();
            this.txtTotalDiscount = new System.Windows.Forms.TextBox();
            this.txtTotalItem = new System.Windows.Forms.TextBox();
            this.txtTotalAmount = new System.Windows.Forms.TextBox();
            this.comboPaymentMethod = new System.Windows.Forms.ComboBox();
            this.btnFinish = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.dtpCreatedTime = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.txtCart_ID = new System.Windows.Forms.TextBox();
            this.comboStaffName = new System.Windows.Forms.ComboBox();
            this.txtCus_id = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(65, 143);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(133, 21);
            this.label1.TabIndex = 0;
            this.label1.Text = "Tên khách hàng:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(65, 195);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 21);
            this.label2.TabIndex = 1;
            this.label2.Text = "Tên nhân viên:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(65, 253);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 21);
            this.label3.TabIndex = 2;
            this.label3.Text = "Thời gian:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(65, 307);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(130, 21);
            this.label4.TabIndex = 3;
            this.label4.Text = "Tổng sản phẩm:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(65, 369);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(142, 21);
            this.label5.TabIndex = 4;
            this.label5.Text = "Tổng khuyến mãi:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(65, 426);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(86, 21);
            this.label6.TabIndex = 5;
            this.label6.Text = "Tổng tiền:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(65, 479);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(172, 21);
            this.label7.TabIndex = 6;
            this.label7.Text = "Hình thức thanh toán:";
            // 
            // txtCustomerName
            // 
            this.txtCustomerName.Enabled = false;
            this.txtCustomerName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.txtCustomerName.Location = new System.Drawing.Point(266, 135);
            this.txtCustomerName.Margin = new System.Windows.Forms.Padding(4);
            this.txtCustomerName.Name = "txtCustomerName";
            this.txtCustomerName.Size = new System.Drawing.Size(296, 28);
            this.txtCustomerName.TabIndex = 7;
            // 
            // txtTotalDiscount
            // 
            this.txtTotalDiscount.Enabled = false;
            this.txtTotalDiscount.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.txtTotalDiscount.Location = new System.Drawing.Point(266, 361);
            this.txtTotalDiscount.Margin = new System.Windows.Forms.Padding(4);
            this.txtTotalDiscount.Name = "txtTotalDiscount";
            this.txtTotalDiscount.Size = new System.Drawing.Size(296, 28);
            this.txtTotalDiscount.TabIndex = 8;
            // 
            // txtTotalItem
            // 
            this.txtTotalItem.Enabled = false;
            this.txtTotalItem.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.txtTotalItem.Location = new System.Drawing.Point(266, 299);
            this.txtTotalItem.Margin = new System.Windows.Forms.Padding(4);
            this.txtTotalItem.Name = "txtTotalItem";
            this.txtTotalItem.Size = new System.Drawing.Size(296, 28);
            this.txtTotalItem.TabIndex = 9;
            // 
            // txtTotalAmount
            // 
            this.txtTotalAmount.Enabled = false;
            this.txtTotalAmount.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.txtTotalAmount.Location = new System.Drawing.Point(266, 418);
            this.txtTotalAmount.Margin = new System.Windows.Forms.Padding(4);
            this.txtTotalAmount.Name = "txtTotalAmount";
            this.txtTotalAmount.Size = new System.Drawing.Size(296, 28);
            this.txtTotalAmount.TabIndex = 10;
            // 
            // comboPaymentMethod
            // 
            this.comboPaymentMethod.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.comboPaymentMethod.FormattingEnabled = true;
            this.comboPaymentMethod.Items.AddRange(new object[] {
            "Thanh toán tiền mặt",
            "Thanh toán chuyển khoản",
            "Thanh toán quẹt thẻ"});
            this.comboPaymentMethod.Location = new System.Drawing.Point(266, 479);
            this.comboPaymentMethod.Margin = new System.Windows.Forms.Padding(4);
            this.comboPaymentMethod.Name = "comboPaymentMethod";
            this.comboPaymentMethod.Size = new System.Drawing.Size(296, 29);
            this.comboPaymentMethod.TabIndex = 14;
            this.comboPaymentMethod.Validating += new System.ComponentModel.CancelEventHandler(this.comboPaymentMethod_Validating);
            this.comboPaymentMethod.Validated += new System.EventHandler(this.comboPaymentMethod_Validated);
            // 
            // btnFinish
            // 
            this.btnFinish.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnFinish.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.btnFinish.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFinish.Location = new System.Drawing.Point(121, 533);
            this.btnFinish.Margin = new System.Windows.Forms.Padding(4);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(136, 50);
            this.btnFinish.TabIndex = 15;
            this.btnFinish.Text = "Hoàn tất";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Click += new System.EventHandler(this.PaymentClick);
            // 
            // btnCancel
            // 
            this.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCancel.Image = global::CSharpCounterFinalProject.Properties.Resources.remove;
            this.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancel.Location = new System.Drawing.Point(360, 533);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(136, 50);
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "Hủy bỏ";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancelPayment);
            // 
            // dtpCreatedTime
            // 
            this.dtpCreatedTime.CustomFormat = "MM/dd/yyyy";
            this.dtpCreatedTime.Enabled = false;
            this.dtpCreatedTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpCreatedTime.Location = new System.Drawing.Point(266, 246);
            this.dtpCreatedTime.Margin = new System.Windows.Forms.Padding(4);
            this.dtpCreatedTime.Name = "dtpCreatedTime";
            this.dtpCreatedTime.Size = new System.Drawing.Size(296, 28);
            this.dtpCreatedTime.TabIndex = 17;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(65, 86);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(111, 21);
            this.label8.TabIndex = 18;
            this.label8.Text = "Mã giỏ hàng: ";
            // 
            // txtCart_ID
            // 
            this.txtCart_ID.Enabled = false;
            this.txtCart_ID.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.txtCart_ID.Location = new System.Drawing.Point(266, 79);
            this.txtCart_ID.Margin = new System.Windows.Forms.Padding(4);
            this.txtCart_ID.Name = "txtCart_ID";
            this.txtCart_ID.Size = new System.Drawing.Size(296, 28);
            this.txtCart_ID.TabIndex = 19;
            // 
            // comboStaffName
            // 
            this.comboStaffName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.comboStaffName.FormattingEnabled = true;
            this.comboStaffName.Items.AddRange(new object[] {
            "Vũ Nguyễn Trường Giang",
            "Trịnh Viết hân",
            "Nguyễn Văn Hải",
            "Nguyễn Tiến Toàn"});
            this.comboStaffName.Location = new System.Drawing.Point(266, 186);
            this.comboStaffName.Margin = new System.Windows.Forms.Padding(4);
            this.comboStaffName.Name = "comboStaffName";
            this.comboStaffName.Size = new System.Drawing.Size(296, 29);
            this.comboStaffName.TabIndex = 20;
            this.comboStaffName.Validating += new System.ComponentModel.CancelEventHandler(this.comboStaffName_Validating);
            this.comboStaffName.Validated += new System.EventHandler(this.comboStaffName_Validated);
            // 
            // txtCus_id
            // 
            this.txtCus_id.Enabled = false;
            this.txtCus_id.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.txtCus_id.Location = new System.Drawing.Point(266, 33);
            this.txtCus_id.Margin = new System.Windows.Forms.Padding(4);
            this.txtCus_id.Name = "txtCus_id";
            this.txtCus_id.Size = new System.Drawing.Size(296, 28);
            this.txtCus_id.TabIndex = 22;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(65, 40);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(132, 21);
            this.label9.TabIndex = 21;
            this.label9.Text = "Mã khách hàng: ";
            // 
            // PaymentFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.ClientSize = new System.Drawing.Size(625, 596);
            this.Controls.Add(this.txtCus_id);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.comboStaffName);
            this.Controls.Add(this.txtCart_ID);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.dtpCreatedTime);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.comboPaymentMethod);
            this.Controls.Add(this.txtTotalAmount);
            this.Controls.Add(this.txtTotalItem);
            this.Controls.Add(this.txtTotalDiscount);
            this.Controls.Add(this.txtCustomerName);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PaymentFrm";
            this.Text = "Thanh Toán";
            this.Load += new System.EventHandler(this.PaymentFrm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtCustomerName;
        private System.Windows.Forms.TextBox txtTotalDiscount;
        private System.Windows.Forms.TextBox txtTotalItem;
        private System.Windows.Forms.TextBox txtTotalAmount;
        private System.Windows.Forms.ComboBox comboPaymentMethod;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.DateTimePicker dtpCreatedTime;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtCart_ID;
        private System.Windows.Forms.ComboBox comboStaffName;
		private System.Windows.Forms.TextBox txtCus_id;
		private System.Windows.Forms.Label label9;
	}
}