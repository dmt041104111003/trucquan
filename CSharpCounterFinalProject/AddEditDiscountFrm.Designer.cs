namespace CSharpCounterFinalProject
{
    partial class AddEditDiscountFrm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddEditDiscountFrm));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btnAddNew = new System.Windows.Forms.Button();
            this.txtDiscountId = new System.Windows.Forms.TextBox();
            this.txtDiscountName = new System.Windows.Forms.TextBox();
            this.dtPickerStart = new System.Windows.Forms.DateTimePicker();
            this.dtPickerEnd = new System.Windows.Forms.DateTimePicker();
            this.comboDiscountType = new System.Windows.Forms.ComboBox();
            this.numericDiscountPercent = new System.Windows.Forms.NumericUpDown();
            this.numericDiscountAmount = new System.Windows.Forms.NumericUpDown();
            this.btnCancel = new System.Windows.Forms.Button();
            this.numericMinPriceToUseDiscount = new System.Windows.Forms.NumericUpDown();
            this.label8 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numericDiscountPercent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericDiscountAmount)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericMinPriceToUseDiscount)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(54, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Mã khuyễn mãi:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(54, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Tên khuyến mãi:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(54, 130);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 16);
            this.label3.TabIndex = 2;
            this.label3.Text = "Bắt đầu:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(54, 176);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 16);
            this.label4.TabIndex = 3;
            this.label4.Text = "Kết thúc:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(54, 225);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 16);
            this.label5.TabIndex = 4;
            this.label5.Text = "Loại khuyến mãi:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(54, 269);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 16);
            this.label6.TabIndex = 5;
            this.label6.Text = "% giá giảm:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(54, 319);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(115, 16);
            this.label7.TabIndex = 6;
            this.label7.Text = "Giá trị khuyến mãi:";
            // 
            // btnAddNew
            // 
            this.btnAddNew.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnAddNew.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAddNew.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddNew.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.btnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddNew.Location = new System.Drawing.Point(79, 413);
            this.btnAddNew.Name = "btnAddNew";
            this.btnAddNew.Size = new System.Drawing.Size(164, 46);
            this.btnAddNew.TabIndex = 7;
            this.btnAddNew.Text = "Thêm mới";
            this.btnAddNew.UseVisualStyleBackColor = false;
            this.btnAddNew.Click += new System.EventHandler(this.BtnAddNewDiscountClick);
            // 
            // txtDiscountId
            // 
            this.txtDiscountId.Enabled = false;
            this.txtDiscountId.Location = new System.Drawing.Point(200, 33);
            this.txtDiscountId.Name = "txtDiscountId";
            this.txtDiscountId.Size = new System.Drawing.Size(266, 22);
            this.txtDiscountId.TabIndex = 9;
            this.txtDiscountId.Text = "Code";
            // 
            // txtDiscountName
            // 
            this.txtDiscountName.Location = new System.Drawing.Point(200, 76);
            this.txtDiscountName.Name = "txtDiscountName";
            this.txtDiscountName.Size = new System.Drawing.Size(266, 22);
            this.txtDiscountName.TabIndex = 10;
            // 
            // dtPickerStart
            // 
            this.dtPickerStart.CustomFormat = "MM/dd/yyyy";
            this.dtPickerStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtPickerStart.Location = new System.Drawing.Point(200, 123);
            this.dtPickerStart.Name = "dtPickerStart";
            this.dtPickerStart.Size = new System.Drawing.Size(266, 22);
            this.dtPickerStart.TabIndex = 11;
            // 
            // dtPickerEnd
            // 
            this.dtPickerEnd.CustomFormat = "MM/dd/yyyy";
            this.dtPickerEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtPickerEnd.Location = new System.Drawing.Point(200, 170);
            this.dtPickerEnd.Name = "dtPickerEnd";
            this.dtPickerEnd.Size = new System.Drawing.Size(266, 22);
            this.dtPickerEnd.TabIndex = 12;
            // 
            // comboDiscountType
            // 
            this.comboDiscountType.FormattingEnabled = true;
            this.comboDiscountType.Items.AddRange(new object[] {
            "Không áp dụng",
            "Giảm giá trực tiếp",
            "Khuyến mãi theo % giá bán"});
            this.comboDiscountType.Location = new System.Drawing.Point(200, 216);
            this.comboDiscountType.Name = "comboDiscountType";
            this.comboDiscountType.Size = new System.Drawing.Size(266, 24);
            this.comboDiscountType.TabIndex = 13;
            // 
            // numericDiscountPercent
            // 
            this.numericDiscountPercent.Location = new System.Drawing.Point(200, 262);
            this.numericDiscountPercent.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericDiscountPercent.Name = "numericDiscountPercent";
            this.numericDiscountPercent.Size = new System.Drawing.Size(266, 22);
            this.numericDiscountPercent.TabIndex = 14;
            // 
            // numericDiscountAmount
            // 
            this.numericDiscountAmount.Location = new System.Drawing.Point(200, 313);
            this.numericDiscountAmount.Maximum = new decimal(new int[] {
            99999999,
            0,
            0,
            0});
            this.numericDiscountAmount.Name = "numericDiscountAmount";
            this.numericDiscountAmount.Size = new System.Drawing.Size(266, 22);
            this.numericDiscountAmount.TabIndex = 15;
            this.numericDiscountAmount.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCancel.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Image = global::CSharpCounterFinalProject.Properties.Resources.remove;
            this.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancel.Location = new System.Drawing.Point(286, 413);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(164, 46);
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "Hủy bỏ";
            this.btnCancel.UseVisualStyleBackColor = false;
            // 
            // numericMinPriceToUseDiscount
            // 
            this.numericMinPriceToUseDiscount.Location = new System.Drawing.Point(200, 354);
            this.numericMinPriceToUseDiscount.Maximum = new decimal(new int[] {
            99999999,
            0,
            0,
            0});
            this.numericMinPriceToUseDiscount.Name = "numericMinPriceToUseDiscount";
            this.numericMinPriceToUseDiscount.Size = new System.Drawing.Size(266, 22);
            this.numericMinPriceToUseDiscount.TabIndex = 18;
            this.numericMinPriceToUseDiscount.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(54, 360);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(131, 16);
            this.label8.TabIndex = 17;
            this.label8.Text = "Giá tối thiểu để dùng:";
            // 
            // AddEditDiscountFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.ClientSize = new System.Drawing.Size(532, 487);
            this.Controls.Add(this.numericMinPriceToUseDiscount);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.numericDiscountAmount);
            this.Controls.Add(this.numericDiscountPercent);
            this.Controls.Add(this.comboDiscountType);
            this.Controls.Add(this.dtPickerEnd);
            this.Controls.Add(this.dtPickerStart);
            this.Controls.Add(this.txtDiscountName);
            this.Controls.Add(this.txtDiscountId);
            this.Controls.Add(this.btnAddNew);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddEditDiscountFrm";
            this.Text = "THÊM MỚI CHƯƠNG TRÌNH KHUYẾN MÃI";
            this.Load += new System.EventHandler(this.AddEditDiscountFrm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numericDiscountPercent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericDiscountAmount)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericMinPriceToUseDiscount)).EndInit();
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
        private System.Windows.Forms.Button btnAddNew;
        private System.Windows.Forms.TextBox txtDiscountId;
        private System.Windows.Forms.TextBox txtDiscountName;
        private System.Windows.Forms.DateTimePicker dtPickerStart;
        private System.Windows.Forms.DateTimePicker dtPickerEnd;
        private System.Windows.Forms.ComboBox comboDiscountType;
        private System.Windows.Forms.NumericUpDown numericDiscountPercent;
        private System.Windows.Forms.NumericUpDown numericDiscountAmount;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.NumericUpDown numericMinPriceToUseDiscount;
        private System.Windows.Forms.Label label8;
    }
}