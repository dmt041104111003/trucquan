namespace CSharpCounterFinalProject
{
    partial class AddEditItemFrm
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddEditItemFrm));
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtId = new System.Windows.Forms.TextBox();
			this.txtItemName = new System.Windows.Forms.TextBox();
			this.txtBrand = new System.Windows.Forms.TextBox();
			this.numericQuantity = new System.Windows.Forms.NumericUpDown();
			this.comboItemType = new System.Windows.Forms.ComboBox();
			this.datePickerReleaseDate = new System.Windows.Forms.DateTimePicker();
			this.numericPrice = new System.Windows.Forms.NumericUpDown();
			this.comboDiscount = new System.Windows.Forms.ComboBox();
			this.btnChooseImage = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnAddNew = new System.Windows.Forms.Button();
			this.pictureImage = new System.Windows.Forms.PictureBox();
			((System.ComponentModel.ISupportInitialize)(this.numericQuantity)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.numericPrice)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureImage)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(66, 29);
			this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(95, 18);
			this.label1.TabIndex = 2;
			this.label1.Text = "Mã mặt hàng";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(66, 80);
			this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(98, 18);
			this.label2.TabIndex = 3;
			this.label2.Text = "Tên mặt hàng";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(66, 121);
			this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(101, 18);
			this.label3.TabIndex = 4;
			this.label3.Text = "Loại mặt hàng";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(66, 168);
			this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(69, 18);
			this.label4.TabIndex = 5;
			this.label4.Text = "Số lưỡng";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(66, 214);
			this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(102, 18);
			this.label5.TabIndex = 6;
			this.label5.Text = "Hãng sản xuất";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(66, 259);
			this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(99, 18);
			this.label6.TabIndex = 7;
			this.label6.Text = "Ngày sản xuất";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(66, 310);
			this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(58, 18);
			this.label7.TabIndex = 8;
			this.label7.Text = "Giá bán";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(66, 354);
			this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(87, 18);
			this.label8.TabIndex = 9;
			this.label8.Text = "Khuyến mãi";
			// 
			// txtId
			// 
			this.txtId.Enabled = false;
			this.txtId.Location = new System.Drawing.Point(199, 29);
			this.txtId.Margin = new System.Windows.Forms.Padding(4);
			this.txtId.Name = "txtId";
			this.txtId.Size = new System.Drawing.Size(239, 26);
			this.txtId.TabIndex = 10;
			this.txtId.Text = "0";
			// 
			// txtItemName
			// 
			this.txtItemName.Location = new System.Drawing.Point(199, 80);
			this.txtItemName.Margin = new System.Windows.Forms.Padding(4);
			this.txtItemName.Name = "txtItemName";
			this.txtItemName.Size = new System.Drawing.Size(239, 26);
			this.txtItemName.TabIndex = 11;
			// 
			// txtBrand
			// 
			this.txtBrand.Location = new System.Drawing.Point(199, 214);
			this.txtBrand.Margin = new System.Windows.Forms.Padding(4);
			this.txtBrand.Name = "txtBrand";
			this.txtBrand.Size = new System.Drawing.Size(239, 26);
			this.txtBrand.TabIndex = 12;
			// 
			// numericQuantity
			// 
			this.numericQuantity.Location = new System.Drawing.Point(199, 168);
			this.numericQuantity.Margin = new System.Windows.Forms.Padding(4);
			this.numericQuantity.Maximum = new decimal(new int[] {
            999999,
            0,
            0,
            0});
			this.numericQuantity.Name = "numericQuantity";
			this.numericQuantity.Size = new System.Drawing.Size(239, 26);
			this.numericQuantity.TabIndex = 13;
			// 
			// comboItemType
			// 
			this.comboItemType.FormattingEnabled = true;
			this.comboItemType.Items.AddRange(new object[] {
            "Đồ ăn",
            "Đồ uống",
            "Gia vị",
            "Đồ trang trí"});
			this.comboItemType.Location = new System.Drawing.Point(199, 121);
			this.comboItemType.Margin = new System.Windows.Forms.Padding(4);
			this.comboItemType.Name = "comboItemType";
			this.comboItemType.Size = new System.Drawing.Size(239, 26);
			this.comboItemType.TabIndex = 14;
			this.comboItemType.SelectedIndexChanged += new System.EventHandler(this.comboItemType_SelectedIndexChanged);
			// 
			// datePickerReleaseDate
			// 
			this.datePickerReleaseDate.CustomFormat = "MM/dd/yyyy";
			this.datePickerReleaseDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.datePickerReleaseDate.Location = new System.Drawing.Point(199, 259);
			this.datePickerReleaseDate.Margin = new System.Windows.Forms.Padding(4);
			this.datePickerReleaseDate.Name = "datePickerReleaseDate";
			this.datePickerReleaseDate.Size = new System.Drawing.Size(239, 26);
			this.datePickerReleaseDate.TabIndex = 15;
			// 
			// numericPrice
			// 
			this.numericPrice.Location = new System.Drawing.Point(199, 310);
			this.numericPrice.Margin = new System.Windows.Forms.Padding(4);
			this.numericPrice.Maximum = new decimal(new int[] {
            99999999,
            0,
            0,
            0});
			this.numericPrice.Name = "numericPrice";
			this.numericPrice.Size = new System.Drawing.Size(239, 26);
			this.numericPrice.TabIndex = 16;
			// 
			// comboDiscount
			// 
			this.comboDiscount.FormattingEnabled = true;
			this.comboDiscount.Items.AddRange(new object[] {
            "Summer Sale",
            "Back to School",
            "Holiday Clearance",
            "Spring Discount",
            "Black Friday"});
			this.comboDiscount.Location = new System.Drawing.Point(199, 354);
			this.comboDiscount.Margin = new System.Windows.Forms.Padding(4);
			this.comboDiscount.Name = "comboDiscount";
			this.comboDiscount.Size = new System.Drawing.Size(239, 26);
			this.comboDiscount.TabIndex = 17;
			// 
			// btnChooseImage
			// 
			this.btnChooseImage.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
			this.btnChooseImage.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnChooseImage.ForeColor = System.Drawing.Color.White;
			this.btnChooseImage.Location = new System.Drawing.Point(530, 343);
			this.btnChooseImage.Margin = new System.Windows.Forms.Padding(4);
			this.btnChooseImage.Name = "btnChooseImage";
			this.btnChooseImage.Size = new System.Drawing.Size(109, 38);
			this.btnChooseImage.TabIndex = 18;
			this.btnChooseImage.Text = "Chọn ảnh";
			this.btnChooseImage.UseVisualStyleBackColor = false;
			this.btnChooseImage.Click += new System.EventHandler(this.BtnChooseImageClick);
			// 
			// btnCancel
			// 
			this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
			this.btnCancel.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnCancel.ForeColor = System.Drawing.Color.White;
			this.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancel.Location = new System.Drawing.Point(516, 418);
			this.btnCancel.Margin = new System.Windows.Forms.Padding(4);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(139, 38);
			this.btnCancel.TabIndex = 1;
			this.btnCancel.Text = "Hủy Bỏ";
			this.btnCancel.UseVisualStyleBackColor = false;
			this.btnCancel.Click += new System.EventHandler(this.BtnCancelClick);
			// 
			// btnAddNew
			// 
			this.btnAddNew.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
			this.btnAddNew.Cursor = System.Windows.Forms.Cursors.Hand;
			this.btnAddNew.ForeColor = System.Drawing.Color.White;
			this.btnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAddNew.Location = new System.Drawing.Point(228, 416);
			this.btnAddNew.Margin = new System.Windows.Forms.Padding(4);
			this.btnAddNew.Name = "btnAddNew";
			this.btnAddNew.Size = new System.Drawing.Size(150, 40);
			this.btnAddNew.TabIndex = 0;
			this.btnAddNew.Text = "Thêm Mới";
			this.btnAddNew.UseVisualStyleBackColor = false;
			this.btnAddNew.Click += new System.EventHandler(this.BtnAddNewItemClick);
			// 
			// pictureImage
			// 
			this.pictureImage.Image = global::CSharpCounterFinalProject.Properties.Resources.store__1_;
			this.pictureImage.Location = new System.Drawing.Point(530, 29);
			this.pictureImage.Margin = new System.Windows.Forms.Padding(4);
			this.pictureImage.Name = "pictureImage";
			this.pictureImage.Size = new System.Drawing.Size(282, 281);
			this.pictureImage.TabIndex = 19;
			this.pictureImage.TabStop = false;
			// 
			// AddEditItemFrm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
			this.ClientSize = new System.Drawing.Size(879, 482);
			this.Controls.Add(this.pictureImage);
			this.Controls.Add(this.btnChooseImage);
			this.Controls.Add(this.comboDiscount);
			this.Controls.Add(this.numericPrice);
			this.Controls.Add(this.datePickerReleaseDate);
			this.Controls.Add(this.comboItemType);
			this.Controls.Add(this.numericQuantity);
			this.Controls.Add(this.txtBrand);
			this.Controls.Add(this.txtItemName);
			this.Controls.Add(this.txtId);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnAddNew);
			this.Font = new System.Drawing.Font("Constantia", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Margin = new System.Windows.Forms.Padding(4);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "AddEditItemFrm";
			this.Text = "THÊM MỚI MẶT HÀNG";
			this.Load += new System.EventHandler(this.AddEditItemFrm_Load);
			((System.ComponentModel.ISupportInitialize)(this.numericQuantity)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.numericPrice)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureImage)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAddNew;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtId;
        private System.Windows.Forms.TextBox txtItemName;
        private System.Windows.Forms.TextBox txtBrand;
        private System.Windows.Forms.NumericUpDown numericQuantity;
        private System.Windows.Forms.ComboBox comboItemType;
        private System.Windows.Forms.DateTimePicker datePickerReleaseDate;
        private System.Windows.Forms.NumericUpDown numericPrice;
        private System.Windows.Forms.ComboBox comboDiscount;
        private System.Windows.Forms.Button btnChooseImage;
        private System.Windows.Forms.PictureBox pictureImage;
    }
}