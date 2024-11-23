namespace CSharpCounterFinalProject
{
    partial class AddEditBillFrm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddEditBillFrm));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tblSearchedCustomer = new System.Windows.Forms.DataGridView();
            this.Customer_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FullName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tblCustomerColSelect = new System.Windows.Forms.DataGridViewButtonColumn();
            this.btnSearchCustomer = new System.Windows.Forms.Button();
            this.txtSearchCustomer = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtBillID = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnEditBillItem = new System.Windows.Forms.Button();
            this.numericSelectedQuantity = new System.Windows.Forms.NumericUpDown();
            this.txtSearchItem = new System.Windows.Forms.TextBox();
            this.btnSearchItem = new System.Windows.Forms.Button();
            this.tblSearchedItem = new System.Windows.Forms.DataGridView();
            this.Item_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ItemName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tblItemColSelect = new System.Windows.Forms.DataGridViewButtonColumn();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnCreateCartBill = new System.Windows.Forms.Button();
            this.btnRefreshCart = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.labelCustomer_IDBill = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.labelCustomerNameBill = new System.Windows.Forms.Label();
            this.labelCart_ID = new System.Windows.Forms.Label();
            this.labelTotalAmount = new System.Windows.Forms.Label();
            this.labelTotalDiscount = new System.Windows.Forms.Label();
            this.labelTotalItem = new System.Windows.Forms.Label();
            this.labelCustomerName = new System.Windows.Forms.Label();
            this.tblBillDetail = new System.Windows.Forms.DataGridView();
            this.SelectedItem_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Item_IDBill = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ItemNameBill = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NumberOfSelectedItem = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PriceBill = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PriceAfterDiscount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TotalAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnEditBill = new System.Windows.Forms.DataGridViewButtonColumn();
            this.btnRemoveBill = new System.Windows.Forms.DataGridViewButtonColumn();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnPay = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblSearchedCustomer)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericSelectedQuantity)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblSearchedItem)).BeginInit();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblBillDetail)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tblSearchedCustomer);
            this.groupBox1.Controls.Add(this.btnSearchCustomer);
            this.groupBox1.Controls.Add(this.txtSearchCustomer);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(2, 2);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(462, 234);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Thông tin khách hàng";
            // 
            // tblSearchedCustomer
            // 
            this.tblSearchedCustomer.AllowUserToAddRows = false;
            this.tblSearchedCustomer.AllowUserToDeleteRows = false;
            this.tblSearchedCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tblSearchedCustomer.BackgroundColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.tblSearchedCustomer.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.tblSearchedCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblSearchedCustomer.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Customer_ID,
            this.FullName,
            this.tblCustomerColSelect});
            this.tblSearchedCustomer.Location = new System.Drawing.Point(0, 92);
            this.tblSearchedCustomer.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tblSearchedCustomer.Name = "tblSearchedCustomer";
            this.tblSearchedCustomer.RowHeadersWidth = 51;
            this.tblSearchedCustomer.RowTemplate.Height = 24;
            this.tblSearchedCustomer.Size = new System.Drawing.Size(462, 141);
            this.tblSearchedCustomer.TabIndex = 5;
            this.tblSearchedCustomer.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TblCustomerCellClick);
            // 
            // Customer_ID
            // 
            this.Customer_ID.HeaderText = "Mã khách hàng";
            this.Customer_ID.MinimumWidth = 6;
            this.Customer_ID.Name = "Customer_ID";
            this.Customer_ID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // FullName
            // 
            this.FullName.HeaderText = "Họ và tên";
            this.FullName.MinimumWidth = 6;
            this.FullName.Name = "FullName";
            this.FullName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // tblCustomerColSelect
            // 
            this.tblCustomerColSelect.HeaderText = "Chọn";
            this.tblCustomerColSelect.MinimumWidth = 6;
            this.tblCustomerColSelect.Name = "tblCustomerColSelect";
            this.tblCustomerColSelect.Text = "Chọn";
            this.tblCustomerColSelect.UseColumnTextForButtonValue = true;
            // 
            // btnSearchCustomer
            // 
            this.btnSearchCustomer.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnSearchCustomer.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSearchCustomer.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearchCustomer.ForeColor = System.Drawing.Color.White;
            this.btnSearchCustomer.Image = global::CSharpCounterFinalProject.Properties.Resources.search;
            this.btnSearchCustomer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSearchCustomer.Location = new System.Drawing.Point(348, 38);
            this.btnSearchCustomer.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnSearchCustomer.Name = "btnSearchCustomer";
            this.btnSearchCustomer.Size = new System.Drawing.Size(90, 32);
            this.btnSearchCustomer.TabIndex = 2;
            this.btnSearchCustomer.Text = "Tìm";
            this.btnSearchCustomer.UseVisualStyleBackColor = false;
            this.btnSearchCustomer.Click += new System.EventHandler(this.BtnSearchCustomerBillClick);
            // 
            // txtSearchCustomer
            // 
            this.txtSearchCustomer.Location = new System.Drawing.Point(116, 42);
            this.txtSearchCustomer.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtSearchCustomer.Name = "txtSearchCustomer";
            this.txtSearchCustomer.Size = new System.Drawing.Size(205, 23);
            this.txtSearchCustomer.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 49);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Tên khách hàng:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtBillID);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.btnEditBillItem);
            this.groupBox2.Controls.Add(this.numericSelectedQuantity);
            this.groupBox2.Controls.Add(this.txtSearchItem);
            this.groupBox2.Controls.Add(this.btnSearchItem);
            this.groupBox2.Controls.Add(this.tblSearchedItem);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Location = new System.Drawing.Point(470, 2);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Size = new System.Drawing.Size(468, 234);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Thông tin mặt hàng";
            // 
            // txtBillID
            // 
            this.txtBillID.Enabled = false;
            this.txtBillID.Location = new System.Drawing.Point(239, 62);
            this.txtBillID.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtBillID.Name = "txtBillID";
            this.txtBillID.Size = new System.Drawing.Size(77, 23);
            this.txtBillID.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(188, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 15);
            this.label5.TabIndex = 8;
            this.label5.Text = "Mã HĐ:";
            // 
            // btnEditBillItem
            // 
            this.btnEditBillItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnEditBillItem.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnEditBillItem.Enabled = false;
            this.btnEditBillItem.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEditBillItem.ForeColor = System.Drawing.Color.White;
            this.btnEditBillItem.Image = global::CSharpCounterFinalProject.Properties.Resources.loop;
            this.btnEditBillItem.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEditBillItem.Location = new System.Drawing.Point(341, 58);
            this.btnEditBillItem.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnEditBillItem.Name = "btnEditBillItem";
            this.btnEditBillItem.Size = new System.Drawing.Size(108, 26);
            this.btnEditBillItem.TabIndex = 7;
            this.btnEditBillItem.Text = "Cập nhật";
            this.btnEditBillItem.UseVisualStyleBackColor = false;
            this.btnEditBillItem.Click += new System.EventHandler(this.BtnEditBillClick);
            // 
            // numericSelectedQuantity
            // 
            this.numericSelectedQuantity.Location = new System.Drawing.Point(110, 62);
            this.numericSelectedQuantity.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.numericSelectedQuantity.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.numericSelectedQuantity.Name = "numericSelectedQuantity";
            this.numericSelectedQuantity.Size = new System.Drawing.Size(52, 23);
            this.numericSelectedQuantity.TabIndex = 6;
            // 
            // txtSearchItem
            // 
            this.txtSearchItem.Location = new System.Drawing.Point(110, 29);
            this.txtSearchItem.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtSearchItem.Name = "txtSearchItem";
            this.txtSearchItem.Size = new System.Drawing.Size(205, 23);
            this.txtSearchItem.TabIndex = 4;
            // 
            // btnSearchItem
            // 
            this.btnSearchItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnSearchItem.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSearchItem.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearchItem.ForeColor = System.Drawing.Color.White;
            this.btnSearchItem.Image = global::CSharpCounterFinalProject.Properties.Resources.search;
            this.btnSearchItem.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSearchItem.Location = new System.Drawing.Point(341, 21);
            this.btnSearchItem.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnSearchItem.Name = "btnSearchItem";
            this.btnSearchItem.Size = new System.Drawing.Size(108, 29);
            this.btnSearchItem.TabIndex = 5;
            this.btnSearchItem.Text = "Tìm";
            this.btnSearchItem.UseVisualStyleBackColor = false;
            this.btnSearchItem.Click += new System.EventHandler(this.BtnSearchItemBillClick);
            // 
            // tblSearchedItem
            // 
            this.tblSearchedItem.AllowUserToAddRows = false;
            this.tblSearchedItem.AllowUserToDeleteRows = false;
            this.tblSearchedItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tblSearchedItem.BackgroundColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.tblSearchedItem.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.tblSearchedItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblSearchedItem.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Item_ID,
            this.ItemName,
            this.Quantity,
            this.tblItemColSelect});
            this.tblSearchedItem.Location = new System.Drawing.Point(0, 92);
            this.tblSearchedItem.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tblSearchedItem.Name = "tblSearchedItem";
            this.tblSearchedItem.RowHeadersWidth = 51;
            this.tblSearchedItem.RowTemplate.Height = 24;
            this.tblSearchedItem.Size = new System.Drawing.Size(467, 141);
            this.tblSearchedItem.TabIndex = 6;
            this.tblSearchedItem.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TblItemCellClick);
            // 
            // Item_ID
            // 
            this.Item_ID.HeaderText = "Mã MH";
            this.Item_ID.MinimumWidth = 6;
            this.Item_ID.Name = "Item_ID";
            this.Item_ID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ItemName
            // 
            this.ItemName.HeaderText = "Tên MH";
            this.ItemName.MinimumWidth = 6;
            this.ItemName.Name = "ItemName";
            this.ItemName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Quantity
            // 
            this.Quantity.HeaderText = "Số lượng";
            this.Quantity.MinimumWidth = 6;
            this.Quantity.Name = "Quantity";
            this.Quantity.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // tblItemColSelect
            // 
            this.tblItemColSelect.HeaderText = "Chọn";
            this.tblItemColSelect.MinimumWidth = 6;
            this.tblItemColSelect.Name = "tblItemColSelect";
            this.tblItemColSelect.Text = "Chọn";
            this.tblItemColSelect.UseColumnTextForButtonValue = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(26, 68);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 15);
            this.label4.TabIndex = 1;
            this.label4.Text = "Số lượng:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 32);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 15);
            this.label3.TabIndex = 0;
            this.label3.Text = "Tên mặt hàng:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnCreateCartBill);
            this.groupBox3.Controls.Add(this.btnRefreshCart);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.labelCustomer_IDBill);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.labelCustomerNameBill);
            this.groupBox3.Controls.Add(this.labelCart_ID);
            this.groupBox3.Controls.Add(this.labelTotalAmount);
            this.groupBox3.Controls.Add(this.labelTotalDiscount);
            this.groupBox3.Controls.Add(this.labelTotalItem);
            this.groupBox3.Controls.Add(this.labelCustomerName);
            this.groupBox3.Controls.Add(this.tblBillDetail);
            this.groupBox3.Location = new System.Drawing.Point(2, 242);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Size = new System.Drawing.Size(936, 308);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Chi tiết mặt hàng trong hóa đơn";
            // 
            // btnCreateCartBill
            // 
            this.btnCreateCartBill.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnCreateCartBill.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateCartBill.ForeColor = System.Drawing.Color.White;
            this.btnCreateCartBill.Image = global::CSharpCounterFinalProject.Properties.Resources.trolley;
            this.btnCreateCartBill.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCreateCartBill.Location = new System.Drawing.Point(668, 250);
            this.btnCreateCartBill.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnCreateCartBill.Name = "btnCreateCartBill";
            this.btnCreateCartBill.Size = new System.Drawing.Size(130, 42);
            this.btnCreateCartBill.TabIndex = 14;
            this.btnCreateCartBill.Text = "Tạo giỏ hàng";
            this.btnCreateCartBill.UseVisualStyleBackColor = false;
            this.btnCreateCartBill.Click += new System.EventHandler(this.BtnCreateCart);
            // 
            // btnRefreshCart
            // 
            this.btnRefreshCart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnRefreshCart.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRefreshCart.ForeColor = System.Drawing.Color.White;
            this.btnRefreshCart.Image = global::CSharpCounterFinalProject.Properties.Resources.loop;
            this.btnRefreshCart.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRefreshCart.Location = new System.Drawing.Point(805, 250);
            this.btnRefreshCart.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnRefreshCart.Name = "btnRefreshCart";
            this.btnRefreshCart.Size = new System.Drawing.Size(122, 42);
            this.btnRefreshCart.TabIndex = 13;
            this.btnRefreshCart.Text = "Cập nhật";
            this.btnRefreshCart.UseVisualStyleBackColor = false;
            this.btnRefreshCart.Click += new System.EventHandler(this.btnRefreshCart_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(565, 264);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(36, 15);
            this.label8.TabIndex = 12;
            this.label8.Text = "Tiền:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(456, 264);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(63, 15);
            this.label7.TabIndex = 11;
            this.label7.Text = "Tổng KM:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(371, 264);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(42, 15);
            this.label2.TabIndex = 10;
            this.label2.Text = "Số SP:";
            // 
            // labelCustomer_IDBill
            // 
            this.labelCustomer_IDBill.AutoSize = true;
            this.labelCustomer_IDBill.Location = new System.Drawing.Point(146, 264);
            this.labelCustomer_IDBill.Name = "labelCustomer_IDBill";
            this.labelCustomer_IDBill.Size = new System.Drawing.Size(35, 15);
            this.labelCustomer_IDBill.TabIndex = 8;
            this.labelCustomer_IDBill.Text = "Code";
            this.labelCustomer_IDBill.Click += new System.EventHandler(this.labelCustomer_IDBill_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(97, 264);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(50, 15);
            this.label6.TabIndex = 7;
            this.label6.Text = "Mã KH:";
            // 
            // labelCustomerNameBill
            // 
            this.labelCustomerNameBill.AutoSize = true;
            this.labelCustomerNameBill.Location = new System.Drawing.Point(298, 264);
            this.labelCustomerNameBill.Name = "labelCustomerNameBill";
            this.labelCustomerNameBill.Size = new System.Drawing.Size(35, 15);
            this.labelCustomerNameBill.TabIndex = 6;
            this.labelCustomerNameBill.Text = "Code";
            // 
            // labelCart_ID
            // 
            this.labelCart_ID.AutoSize = true;
            this.labelCart_ID.Location = new System.Drawing.Point(20, 264);
            this.labelCart_ID.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelCart_ID.Name = "labelCart_ID";
            this.labelCart_ID.Size = new System.Drawing.Size(14, 15);
            this.labelCart_ID.TabIndex = 5;
            this.labelCart_ID.Text = "0";
            // 
            // labelTotalAmount
            // 
            this.labelTotalAmount.AutoSize = true;
            this.labelTotalAmount.Location = new System.Drawing.Point(606, 264);
            this.labelTotalAmount.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelTotalAmount.Name = "labelTotalAmount";
            this.labelTotalAmount.Size = new System.Drawing.Size(14, 15);
            this.labelTotalAmount.TabIndex = 4;
            this.labelTotalAmount.Text = "0";
            // 
            // labelTotalDiscount
            // 
            this.labelTotalDiscount.AutoSize = true;
            this.labelTotalDiscount.Location = new System.Drawing.Point(516, 264);
            this.labelTotalDiscount.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelTotalDiscount.Name = "labelTotalDiscount";
            this.labelTotalDiscount.Size = new System.Drawing.Size(14, 15);
            this.labelTotalDiscount.TabIndex = 3;
            this.labelTotalDiscount.Text = "0";
            // 
            // labelTotalItem
            // 
            this.labelTotalItem.AutoSize = true;
            this.labelTotalItem.Location = new System.Drawing.Point(417, 264);
            this.labelTotalItem.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelTotalItem.Name = "labelTotalItem";
            this.labelTotalItem.Size = new System.Drawing.Size(14, 15);
            this.labelTotalItem.TabIndex = 2;
            this.labelTotalItem.Text = "0";
            // 
            // labelCustomerName
            // 
            this.labelCustomerName.AutoSize = true;
            this.labelCustomerName.Location = new System.Drawing.Point(214, 264);
            this.labelCustomerName.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelCustomerName.Name = "labelCustomerName";
            this.labelCustomerName.Size = new System.Drawing.Size(79, 15);
            this.labelCustomerName.TabIndex = 1;
            this.labelCustomerName.Text = "Khách hàng:";
            // 
            // tblBillDetail
            // 
            this.tblBillDetail.AllowUserToAddRows = false;
            this.tblBillDetail.AllowUserToDeleteRows = false;
            this.tblBillDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tblBillDetail.BackgroundColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.tblBillDetail.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.tblBillDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblBillDetail.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SelectedItem_ID,
            this.Item_IDBill,
            this.ItemNameBill,
            this.NumberOfSelectedItem,
            this.PriceBill,
            this.PriceAfterDiscount,
            this.TotalAmount,
            this.btnEditBill,
            this.btnRemoveBill});
            this.tblBillDetail.Location = new System.Drawing.Point(0, 21);
            this.tblBillDetail.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tblBillDetail.Name = "tblBillDetail";
            this.tblBillDetail.RowHeadersWidth = 51;
            this.tblBillDetail.RowTemplate.Height = 24;
            this.tblBillDetail.Size = new System.Drawing.Size(934, 212);
            this.tblBillDetail.TabIndex = 0;
            this.tblBillDetail.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TblBillCellClick);
            // 
            // SelectedItem_ID
            // 
            this.SelectedItem_ID.HeaderText = "Mã HĐ";
            this.SelectedItem_ID.MinimumWidth = 6;
            this.SelectedItem_ID.Name = "SelectedItem_ID";
            this.SelectedItem_ID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Item_IDBill
            // 
            this.Item_IDBill.HeaderText = "Mã MH";
            this.Item_IDBill.MinimumWidth = 6;
            this.Item_IDBill.Name = "Item_IDBill";
            this.Item_IDBill.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ItemNameBill
            // 
            this.ItemNameBill.HeaderText = "Tên MH";
            this.ItemNameBill.MinimumWidth = 6;
            this.ItemNameBill.Name = "ItemNameBill";
            this.ItemNameBill.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // NumberOfSelectedItem
            // 
            this.NumberOfSelectedItem.HeaderText = "Số lượng";
            this.NumberOfSelectedItem.MinimumWidth = 6;
            this.NumberOfSelectedItem.Name = "NumberOfSelectedItem";
            this.NumberOfSelectedItem.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // PriceBill
            // 
            this.PriceBill.HeaderText = "Giá bán";
            this.PriceBill.MinimumWidth = 6;
            this.PriceBill.Name = "PriceBill";
            this.PriceBill.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // PriceAfterDiscount
            // 
            this.PriceAfterDiscount.HeaderText = "Giá sau KM";
            this.PriceAfterDiscount.MinimumWidth = 6;
            this.PriceAfterDiscount.Name = "PriceAfterDiscount";
            this.PriceAfterDiscount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // TotalAmount
            // 
            this.TotalAmount.HeaderText = "Thành tiền";
            this.TotalAmount.MinimumWidth = 6;
            this.TotalAmount.Name = "TotalAmount";
            this.TotalAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // btnEditBill
            // 
            this.btnEditBill.HeaderText = "Sửa";
            this.btnEditBill.MinimumWidth = 6;
            this.btnEditBill.Name = "btnEditBill";
            this.btnEditBill.Text = "Sửa";
            this.btnEditBill.UseColumnTextForButtonValue = true;
            // 
            // btnRemoveBill
            // 
            this.btnRemoveBill.HeaderText = "Xóa bỏ";
            this.btnRemoveBill.MinimumWidth = 6;
            this.btnRemoveBill.Name = "btnRemoveBill";
            this.btnRemoveBill.Text = "Xóa bỏ";
            this.btnRemoveBill.UseColumnTextForButtonValue = true;
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnSave.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.ForeColor = System.Drawing.Color.White;
            this.btnSave.Image = global::CSharpCounterFinalProject.Properties.Resources.floppy_disk;
            this.btnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSave.Location = new System.Drawing.Point(168, 571);
            this.btnSave.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(102, 38);
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "Trả sau";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.BtnPayClick);
            // 
            // btnPay
            // 
            this.btnPay.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnPay.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPay.ForeColor = System.Drawing.Color.White;
            this.btnPay.Image = global::CSharpCounterFinalProject.Properties.Resources.payment_method;
            this.btnPay.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPay.Location = new System.Drawing.Point(338, 571);
            this.btnPay.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnPay.Name = "btnPay";
            this.btnPay.Size = new System.Drawing.Size(102, 38);
            this.btnPay.TabIndex = 8;
            this.btnPay.Text = "Thanh toán";
            this.btnPay.UseVisualStyleBackColor = false;
            this.btnPay.Click += new System.EventHandler(this.BtnPayClick);
            // 
            // btnRemove
            // 
            this.btnRemove.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnRemove.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRemove.ForeColor = System.Drawing.Color.White;
            this.btnRemove.Image = global::CSharpCounterFinalProject.Properties.Resources.remove__1_;
            this.btnRemove.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRemove.Location = new System.Drawing.Point(502, 571);
            this.btnRemove.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(102, 38);
            this.btnRemove.TabIndex = 9;
            this.btnRemove.Text = "Xóa bỏ";
            this.btnRemove.UseVisualStyleBackColor = false;
            this.btnRemove.Click += new System.EventHandler(this.BtnRemoveCart);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnCancel.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Image = global::CSharpCounterFinalProject.Properties.Resources.remove;
            this.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancel.Location = new System.Drawing.Point(670, 571);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(102, 38);
            this.btnCancel.TabIndex = 10;
            this.btnCancel.Text = "Hủy bỏ";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancelCart);
            // 
            // AddEditBillFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.ClientSize = new System.Drawing.Size(939, 618);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnPay);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddEditBillFrm";
            this.Text = "THÊM MỚI HÓA ĐƠN";
            this.Load += new System.EventHandler(this.AddEditBillFrm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblSearchedCustomer)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericSelectedQuantity)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblSearchedItem)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblBillDetail)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtSearchCustomer;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.DataGridView tblSearchedCustomer;
        private System.Windows.Forms.Button btnSearchCustomer;
        private System.Windows.Forms.NumericUpDown numericSelectedQuantity;
        private System.Windows.Forms.TextBox txtSearchItem;
        private System.Windows.Forms.Button btnSearchItem;
        private System.Windows.Forms.DataGridView tblSearchedItem;
        private System.Windows.Forms.Label labelCart_ID;
        private System.Windows.Forms.Label labelTotalAmount;
        private System.Windows.Forms.Label labelTotalDiscount;
        private System.Windows.Forms.Label labelTotalItem;
        private System.Windows.Forms.Label labelCustomerName;
        private System.Windows.Forms.DataGridView tblBillDetail;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnPay;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.DataGridViewTextBoxColumn Customer_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn FullName;
        private System.Windows.Forms.DataGridViewButtonColumn tblCustomerColSelect;
        private System.Windows.Forms.DataGridViewTextBoxColumn Item_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn ItemName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Quantity;
        private System.Windows.Forms.DataGridViewButtonColumn tblItemColSelect;
        private System.Windows.Forms.DataGridViewTextBoxColumn SelectedItem_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn Item_IDBill;
        private System.Windows.Forms.DataGridViewTextBoxColumn ItemNameBill;
        private System.Windows.Forms.DataGridViewTextBoxColumn NumberOfSelectedItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn PriceBill;
        private System.Windows.Forms.DataGridViewTextBoxColumn PriceAfterDiscount;
        private System.Windows.Forms.DataGridViewTextBoxColumn TotalAmount;
        private System.Windows.Forms.DataGridViewButtonColumn btnEditBill;
        private System.Windows.Forms.DataGridViewButtonColumn btnRemoveBill;
        private System.Windows.Forms.Button btnEditBillItem;
        private System.Windows.Forms.TextBox txtBillID;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label labelCustomerNameBill;
        private System.Windows.Forms.Label labelCustomer_IDBill;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRefreshCart;
        private System.Windows.Forms.Button btnCreateCartBill;
    }
}