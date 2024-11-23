using System;
using System.Windows.Forms;

namespace CSharpCounterFinalProject
{
    partial class HomeFrm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HomeFrm));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.thaoTácToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.thêmMớiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mAddItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mAddCustomer = new System.Windows.Forms.ToolStripMenuItem();
            this.mAddDiscount = new System.Windows.Forms.ToolStripMenuItem();
            this.mAddBill = new System.Windows.Forms.ToolStripMenuItem();
            this.btn_SaveFile = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.thoátToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.xemToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mItemTab = new System.Windows.Forms.ToolStripMenuItem();
            this.mCustomerTab = new System.Windows.Forms.ToolStripMenuItem();
            this.mDiscountTab = new System.Windows.Forms.ToolStripMenuItem();
            this.mBillTab = new System.Windows.Forms.ToolStripMenuItem();
            this.tabStat = new System.Windows.Forms.TabPage();
            this.tbl_month_bestsold = new System.Windows.Forms.DataGridView();
            this.Month_bestsold = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tbl_revenue_month = new System.Windows.Forms.DataGridView();
            this.revenue_month = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tblStatCustomer = new System.Windows.Forms.DataGridView();
            this.stat_customer_Id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.stat_customer_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.stat_customer_times = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.stat_customer_quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.stat_customer_total = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tblStatItem = new System.Windows.Forms.DataGridView();
            this.inven_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.inven_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.inven_quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.inven_revenue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.nudYear = new System.Windows.Forms.NumericUpDown();
            this.nudMonth = new System.Windows.Forms.NumericUpDown();
            this.nudDay = new System.Windows.Forms.NumericUpDown();
            this.btnStatResult = new System.Windows.Forms.Button();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.comboStat = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.tabBill = new System.Windows.Forms.TabPage();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.btnSearchBill = new System.Windows.Forms.Button();
            this.txtSearchBill = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.comboSearchBill = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.btnRefreshBill = new System.Windows.Forms.Button();
            this.btnAddNewBill = new System.Windows.Forms.Button();
            this.tblHoaDon = new System.Windows.Forms.DataGridView();
            this.Bill_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FullNameBill = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StaffName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CreateTimeBill = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TotalItem = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SubTotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TotalDiscountAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TotalAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Status = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.outFile = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tabDiscount = new System.Windows.Forms.TabPage();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.btnSearchDiscount = new System.Windows.Forms.Button();
            this.txtSearchDiscount = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.comboSearchDiscount = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.btnRefreshDiscount = new System.Windows.Forms.Button();
            this.btnAddNeuDiscount = new System.Windows.Forms.Button();
            this.tblKhuyenMai = new System.Windows.Forms.DataGridView();
            this.Discount_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NameDiscount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StartTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EndTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DiscountType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DiscountPercent = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DiscountPriceAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MinPriceToUseDiscount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tblDiscountEdit = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tblDiscountRemove = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tabCustomer = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnSearchCustomer = new System.Windows.Forms.Button();
            this.txtSearchCustomer = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.comboSearchCustomer = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.radioSortCustomerByCreatedDate = new System.Windows.Forms.RadioButton();
            this.radioSortCustomerByPoint = new System.Windows.Forms.RadioButton();
            this.radioSortCustomerByBirthDate = new System.Windows.Forms.RadioButton();
            this.radioSortCustomerByName = new System.Windows.Forms.RadioButton();
            this.radioSortCustomerById = new System.Windows.Forms.RadioButton();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.btnRefreshCustomer = new System.Windows.Forms.Button();
            this.btnAddNewCustomer = new System.Windows.Forms.Button();
            this.tblKhachHang = new System.Windows.Forms.DataGridView();
            this.Customer_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FullName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BirthDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Address = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PhoneNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CustomerType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Point = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CreateTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Email = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tblCustomerEdit = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tblCustomerRemove = new System.Windows.Forms.DataGridViewButtonColumn();
            this.tabItem = new System.Windows.Forms.TabPage();
            this.txtMoTa = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.btnAnh = new System.Windows.Forms.Button();
            this.picAnh = new System.Windows.Forms.PictureBox();
            this.txtPhanLoai = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.txtGia = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.txtTonKho = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.txtTenSP = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.txtHangSX = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.txtMaSP = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.tblDuLieu = new System.Windows.Forms.DataGridView();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.numericItemTo = new System.Windows.Forms.NumericUpDown();
            this.numericItemFrom = new System.Windows.Forms.NumericUpDown();
            this.btnSearchItem = new System.Windows.Forms.Button();
            this.txtSearchItem = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.comboSearchItem = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.radioSortItemByProductDate = new System.Windows.Forms.RadioButton();
            this.radioSortItemByQuantity = new System.Windows.Forms.RadioButton();
            this.radioSortItemByName = new System.Windows.Forms.RadioButton();
            this.radioSortItemByPriceDESC = new System.Windows.Forms.RadioButton();
            this.radioSortItemByPriceASC = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnXoaSP = new System.Windows.Forms.Button();
            this.btnSuaSP = new System.Windows.Forms.Button();
            this.btnFreshItem = new System.Windows.Forms.Button();
            this.btnAddNewItem = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.groupBox12 = new System.Windows.Forms.GroupBox();
            this.menuStrip1.SuspendLayout();
            this.tabStat.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tbl_month_bestsold)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbl_revenue_month)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblStatCustomer)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblStatItem)).BeginInit();
            this.groupBox11.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMonth)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDay)).BeginInit();
            this.tabBill.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox10.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblHoaDon)).BeginInit();
            this.tabDiscount.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblKhuyenMai)).BeginInit();
            this.tabCustomer.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblKhachHang)).BeginInit();
            this.tabItem.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picAnh)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblDuLieu)).BeginInit();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericItemTo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericItemFrom)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.thaoTácToolStripMenuItem,
            this.xemToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(7, 2, 0, 2);
            this.menuStrip1.Size = new System.Drawing.Size(935, 28);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // thaoTácToolStripMenuItem
            // 
            this.thaoTácToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.thêmMớiToolStripMenuItem,
            this.btn_SaveFile,
            this.toolStripSeparator1,
            this.thoátToolStripMenuItem});
            this.thaoTácToolStripMenuItem.Image = global::CSharpCounterFinalProject.Properties.Resources.action;
            this.thaoTácToolStripMenuItem.Name = "thaoTácToolStripMenuItem";
            this.thaoTácToolStripMenuItem.Size = new System.Drawing.Size(84, 24);
            this.thaoTácToolStripMenuItem.Text = "Thao tác";
            // 
            // thêmMớiToolStripMenuItem
            // 
            this.thêmMớiToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mAddItem,
            this.mAddCustomer,
            this.mAddDiscount,
            this.mAddBill});
            this.thêmMớiToolStripMenuItem.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.thêmMớiToolStripMenuItem.Name = "thêmMớiToolStripMenuItem";
            this.thêmMớiToolStripMenuItem.Size = new System.Drawing.Size(132, 26);
            this.thêmMớiToolStripMenuItem.Text = "Thêm mới";
            // 
            // mAddItem
            // 
            this.mAddItem.Image = global::CSharpCounterFinalProject.Properties.Resources.shopping_bag;
            this.mAddItem.Name = "mAddItem";
            this.mAddItem.Size = new System.Drawing.Size(141, 26);
            this.mAddItem.Text = "Mặt hàng";
            this.mAddItem.Click += new System.EventHandler(this.MenuAddView);
            // 
            // mAddCustomer
            // 
            this.mAddCustomer.Image = global::CSharpCounterFinalProject.Properties.Resources.rating;
            this.mAddCustomer.Name = "mAddCustomer";
            this.mAddCustomer.Size = new System.Drawing.Size(141, 26);
            this.mAddCustomer.Text = "Khách hàng";
            this.mAddCustomer.Click += new System.EventHandler(this.MenuAddView);
            // 
            // mAddDiscount
            // 
            this.mAddDiscount.Image = global::CSharpCounterFinalProject.Properties.Resources.discount;
            this.mAddDiscount.Name = "mAddDiscount";
            this.mAddDiscount.Size = new System.Drawing.Size(141, 26);
            this.mAddDiscount.Text = "Khuyến mãi";
            this.mAddDiscount.Click += new System.EventHandler(this.MenuAddView);
            // 
            // mAddBill
            // 
            this.mAddBill.Image = global::CSharpCounterFinalProject.Properties.Resources.bill;
            this.mAddBill.Name = "mAddBill";
            this.mAddBill.Size = new System.Drawing.Size(141, 26);
            this.mAddBill.Text = "Hóa đơn";
            this.mAddBill.Click += new System.EventHandler(this.MenuAddView);
            // 
            // btn_SaveFile
            // 
            this.btn_SaveFile.Image = global::CSharpCounterFinalProject.Properties.Resources.floppy_disk;
            this.btn_SaveFile.Name = "btn_SaveFile";
            this.btn_SaveFile.Size = new System.Drawing.Size(132, 26);
            this.btn_SaveFile.Text = "Lưu file";
            this.btn_SaveFile.Click += new System.EventHandler(this.btnInHDClick);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(129, 6);
            // 
            // thoátToolStripMenuItem
            // 
            this.thoátToolStripMenuItem.Image = global::CSharpCounterFinalProject.Properties.Resources.remove;
            this.thoátToolStripMenuItem.Name = "thoátToolStripMenuItem";
            this.thoátToolStripMenuItem.Size = new System.Drawing.Size(132, 26);
            this.thoátToolStripMenuItem.Text = "Thoát";
            this.thoátToolStripMenuItem.Click += new System.EventHandler(this.BtnCloseApp);
            // 
            // xemToolStripMenuItem
            // 
            this.xemToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mItemTab,
            this.mCustomerTab,
            this.mDiscountTab,
            this.mBillTab});
            this.xemToolStripMenuItem.Image = global::CSharpCounterFinalProject.Properties.Resources.research_center;
            this.xemToolStripMenuItem.Name = "xemToolStripMenuItem";
            this.xemToolStripMenuItem.Size = new System.Drawing.Size(63, 24);
            this.xemToolStripMenuItem.Text = "Xem";
            // 
            // mItemTab
            // 
            this.mItemTab.Image = global::CSharpCounterFinalProject.Properties.Resources.shopping_bag;
            this.mItemTab.Name = "mItemTab";
            this.mItemTab.Size = new System.Drawing.Size(141, 26);
            this.mItemTab.Text = "Mặt hàng";
            this.mItemTab.Click += new System.EventHandler(this.MenuItemViewTabClick);
            // 
            // mCustomerTab
            // 
            this.mCustomerTab.Image = global::CSharpCounterFinalProject.Properties.Resources.rating;
            this.mCustomerTab.Name = "mCustomerTab";
            this.mCustomerTab.Size = new System.Drawing.Size(141, 26);
            this.mCustomerTab.Text = "Khách hàng";
            this.mCustomerTab.Click += new System.EventHandler(this.MenuItemViewTabClick);
            // 
            // mDiscountTab
            // 
            this.mDiscountTab.Image = global::CSharpCounterFinalProject.Properties.Resources.discount;
            this.mDiscountTab.Name = "mDiscountTab";
            this.mDiscountTab.Size = new System.Drawing.Size(141, 26);
            this.mDiscountTab.Text = "Khuyến mãi";
            this.mDiscountTab.Click += new System.EventHandler(this.MenuItemViewTabClick);
            // 
            // mBillTab
            // 
            this.mBillTab.Image = global::CSharpCounterFinalProject.Properties.Resources.bill;
            this.mBillTab.Name = "mBillTab";
            this.mBillTab.Size = new System.Drawing.Size(141, 26);
            this.mBillTab.Text = "Hóa đơn";
            this.mBillTab.Click += new System.EventHandler(this.MenuItemViewTabClick);
            // 
            // tabStat
            // 
            this.tabStat.Controls.Add(this.tbl_month_bestsold);
            this.tabStat.Controls.Add(this.tbl_revenue_month);
            this.tabStat.Controls.Add(this.tblStatCustomer);
            this.tabStat.Controls.Add(this.tblStatItem);
            this.tabStat.Controls.Add(this.groupBox11);
            this.tabStat.Location = new System.Drawing.Point(4, 22);
            this.tabStat.Name = "tabStat";
            this.tabStat.Padding = new System.Windows.Forms.Padding(3);
            this.tabStat.Size = new System.Drawing.Size(931, 580);
            this.tabStat.TabIndex = 4;
            this.tabStat.Text = "Thống Kê";
            this.tabStat.UseVisualStyleBackColor = true;
            // 
            // tbl_month_bestsold
            // 
            this.tbl_month_bestsold.AllowUserToAddRows = false;
            this.tbl_month_bestsold.AllowUserToDeleteRows = false;
            this.tbl_month_bestsold.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tbl_month_bestsold.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tbl_month_bestsold.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Month_bestsold});
            this.tbl_month_bestsold.Location = new System.Drawing.Point(694, 0);
            this.tbl_month_bestsold.Name = "tbl_month_bestsold";
            this.tbl_month_bestsold.RowHeadersWidth = 51;
            this.tbl_month_bestsold.RowTemplate.Height = 24;
            this.tbl_month_bestsold.Size = new System.Drawing.Size(236, 108);
            this.tbl_month_bestsold.TabIndex = 18;
            // 
            // Month_bestsold
            // 
            this.Month_bestsold.HeaderText = "Ngày bán chạy nhất của tháng";
            this.Month_bestsold.MinimumWidth = 6;
            this.Month_bestsold.Name = "Month_bestsold";
            // 
            // tbl_revenue_month
            // 
            this.tbl_revenue_month.AllowUserToAddRows = false;
            this.tbl_revenue_month.AllowUserToDeleteRows = false;
            this.tbl_revenue_month.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tbl_revenue_month.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tbl_revenue_month.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.revenue_month});
            this.tbl_revenue_month.Location = new System.Drawing.Point(694, 103);
            this.tbl_revenue_month.Name = "tbl_revenue_month";
            this.tbl_revenue_month.RowHeadersWidth = 51;
            this.tbl_revenue_month.RowTemplate.Height = 24;
            this.tbl_revenue_month.Size = new System.Drawing.Size(238, 117);
            this.tbl_revenue_month.TabIndex = 17;
            // 
            // revenue_month
            // 
            this.revenue_month.HeaderText = "Danh thu ";
            this.revenue_month.MinimumWidth = 6;
            this.revenue_month.Name = "revenue_month";
            // 
            // tblStatCustomer
            // 
            this.tblStatCustomer.AllowUserToAddRows = false;
            this.tblStatCustomer.AllowUserToDeleteRows = false;
            this.tblStatCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tblStatCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblStatCustomer.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.stat_customer_Id,
            this.stat_customer_name,
            this.stat_customer_times,
            this.stat_customer_quantity,
            this.stat_customer_total});
            this.tblStatCustomer.Location = new System.Drawing.Point(0, 219);
            this.tblStatCustomer.Name = "tblStatCustomer";
            this.tblStatCustomer.RowHeadersWidth = 51;
            this.tblStatCustomer.RowTemplate.Height = 24;
            this.tblStatCustomer.Size = new System.Drawing.Size(932, 215);
            this.tblStatCustomer.TabIndex = 16;
            // 
            // stat_customer_Id
            // 
            this.stat_customer_Id.HeaderText = "Mã KH";
            this.stat_customer_Id.MinimumWidth = 6;
            this.stat_customer_Id.Name = "stat_customer_Id";
            this.stat_customer_Id.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // stat_customer_name
            // 
            this.stat_customer_name.HeaderText = "Tên KH";
            this.stat_customer_name.MinimumWidth = 6;
            this.stat_customer_name.Name = "stat_customer_name";
            this.stat_customer_name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // stat_customer_times
            // 
            this.stat_customer_times.HeaderText = "Số lần mua hàng";
            this.stat_customer_times.MinimumWidth = 6;
            this.stat_customer_times.Name = "stat_customer_times";
            this.stat_customer_times.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // stat_customer_quantity
            // 
            this.stat_customer_quantity.HeaderText = "Số mặt hàng đã mua";
            this.stat_customer_quantity.MinimumWidth = 6;
            this.stat_customer_quantity.Name = "stat_customer_quantity";
            // 
            // stat_customer_total
            // 
            this.stat_customer_total.HeaderText = "Tổng tiền đã mua ";
            this.stat_customer_total.MinimumWidth = 6;
            this.stat_customer_total.Name = "stat_customer_total";
            // 
            // tblStatItem
            // 
            this.tblStatItem.AllowUserToAddRows = false;
            this.tblStatItem.AllowUserToDeleteRows = false;
            this.tblStatItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tblStatItem.BackgroundColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.tblStatItem.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.tblStatItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblStatItem.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.inven_id,
            this.inven_name,
            this.inven_quantity,
            this.inven_revenue});
            this.tblStatItem.Location = new System.Drawing.Point(0, 0);
            this.tblStatItem.Name = "tblStatItem";
            this.tblStatItem.RowHeadersWidth = 51;
            this.tblStatItem.RowTemplate.Height = 24;
            this.tblStatItem.Size = new System.Drawing.Size(693, 221);
            this.tblStatItem.TabIndex = 15;
            // 
            // inven_id
            // 
            this.inven_id.HeaderText = "Mã MH";
            this.inven_id.MinimumWidth = 6;
            this.inven_id.Name = "inven_id";
            this.inven_id.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // inven_name
            // 
            this.inven_name.HeaderText = "Tên mặt hàng";
            this.inven_name.MinimumWidth = 6;
            this.inven_name.Name = "inven_name";
            this.inven_name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // inven_quantity
            // 
            this.inven_quantity.HeaderText = "Số lượng";
            this.inven_quantity.MinimumWidth = 6;
            this.inven_quantity.Name = "inven_quantity";
            this.inven_quantity.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // inven_revenue
            // 
            this.inven_revenue.HeaderText = "Tổng doanh thu ";
            this.inven_revenue.MinimumWidth = 6;
            this.inven_revenue.Name = "inven_revenue";
            this.inven_revenue.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.button2);
            this.groupBox11.Controls.Add(this.button1);
            this.groupBox11.Controls.Add(this.nudYear);
            this.groupBox11.Controls.Add(this.nudMonth);
            this.groupBox11.Controls.Add(this.nudDay);
            this.groupBox11.Controls.Add(this.btnStatResult);
            this.groupBox11.Controls.Add(this.label14);
            this.groupBox11.Controls.Add(this.label13);
            this.groupBox11.Controls.Add(this.label11);
            this.groupBox11.Controls.Add(this.comboStat);
            this.groupBox11.Controls.Add(this.label12);
            this.groupBox11.Location = new System.Drawing.Point(0, 439);
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.Size = new System.Drawing.Size(930, 150);
            this.groupBox11.TabIndex = 14;
            this.groupBox11.TabStop = false;
            this.groupBox11.Text = "Thống kê theo tiêu chí";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(592, 71);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(131, 30);
            this.button2.TabIndex = 18;
            this.button2.Text = "Vẽ biểu đồ tròn";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.BtnCreateChartPieClick);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(592, 32);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(131, 30);
            this.button1.TabIndex = 17;
            this.button1.Text = "Vẽ biểu đồ cột";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.BtnCreateChartColumnClick);
            // 
            // nudYear
            // 
            this.nudYear.Location = new System.Drawing.Point(501, 80);
            this.nudYear.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.nudYear.Name = "nudYear";
            this.nudYear.Size = new System.Drawing.Size(70, 20);
            this.nudYear.TabIndex = 16;
            // 
            // nudMonth
            // 
            this.nudMonth.Location = new System.Drawing.Point(333, 80);
            this.nudMonth.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.nudMonth.Name = "nudMonth";
            this.nudMonth.Size = new System.Drawing.Size(70, 20);
            this.nudMonth.TabIndex = 15;
            // 
            // nudDay
            // 
            this.nudDay.Location = new System.Drawing.Point(138, 80);
            this.nudDay.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.nudDay.Name = "nudDay";
            this.nudDay.Size = new System.Drawing.Size(70, 20);
            this.nudDay.TabIndex = 14;
            // 
            // btnStatResult
            // 
            this.btnStatResult.Location = new System.Drawing.Point(744, 54);
            this.btnStatResult.Name = "btnStatResult";
            this.btnStatResult.Size = new System.Drawing.Size(112, 48);
            this.btnStatResult.TabIndex = 7;
            this.btnStatResult.Text = "Xem kết quả";
            this.btnStatResult.UseVisualStyleBackColor = true;
            this.btnStatResult.Click += new System.EventHandler(this.BtnSearchStatClick);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(270, 86);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(40, 13);
            this.label14.TabIndex = 5;
            this.label14.Text = "Tháng:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(446, 86);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(30, 13);
            this.label13.TabIndex = 3;
            this.label13.Text = "Năm";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(75, 86);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(34, 13);
            this.label11.TabIndex = 2;
            this.label11.Text = "Ngày:";
            // 
            // comboStat
            // 
            this.comboStat.FormattingEnabled = true;
            this.comboStat.Items.AddRange(new object[] {
            "Các mặt hàng có danh thu tốt nhất",
            "Các khách hàng mua nhiều hàng nhất",
            "Những ngày của tháng bán được nhiều hàng nhất",
            "Danh thu theo tháng của năm cho trước",
            "Danh thu theo ngày của tháng cho trước"});
            this.comboStat.Location = new System.Drawing.Point(138, 40);
            this.comboStat.Name = "comboStat";
            this.comboStat.Size = new System.Drawing.Size(434, 21);
            this.comboStat.TabIndex = 1;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(74, 47);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(45, 13);
            this.label12.TabIndex = 0;
            this.label12.Text = "Tiêu chí";
            // 
            // tabBill
            // 
            this.tabBill.Controls.Add(this.groupBox8);
            this.tabBill.Controls.Add(this.groupBox10);
            this.tabBill.Controls.Add(this.tblHoaDon);
            this.tabBill.Location = new System.Drawing.Point(4, 22);
            this.tabBill.Name = "tabBill";
            this.tabBill.Padding = new System.Windows.Forms.Padding(3);
            this.tabBill.Size = new System.Drawing.Size(931, 580);
            this.tabBill.TabIndex = 3;
            this.tabBill.Text = "QL Hóa Đơn";
            this.tabBill.UseVisualStyleBackColor = true;
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.btnSearchBill);
            this.groupBox8.Controls.Add(this.txtSearchBill);
            this.groupBox8.Controls.Add(this.label9);
            this.groupBox8.Controls.Add(this.comboSearchBill);
            this.groupBox8.Controls.Add(this.label10);
            this.groupBox8.Location = new System.Drawing.Point(480, 411);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(448, 175);
            this.groupBox8.TabIndex = 14;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Tìm kiếm";
            // 
            // btnSearchBill
            // 
            this.btnSearchBill.Image = global::CSharpCounterFinalProject.Properties.Resources.search;
            this.btnSearchBill.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSearchBill.Location = new System.Drawing.Point(137, 111);
            this.btnSearchBill.Name = "btnSearchBill";
            this.btnSearchBill.Size = new System.Drawing.Size(131, 41);
            this.btnSearchBill.TabIndex = 6;
            this.btnSearchBill.Text = "Tìm kiếm";
            this.btnSearchBill.UseVisualStyleBackColor = true;
            this.btnSearchBill.Click += new System.EventHandler(this.BtnSearchBillClick);
            // 
            // txtSearchBill
            // 
            this.txtSearchBill.Location = new System.Drawing.Point(137, 64);
            this.txtSearchBill.Name = "txtSearchBill";
            this.txtSearchBill.Size = new System.Drawing.Size(240, 20);
            this.txtSearchBill.TabIndex = 5;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(74, 64);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(51, 13);
            this.label9.TabIndex = 2;
            this.label9.Text = "Nội dung";
            // 
            // comboSearchBill
            // 
            this.comboSearchBill.FormattingEnabled = true;
            this.comboSearchBill.Items.AddRange(new object[] {
            "Tìm theo số lượng hàng",
            "Tìm theo tên khách hàng.",
            "Tìm theo trạng thái hóa đơn.",
            "Tìm theo tên nhân viên"});
            this.comboSearchBill.Location = new System.Drawing.Point(137, 26);
            this.comboSearchBill.Name = "comboSearchBill";
            this.comboSearchBill.Size = new System.Drawing.Size(240, 21);
            this.comboSearchBill.TabIndex = 1;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(74, 34);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(45, 13);
            this.label10.TabIndex = 0;
            this.label10.Text = "Tiêu chí";
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.btnRefreshBill);
            this.groupBox10.Controls.Add(this.btnAddNewBill);
            this.groupBox10.Location = new System.Drawing.Point(-7, 411);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(490, 178);
            this.groupBox10.TabIndex = 13;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "Hành động";
            // 
            // btnRefreshBill
            // 
            this.btnRefreshBill.Image = global::CSharpCounterFinalProject.Properties.Resources.refresh;
            this.btnRefreshBill.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRefreshBill.Location = new System.Drawing.Point(263, 64);
            this.btnRefreshBill.Name = "btnRefreshBill";
            this.btnRefreshBill.Size = new System.Drawing.Size(185, 42);
            this.btnRefreshBill.TabIndex = 1;
            this.btnRefreshBill.Text = "Làm mới";
            this.btnRefreshBill.UseVisualStyleBackColor = true;
            this.btnRefreshBill.Click += new System.EventHandler(this.BtnRefreshClick);
            // 
            // btnAddNewBill
            // 
            this.btnAddNewBill.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.btnAddNewBill.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddNewBill.Location = new System.Drawing.Point(34, 64);
            this.btnAddNewBill.Name = "btnAddNewBill";
            this.btnAddNewBill.Size = new System.Drawing.Size(185, 42);
            this.btnAddNewBill.TabIndex = 0;
            this.btnAddNewBill.Text = "Thêm mới";
            this.btnAddNewBill.UseVisualStyleBackColor = true;
            this.btnAddNewBill.Click += new System.EventHandler(this.BtnAddNewBillClick);
            // 
            // tblHoaDon
            // 
            this.tblHoaDon.AllowUserToAddRows = false;
            this.tblHoaDon.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tblHoaDon.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.tblHoaDon.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.tblHoaDon.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblHoaDon.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Bill_ID,
            this.FullNameBill,
            this.StaffName,
            this.CreateTimeBill,
            this.TotalItem,
            this.SubTotal,
            this.TotalDiscountAmount,
            this.TotalAmount,
            this.Status,
            this.outFile});
            this.tblHoaDon.Location = new System.Drawing.Point(-2, -7);
            this.tblHoaDon.Name = "tblHoaDon";
            this.tblHoaDon.RowHeadersWidth = 51;
            this.tblHoaDon.RowTemplate.Height = 24;
            this.tblHoaDon.Size = new System.Drawing.Size(932, 414);
            this.tblHoaDon.TabIndex = 12;
            this.tblHoaDon.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TblBillCellClick);
            // 
            // Bill_ID
            // 
            this.Bill_ID.HeaderText = "Mã HĐ";
            this.Bill_ID.MinimumWidth = 6;
            this.Bill_ID.Name = "Bill_ID";
            this.Bill_ID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // FullNameBill
            // 
            this.FullNameBill.HeaderText = "Tên KH";
            this.FullNameBill.MinimumWidth = 6;
            this.FullNameBill.Name = "FullNameBill";
            this.FullNameBill.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // StaffName
            // 
            this.StaffName.HeaderText = "Tên NV";
            this.StaffName.MinimumWidth = 6;
            this.StaffName.Name = "StaffName";
            this.StaffName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // CreateTimeBill
            // 
            this.CreateTimeBill.HeaderText = "TG lập HĐ";
            this.CreateTimeBill.MinimumWidth = 6;
            this.CreateTimeBill.Name = "CreateTimeBill";
            this.CreateTimeBill.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // TotalItem
            // 
            this.TotalItem.HeaderText = "Tổng SL hàng";
            this.TotalItem.MinimumWidth = 6;
            this.TotalItem.Name = "TotalItem";
            this.TotalItem.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // SubTotal
            // 
            this.SubTotal.HeaderText = "Tạm tính";
            this.SubTotal.MinimumWidth = 6;
            this.SubTotal.Name = "SubTotal";
            this.SubTotal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // TotalDiscountAmount
            // 
            this.TotalDiscountAmount.HeaderText = "Tổng KM";
            this.TotalDiscountAmount.MinimumWidth = 6;
            this.TotalDiscountAmount.Name = "TotalDiscountAmount";
            this.TotalDiscountAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // TotalAmount
            // 
            this.TotalAmount.HeaderText = "Tổng tiền ";
            this.TotalAmount.MinimumWidth = 6;
            this.TotalAmount.Name = "TotalAmount";
            this.TotalAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Status
            // 
            this.Status.HeaderText = "Trạng thái";
            this.Status.MinimumWidth = 6;
            this.Status.Name = "Status";
            this.Status.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // outFile
            // 
            this.outFile.HeaderText = "Chi tiết";
            this.outFile.MinimumWidth = 6;
            this.outFile.Name = "outFile";
            this.outFile.Text = "In HĐ";
            this.outFile.UseColumnTextForButtonValue = true;
            // 
            // tabDiscount
            // 
            this.tabDiscount.Controls.Add(this.groupBox7);
            this.tabDiscount.Controls.Add(this.groupBox9);
            this.tabDiscount.Controls.Add(this.tblKhuyenMai);
            this.tabDiscount.Location = new System.Drawing.Point(4, 22);
            this.tabDiscount.Name = "tabDiscount";
            this.tabDiscount.Padding = new System.Windows.Forms.Padding(3);
            this.tabDiscount.Size = new System.Drawing.Size(931, 580);
            this.tabDiscount.TabIndex = 2;
            this.tabDiscount.Text = "QL Khuyết Mãi ";
            this.tabDiscount.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.btnSearchDiscount);
            this.groupBox7.Controls.Add(this.txtSearchDiscount);
            this.groupBox7.Controls.Add(this.label5);
            this.groupBox7.Controls.Add(this.comboSearchDiscount);
            this.groupBox7.Controls.Add(this.label6);
            this.groupBox7.Location = new System.Drawing.Point(480, 411);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(450, 178);
            this.groupBox7.TabIndex = 11;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Tìm kiếm";
            // 
            // btnSearchDiscount
            // 
            this.btnSearchDiscount.Image = global::CSharpCounterFinalProject.Properties.Resources.search;
            this.btnSearchDiscount.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSearchDiscount.Location = new System.Drawing.Point(137, 111);
            this.btnSearchDiscount.Name = "btnSearchDiscount";
            this.btnSearchDiscount.Size = new System.Drawing.Size(131, 41);
            this.btnSearchDiscount.TabIndex = 6;
            this.btnSearchDiscount.Text = "Tìm kiếm";
            this.btnSearchDiscount.UseVisualStyleBackColor = true;
            this.btnSearchDiscount.Click += new System.EventHandler(this.BtnSearchDiscountClick);
            // 
            // txtSearchDiscount
            // 
            this.txtSearchDiscount.Location = new System.Drawing.Point(137, 64);
            this.txtSearchDiscount.Name = "txtSearchDiscount";
            this.txtSearchDiscount.Size = new System.Drawing.Size(240, 20);
            this.txtSearchDiscount.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(74, 64);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Nội dung";
            // 
            // comboSearchDiscount
            // 
            this.comboSearchDiscount.FormattingEnabled = true;
            this.comboSearchDiscount.Items.AddRange(new object[] {
            "Tên khuyến mãi gần đúng",
            "Mã khuyến mãi",
            "Loại khuyến mãi",
            "Số tiền KM"});
            this.comboSearchDiscount.Location = new System.Drawing.Point(137, 26);
            this.comboSearchDiscount.Name = "comboSearchDiscount";
            this.comboSearchDiscount.Size = new System.Drawing.Size(240, 21);
            this.comboSearchDiscount.TabIndex = 1;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(74, 34);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(45, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Tiêu chí";
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.button5);
            this.groupBox9.Controls.Add(this.button6);
            this.groupBox9.Controls.Add(this.btnRefreshDiscount);
            this.groupBox9.Controls.Add(this.btnAddNeuDiscount);
            this.groupBox9.Location = new System.Drawing.Point(2, 411);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(481, 178);
            this.groupBox9.TabIndex = 9;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Hành động";
            // 
            // button5
            // 
            this.button5.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button5.Location = new System.Drawing.Point(263, 14);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(185, 42);
            this.button5.TabIndex = 5;
            this.button5.Text = "Xoá";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // button6
            // 
            this.button6.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button6.Location = new System.Drawing.Point(36, 14);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(183, 42);
            this.button6.TabIndex = 4;
            this.button6.Text = "Sửa";
            this.button6.UseVisualStyleBackColor = true;
            // 
            // btnRefreshDiscount
            // 
            this.btnRefreshDiscount.Image = global::CSharpCounterFinalProject.Properties.Resources.refresh;
            this.btnRefreshDiscount.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRefreshDiscount.Location = new System.Drawing.Point(263, 64);
            this.btnRefreshDiscount.Name = "btnRefreshDiscount";
            this.btnRefreshDiscount.Size = new System.Drawing.Size(185, 42);
            this.btnRefreshDiscount.TabIndex = 1;
            this.btnRefreshDiscount.Text = "Làm mới";
            this.btnRefreshDiscount.UseVisualStyleBackColor = true;
            this.btnRefreshDiscount.Click += new System.EventHandler(this.BtnRefreshClick);
            // 
            // btnAddNeuDiscount
            // 
            this.btnAddNeuDiscount.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.btnAddNeuDiscount.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddNeuDiscount.Location = new System.Drawing.Point(34, 64);
            this.btnAddNeuDiscount.Name = "btnAddNeuDiscount";
            this.btnAddNeuDiscount.Size = new System.Drawing.Size(185, 42);
            this.btnAddNeuDiscount.TabIndex = 0;
            this.btnAddNeuDiscount.Text = "Thêm mới";
            this.btnAddNeuDiscount.UseVisualStyleBackColor = true;
            this.btnAddNeuDiscount.Click += new System.EventHandler(this.BtnAddNewDiscountClick);
            // 
            // tblKhuyenMai
            // 
            this.tblKhuyenMai.AllowUserToAddRows = false;
            this.tblKhuyenMai.AllowUserToDeleteRows = false;
            this.tblKhuyenMai.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tblKhuyenMai.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.tblKhuyenMai.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.tblKhuyenMai.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblKhuyenMai.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Discount_ID,
            this.NameDiscount,
            this.StartTime,
            this.EndTime,
            this.DiscountType,
            this.DiscountPercent,
            this.DiscountPriceAmount,
            this.MinPriceToUseDiscount,
            this.tblDiscountEdit,
            this.tblDiscountRemove});
            this.tblKhuyenMai.Location = new System.Drawing.Point(-4, 0);
            this.tblKhuyenMai.Name = "tblKhuyenMai";
            this.tblKhuyenMai.RowHeadersWidth = 51;
            this.tblKhuyenMai.RowTemplate.Height = 24;
            this.tblKhuyenMai.Size = new System.Drawing.Size(932, 406);
            this.tblKhuyenMai.TabIndex = 8;
            this.tblKhuyenMai.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TblDiscountCellClick);
            // 
            // Discount_ID
            // 
            this.Discount_ID.HeaderText = "Mã KM";
            this.Discount_ID.MinimumWidth = 6;
            this.Discount_ID.Name = "Discount_ID";
            this.Discount_ID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // NameDiscount
            // 
            this.NameDiscount.HeaderText = "Tên KM";
            this.NameDiscount.MinimumWidth = 6;
            this.NameDiscount.Name = "NameDiscount";
            this.NameDiscount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // StartTime
            // 
            this.StartTime.HeaderText = "TG bắt đầu";
            this.StartTime.MinimumWidth = 6;
            this.StartTime.Name = "StartTime";
            this.StartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // EndTime
            // 
            this.EndTime.HeaderText = "TG kết thúc";
            this.EndTime.MinimumWidth = 6;
            this.EndTime.Name = "EndTime";
            this.EndTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // DiscountType
            // 
            this.DiscountType.HeaderText = "Loại KM";
            this.DiscountType.MinimumWidth = 6;
            this.DiscountType.Name = "DiscountType";
            this.DiscountType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // DiscountPercent
            // 
            this.DiscountPercent.HeaderText = "% KM";
            this.DiscountPercent.MinimumWidth = 6;
            this.DiscountPercent.Name = "DiscountPercent";
            this.DiscountPercent.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // DiscountPriceAmount
            // 
            this.DiscountPriceAmount.HeaderText = "Số tiền KM";
            this.DiscountPriceAmount.MinimumWidth = 6;
            this.DiscountPriceAmount.Name = "DiscountPriceAmount";
            this.DiscountPriceAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // MinPriceToUseDiscount
            // 
            this.MinPriceToUseDiscount.HeaderText = "Giá tối thiểu";
            this.MinPriceToUseDiscount.MinimumWidth = 6;
            this.MinPriceToUseDiscount.Name = "MinPriceToUseDiscount";
            this.MinPriceToUseDiscount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // tblDiscountEdit
            // 
            this.tblDiscountEdit.HeaderText = "Sửa";
            this.tblDiscountEdit.MinimumWidth = 6;
            this.tblDiscountEdit.Name = "tblDiscountEdit";
            this.tblDiscountEdit.Text = "Sửa";
            this.tblDiscountEdit.UseColumnTextForButtonValue = true;
            // 
            // tblDiscountRemove
            // 
            this.tblDiscountRemove.HeaderText = "Xóa";
            this.tblDiscountRemove.MinimumWidth = 6;
            this.tblDiscountRemove.Name = "tblDiscountRemove";
            this.tblDiscountRemove.Text = "Xóa";
            this.tblDiscountRemove.UseColumnTextForButtonValue = true;
            // 
            // tabCustomer
            // 
            this.tabCustomer.Controls.Add(this.groupBox12);
            this.tabCustomer.Controls.Add(this.groupBox4);
            this.tabCustomer.Controls.Add(this.groupBox5);
            this.tabCustomer.Controls.Add(this.groupBox6);
            this.tabCustomer.Controls.Add(this.tblKhachHang);
            this.tabCustomer.Location = new System.Drawing.Point(4, 22);
            this.tabCustomer.Name = "tabCustomer";
            this.tabCustomer.Padding = new System.Windows.Forms.Padding(3);
            this.tabCustomer.Size = new System.Drawing.Size(931, 580);
            this.tabCustomer.TabIndex = 1;
            this.tabCustomer.Text = "QL Khách Hàng";
            this.tabCustomer.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btnSearchCustomer);
            this.groupBox4.Controls.Add(this.txtSearchCustomer);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Controls.Add(this.comboSearchCustomer);
            this.groupBox4.Controls.Add(this.label8);
            this.groupBox4.Location = new System.Drawing.Point(648, 411);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(282, 178);
            this.groupBox4.TabIndex = 7;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Tìm kiếm";
            // 
            // btnSearchCustomer
            // 
            this.btnSearchCustomer.Image = global::CSharpCounterFinalProject.Properties.Resources.search;
            this.btnSearchCustomer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSearchCustomer.Location = new System.Drawing.Point(83, 109);
            this.btnSearchCustomer.Name = "btnSearchCustomer";
            this.btnSearchCustomer.Size = new System.Drawing.Size(131, 41);
            this.btnSearchCustomer.TabIndex = 6;
            this.btnSearchCustomer.Text = "Tìm kiếm";
            this.btnSearchCustomer.UseVisualStyleBackColor = true;
            this.btnSearchCustomer.Click += new System.EventHandler(this.BtnSearchCustomerClick);
            // 
            // txtSearchCustomer
            // 
            this.txtSearchCustomer.Location = new System.Drawing.Point(83, 63);
            this.txtSearchCustomer.Name = "txtSearchCustomer";
            this.txtSearchCustomer.Size = new System.Drawing.Size(180, 20);
            this.txtSearchCustomer.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(19, 63);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(51, 13);
            this.label7.TabIndex = 2;
            this.label7.Text = "Nội dung";
            // 
            // comboSearchCustomer
            // 
            this.comboSearchCustomer.FormattingEnabled = true;
            this.comboSearchCustomer.Items.AddRange(new object[] {
            "Tên khách hàng gần đúng",
            "Mã khách hàng",
            "Loại khách hàng",
            "Địa chỉ",
            "Số điện thoại"});
            this.comboSearchCustomer.Location = new System.Drawing.Point(83, 25);
            this.comboSearchCustomer.Name = "comboSearchCustomer";
            this.comboSearchCustomer.Size = new System.Drawing.Size(180, 21);
            this.comboSearchCustomer.TabIndex = 1;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(19, 32);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(45, 13);
            this.label8.TabIndex = 0;
            this.label8.Text = "Tiêu chí";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.radioSortCustomerByCreatedDate);
            this.groupBox5.Controls.Add(this.radioSortCustomerByPoint);
            this.groupBox5.Controls.Add(this.radioSortCustomerByBirthDate);
            this.groupBox5.Controls.Add(this.radioSortCustomerByName);
            this.groupBox5.Controls.Add(this.radioSortCustomerById);
            this.groupBox5.Location = new System.Drawing.Point(304, 411);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(346, 178);
            this.groupBox5.TabIndex = 6;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Sắp xếp theo";
            // 
            // radioSortCustomerByCreatedDate
            // 
            this.radioSortCustomerByCreatedDate.AutoSize = true;
            this.radioSortCustomerByCreatedDate.Location = new System.Drawing.Point(192, 89);
            this.radioSortCustomerByCreatedDate.Name = "radioSortCustomerByCreatedDate";
            this.radioSortCustomerByCreatedDate.Size = new System.Drawing.Size(114, 17);
            this.radioSortCustomerByCreatedDate.TabIndex = 4;
            this.radioSortCustomerByCreatedDate.TabStop = true;
            this.radioSortCustomerByCreatedDate.Text = "Ngày tạo tài khoản";
            this.radioSortCustomerByCreatedDate.UseVisualStyleBackColor = true;
            this.radioSortCustomerByCreatedDate.CheckedChanged += new System.EventHandler(this.RadioSortCustomerCheckedChanged);
            // 
            // radioSortCustomerByPoint
            // 
            this.radioSortCustomerByPoint.AutoSize = true;
            this.radioSortCustomerByPoint.Location = new System.Drawing.Point(192, 47);
            this.radioSortCustomerByPoint.Name = "radioSortCustomerByPoint";
            this.radioSortCustomerByPoint.Size = new System.Drawing.Size(133, 17);
            this.radioSortCustomerByPoint.TabIndex = 3;
            this.radioSortCustomerByPoint.TabStop = true;
            this.radioSortCustomerByPoint.Text = "Diểm tích lũy tăng dần";
            this.radioSortCustomerByPoint.UseVisualStyleBackColor = true;
            this.radioSortCustomerByPoint.CheckedChanged += new System.EventHandler(this.RadioSortCustomerCheckedChanged);
            // 
            // radioSortCustomerByBirthDate
            // 
            this.radioSortCustomerByBirthDate.AutoSize = true;
            this.radioSortCustomerByBirthDate.Location = new System.Drawing.Point(30, 132);
            this.radioSortCustomerByBirthDate.Name = "radioSortCustomerByBirthDate";
            this.radioSortCustomerByBirthDate.Size = new System.Drawing.Size(115, 17);
            this.radioSortCustomerByBirthDate.TabIndex = 2;
            this.radioSortCustomerByBirthDate.TabStop = true;
            this.radioSortCustomerByBirthDate.Text = "Ngày sinh tăng dần";
            this.radioSortCustomerByBirthDate.UseVisualStyleBackColor = true;
            this.radioSortCustomerByBirthDate.CheckedChanged += new System.EventHandler(this.RadioSortCustomerCheckedChanged);
            // 
            // radioSortCustomerByName
            // 
            this.radioSortCustomerByName.AutoSize = true;
            this.radioSortCustomerByName.Location = new System.Drawing.Point(30, 89);
            this.radioSortCustomerByName.Name = "radioSortCustomerByName";
            this.radioSortCustomerByName.Size = new System.Drawing.Size(143, 17);
            this.radioSortCustomerByName.TabIndex = 1;
            this.radioSortCustomerByName.TabStop = true;
            this.radioSortCustomerByName.Text = "Tên khách hàng tăng dần";
            this.radioSortCustomerByName.UseVisualStyleBackColor = true;
            this.radioSortCustomerByName.CheckedChanged += new System.EventHandler(this.RadioSortCustomerCheckedChanged);
            // 
            // radioSortCustomerById
            // 
            this.radioSortCustomerById.AutoSize = true;
            this.radioSortCustomerById.Location = new System.Drawing.Point(30, 47);
            this.radioSortCustomerById.Name = "radioSortCustomerById";
            this.radioSortCustomerById.Size = new System.Drawing.Size(141, 17);
            this.radioSortCustomerById.TabIndex = 0;
            this.radioSortCustomerById.TabStop = true;
            this.radioSortCustomerById.Text = "Mã khách hàng tăng dần";
            this.radioSortCustomerById.UseVisualStyleBackColor = true;
            this.radioSortCustomerById.CheckedChanged += new System.EventHandler(this.RadioSortCustomerCheckedChanged);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.button3);
            this.groupBox6.Controls.Add(this.button4);
            this.groupBox6.Controls.Add(this.btnRefreshCustomer);
            this.groupBox6.Controls.Add(this.btnAddNewCustomer);
            this.groupBox6.Location = new System.Drawing.Point(2, 411);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(304, 178);
            this.groupBox6.TabIndex = 5;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Hành động";
            // 
            // button3
            // 
            this.button3.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(155, 25);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(131, 35);
            this.button3.TabIndex = 5;
            this.button3.Text = "Xoá";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button4.Location = new System.Drawing.Point(18, 25);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(131, 35);
            this.button4.TabIndex = 4;
            this.button4.Text = "Sửa";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // btnRefreshCustomer
            // 
            this.btnRefreshCustomer.Image = global::CSharpCounterFinalProject.Properties.Resources.refresh;
            this.btnRefreshCustomer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRefreshCustomer.Location = new System.Drawing.Point(155, 71);
            this.btnRefreshCustomer.Name = "btnRefreshCustomer";
            this.btnRefreshCustomer.Size = new System.Drawing.Size(131, 35);
            this.btnRefreshCustomer.TabIndex = 1;
            this.btnRefreshCustomer.Text = "Làm mới";
            this.btnRefreshCustomer.UseVisualStyleBackColor = true;
            this.btnRefreshCustomer.Click += new System.EventHandler(this.BtnRefreshClick);
            // 
            // btnAddNewCustomer
            // 
            this.btnAddNewCustomer.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.btnAddNewCustomer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddNewCustomer.Location = new System.Drawing.Point(18, 71);
            this.btnAddNewCustomer.Name = "btnAddNewCustomer";
            this.btnAddNewCustomer.Size = new System.Drawing.Size(131, 35);
            this.btnAddNewCustomer.TabIndex = 0;
            this.btnAddNewCustomer.Text = "Thêm mới";
            this.btnAddNewCustomer.UseVisualStyleBackColor = true;
            this.btnAddNewCustomer.Click += new System.EventHandler(this.BtnAddNewCustomerClick);
            // 
            // tblKhachHang
            // 
            this.tblKhachHang.AllowUserToAddRows = false;
            this.tblKhachHang.AllowUserToDeleteRows = false;
            this.tblKhachHang.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.tblKhachHang.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.tblKhachHang.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.tblKhachHang.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblKhachHang.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Customer_ID,
            this.FullName,
            this.BirthDate,
            this.Address,
            this.PhoneNumber,
            this.CustomerType,
            this.Point,
            this.CreateTime,
            this.Email,
            this.tblCustomerEdit,
            this.tblCustomerRemove});
            this.tblKhachHang.Location = new System.Drawing.Point(-9, 0);
            this.tblKhachHang.Name = "tblKhachHang";
            this.tblKhachHang.RowHeadersWidth = 51;
            this.tblKhachHang.RowTemplate.Height = 24;
            this.tblKhachHang.Size = new System.Drawing.Size(937, 272);
            this.tblKhachHang.TabIndex = 4;
            this.tblKhachHang.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TblCustomerCellClick);
            // 
            // Customer_ID
            // 
            this.Customer_ID.HeaderText = "Ma KH";
            this.Customer_ID.MinimumWidth = 6;
            this.Customer_ID.Name = "Customer_ID";
            this.Customer_ID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // FullName
            // 
            this.FullName.HeaderText = "Họ tên KH";
            this.FullName.MinimumWidth = 6;
            this.FullName.Name = "FullName";
            this.FullName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // BirthDate
            // 
            this.BirthDate.HeaderText = "Ngày sinh";
            this.BirthDate.MinimumWidth = 6;
            this.BirthDate.Name = "BirthDate";
            this.BirthDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Address
            // 
            this.Address.HeaderText = "Địa chỉ";
            this.Address.MinimumWidth = 6;
            this.Address.Name = "Address";
            this.Address.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // PhoneNumber
            // 
            this.PhoneNumber.HeaderText = "Số ĐT";
            this.PhoneNumber.MinimumWidth = 6;
            this.PhoneNumber.Name = "PhoneNumber";
            this.PhoneNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // CustomerType
            // 
            this.CustomerType.HeaderText = "Loại KH";
            this.CustomerType.MinimumWidth = 6;
            this.CustomerType.Name = "CustomerType";
            this.CustomerType.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Point
            // 
            this.Point.HeaderText = "Điểm TL";
            this.Point.MinimumWidth = 6;
            this.Point.Name = "Point";
            this.Point.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // CreateTime
            // 
            this.CreateTime.HeaderText = "Ngày tạo TK";
            this.CreateTime.MinimumWidth = 6;
            this.CreateTime.Name = "CreateTime";
            this.CreateTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Email
            // 
            this.Email.HeaderText = "Email";
            this.Email.MinimumWidth = 6;
            this.Email.Name = "Email";
            this.Email.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // tblCustomerEdit
            // 
            this.tblCustomerEdit.HeaderText = "Sửa";
            this.tblCustomerEdit.MinimumWidth = 6;
            this.tblCustomerEdit.Name = "tblCustomerEdit";
            this.tblCustomerEdit.Text = "Sửa";
            this.tblCustomerEdit.UseColumnTextForButtonValue = true;
            // 
            // tblCustomerRemove
            // 
            this.tblCustomerRemove.HeaderText = "Xóa";
            this.tblCustomerRemove.MinimumWidth = 6;
            this.tblCustomerRemove.Name = "tblCustomerRemove";
            this.tblCustomerRemove.Text = "Xóa";
            this.tblCustomerRemove.UseColumnTextForButtonValue = true;
            // 
            // tabItem
            // 
            this.tabItem.Controls.Add(this.txtMoTa);
            this.tabItem.Controls.Add(this.label21);
            this.tabItem.Controls.Add(this.btnAnh);
            this.tabItem.Controls.Add(this.picAnh);
            this.tabItem.Controls.Add(this.txtPhanLoai);
            this.tabItem.Controls.Add(this.label20);
            this.tabItem.Controls.Add(this.txtGia);
            this.tabItem.Controls.Add(this.label19);
            this.tabItem.Controls.Add(this.txtTonKho);
            this.tabItem.Controls.Add(this.label18);
            this.tabItem.Controls.Add(this.txtTenSP);
            this.tabItem.Controls.Add(this.label17);
            this.tabItem.Controls.Add(this.txtHangSX);
            this.tabItem.Controls.Add(this.label16);
            this.tabItem.Controls.Add(this.txtMaSP);
            this.tabItem.Controls.Add(this.label15);
            this.tabItem.Controls.Add(this.tblDuLieu);
            this.tabItem.Controls.Add(this.groupBox3);
            this.tabItem.Controls.Add(this.groupBox2);
            this.tabItem.Controls.Add(this.groupBox1);
            this.tabItem.Location = new System.Drawing.Point(4, 22);
            this.tabItem.Name = "tabItem";
            this.tabItem.Padding = new System.Windows.Forms.Padding(3);
            this.tabItem.Size = new System.Drawing.Size(931, 580);
            this.tabItem.TabIndex = 0;
            this.tabItem.Text = "QL Mặt Hàng";
            this.tabItem.UseVisualStyleBackColor = true;
            // 
            // txtMoTa
            // 
            this.txtMoTa.Location = new System.Drawing.Point(448, 320);
            this.txtMoTa.Multiline = true;
            this.txtMoTa.Name = "txtMoTa";
            this.txtMoTa.Size = new System.Drawing.Size(172, 84);
            this.txtMoTa.TabIndex = 20;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(398, 323);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(44, 13);
            this.label21.TabIndex = 19;
            this.label21.Text = "Mô tả : ";
            // 
            // btnAnh
            // 
            this.btnAnh.Location = new System.Drawing.Point(678, 320);
            this.btnAnh.Name = "btnAnh";
            this.btnAnh.Size = new System.Drawing.Size(75, 23);
            this.btnAnh.TabIndex = 18;
            this.btnAnh.Text = "Ảnh";
            this.btnAnh.UseVisualStyleBackColor = true;
            // 
            // picAnh
            // 
            this.picAnh.Location = new System.Drawing.Point(780, 307);
            this.picAnh.Name = "picAnh";
            this.picAnh.Size = new System.Drawing.Size(100, 93);
            this.picAnh.TabIndex = 17;
            this.picAnh.TabStop = false;
            // 
            // txtPhanLoai
            // 
            this.txtPhanLoai.Location = new System.Drawing.Point(93, 388);
            this.txtPhanLoai.Name = "txtPhanLoai";
            this.txtPhanLoai.Size = new System.Drawing.Size(97, 20);
            this.txtPhanLoai.TabIndex = 16;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(15, 391);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(50, 13);
            this.label20.TabIndex = 15;
            this.label20.Text = "Phân loại";
            // 
            // txtGia
            // 
            this.txtGia.Location = new System.Drawing.Point(274, 384);
            this.txtGia.Name = "txtGia";
            this.txtGia.Size = new System.Drawing.Size(68, 20);
            this.txtGia.TabIndex = 14;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(221, 387);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(23, 13);
            this.label19.TabIndex = 13;
            this.label19.Text = "Giá";
            // 
            // txtTonKho
            // 
            this.txtTonKho.Location = new System.Drawing.Point(274, 351);
            this.txtTonKho.Name = "txtTonKho";
            this.txtTonKho.Size = new System.Drawing.Size(68, 20);
            this.txtTonKho.TabIndex = 12;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(221, 354);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(47, 13);
            this.label18.TabIndex = 11;
            this.label18.Text = "Tồn kho";
            // 
            // txtTenSP
            // 
            this.txtTenSP.Location = new System.Drawing.Point(93, 351);
            this.txtTenSP.Name = "txtTenSP";
            this.txtTenSP.Size = new System.Drawing.Size(97, 20);
            this.txtTenSP.TabIndex = 10;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(15, 354);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(72, 13);
            this.label17.TabIndex = 9;
            this.label17.Text = "Tên sản phẩm";
            // 
            // txtHangSX
            // 
            this.txtHangSX.Location = new System.Drawing.Point(274, 317);
            this.txtHangSX.Name = "txtHangSX";
            this.txtHangSX.Size = new System.Drawing.Size(68, 20);
            this.txtHangSX.TabIndex = 8;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(221, 320);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(49, 13);
            this.label16.TabIndex = 7;
            this.label16.Text = "Hãng SX";
            // 
            // txtMaSP
            // 
            this.txtMaSP.Location = new System.Drawing.Point(93, 314);
            this.txtMaSP.Name = "txtMaSP";
            this.txtMaSP.Size = new System.Drawing.Size(68, 20);
            this.txtMaSP.TabIndex = 6;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(15, 317);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(70, 13);
            this.label15.TabIndex = 5;
            this.label15.Text = "Mã sản phẩm";
            // 
            // tblDuLieu
            // 
            this.tblDuLieu.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.tblDuLieu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tblDuLieu.Location = new System.Drawing.Point(3, 6);
            this.tblDuLieu.Name = "tblDuLieu";
            this.tblDuLieu.ReadOnly = true;
            this.tblDuLieu.RowHeadersWidth = 92;
            this.tblDuLieu.RowTemplate.Height = 37;
            this.tblDuLieu.Size = new System.Drawing.Size(910, 295);
            this.tblDuLieu.TabIndex = 4;
            this.tblDuLieu.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.tblDuLieu_CellDoubleClick);
            // 
            // groupBox3
            // 
            this.groupBox3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.groupBox3.Controls.Add(this.numericItemTo);
            this.groupBox3.Controls.Add(this.numericItemFrom);
            this.groupBox3.Controls.Add(this.btnSearchItem);
            this.groupBox3.Controls.Add(this.txtSearchItem);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.comboSearchItem);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Location = new System.Drawing.Point(647, 414);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(285, 169);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Tìm kiếm";
            // 
            // numericItemTo
            // 
            this.numericItemTo.Location = new System.Drawing.Point(198, 94);
            this.numericItemTo.Maximum = new decimal(new int[] {
            99999999,
            0,
            0,
            0});
            this.numericItemTo.Name = "numericItemTo";
            this.numericItemTo.Size = new System.Drawing.Size(66, 20);
            this.numericItemTo.TabIndex = 8;
            // 
            // numericItemFrom
            // 
            this.numericItemFrom.Location = new System.Drawing.Point(83, 93);
            this.numericItemFrom.Maximum = new decimal(new int[] {
            99999999,
            0,
            0,
            0});
            this.numericItemFrom.Name = "numericItemFrom";
            this.numericItemFrom.Size = new System.Drawing.Size(63, 20);
            this.numericItemFrom.TabIndex = 7;
            // 
            // btnSearchItem
            // 
            this.btnSearchItem.Image = global::CSharpCounterFinalProject.Properties.Resources.search;
            this.btnSearchItem.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSearchItem.Location = new System.Drawing.Point(83, 122);
            this.btnSearchItem.Name = "btnSearchItem";
            this.btnSearchItem.Size = new System.Drawing.Size(131, 41);
            this.btnSearchItem.TabIndex = 6;
            this.btnSearchItem.Text = "Tìm kiếm";
            this.btnSearchItem.UseVisualStyleBackColor = true;
            this.btnSearchItem.Click += new System.EventHandler(this.BtnSearchItemClick);
            // 
            // txtSearchItem
            // 
            this.txtSearchItem.Location = new System.Drawing.Point(83, 63);
            this.txtSearchItem.Name = "txtSearchItem";
            this.txtSearchItem.Size = new System.Drawing.Size(180, 20);
            this.txtSearchItem.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(163, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "Đến:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(50, 93);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(23, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Từ:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Nội dung";
            // 
            // comboSearchItem
            // 
            this.comboSearchItem.FormattingEnabled = true;
            this.comboSearchItem.Items.AddRange(new object[] {
            "Tìm theo tên mặt hàng.",
            "Tìm theo khoảng giá a-b.",
            "Tìm theo loại mặt hàng.",
            "Tìm theo hãng sản xuất.",
            "Theo số lượng khoảng a-b."});
            this.comboSearchItem.Location = new System.Drawing.Point(83, 25);
            this.comboSearchItem.Name = "comboSearchItem";
            this.comboSearchItem.Size = new System.Drawing.Size(180, 21);
            this.comboSearchItem.TabIndex = 1;
            this.comboSearchItem.SelectedIndexChanged += new System.EventHandler(this.ComboSearchItemSelectefIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Tiêu chí";
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.groupBox2.Controls.Add(this.radioSortItemByProductDate);
            this.groupBox2.Controls.Add(this.radioSortItemByQuantity);
            this.groupBox2.Controls.Add(this.radioSortItemByName);
            this.groupBox2.Controls.Add(this.radioSortItemByPriceDESC);
            this.groupBox2.Controls.Add(this.radioSortItemByPriceASC);
            this.groupBox2.Location = new System.Drawing.Point(303, 414);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(346, 178);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Sắp xếp theo";
            // 
            // radioSortItemByProductDate
            // 
            this.radioSortItemByProductDate.AutoSize = true;
            this.radioSortItemByProductDate.Location = new System.Drawing.Point(192, 89);
            this.radioSortItemByProductDate.Name = "radioSortItemByProductDate";
            this.radioSortItemByProductDate.Size = new System.Drawing.Size(134, 17);
            this.radioSortItemByProductDate.TabIndex = 4;
            this.radioSortItemByProductDate.TabStop = true;
            this.radioSortItemByProductDate.Text = "Ngày sản xuất tăng dần";
            this.radioSortItemByProductDate.UseVisualStyleBackColor = true;
            this.radioSortItemByProductDate.CheckedChanged += new System.EventHandler(this.SortItemHandler);
            // 
            // radioSortItemByQuantity
            // 
            this.radioSortItemByQuantity.AutoSize = true;
            this.radioSortItemByQuantity.Location = new System.Drawing.Point(192, 47);
            this.radioSortItemByQuantity.Name = "radioSortItemByQuantity";
            this.radioSortItemByQuantity.Size = new System.Drawing.Size(116, 17);
            this.radioSortItemByQuantity.TabIndex = 3;
            this.radioSortItemByQuantity.TabStop = true;
            this.radioSortItemByQuantity.Text = "Số lượng giảm dần";
            this.radioSortItemByQuantity.UseVisualStyleBackColor = true;
            this.radioSortItemByQuantity.CheckedChanged += new System.EventHandler(this.SortItemHandler);
            // 
            // radioSortItemByName
            // 
            this.radioSortItemByName.AutoSize = true;
            this.radioSortItemByName.Location = new System.Drawing.Point(30, 132);
            this.radioSortItemByName.Name = "radioSortItemByName";
            this.radioSortItemByName.Size = new System.Drawing.Size(107, 17);
            this.radioSortItemByName.TabIndex = 2;
            this.radioSortItemByName.TabStop = true;
            this.radioSortItemByName.Text = "Tên mặt hàng a-z";
            this.radioSortItemByName.UseVisualStyleBackColor = true;
            this.radioSortItemByName.CheckedChanged += new System.EventHandler(this.SortItemHandler);
            // 
            // radioSortItemByPriceDESC
            // 
            this.radioSortItemByPriceDESC.AutoSize = true;
            this.radioSortItemByPriceDESC.Location = new System.Drawing.Point(30, 89);
            this.radioSortItemByPriceDESC.Name = "radioSortItemByPriceDESC";
            this.radioSortItemByPriceDESC.Size = new System.Drawing.Size(134, 17);
            this.radioSortItemByPriceDESC.TabIndex = 1;
            this.radioSortItemByPriceDESC.TabStop = true;
            this.radioSortItemByPriceDESC.Text = "Giá niêm yếu giảm dần";
            this.radioSortItemByPriceDESC.UseVisualStyleBackColor = true;
            this.radioSortItemByPriceDESC.CheckedChanged += new System.EventHandler(this.SortItemHandler);
            // 
            // radioSortItemByPriceASC
            // 
            this.radioSortItemByPriceASC.AutoSize = true;
            this.radioSortItemByPriceASC.Location = new System.Drawing.Point(30, 47);
            this.radioSortItemByPriceASC.Name = "radioSortItemByPriceASC";
            this.radioSortItemByPriceASC.Size = new System.Drawing.Size(129, 17);
            this.radioSortItemByPriceASC.TabIndex = 0;
            this.radioSortItemByPriceASC.TabStop = true;
            this.radioSortItemByPriceASC.Text = "Giá niêm yết tăng dần";
            this.radioSortItemByPriceASC.UseVisualStyleBackColor = true;
            this.radioSortItemByPriceASC.CheckedChanged += new System.EventHandler(this.SortItemHandler);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.groupBox1.Controls.Add(this.btnXoaSP);
            this.groupBox1.Controls.Add(this.btnSuaSP);
            this.groupBox1.Controls.Add(this.btnFreshItem);
            this.groupBox1.Controls.Add(this.btnAddNewItem);
            this.groupBox1.Location = new System.Drawing.Point(0, 414);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(304, 178);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Hành động";
            // 
            // btnXoaSP
            // 
            this.btnXoaSP.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.btnXoaSP.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnXoaSP.Location = new System.Drawing.Point(155, 21);
            this.btnXoaSP.Name = "btnXoaSP";
            this.btnXoaSP.Size = new System.Drawing.Size(131, 35);
            this.btnXoaSP.TabIndex = 3;
            this.btnXoaSP.Text = "Xoá";
            this.btnXoaSP.UseVisualStyleBackColor = true;
            // 
            // btnSuaSP
            // 
            this.btnSuaSP.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.btnSuaSP.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSuaSP.Location = new System.Drawing.Point(18, 21);
            this.btnSuaSP.Name = "btnSuaSP";
            this.btnSuaSP.Size = new System.Drawing.Size(131, 35);
            this.btnSuaSP.TabIndex = 2;
            this.btnSuaSP.Text = "Sửa";
            this.btnSuaSP.UseVisualStyleBackColor = true;
            // 
            // btnFreshItem
            // 
            this.btnFreshItem.Image = global::CSharpCounterFinalProject.Properties.Resources.refresh;
            this.btnFreshItem.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFreshItem.Location = new System.Drawing.Point(155, 71);
            this.btnFreshItem.Name = "btnFreshItem";
            this.btnFreshItem.Size = new System.Drawing.Size(131, 35);
            this.btnFreshItem.TabIndex = 1;
            this.btnFreshItem.Text = "Làm mới";
            this.btnFreshItem.UseVisualStyleBackColor = true;
            this.btnFreshItem.Click += new System.EventHandler(this.BtnRefreshClick);
            // 
            // btnAddNewItem
            // 
            this.btnAddNewItem.Image = global::CSharpCounterFinalProject.Properties.Resources.plus;
            this.btnAddNewItem.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddNewItem.Location = new System.Drawing.Point(18, 71);
            this.btnAddNewItem.Name = "btnAddNewItem";
            this.btnAddNewItem.Size = new System.Drawing.Size(131, 35);
            this.btnAddNewItem.TabIndex = 0;
            this.btnAddNewItem.Text = "Thêm mới";
            this.btnAddNewItem.UseVisualStyleBackColor = true;
            this.btnAddNewItem.Click += new System.EventHandler(this.BtnAddNewClick);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabItem);
            this.tabControl1.Controls.Add(this.tabCustomer);
            this.tabControl1.Controls.Add(this.tabDiscount);
            this.tabControl1.Controls.Add(this.tabBill);
            this.tabControl1.Controls.Add(this.tabStat);
            this.tabControl1.Location = new System.Drawing.Point(0, 27);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(939, 606);
            this.tabControl1.TabIndex = 0;
            // 
            // groupBox12
            // 
            this.groupBox12.Location = new System.Drawing.Point(2, 305);
            this.groupBox12.Name = "groupBox12";
            this.groupBox12.Size = new System.Drawing.Size(923, 100);
            this.groupBox12.TabIndex = 8;
            this.groupBox12.TabStop = false;
            this.groupBox12.Text = "Thông tin chung";
            // 
            // HomeFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.ClientSize = new System.Drawing.Size(935, 633);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Constantia", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "HomeFrm";
            this.Text = "Quản Lý Bán Hàng";
            this.Load += new System.EventHandler(this.HomeFrm_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.tabStat.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tbl_month_bestsold)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tbl_revenue_month)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblStatCustomer)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblStatItem)).EndInit();
            this.groupBox11.ResumeLayout(false);
            this.groupBox11.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMonth)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDay)).EndInit();
            this.tabBill.ResumeLayout(false);
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.groupBox10.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tblHoaDon)).EndInit();
            this.tabDiscount.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tblKhuyenMai)).EndInit();
            this.tabCustomer.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tblKhachHang)).EndInit();
            this.tabItem.ResumeLayout(false);
            this.tabItem.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picAnh)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblDuLieu)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericItemTo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericItemFrom)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        

        #endregion
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem thaoTácToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem thêmMớiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mAddItem;
        private System.Windows.Forms.ToolStripMenuItem mAddCustomer;
        private System.Windows.Forms.ToolStripMenuItem mAddDiscount;
        private System.Windows.Forms.ToolStripMenuItem mAddBill;
        private System.Windows.Forms.ToolStripMenuItem btn_SaveFile;
        private System.Windows.Forms.ToolStripMenuItem xemToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem thoátToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mItemTab;
        private System.Windows.Forms.ToolStripMenuItem mCustomerTab;
        private System.Windows.Forms.ToolStripMenuItem mDiscountTab;
        private System.Windows.Forms.ToolStripMenuItem mBillTab;
        private TabPage tabStat;
        private DataGridView tbl_month_bestsold;
        private DataGridViewTextBoxColumn Month_bestsold;
        private DataGridView tbl_revenue_month;
        private DataGridViewTextBoxColumn revenue_month;
        private DataGridView tblStatCustomer;
        private DataGridViewTextBoxColumn stat_customer_Id;
        private DataGridViewTextBoxColumn stat_customer_name;
        private DataGridViewTextBoxColumn stat_customer_times;
        private DataGridViewTextBoxColumn stat_customer_quantity;
        private DataGridViewTextBoxColumn stat_customer_total;
        private DataGridView tblStatItem;
        private DataGridViewTextBoxColumn inven_id;
        private DataGridViewTextBoxColumn inven_name;
        private DataGridViewTextBoxColumn inven_quantity;
        private DataGridViewTextBoxColumn inven_revenue;
        private GroupBox groupBox11;
        private Button button2;
        private Button button1;
        private NumericUpDown nudYear;
        private NumericUpDown nudMonth;
        private NumericUpDown nudDay;
        private Button btnStatResult;
        private Label label14;
        private Label label13;
        private Label label11;
        private ComboBox comboStat;
        private Label label12;
        private TabPage tabBill;
        private GroupBox groupBox8;
        private Button btnSearchBill;
        private TextBox txtSearchBill;
        private Label label9;
        private ComboBox comboSearchBill;
        private Label label10;
        private GroupBox groupBox10;
        private Button btnRefreshBill;
        private Button btnAddNewBill;
        private DataGridView tblHoaDon;
        private TabPage tabDiscount;
        private GroupBox groupBox7;
        private Button btnSearchDiscount;
        private TextBox txtSearchDiscount;
        private Label label5;
        private ComboBox comboSearchDiscount;
        private Label label6;
        private GroupBox groupBox9;
        private Button btnRefreshDiscount;
        private Button btnAddNeuDiscount;
        private DataGridView tblKhuyenMai;
        private DataGridViewTextBoxColumn Discount_ID;
        private DataGridViewTextBoxColumn NameDiscount;
        private DataGridViewTextBoxColumn StartTime;
        private DataGridViewTextBoxColumn EndTime;
        private DataGridViewTextBoxColumn DiscountType;
        private DataGridViewTextBoxColumn DiscountPercent;
        private DataGridViewTextBoxColumn DiscountPriceAmount;
        private DataGridViewTextBoxColumn MinPriceToUseDiscount;
        private DataGridViewButtonColumn tblDiscountEdit;
        private DataGridViewButtonColumn tblDiscountRemove;
        private TabPage tabCustomer;
        private GroupBox groupBox4;
        private Button btnSearchCustomer;
        private TextBox txtSearchCustomer;
        private Label label7;
        private ComboBox comboSearchCustomer;
        private Label label8;
        private GroupBox groupBox5;
        private RadioButton radioSortCustomerByCreatedDate;
        private RadioButton radioSortCustomerByPoint;
        private RadioButton radioSortCustomerByBirthDate;
        private RadioButton radioSortCustomerByName;
        private RadioButton radioSortCustomerById;
        private GroupBox groupBox6;
        private Button btnRefreshCustomer;
        private Button btnAddNewCustomer;
        private DataGridView tblKhachHang;
        private DataGridViewTextBoxColumn Customer_ID;
        private DataGridViewTextBoxColumn FullName;
        private DataGridViewTextBoxColumn BirthDate;
        private DataGridViewTextBoxColumn Address;
        private DataGridViewTextBoxColumn PhoneNumber;
        private DataGridViewTextBoxColumn CustomerType;
        private DataGridViewTextBoxColumn Point;
        private DataGridViewTextBoxColumn CreateTime;
        private DataGridViewTextBoxColumn Email;
        private DataGridViewButtonColumn tblCustomerEdit;
        private DataGridViewButtonColumn tblCustomerRemove;
        private TabPage tabItem;
        private GroupBox groupBox3;
        private NumericUpDown numericItemTo;
        private NumericUpDown numericItemFrom;
        private Button btnSearchItem;
        private TextBox txtSearchItem;
        private Label label4;
        private Label label3;
        private Label label2;
        private ComboBox comboSearchItem;
        private Label label1;
        private GroupBox groupBox2;
        private RadioButton radioSortItemByProductDate;
        private RadioButton radioSortItemByQuantity;
        private RadioButton radioSortItemByName;
        private RadioButton radioSortItemByPriceDESC;
        private RadioButton radioSortItemByPriceASC;
        private GroupBox groupBox1;
        private Button btnFreshItem;
        private Button btnAddNewItem;
        private TabControl tabControl1;
        private DataGridViewTextBoxColumn Bill_ID;
        private DataGridViewTextBoxColumn FullNameBill;
        private DataGridViewTextBoxColumn StaffName;
        private DataGridViewTextBoxColumn CreateTimeBill;
        private DataGridViewTextBoxColumn TotalItem;
        private DataGridViewTextBoxColumn SubTotal;
        private DataGridViewTextBoxColumn TotalDiscountAmount;
        private DataGridViewTextBoxColumn TotalAmount;
        private DataGridViewTextBoxColumn Status;
        private DataGridViewButtonColumn outFile;
        private DataGridView tblDuLieu;
        private Button button3;
        private Button button4;
        private Button btnXoaSP;
        private Button btnSuaSP;
        private Button button5;
        private Button button6;
        private TextBox txtTonKho;
        private Label label18;
        private TextBox txtTenSP;
        private Label label17;
        private TextBox txtHangSX;
        private Label label16;
        private TextBox txtMaSP;
        private Label label15;
        private TextBox txtMoTa;
        private Label label21;
        private Button btnAnh;
        private PictureBox picAnh;
        private TextBox txtPhanLoai;
        private Label label20;
        private TextBox txtGia;
        private Label label19;
        private GroupBox groupBox12;
    }
}

