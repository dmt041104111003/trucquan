namespace CSharpCounterFinalProject.ViewCustomer
{
    partial class HomeCustomerView
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HomeCustomerView));
            this.btnBuyItem = new System.Windows.Forms.Button();
            this.btnInfor = new System.Windows.Forms.Button();
            this.btnSignOut = new System.Windows.Forms.Button();
            this.labelName = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.WMPlayer = new AxWMPLib.AxWindowsMediaPlayer();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.WMPlayer)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBuyItem
            // 
            this.btnBuyItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnBuyItem.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnBuyItem.FlatAppearance.BorderSize = 2;
            this.btnBuyItem.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBuyItem.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnBuyItem.Location = new System.Drawing.Point(28, 139);
            this.btnBuyItem.Name = "btnBuyItem";
            this.btnBuyItem.Size = new System.Drawing.Size(206, 50);
            this.btnBuyItem.TabIndex = 1;
            this.btnBuyItem.Text = "Mua sắm";
            this.btnBuyItem.UseVisualStyleBackColor = false;
            this.btnBuyItem.Click += new System.EventHandler(this.btnBuyItem_Click);
            // 
            // btnInfor
            // 
            this.btnInfor.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnInfor.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnInfor.FlatAppearance.BorderSize = 2;
            this.btnInfor.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInfor.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnInfor.Location = new System.Drawing.Point(28, 229);
            this.btnInfor.Name = "btnInfor";
            this.btnInfor.Size = new System.Drawing.Size(206, 50);
            this.btnInfor.TabIndex = 2;
            this.btnInfor.Text = "Thông tin tài khoản";
            this.btnInfor.UseVisualStyleBackColor = false;
            this.btnInfor.Click += new System.EventHandler(this.btnInfor_Click);
            // 
            // btnSignOut
            // 
            this.btnSignOut.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(157)))), ((int)(((byte)(192)))), ((int)(((byte)(139)))));
            this.btnSignOut.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSignOut.FlatAppearance.BorderSize = 2;
            this.btnSignOut.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSignOut.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnSignOut.Location = new System.Drawing.Point(28, 324);
            this.btnSignOut.Name = "btnSignOut";
            this.btnSignOut.Size = new System.Drawing.Size(206, 50);
            this.btnSignOut.TabIndex = 3;
            this.btnSignOut.Text = "Đăng xuất";
            this.btnSignOut.UseVisualStyleBackColor = false;
            this.btnSignOut.Click += new System.EventHandler(this.btnSignOut_Click);
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Font = new System.Drawing.Font("Tahoma", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.labelName.Location = new System.Drawing.Point(182, 50);
            this.labelName.Name = "labelName";
            this.labelName.Size = new System.Drawing.Size(52, 27);
            this.labelName.TabIndex = 4;
            this.labelName.Text = "<:))";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.label1.Location = new System.Drawing.Point(73, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(103, 27);
            this.label1.TabIndex = 5;
            this.label1.Text = "Xin chào:";
            // 
            // WMPlayer
            // 
            this.WMPlayer.Enabled = true;
            this.WMPlayer.Location = new System.Drawing.Point(279, 139);
            this.WMPlayer.Name = "WMPlayer";
            this.WMPlayer.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("WMPlayer.OcxState")));
            this.WMPlayer.Size = new System.Drawing.Size(780, 488);
            this.WMPlayer.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(96)))), ((int)(((byte)(153)))), ((int)(((byte)(102)))));
            this.label2.Location = new System.Drawing.Point(346, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(662, 51);
            this.label2.TabIndex = 7;
            this.label2.Text = "Chào mừng bạn đến với cửa hàng!";
            // 
            // HomeCustomerView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(241)))), ((int)(((byte)(214)))));
            this.ClientSize = new System.Drawing.Size(1169, 653);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.WMPlayer);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.labelName);
            this.Controls.Add(this.btnSignOut);
            this.Controls.Add(this.btnInfor);
            this.Controls.Add(this.btnBuyItem);
            this.Font = new System.Drawing.Font("Constantia", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "HomeCustomerView";
            this.Text = "Trang chủ";
            this.Load += new System.EventHandler(this.HomeCustomerView_Load);
            ((System.ComponentModel.ISupportInitialize)(this.WMPlayer)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnBuyItem;
        private System.Windows.Forms.Button btnInfor;
        private System.Windows.Forms.Button btnSignOut;
        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.Label label1;
        private AxWMPLib.AxWindowsMediaPlayer WMPlayer;
        private System.Windows.Forms.Label label2;
    }
}