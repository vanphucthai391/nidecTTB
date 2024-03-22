namespace BoxIdDb
{
    partial class frmBox
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmBox));
            this.btnAddBoxId = new System.Windows.Forms.Button();
            this.btnSearchBoxId = new System.Windows.Forms.Button();
            this.dgvBoxId = new System.Windows.Forms.DataGridView();
            this.colUpdateInvoice = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.col_boxid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_user = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_regist_date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_update_ship = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.col_ship_date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.col_invoice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label12 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtBoxIdFrom = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtProductSerial = new System.Windows.Forms.TextBox();
            this.dtpRegistDate = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.pnlBarcode = new System.Windows.Forms.Panel();
            this.btnCancel = new System.Windows.Forms.Button();
            this.rdbBoxId = new System.Windows.Forms.RadioButton();
            this.rdbPrintDate = new System.Windows.Forms.RadioButton();
            this.rdbProductSerial = new System.Windows.Forms.RadioButton();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnEditShipping = new System.Windows.Forms.Button();
            this.rdbShipDate = new System.Windows.Forms.RadioButton();
            this.dtpShipDate = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.txtBoxIdTo = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnUpInv = new System.Windows.Forms.Button();
            this.pnlInvoice = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.txtInvoice = new System.Windows.Forms.TextBox();
            this.btnAdd517 = new System.Windows.Forms.Button();
            this.btnAddBoxID523 = new System.Windows.Forms.Button();
            this.btnAddBoxID517FB = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnAddBoxBFB_0025 = new System.Windows.Forms.Button();
            this.btnAddBoxBMA_0148 = new System.Windows.Forms.Button();
            this.btnAddBoxLD = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBoxId)).BeginInit();
            this.pnlInvoice.SuspendLayout();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnAddBoxId
            // 
            this.btnAddBoxId.Location = new System.Drawing.Point(30, 53);
            this.btnAddBoxId.Name = "btnAddBoxId";
            this.btnAddBoxId.Size = new System.Drawing.Size(114, 25);
            this.btnAddBoxId.TabIndex = 6;
            this.btnAddBoxId.Text = "Add Box ID";
            this.btnAddBoxId.UseVisualStyleBackColor = true;
            this.btnAddBoxId.Click += new System.EventHandler(this.btnAddBoxId_Click);
            // 
            // btnSearchBoxId
            // 
            this.btnSearchBoxId.Location = new System.Drawing.Point(30, 22);
            this.btnSearchBoxId.Name = "btnSearchBoxId";
            this.btnSearchBoxId.Size = new System.Drawing.Size(80, 25);
            this.btnSearchBoxId.TabIndex = 2;
            this.btnSearchBoxId.Text = "Search";
            this.btnSearchBoxId.UseVisualStyleBackColor = true;
            this.btnSearchBoxId.Click += new System.EventHandler(this.btnSearchBoxId_Click);
            // 
            // dgvBoxId
            // 
            this.dgvBoxId.AllowUserToAddRows = false;
            this.dgvBoxId.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvBoxId.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvBoxId.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvBoxId.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colUpdateInvoice,
            this.col_boxid,
            this.col_user,
            this.col_regist_date,
            this.col_update_ship,
            this.col_ship_date,
            this.col_invoice});
            this.dgvBoxId.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvBoxId.Location = new System.Drawing.Point(0, 0);
            this.dgvBoxId.Name = "dgvBoxId";
            this.dgvBoxId.RowTemplate.Height = 21;
            this.dgvBoxId.Size = new System.Drawing.Size(980, 409);
            this.dgvBoxId.TabIndex = 9;
            this.dgvBoxId.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBoxId_CellContentClick);
            // 
            // colUpdateInvoice
            // 
            this.colUpdateInvoice.HeaderText = "";
            this.colUpdateInvoice.Name = "colUpdateInvoice";
            // 
            // col_boxid
            // 
            this.col_boxid.DataPropertyName = "boxid";
            this.col_boxid.HeaderText = "BoxID";
            this.col_boxid.Name = "col_boxid";
            // 
            // col_user
            // 
            this.col_user.DataPropertyName = "suser";
            this.col_user.HeaderText = "User";
            this.col_user.Name = "col_user";
            // 
            // col_regist_date
            // 
            this.col_regist_date.DataPropertyName = "regist_date";
            this.col_regist_date.HeaderText = "Regist Date";
            this.col_regist_date.Name = "col_regist_date";
            // 
            // col_update_ship
            // 
            this.col_update_ship.HeaderText = "";
            this.col_update_ship.Name = "col_update_ship";
            this.col_update_ship.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.col_update_ship.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // col_ship_date
            // 
            this.col_ship_date.DataPropertyName = "shipdate";
            this.col_ship_date.HeaderText = "Ship Date";
            this.col_ship_date.Name = "col_ship_date";
            // 
            // col_invoice
            // 
            this.col_invoice.DataPropertyName = "invoice";
            this.col_invoice.HeaderText = "Invoice";
            this.col_invoice.Name = "col_invoice";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(27, 50);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(69, 13);
            this.label12.TabIndex = 6;
            this.label12.Text = "Regist Date: ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Box ID from: ";
            // 
            // txtBoxIdFrom
            // 
            this.txtBoxIdFrom.Location = new System.Drawing.Point(125, 10);
            this.txtBoxIdFrom.Name = "txtBoxIdFrom";
            this.txtBoxIdFrom.Size = new System.Drawing.Size(166, 20);
            this.txtBoxIdFrom.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Product Serial: ";
            // 
            // txtProductSerial
            // 
            this.txtProductSerial.Location = new System.Drawing.Point(125, 81);
            this.txtProductSerial.Name = "txtProductSerial";
            this.txtProductSerial.Size = new System.Drawing.Size(166, 20);
            this.txtProductSerial.TabIndex = 1;
            // 
            // dtpRegistDate
            // 
            this.dtpRegistDate.CustomFormat = "yyyy/MM/dd";
            this.dtpRegistDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpRegistDate.Location = new System.Drawing.Point(125, 44);
            this.dtpRegistDate.Name = "dtpRegistDate";
            this.dtpRegistDate.Size = new System.Drawing.Size(166, 20);
            this.dtpRegistDate.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(405, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "User: ";
            // 
            // txtUser
            // 
            this.txtUser.Enabled = false;
            this.txtUser.Location = new System.Drawing.Point(446, 47);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(166, 20);
            this.txtUser.TabIndex = 1;
            // 
            // pnlBarcode
            // 
            this.pnlBarcode.BackColor = System.Drawing.Color.White;
            this.pnlBarcode.Location = new System.Drawing.Point(406, 85);
            this.pnlBarcode.Name = "pnlBarcode";
            this.pnlBarcode.Size = new System.Drawing.Size(206, 57);
            this.pnlBarcode.TabIndex = 11;
            this.pnlBarcode.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlBarcode_Paint);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(536, 23);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 25);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Close";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // rdbBoxId
            // 
            this.rdbBoxId.AutoSize = true;
            this.rdbBoxId.Location = new System.Drawing.Point(312, 12);
            this.rdbBoxId.Name = "rdbBoxId";
            this.rdbBoxId.Size = new System.Drawing.Size(14, 13);
            this.rdbBoxId.TabIndex = 12;
            this.rdbBoxId.UseVisualStyleBackColor = true;
            this.rdbBoxId.CheckedChanged += new System.EventHandler(this.rdbBoxId_CheckedChanged);
            // 
            // rdbPrintDate
            // 
            this.rdbPrintDate.AutoSize = true;
            this.rdbPrintDate.Checked = true;
            this.rdbPrintDate.Location = new System.Drawing.Point(313, 48);
            this.rdbPrintDate.Name = "rdbPrintDate";
            this.rdbPrintDate.Size = new System.Drawing.Size(14, 13);
            this.rdbPrintDate.TabIndex = 12;
            this.rdbPrintDate.TabStop = true;
            this.rdbPrintDate.UseVisualStyleBackColor = true;
            this.rdbPrintDate.CheckedChanged += new System.EventHandler(this.rdbPrintDate_CheckedChanged);
            // 
            // rdbProductSerial
            // 
            this.rdbProductSerial.AutoSize = true;
            this.rdbProductSerial.Location = new System.Drawing.Point(313, 82);
            this.rdbProductSerial.Name = "rdbProductSerial";
            this.rdbProductSerial.Size = new System.Drawing.Size(14, 13);
            this.rdbProductSerial.TabIndex = 12;
            this.rdbProductSerial.UseVisualStyleBackColor = true;
            this.rdbProductSerial.CheckedChanged += new System.EventHandler(this.rdbProductSerial_CheckedChanged);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(156, 22);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(80, 25);
            this.btnExport.TabIndex = 42;
            this.btnExport.Text = "Excel Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnEditShipping
            // 
            this.btnEditShipping.Location = new System.Drawing.Point(105, 2);
            this.btnEditShipping.Name = "btnEditShipping";
            this.btnEditShipping.Size = new System.Drawing.Size(100, 25);
            this.btnEditShipping.TabIndex = 43;
            this.btnEditShipping.Text = "Edit Shipping";
            this.btnEditShipping.UseVisualStyleBackColor = true;
            this.btnEditShipping.Click += new System.EventHandler(this.btnEditShipping_Click);
            // 
            // rdbShipDate
            // 
            this.rdbShipDate.AutoSize = true;
            this.rdbShipDate.Location = new System.Drawing.Point(313, 119);
            this.rdbShipDate.Name = "rdbShipDate";
            this.rdbShipDate.Size = new System.Drawing.Size(14, 13);
            this.rdbShipDate.TabIndex = 46;
            this.rdbShipDate.UseVisualStyleBackColor = true;
            this.rdbShipDate.CheckedChanged += new System.EventHandler(this.rdbShipDate_CheckedChanged_1);
            // 
            // dtpShipDate
            // 
            this.dtpShipDate.CustomFormat = "yyyy/MM/dd HH:mm:ss";
            this.dtpShipDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpShipDate.Location = new System.Drawing.Point(125, 119);
            this.dtpShipDate.Name = "dtpShipDate";
            this.dtpShipDate.ShowUpDown = true;
            this.dtpShipDate.Size = new System.Drawing.Size(166, 20);
            this.dtpShipDate.TabIndex = 45;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(27, 121);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(86, 13);
            this.label4.TabIndex = 44;
            this.label4.Text = "OQC Ship Date: ";
            // 
            // txtBoxIdTo
            // 
            this.txtBoxIdTo.Location = new System.Drawing.Point(446, 10);
            this.txtBoxIdTo.Name = "txtBoxIdTo";
            this.txtBoxIdTo.Size = new System.Drawing.Size(166, 20);
            this.txtBoxIdTo.TabIndex = 47;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(404, 13);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(22, 13);
            this.label5.TabIndex = 48;
            this.label5.Text = "to: ";
            // 
            // btnUpInv
            // 
            this.btnUpInv.Enabled = false;
            this.btnUpInv.Location = new System.Drawing.Point(4, 2);
            this.btnUpInv.Name = "btnUpInv";
            this.btnUpInv.Size = new System.Drawing.Size(97, 25);
            this.btnUpInv.TabIndex = 43;
            this.btnUpInv.Text = "Update Invoice";
            this.btnUpInv.UseVisualStyleBackColor = true;
            this.btnUpInv.Click += new System.EventHandler(this.btnUpInv_Click);
            // 
            // pnlInvoice
            // 
            this.pnlInvoice.BackColor = System.Drawing.Color.Lime;
            this.pnlInvoice.Controls.Add(this.btnEditShipping);
            this.pnlInvoice.Controls.Add(this.btnUpInv);
            this.pnlInvoice.Enabled = false;
            this.pnlInvoice.Location = new System.Drawing.Point(282, 19);
            this.pnlInvoice.Name = "pnlInvoice";
            this.pnlInvoice.Size = new System.Drawing.Size(208, 29);
            this.pnlInvoice.TabIndex = 49;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(27, 161);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(45, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "Invoice:";
            // 
            // txtInvoice
            // 
            this.txtInvoice.Location = new System.Drawing.Point(125, 158);
            this.txtInvoice.Name = "txtInvoice";
            this.txtInvoice.Size = new System.Drawing.Size(166, 20);
            this.txtInvoice.TabIndex = 1;
            // 
            // btnAdd517
            // 
            this.btnAdd517.Location = new System.Drawing.Point(150, 53);
            this.btnAdd517.Name = "btnAdd517";
            this.btnAdd517.Size = new System.Drawing.Size(114, 25);
            this.btnAdd517.TabIndex = 6;
            this.btnAdd517.Text = "Add Box ID 517EB";
            this.btnAdd517.UseVisualStyleBackColor = true;
            this.btnAdd517.Click += new System.EventHandler(this.btnAdd517_Click);
            // 
            // btnAddBoxID523
            // 
            this.btnAddBoxID523.Location = new System.Drawing.Point(390, 53);
            this.btnAddBoxID523.Name = "btnAddBoxID523";
            this.btnAddBoxID523.Size = new System.Drawing.Size(114, 25);
            this.btnAddBoxID523.TabIndex = 50;
            this.btnAddBoxID523.Text = "Add Box ID 523";
            this.btnAddBoxID523.UseVisualStyleBackColor = true;
            this.btnAddBoxID523.Click += new System.EventHandler(this.btnAddBoxID523_Click);
            // 
            // btnAddBoxID517FB
            // 
            this.btnAddBoxID517FB.Location = new System.Drawing.Point(270, 53);
            this.btnAddBoxID517FB.Name = "btnAddBoxID517FB";
            this.btnAddBoxID517FB.Size = new System.Drawing.Size(114, 25);
            this.btnAddBoxID517FB.TabIndex = 51;
            this.btnAddBoxID517FB.Text = "Add Box ID 517FB";
            this.btnAddBoxID517FB.UseVisualStyleBackColor = true;
            this.btnAddBoxID517FB.Click += new System.EventHandler(this.btnAddBoxID517FB_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txtBoxIdTo);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.txtProductSerial);
            this.panel1.Controls.Add(this.rdbShipDate);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.dtpShipDate);
            this.panel1.Controls.Add(this.txtUser);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txtBoxIdFrom);
            this.panel1.Controls.Add(this.txtInvoice);
            this.panel1.Controls.Add(this.rdbProductSerial);
            this.panel1.Controls.Add(this.dtpRegistDate);
            this.panel1.Controls.Add(this.rdbPrintDate);
            this.panel1.Controls.Add(this.pnlBarcode);
            this.panel1.Controls.Add(this.rdbBoxId);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(980, 191);
            this.panel1.TabIndex = 52;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnAddBoxBFB_0025);
            this.panel2.Controls.Add(this.btnAddBoxBMA_0148);
            this.panel2.Controls.Add(this.btnAddBoxLD);
            this.panel2.Controls.Add(this.btnSearchBoxId);
            this.panel2.Controls.Add(this.pnlInvoice);
            this.panel2.Controls.Add(this.btnAddBoxID517FB);
            this.panel2.Controls.Add(this.button2);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Controls.Add(this.btnAddBoxId);
            this.panel2.Controls.Add(this.btnAddBoxID523);
            this.panel2.Controls.Add(this.btnAdd517);
            this.panel2.Controls.Add(this.btnExport);
            this.panel2.Controls.Add(this.btnCancel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 191);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(980, 130);
            this.panel2.TabIndex = 53;
            // 
            // btnAddBoxBFB_0025
            // 
            this.btnAddBoxBFB_0025.Location = new System.Drawing.Point(286, 84);
            this.btnAddBoxBFB_0025.Name = "btnAddBoxBFB_0025";
            this.btnAddBoxBFB_0025.Size = new System.Drawing.Size(130, 25);
            this.btnAddBoxBFB_0025.TabIndex = 54;
            this.btnAddBoxBFB_0025.Text = "Add Box ID BFB_0025";
            this.btnAddBoxBFB_0025.UseVisualStyleBackColor = true;
            this.btnAddBoxBFB_0025.Click += new System.EventHandler(this.btnAddBoxBFB_0025_Click);
            // 
            // btnAddBoxBMA_0148
            // 
            this.btnAddBoxBMA_0148.Location = new System.Drawing.Point(150, 84);
            this.btnAddBoxBMA_0148.Name = "btnAddBoxBMA_0148";
            this.btnAddBoxBMA_0148.Size = new System.Drawing.Size(130, 25);
            this.btnAddBoxBMA_0148.TabIndex = 53;
            this.btnAddBoxBMA_0148.Text = "Add Box ID BMA_0148";
            this.btnAddBoxBMA_0148.UseVisualStyleBackColor = true;
            this.btnAddBoxBMA_0148.Click += new System.EventHandler(this.btnAddBoxBMA_0148_Click);
            // 
            // btnAddBoxLD
            // 
            this.btnAddBoxLD.Location = new System.Drawing.Point(30, 84);
            this.btnAddBoxLD.Name = "btnAddBoxLD";
            this.btnAddBoxLD.Size = new System.Drawing.Size(114, 25);
            this.btnAddBoxLD.TabIndex = 52;
            this.btnAddBoxLD.Text = "Add Box ID Deus";
            this.btnAddBoxLD.UseVisualStyleBackColor = true;
            this.btnAddBoxLD.Click += new System.EventHandler(this.btnAddBoxLD_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(562, 84);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(134, 25);
            this.button2.TabIndex = 6;
            this.button2.Text = "Add Box ID BMA_0259";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.btnAddBoxId0259_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(422, 84);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(134, 25);
            this.button1.TabIndex = 6;
            this.button1.Text = "Add Box ID BMA_0241";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnAddBoxId0241_Click);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dgvBoxId);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 321);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(980, 409);
            this.panel3.TabIndex = 54;
            // 
            // frmBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(980, 730);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmBox";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Box ID";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmBox_FormClosed);
            this.Load += new System.EventHandler(this.frmBox_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvBoxId)).EndInit();
            this.pnlInvoice.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView dgvBoxId;
        private System.Windows.Forms.Button btnSearchBoxId;
        private System.Windows.Forms.Button btnAddBoxId;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtBoxIdFrom;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtProductSerial;
        private System.Windows.Forms.DateTimePicker dtpRegistDate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtUser;
        private System.Windows.Forms.Panel pnlBarcode;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.RadioButton rdbBoxId;
        private System.Windows.Forms.RadioButton rdbPrintDate;
        private System.Windows.Forms.RadioButton rdbProductSerial;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnEditShipping;
        private System.Windows.Forms.RadioButton rdbShipDate;
        private System.Windows.Forms.DateTimePicker dtpShipDate;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtBoxIdTo;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnUpInv;
        private System.Windows.Forms.Panel pnlInvoice;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtInvoice;
        private System.Windows.Forms.CheckBox ckbInvoice;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colUpdateInvoice;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_boxid;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_user;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_regist_date;
        private System.Windows.Forms.DataGridViewCheckBoxColumn col_update_ship;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_ship_date;
        private System.Windows.Forms.DataGridViewTextBoxColumn col_invoice;
        private System.Windows.Forms.Button btnAdd517;
        private System.Windows.Forms.Button btnAddBoxID523;
        private System.Windows.Forms.Button btnAddBoxID517FB;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button btnAddBoxLD;
        private System.Windows.Forms.Button btnAddBoxBMA_0148;
        private System.Windows.Forms.Button btnAddBoxBFB_0025;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

