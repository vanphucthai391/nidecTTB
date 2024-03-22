using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Security.Permissions;
using Npgsql;
using BoxIdDB;

namespace BoxIdDb
{
    public partial class frmBox : Form
    {
        //�e�t�H�[��frmLogin�ցA�C�x���g������A���i�f���Q�[�g�j
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        CheckBox ckbShipDate;
        DataGridViewButtonColumn openBoxId;
        DataGridViewButtonColumn editShipDate;
        //���̑��񃍁[�J���ϐ�

        // �R���X�g���N�^
        public frmBox()
        {
            InitializeComponent();
        }
        // ���[�h���̏���
        private void frmBox_Load(object sender, EventArgs e)
        {
            //dgvBoxId.Columns[0].ReadOnly = true;
            dgvBoxId.Columns[1].ReadOnly = true;
            dgvBoxId.Columns[2].ReadOnly = true;
            dgvBoxId.Columns[3].ReadOnly = true;
            dgvBoxId.Columns[5].ReadOnly = true;
            dgvBoxId.Columns[6].ReadOnly = true;

            ckbInvoice = new CheckBox();
            //Get the column header cell bounds
            Rectangle rect = this.dgvBoxId.GetCellDisplayRectangle(0, -1, true);
            ckbInvoice.Size = new Size(14, 14);
            //Change the location of the CheckBox to make it stay on the header
            ckbInvoice.Location = rect.Location;
            ckbInvoice.CheckedChanged += new EventHandler(ckbInvoice_CheckedChanged);
            //Add the CheckBox into the DataGridView
            this.dgvBoxId.Controls.Add(ckbInvoice);

            ckbShipDate = new CheckBox();
            //Get the column header cell bounds
            Rectangle rect1 = this.dgvBoxId.GetCellDisplayRectangle(3, -1, true);
            ckbShipDate.Size = new Size(14, 14);
            //Change the location of the CheckBox to make it stay on the header
            ckbShipDate.Location = rect1.Location;
            ckbShipDate.CheckedChanged += new EventHandler(ckbShipDate_CheckedChanged);
            //Add the CheckBox into the DataGridView
            this.dgvBoxId.Controls.Add(ckbShipDate);

            //�t�H�[���̏ꏊ���w��
            Left = 50;
            Top = 10;
            updateDataGridViews(ref dgvBoxId, true);
            //dgvBoxId["col_invoice", ].ReadOnly = false;

            // �c�`�s�d�s�h�l�d�o�h�b�j�d�q�̕��ȉ���������
            dtpRounddownHour(dtpShipDate);

            for (int i = 0; i < dgvBoxId.Rows.Count; i++)
            {
                if (!String.IsNullOrEmpty(dgvBoxId["col_invoice", i].Value.ToString()))
                {
                    dgvBoxId["colUpdateInvoice", i].Value = true;
                }
                if (!String.IsNullOrEmpty(dgvBoxId["col_ship_date", i].Value.ToString()))
                {
                    dgvBoxId["col_update_ship", i].Value = true;
                }
            }

            if (txtUser.Text == "admin" || txtUser.Text == "Ms.Ngoan")
            {
                txtBoxIdTo.Enabled = true;
                pnlInvoice.Enabled = true;
                btnUpInv.Enabled = true;
            }
            else
            {
                txtBoxIdTo.Enabled = false;
            }
        }

        // �T�u�v���V�[�W���F�f�[�^�O���b�g�r���[�̍X�V�B�e�t�H�[���ŌĂяo���A�e�t�H�[���̏��������p��
        public void updateControls(string user)
        {
            txtUser.Text = user;
        }

        // �T�u�v���V�[�W���F�f�[�^�e�[�u���̒�`
        private void defineAndReadDatatable(ref DataTable dt)
        {
            dt.Columns.Add("Boxid", Type.GetType("System.String"));
            dt.Columns.Add("User", Type.GetType("System.String"));
            dt.Columns.Add("Regist Date", Type.GetType("System.DateTime"));
            dt.Columns.Add("Ship Date", Type.GetType("System.DateTime"));
            dt.Columns.Add("Invoice", Type.GetType("System.String"));
        }


        // �T�u�v���V�[�W���F�f�[�^�O���b�g�r���[�̍X�V
        public void updateDataGridViews(ref DataGridView dgv, bool load)
        {
            string boxId = txtBoxIdFrom.Text;
            DateTime printDate = dtpRegistDate.Value;
            DateTime shipDate = dtpShipDate.Value;
            string serialNo = txtProductSerial.Text;
            string sql = String.Empty;

            // �r�p�k���ʂ��A�c�s�`�`�s�`�a�k�d�֊i�[
            TfSQL tf = new TfSQL();
            if (rdbBoxId.Checked)
            {
                if (boxId.Length < 6)
                {
                    MessageBox.Show("Please select at least 6 characters like LM1601", "BoxId DB",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                    return;
                }

                sql = "select boxid, suser, regist_date, shipdate, invoice FROM box_id_rt" +
                    (boxId == String.Empty ? String.Empty : " WHERE boxid like '" + boxId + "%'") +
                    " order by boxid";
            }
            else if (rdbPrintDate.Checked)
            {
                sql = "select boxid, suser, regist_date, shipdate, invoice FROM box_id_rt WHERE regist_date " +
                    "BETWEEN '" + printDate.Date + "' AND '" + printDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59) + "'" +
                    " order by boxid";
            }
            else if (rdbProductSerial.Checked)
            {
                sql = "select boxid FROM product_serial_rt WHERE serialno='" + serialNo + "'";
                boxId = tf.sqlExecuteScalarString(sql);
                txtBoxIdFrom.Text = boxId;
                if (boxId == String.Empty)
                {
                    sql = "select boxid, suser, regist_date, shipdate, invoice FROM box_id_rt WHERE printdate " +
                        "BETWEEN '" + printDate.Date + "' AND '" + printDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59) + "'" +
                        " order by boxid";
                }
                else
                {
                    sql = "select boxid, suser, regist_date, shipdate, invoice FROM box_id_rt" +
                        " WHERE boxid='" + boxId + "'";
                }
            }
            else if (dtpShipDate.Checked)
            {
                sql = "select boxid, suser, regist_date, shipdate, invoice FROM box_id_rt WHERE shipdate " +
                    "BETWEEN '" + shipDate.Date + "' AND '" + shipDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59) + "'" +
                    " order by boxid";
            }

            DataTable dt1 = new DataTable();
            tf.sqlDataAdapterFillDatatable(sql, ref dt1);

            // �f�[�^�O���b�g�r���[�ւc�s�`�`�s�`�a�k�d���i�[
            dgv.DataSource = dt1;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            // �O���b�g�r���[�E�[�Ƀ{�^����ǉ��i����̂݁j
            if (load) addButtonsToDataGridView(dgv);

            //�s�w�b�_�[�ɍs�ԍ���\������
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                //dgv.Rows[i].HeaderCell.Value = (i + 1).ToString();
                if (!String.IsNullOrEmpty(dgv["col_invoice", i].Value.ToString()))
                {
                    dgv["colUpdateInvoice", i].Value = true;
                }
                if (!String.IsNullOrEmpty(dgv["col_ship_date", i].Value.ToString()))
                {
                    dgv["col_update_ship", i].Value = true;
                }
            }
            //�s�w�b�_�[�̕����������߂���
            dgv.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            // ��ԉ��̍s��\������
            if (dgv.Rows.Count != 0)
                dgv.FirstDisplayedScrollingRowIndex = dgv.Rows.Count - 1;

            // �p�l���Ƀo�[�R�[�h��\��
            pnlBarcode.Refresh();
        }

        // �T�u�T�u�v���V�[�W���F�O���b�g�r���[�E�[�Ƀ{�^����ǉ�
        private void addButtonsToDataGridView(DataGridView dgv)
        {
            // �J���{�^���͑S�Ẵ��[�U�[
            openBoxId = new DataGridViewButtonColumn();
            openBoxId.HeaderText = "Open";
            openBoxId.Text = "Open";
            openBoxId.UseColumnTextForButtonValue = true;
            openBoxId.Width = 80;
            dgv.Columns.Add(openBoxId);
        }

        // �{�^���P�̓t�H�[���R���{�����[�h�ŊJ���i�f���Q�[�g�Ȃ��j�A�{�^���Q�͏o�ד��̕ҏW
        private void dgvBoxId_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int currentRow = int.Parse(e.RowIndex.ToString());

            if (dgvBoxId.Columns[e.ColumnIndex] == openBoxId && currentRow >= 0)
            {
                //����frmModule ���J����Ă���ꍇ�́A�������悤�ɑ���
                bool inUse = TfGeneral.checkOpenFormExists("frmModule") && TfGeneral.checkOpenFormExists("frmModule517EB") && TfGeneral.checkOpenFormExists("frmModule517FB") && TfGeneral.checkOpenFormExists("frmModule523") && TfGeneral.checkOpenFormExists("frmModuleLD") && TfGeneral.checkOpenFormExists("frmModule0148") && TfGeneral.checkOpenFormExists("frmModule0025") && TfGeneral.checkOpenFormExists("frmModule0241") && TfGeneral.checkOpenFormExists("frmModule0259");
                if (inUse)
                {
                    MessageBox.Show("Please close the other already open window.", "Notice",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    return;
                }

                string frmName = "VIEW";
                string boxId = dgvBoxId["col_boxid", currentRow].Value.ToString();
                DateTime printDate = DateTime.Parse(dgvBoxId["col_regist_date", currentRow].Value.ToString());
                string serialNo = txtProductSerial.Text;
                string user = txtUser.Text;
                string invoice = dgvBoxId["col_invoice", currentRow].Value.ToString();

                if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("517EB"))
                {
                    frmModule517EB f3 = new frmModule517EB();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("523"))
                {
                    frmModule523 f3 = new frmModule523();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("0148"))
                {
                    frmModule0148 f3 = new frmModule0148();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("0025"))
                {
                    frmModule0025 f3 = new frmModule0025();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("517FB"))
                {
                    frmModule517FB f3 = new frmModule517FB();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("LD20"))
                {
                    frmModuleLD f3 = new frmModuleLD();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("BFB_0025"))
                {
                    frmModule0025 f3 = new frmModule0025();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("0241"))
                {
                    frmModule0241 f3 = new frmModule0241();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else if (dgvBoxId.CurrentRow.Cells["col_boxid"].Value.ToString().StartsWith("0259"))
                {
                    frmModule0259 f3 = new frmModule0259();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
                else
                {
                    frmModule f3 = new frmModule();
                    //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                    f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                    {
                        updateDataGridViews(ref dgvBoxId, false);
                        Focus();
                    };
                    f3.updateControls(frmName, boxId, printDate, serialNo, invoice, user, false, false);
                    f3.Show();
                }
            }

            if (dgvBoxId.Columns[e.ColumnIndex] == editShipDate && currentRow >= 0)
            {
                string boxId = dgvBoxId["col_boxid", currentRow].Value.ToString();
                DateTime shipdate = dtpShipDate.Value;

                DialogResult result1 = MessageBox.Show("Do you want to update the shipping date of as follows:" + System.Environment.NewLine +
                    boxId + ": " + shipdate, "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (result1 == DialogResult.Yes)
                {
                    string sql = "update box_id_rt SET shipdate ='" + shipdate + "' " +
                        "WHERE boxid= '" + boxId + "'";
                    System.Diagnostics.Debug.Print(sql);
                    TfSQL tf = new TfSQL();
                    int res = tf.sqlExecuteNonQueryInt(sql, false);
                    updateDataGridViews(ref dgvBoxId, false);
                }
            }
        }

        // �����{�^�������A���ۂ̓O���b�g�r���[�̍X�V�����邾��
        private void btnSearchBoxId_Click(object sender, EventArgs e)
        {
            updateDataGridViews(ref dgvBoxId, false);
        }

        // �t�H�[���R��ҏW���[�h�ŊJ���A�f���Q�[�g����
        private void btnAddBoxId_Click(object sender, EventArgs e)
        {
            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModule");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModule f3 = new frmModule();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };

                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }
        }

        // �o�ד����ꊇ�o�^����
        private void btnEditShipping_Click(object sender, EventArgs e)
        {
            DateTime shipdate = dtpShipDate.Value;
            string boxid, sql, a;
            TfSQL tf = new TfSQL();

            for (int i = 0; i < dgvBoxId.Rows.Count; i++)
            {
                if (dgvBoxId["col_update_ship", i].Value != null)
                {
                    a = dgvBoxId["col_update_ship", i].Value.ToString();
                    if (bool.Parse(a) == true && String.IsNullOrEmpty(dgvBoxId["col_ship_date", i].Value.ToString()))
                    {
                        boxid = dgvBoxId["col_boxid", i].Value.ToString();
                        sql = "UPDATE box_id_rt SET shipdate = '" + shipdate + "' WHERE boxid = '" + boxid + "'";
                        tf.sqlExecuteScalarString(sql);
                    }
                }
            }
            MessageBox.Show("Updated ShipDate!", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            updateDataGridViews(ref dgvBoxId, false);
        }

        // �T�u�v���V�[�W���F�����p�̃o�[�R�[�h�\���p�l���̍X�V�A���ۂ̏o�͂Ƃ͊֌W�̂Ȃ����C�u�������g�p���Ă���
        private void pnlBarcode_Paint(object sender, PaintEventArgs e)
        {
            DotNetBarcode barCode = new DotNetBarcode();
            string barcodeNumber;
            Single x1;
            Single y1;
            Single x2;
            Single y2;
            x1 = 0;
            y1 = 0;
            x2 = pnlBarcode.Size.Width;
            y2 = pnlBarcode.Size.Height;
            barcodeNumber = txtBoxIdFrom.Text;
            barCode.Type = DotNetBarcode.Types.Jan13;

            if (barcodeNumber != String.Empty)
                barCode.WriteBar(barcodeNumber, x1, y1, x2, y2, e.Graphics);
        }

        //frmBox�����ہA��\���ɂȂ��Ă���e�t�H�[��frmLogin�����
        private void frmBox_FormClosed(object sender, FormClosedEventArgs e)
        {
            //�e�t�H�[��frmLogin�����悤�A�f���Q�[�g�C�x���g�𔭐�������
            RefreshEvent(this, new EventArgs());
        }

        // ����{�^����V���[�g�J�b�g�ł̏I���������Ȃ�
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x112;
            const long SC_CLOSE = 0xF060L;
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & 0xFFF0L) == SC_CLOSE) { return; }
            base.WndProc(ref m);
        }

        // �t�H�[���R���J����Ă��Ȃ����Ƃ��m�F���Ă���A����
        private void btnCancel_Click(object sender, EventArgs e)
        {
            string formName = "frmModule";
            string formName1 = "frmModule517EB";
            string formName2 = "frmModule517FB";
            string formName3 = "frmModule523";
            string formName4 = "frmModule0241";
            string formName5 = "frmModule0259";



            bool bl = false;
            foreach (Form buff in Application.OpenForms)
            {
                if (buff.Name == formName || buff.Name == formName1 || buff.Name == formName2 || buff.Name == formName3 || buff.Name == formName4 || buff.Name == formName5)
                { bl = true; }
            }
            if (bl)
            {
                MessageBox.Show("You need to close Form Product Serial first.", "BoxId DB",
                  MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }
            FormCollection forms = Application.OpenForms;

            for (int countForms = forms.Count - 1; countForms >= 0; --countForms)
            {
                if (forms[countForms].GetType().BaseType != typeof(Form))
                    forms[countForms].Close();
            }
            this.Close();
        }

        // ���W�I�{�^���u�{�b�N�X�h�c�v�ύX���̏����i�e�L�X�g�{�b�N�X�ҏW�ɂ�錟�������̕ύX�j
        private void rdbBoxId_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbBoxId.Checked) { txtProductSerial.Text = String.Empty; }
        }
        // ���W�I�{�^���u�v�����g���t�v�ύX���̏����i�e�L�X�g�{�b�N�X�ҏW�ɂ�錟�������̕ύX�j
        private void rdbPrintDate_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbPrintDate.Checked)
            {
                txtBoxIdFrom.Text = String.Empty;
                txtProductSerial.Text = String.Empty;
            }
        }
        // ���W�I�{�^���u���i�V���A���v�ύX���̏����i�e�L�X�g�{�b�N�X�ҏW�ɂ�錟�������̕ύX�j
        private void rdbProductSerial_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbProductSerial.Checked)
            {
                txtBoxIdFrom.Text = String.Empty;
            }
        }
        // ���W�I�{�^���u�o�ד��v�ύX���̏����i�e�L�X�g�{�b�N�X�ҏW�ɂ�錟�������̕ύX�j
        private void rdbShipDate_CheckedChanged_1(object sender, EventArgs e)
        {
            if (rdbShipDate.Checked)
            {
                txtBoxIdFrom.Text = String.Empty;
                txtBoxIdTo.Text = String.Empty;
                txtProductSerial.Text = String.Empty;
            }
        }

        // �T�u�T�u�v���V�[�W���F�c�`�s�d�s�h�l�d�o�h�b�j�d�q�̕��ȉ���������
        private void dtpRounddownHour(DateTimePicker dtp)
        {
            DateTime dt = dtp.Value;
            int hour = dt.Hour;
            int minute = dt.Minute;
            int second = dt.Second;
            int millisecond = dt.Millisecond;
            dtp.Value = dt.AddHours(-hour).AddMinutes(-minute).AddSeconds(-second).AddMilliseconds(-millisecond);
        }

        // �f�[�^���G�N�Z���փG�N�X�|�[�g
        private void btnExport_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)dgvBoxId.DataSource;
            ExcelClass xl = new ExcelClass();
            xl.ExportToExcel(dt);
            //xl.ExportToCsv(dt, System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\ipqcdb.csv");
        }

        private void btnUpInv_Click(object sender, EventArgs e)
        {
            string invoice, boxid, sql, a;
            TfSQL tf = new TfSQL();
            invoice = txtInvoice.Text;
            for (int i = 0; i < dgvBoxId.Rows.Count; i++)
            {
                if (dgvBoxId["colUpdateInvoice", i].Value != null)
                {
                    a = dgvBoxId["colUpdateInvoice", i].Value.ToString();
                    if (bool.Parse(a) == true && String.IsNullOrEmpty(dgvBoxId["col_invoice", i].Value.ToString()))
                    {
                        boxid = dgvBoxId["col_boxid", i].Value.ToString();
                        sql = "UPDATE box_id_rt SET invoice = '" + invoice + "' WHERE boxid = '" + boxid + "'";
                        tf.sqlExecuteScalarString(sql);
                    }
                }
            }
            MessageBox.Show("Updated Invoice!", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            updateDataGridViews(ref dgvBoxId, false);
        }

        private void ckbInvoice_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dgvBoxId.RowCount; i++)
            {
                if (ckbInvoice.Checked)
                {
                    if (dgvBoxId["col_invoice", i].Value.ToString() == "")
                    {
                        dgvBoxId["colUpdateInvoice", i].Value = true;
                    }
                }
                else dgvBoxId["colUpdateInvoice", i].Value = false;
            }
        }

        private void ckbShipDate_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dgvBoxId.RowCount; i++)
            {
                if (ckbShipDate.Checked)
                {
                    if (dgvBoxId["col_ship_date", i].Value.ToString() == "")
                    {
                        dgvBoxId["col_update_ship", i].Value = true;
                    }
                }
                else dgvBoxId["col_update_ship", i].Value = false;
            }
        }

        private void btnAdd517_Click(object sender, EventArgs e)
        {
            //frmModule517EB frm517 = new frmModule517EB();
            //frm517.Show();

            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModule517EB");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModule517EB f3 = new frmModule517EB();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };

                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }
        }

        private void btnAddBoxID523_Click(object sender, EventArgs e)
        {
            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModule523");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModule523 f3 = new frmModule523();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };

                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }
        }

        private void btnAddBoxID517FB_Click(object sender, EventArgs e)
        {
            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModule517FB");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModule517FB f3 = new frmModule517FB();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };

                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }
        }

        private void btnAddBoxLD_Click(object sender, EventArgs e)
        {
            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModuleLD");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModuleLD f3 = new frmModuleLD();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };

                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }

        }

        private void btnAddBoxBMA_0148_Click(object sender, EventArgs e)
        {

            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModule0148");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModule0148 f3 = new frmModule0148();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };

                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }
        }

        private void btnAddBoxBFB_0025_Click(object sender, EventArgs e)
        {

            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModule0025");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModule0025 f3 = new frmModule0025();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };

                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }
        }

        private void btnAddBoxId0241_Click(object sender, EventArgs e)
        {
            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModule0241");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModule0241 f3 = new frmModule0241();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };

                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }
        }

        private void btnAddBoxId0259_Click(object sender, EventArgs e)
        {
            string user = txtUser.Text;

            bool bl = TfGeneral.checkOpenFormExists("frmModule0259");
            if (bl)
            {
                MessageBox.Show("Please close brows-mode form or finish the current edit form.", "BoxId DB",
                MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmModule0259 f3 = new frmModule0259();
                //�q�C�x���g���L���b�`���āA�f�[�^�O���b�h���X�V����
                f3.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    updateDataGridViews(ref dgvBoxId, false);
                    Focus();
                };
                f3.updateControls(String.Empty, String.Empty, DateTime.Now, String.Empty, String.Empty, user, true, false);
                f3.Show();
            }
        }
    }
}
