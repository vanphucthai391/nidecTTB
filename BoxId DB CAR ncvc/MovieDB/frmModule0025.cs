﻿using BoxIdDb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Windows.Forms;

namespace BoxIdDB
{
    public partial class frmModule0025 : Form
    {
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        string appconfig = @"\\192.168.193.1\barcode$\BoxId Printer vc5\info.ini";
        string directory = @"C:\Users\takusuke.fujii\Desktop\Auto Print\\";

        bool formEditMode;
        bool formReturnMode;
        bool formAddMode;
        string user;
        string m_lot;
        int okCount;
        bool inputBoxModeOriginal;
        string productTable;
        string testerTableThisMonth;
        string testerTableLastMonth;
        string tableThisMonth;
        string tableLastMonth;
        string testerTableThisMonth1;
        string testerTableLastMonth1;
        string tableThisMonth1;
        string tableLastMonth1;
        //string tableAssyThisMonth, tableAssyLastMonth;
        DataTable dtOverall;
        DataTable dtAllProcess;
        int limit1 = 500;//testing  actual 500:
        public int limit2 = 0;
        bool sound;
        bool frictiongood = false;
        string pcbbarcode = "";
        bool frictionnorow =false;
        public frmModule0025()
        {
            InitializeComponent();
            dtAllProcess = new DataTable();
        }

        private void frmModule0025_Load(object sender, EventArgs e)
        {
            txtCarton.Enabled = false;
            user = txtUser.Text;
            txtLimit.Text = limit2.ToString();
            directory = readIni("TARGET DIRECTORY", "DIR", appconfig);
            this.Left = 250;
            this.Top = 20;
            dtOverall = new DataTable();
            defineAndReadDtOverall(ref dtOverall);
            if (!formEditMode)
            {
                if (dtOverall.Rows.Count >= 0)
                {
                    limit1 = 500;//testing  actual 500:
                }
            }
            updateDataGridViews(dtOverall, ref dgvInline);
            if (!formEditMode)
            {
                txtProductSerial.Enabled = false;
            }
            txtCarton.Enabled = true;
            cmbModel.SelectedIndex = 0;
            //limit1 = 500;
            txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
            //txtProductSerial.Enabled = true;
            txtProductSerial.Focus();

        }

        private string readIni(string s, string k, string cfs)
        {
            StringBuilder retVal = new StringBuilder(255);
            string section = s;
            string key = k;
            string def = String.Empty;
            int size = 255;
            //get the value from the key in section
            int strref = GetPrivateProfileString(section, key, def, retVal, size, cfs);
            return retVal.ToString();
        }
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filepath);
        public void updateControls(string frmName, string boxId, DateTime registDate, string serialNo, string invoice, string user, bool editMode, bool returnMode)
        {
            lblFrmName.Text = frmName;
            txtBoxId.Text = boxId;
            txtProductSerial.Text = serialNo;
            txtUser.Text = user;
            txtInvoice.Text = invoice;
            if (boxId != "")
            {
                string[] box_arr = boxId.Split('-');
                string model = box_arr[0];
                productTable = "product_serial_0025 ";
                switch (model)
                {
                    case "0025":
                        cmbModel.Text = "BFB_0025";
                        limit1 = 500;
                        break;

                }
                txtCarton.Text = box_arr[2];
            }

            txtCarton.Enabled = editMode;
            txtProductSerial.Enabled = editMode;
            cmbModel.Enabled = editMode;
            btnRegisterBoxId.Enabled = !editMode;
            btnPrint.Visible = !editMode;
            btnDeleteSelection.Visible = editMode;
            formEditMode = editMode;
            formReturnMode = returnMode;

            this.Text = editMode ? "Product Serial - Edit Mode" : "Product Serial - Browse Mode";
            if (editMode && user == "admin" || editMode && user == "User_9" || editMode && user == "Diep")
            {
                btnChangeLimit.Visible = true;
                txtLimit.Visible = true;
            }
            if (!editMode && user == "admin" || !editMode && user == "User_9" || !editMode && user == "Diep")
            {
                //btnAddSerial.Visible = true;
                btnCancelBoxid.Visible = true;
                btnChangeLimit.Visible = true;
                //btnDeleteSerial.Visible = true;
            }
        }
        private void setProductTable()
        {
            if (!string.IsNullOrEmpty(cmbModel.Text))
            {
                //string[] model = cmbModel.Text.Split('-');
                productTable = "product_serial_0025";
            }
        }
        private void defineAndReadDtOverall(ref DataTable dt)
        {
            string boxId = txtBoxId.Text;
            setProductTable();
            dt.Columns.Add("serialno", Type.GetType("System.String"));//format form
            dt.Columns.Add("model", Type.GetType("System.String"));//format form
            dt.Columns.Add("lot", Type.GetType("System.String"));//format form
            //start OQC
            dt.Columns.Add("inspectdate", Type.GetType("System.DateTime")); //datetest NMT
            dt.Columns.Add("qacurrent", Type.GetType("System.String"));//OQC 
            dt.Columns.Add("qafg", Type.GetType("System.String"));//OQC 
            dt.Columns.Add("qaspeed", Type.GetType("System.String"));//OQC 
            dt.Columns.Add("tjudge", Type.GetType("System.String"));//Judge OQC
            //Start Line
            dt.Columns.Add("date_line", Type.GetType("System.DateTime")); //datetest NO41-FAN
            dt.Columns.Add("current", Type.GetType("System.String"));//datetest NO41-FAN
            dt.Columns.Add("fg", Type.GetType("System.String"));//datetest NO41-FAN
            dt.Columns.Add("speed", Type.GetType("System.String"));//datetest NO41-FAN

            dt.Columns.Add("tjudge_line", Type.GetType("System.String"));//format form
            dt.Columns.Add("svfi", Type.GetType("System.String"));//datetest NO41-FAN
            dt.Columns.Add("return", Type.GetType("System.String"));//format form

            if (!formEditMode)
            {
                string sql;
                sql = "select serialno, model, lot, inspectdate, qacurrent, qafg, qaspeed, tjudge, date_line, current, fg, speed, svfi, tjudge_line, return " +
                    "FROM " + productTable + " WHERE boxid='" + boxId + "'";
                TfSQL tf = new TfSQL();
                System.Diagnostics.Debug.Print(sql);
                tf.sqlDataAdapterFillDatatable(sql, ref dt);
            }
        }
        private void updateDataGridViews(DataTable dt1, ref DataGridView dgv1)
        {
            inputBoxModeOriginal = txtProductSerial.Enabled;
            txtProductSerial.Enabled = false;
            updateDataGridViewsSub(dt1, ref dgv1);
            colorViewForFailAndBlank(ref dgv1);
            colorViewForDuplicateSerial(ref dgv1);
            for (int i = 0; i < dgv1.Rows.Count; i++)
                dgv1.Rows[i].HeaderCell.Value = (i + 1).ToString();

            dgv1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            if (dgv1.Rows.Count >= 1)
                dgv1.FirstDisplayedScrollingRowIndex = dgv1.Rows.Count - 1;

            txtProductSerial.Enabled = inputBoxModeOriginal;

            okCount = getOkCount(dt1);
            txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
            if (okCount == limit1)
            {
                txtProductSerial.Enabled = false;
            }
            else
            {
                txtProductSerial.Enabled = true;
            }

            if (okCount == limit1 && dgv1.Rows.Count == limit1)
            {
                btnRegisterBoxId.Enabled = true;
                txtCarton.Enabled = true;
            }
            else
            {
                btnRegisterBoxId.Enabled = false;
            }

        }
        private int getOkCount(DataTable dt)
        {
            if (dt.Rows.Count <= 0) return 0;
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge", "tjudge_line" });
            DataRow[] dr = distinct.Select("tjudge = 'PASS' and tjudge_line = 'PASS'");
            int dist = dr.Length;
            return dist;
        }

        private void updateDataGridViewsSub(DataTable dt1, ref DataGridView dgv1)
        {
            dgv1.DataSource = dt1;
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            string[] criteriaDateCode = getLotArray(dt1);
            makeDatatableSummary(dt1, ref dgvDateCode, criteriaDateCode, "lot");
        }
        private string[] getLotArray(DataTable dt0)
        {
            DataTable dt1 = dt0.Copy();
            DataView dv = dt1.DefaultView;
            dv.Sort = "lot";
            DataTable dt2 = dv.ToTable(true, "lot");
            string[] array = new string[dt2.Rows.Count + 1];
            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                array[i] = dt2.Rows[i]["lot"].ToString();
            }
            array[dt2.Rows.Count] = "Total";
            return array;
        }
        public void makeDatatableSummary(DataTable dt0, ref DataGridView dgv, string[] criteria, string header)
        {
            DataTable dt1 = new DataTable();
            DataRow dr = dt1.NewRow();
            Int32 count;
            Int32 total = 0;
            string condition;

            for (int i = 0; i < criteria.Length; i++)
            {
                dt1.Columns.Add(criteria[i], typeof(Int32));
                condition = header + " = '" + criteria[i] + "'";
                count = dt0.Select(condition).Length;
                total += count;
                dr[criteria[i]] = count;
                if (criteria[i] == "Total") dr[criteria[i]] = total;
                if (criteria[i] == "No Data") dr[criteria[i]] = dgvInline.Rows.Count - total;
            }
            dt1.Rows.Add(dr);

            dgv.Columns.Clear();
            dgv.DataSource = dt1;
            dgv.AllowUserToAddRows = false;
            dgv.ReadOnly = true;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }
        //not used
        private void colorMixedLot(DataTable dt, ref DataGridView dgv)
        {
            if (dt.Rows.Count <= 0) return;

            DataTable distinct1 = dt.DefaultView.ToTable(true, new string[] { "lot" });

            if (distinct1.Rows.Count == 1)
                m_lot = distinct1.Rows[0]["lot"].ToString();

            if (distinct1.Rows.Count >= 2)
            {
                string A = distinct1.Rows[0]["lot"].ToString();
                string B = distinct1.Rows[1]["lot"].ToString();
                int a = distinct1.Select("lot = '" + A + "'").Length;
                int b = distinct1.Select("lot = '" + B + "'").Length;

                m_lot = a > b ? A : B;

                string C = a < b ? A : B;
                int c = -1;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["lot"].ToString() == C) { c = i; }
                }

                if (c != -1)
                {
                    dgv["col_lot", c].Style.BackColor = Color.Red;
                    soundAlarm();
                }
                else
                {
                    dgv.Columns["col_lot"].DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Window);
                }
            }
        }
        //not used
        private void btnPrint_Click(object sender, EventArgs e)
        {
            string boxId = txtBoxId.Text;
            string model = cmbModel.Text;
            string shipKind = dtOverall.Rows[0]["return"].ToString();
            printBarcode(directory, boxId, model, dgvDateCode, ref dgvDateCode2, ref txtBoxIdPrint, shipKind);
        }

        private void btnRegisterBoxId_Click(object sender, EventArgs e)
        {

            if (!String.IsNullOrEmpty(txtCarton.Text))
            {

                btnRegisterBoxId.Enabled = false;
                btnDeleteSelection.Enabled = false;
                btnCancel.Enabled = false;

                string boxId = txtBoxId.Text;

                string checkResult = checkDataTableWithRealTable(dtOverall);

                if (checkResult != String.Empty)
                {
                    MessageBox.Show("The following serials are already registered with box id:" + Environment.NewLine +
                        checkResult + Environment.NewLine + "Please check and delete.", "Notice",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    btnRegisterBoxId.Enabled = true;
                    btnDeleteSelection.Enabled = true;
                    btnCancel.Enabled = true;
                    return;
                }

                TfSQL yn = new TfSQL();
                string sql_box = "INSERT INTO box_id_rt(" +
                    "boxid," +
                    "suser," +
                    "regist_date) " +
                    "VALUES(" +
                    "'" + boxId + "'," +
                    "'" + user + "'," +
                    "'" + DateTime.Now.ToString() + "')";
                System.Diagnostics.Debug.Print(sql_box);
                yn.sqlExecuteNonQuery(sql_box, false);

                DataTable dt = dtOverall.Copy();
                dt.Columns.Add("boxid", Type.GetType("System.String"));
                dt.Columns.Add("carton", Type.GetType("System.String"));
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["boxid"] = boxId;
                    dt.Rows[i]["carton"] = txtCarton.Text;
                }

                TfSQL tf = new TfSQL();
                bool res1= tf.sqlMultipleInsertBFB0025(dt);
                if (res1)
                {
                    string shipKind = dtOverall.Rows[0]["return"].ToString();
                    string prt_model = cmbModel.Text;
                    dtOverall.Clear();
                    dt = null;

                    txtBoxId.Text = boxId;
                    this.RefreshEvent(this, new EventArgs());

                    this.Focus();
                    MessageBox.Show("The box id " + boxId + " and " + Environment.NewLine +
                        "its product serials were registered.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtBoxId.Text = String.Empty;
                    txtProductSerial.Text = String.Empty;
                    updateDataGridViews(dtOverall, ref dgvInline);
                    btnRegisterBoxId.Enabled = false;
                    btnDeleteSelection.Enabled = true;
                    btnCancel.Enabled = true;
                }
                else
                {
                    string sql = "delete from box_id_rt WHERE boxid= '" + boxId + "'";
                    int res = tf.sqlExecuteNonQueryInt(sql, false);

                    MessageBox.Show("Box id and product serials were not registered.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    btnRegisterBoxId.Enabled = true;
                    btnDeleteSelection.Enabled = true;
                    btnCancel.Enabled = true;
                }
            }
            else MessageBox.Show("Please input the carton number!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void colorViewForFailAndBlank(ref DataGridView dgv)
        {
            int row = dgv.Rows.Count;
            for (int i = 0; i < row; ++i)
            {
                //Alarm OQC FAIL or NODATA
                if (dgv["col_judge_qa", i].Value.ToString() == "FAIL" || dgv["col_judge_qa", i].Value.ToString() == "PLS NG" || String.IsNullOrEmpty(dgv["col_judge_qa", i].Value.ToString()))
                {
                    dgv["col_date", i].Style.BackColor = Color.Red;
                    dgv["col_qacurrent", i].Style.BackColor = Color.Red;
                    dgv["col_qafg", i].Style.BackColor = Color.Red;
                    dgv["col_qaspeed", i].Style.BackColor = Color.Red;
                    dgv["col_judge_qa", i].Style.BackColor = Color.Red;

                    if (dgv.Name == "dgvInline") tabControl1.SelectedIndex = 1;
                    else tabControl1.SelectedIndex = 0;

                    soundAlarm();
                }
                else
                {
                    dgv.Rows[i].InheritedStyle.BackColor = Color.FromKnownColor(KnownColor.Window);

                    tabControl1.SelectedIndex = 0;
                }

                //Alarm INLINE FAIL or NODATA
                if (dgv["col_judge_inline", i].Value.ToString() == "FAIL" || dgv["col_judge_inline", i].Value.ToString() == "PLS NG" || String.IsNullOrEmpty(dgv["col_judge_inline", i].Value.ToString()))
                {
                    dgv["col_date_line", i].Style.BackColor = Color.Red;
                    dgv["col_current", i].Style.BackColor = Color.Red;
                    dgv["col_fg", i].Style.BackColor = Color.Red;
                    dgv["col_speed", i].Style.BackColor = Color.Red;
                    dgv["col_svfi", i].Style.BackColor = Color.Red;
                    dgv["col_judge_inline", i].Style.BackColor = Color.Red;

                    if (dgv.Name == "dgvInline") tabControl1.SelectedIndex = 1;
                    else tabControl1.SelectedIndex = 0;

                    soundAlarm();
                }
                else
                {
                    dgv.Rows[i].InheritedStyle.BackColor = Color.FromKnownColor(KnownColor.Window);

                    tabControl1.SelectedIndex = 0;
                }
            }
        }
        private void colorViewForDuplicateSerial(ref DataGridView dgv)
        {
            DataTable dt = ((DataTable)dgv.DataSource).Copy();
            if (dt.Rows.Count <= 0) return;

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                string serial;
                serial = dgv["col_serial_no", i].Value.ToString();

                DataRow[] dr = dt.Select("serialno = '" + serial + "'");
                if (dr.Length >= 2 || dgv["col_serial_no", i].Value.ToString().Length >= 25)
                {
                    if (dgv.Name == "dgvInline") tabControl1.SelectedIndex = 1;
                    else tabControl1.SelectedIndex = 0;

                    dgv["col_serial_no", i].Style.BackColor = Color.Red;
                    soundAlarm();
                }
                else
                {
                    dgv["col_serial_no", i].Style.BackColor = Color.FromKnownColor(KnownColor.Window);
                    tabControl1.SelectedIndex = 0;
                }
            }
        }

        private void btnDeleteSelection_Click(object sender, EventArgs e)
        {
            DataGridView dgv = new DataGridView();

            if (tabControl1.SelectedTab == tabControl1.TabPages["tabInline"])
                dgv = dgvInline;

            if (dgv.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2)
            {
                MessageBox.Show("Please select range with only one columns.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            DialogResult result = MessageBox.Show("Do you really want to delete the selected rows?",
                "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                foreach (DataGridViewCell cell in dgv.SelectedCells)
                {
                    int i = cell.RowIndex;
                    dtOverall.Rows[i].Delete();
                }
                dtOverall.AcceptChanges();
                updateDataGridViews(dtOverall, ref dgvInline);

                txtProductSerial.Focus();
                txtProductSerial.SelectAll();
                txtProductSerial.Enabled = true;
                lbFRICTIONAlarm.Text = "";
                lbFRICTIONAlarm.BackColor = Color.LightGray;
                txtCount.Clear();
                txtCount.BackColor = Color.LightGray;
                txtResultDetail.Clear();
                txtResultDetail.BackColor = Color.LightGray;

            }
        }

        private void btnChangeLimit_Click(object sender, EventArgs e)
        {
            bool bl = TfGeneral.checkOpenFormExists("frmCapacity");
            if (bl)
            {
                MessageBox.Show("Please close or complete another form.", "Warning",
                MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmCapacity f4 = new frmCapacity();
                f4.RefreshEvent += delegate (object sndr, EventArgs excp)
                {
                    int l = f4.getLimit();
                    if (l != 0)
                    {
                        limit2 = f4.getLimit();
                        txtLimit.Text = limit2.ToString();
                        limit1 = limit2;
                    }
                    updateDataGridViews(dtOverall, ref dgvInline);
                    this.Focus();
                };

                f4.updateControls(limit2.ToString());
                f4.Show();
            }
        }

        private void btnAddSerial_Click(object sender, EventArgs e)
        {
            if (!formAddMode)
            {
                formAddMode = true;
                btnAddSerial.Text = "Register";
                btnRegisterBoxId.Enabled = false;
                btnExport.Enabled = false;
                btnCancelBoxid.Enabled = false;
                btnDeleteSerial.Enabled = false;
                txtProductSerial.Enabled = true;
                if (dtOverall.Rows.Count >= 0)
                {
                    formReturnMode = (dtOverall.Rows[0]["return"].ToString() == "R" ? true : false);
                }
            }
            else
            {
                string boxId = txtBoxId.Text;
                string[] model = cmbModel.Text.Split('_');
                //productTable = "product_serial_"+ model[0].ToLower() + model[1];
                productTable = "product_serial_0025";
                string sql = "delete from " + productTable + " where boxid = '" + boxId + "'";
                System.Diagnostics.Debug.Print(sql);
                TfSQL tf = new TfSQL();
                bool res1 = tf.sqlExecuteNonQuery(sql, false);
                DataTable dt = dtOverall.Copy();
                dt.Columns.Add("boxid", Type.GetType("System.String"));
                for (int i = 0; i < dt.Rows.Count; i++)
                    dt.Rows[i]["boxid"] = boxId;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string buff = string.Empty;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        buff += dt.Rows[i][j].ToString() + "  ";
                        System.Diagnostics.Debug.Print(buff);
                    }
                }
                bool res2;
                if (cmbModel.Text == "BFB_0025")
                    res2 = tf.sqlMultipleInsertBFB0025(dt);

                else
                    res2 = tf.sqlMultipleInsertBFB0025(dt);

                if (!res1 || !res2)
                {
                    MessageBox.Show("Error happened in the register process.", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                }
                else
                {
                    MessageBox.Show("Register completed.", "Notice",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                }

                formAddMode = false;
                btnAddSerial.Text = "Add Product";
                btnRegisterBoxId.Enabled = true;
                btnExport.Enabled = true;
                btnCancelBoxid.Enabled = true;
                btnDeleteSerial.Enabled = true;
                txtProductSerial.Enabled = false;
                txtProductSerial.Text = string.Empty;
            }
        }

        private void btnDeleteSerial_Click(object sender, EventArgs e)
        {
            if (dgvInline.Columns.GetColumnCount(DataGridViewElementStates.Selected) >= 2)
            {
                MessageBox.Show("Please select range with only one columns.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                return;
            }

            DialogResult result = MessageBox.Show("Do you really want to delete the selected rows?",
                "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                string boxId = txtBoxId.Text;
                string whereSer = string.Empty;
                string[] model = cmbModel.Text.Split('_');
                productTable = "product_serial_" + model[1];

                foreach (DataGridViewCell cell in dgvInline.SelectedCells)
                {
                    whereSer += "'" + cell.Value.ToString() + "', ";
                }
                string sql = "delete from " + productTable + " where boxid = '" + boxId + "' and  serialno in (" + VBS.Left(whereSer, whereSer.Length - 2) + ")";
                System.Diagnostics.Debug.Print(sql);
                TfSQL tf = new TfSQL();
                int res = tf.sqlExecuteNonQueryInt(sql, false);

                if (res >= 1)
                {
                    foreach (DataGridViewCell cell in dgvInline.SelectedCells)
                    {
                        int i = cell.RowIndex;
                        dtOverall.Rows[i].Delete();
                    }
                    dtOverall.AcceptChanges();
                    updateDataGridViews(dtOverall, ref dgvInline);
                    MessageBox.Show(res.ToString() + " module(s) deleted.", "Notice",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                }
                else
                {
                    MessageBox.Show("Delete failed.", "Notice",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                }
                txtCount.Clear();
                txtCount.BackColor = Color.LightGray;
                txtResultDetail.Clear();
                txtResultDetail.BackColor = Color.LightGray;
            }
        }

        private void btnCancelBoxid_Click(object sender, EventArgs e)
        {
            DialogResult result1 = MessageBox.Show("Do you really delete this box id's all the serial data?",
               "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result1 == DialogResult.Yes)
            {
                DialogResult result2 = MessageBox.Show("Are you really sure? Please select NO if you are not sure.",
                    "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result2 == DialogResult.Yes)
                {
                    string boxid = txtBoxId.Text;
                    TfSQL tf = new TfSQL();
                    int res = tf.sqlDeleteBoxid_0025(boxid);

                    dtOverall.Clear();
                    updateDataGridViews(dtOverall, ref dgvInline);

                    this.RefreshEvent(this, new EventArgs());
                    this.Focus();

                    if (res != -1)
                    {
                        MessageBox.Show("Boxid " + boxid + " and its " + res + " products were deleted.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        btnCancelBoxid.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("An Error has happened in the process and no data has been deleted.", "Process Result", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }
                    btnAddSerial.Enabled = false;
                    btnExport.Enabled = false;
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            string formName = "frmCapacity";
            bool bl = false;
            foreach (Form buff in Application.OpenForms)
            {
                if (buff.Name == formName) { bl = true; }
            }
            if (bl)
            {
                MessageBox.Show("You need to close another form before canceling.", "Notice",
                  MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                return;
            }

            if (dtOverall.Rows.Count == 0 || !formEditMode)
            {
                Application.OpenForms["frmBox"].Focus();
                Close();
                return;
            }

            DialogResult result = MessageBox.Show("The current serial data has not been saved." + System.Environment.NewLine +
                "Do you rally cancel?", "Notice", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                dtOverall.Clear();
                updateDataGridViews(dtOverall, ref dgvInline);
                MessageBox.Show("The temporary serial numbers are deleted.", "Notice",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                Application.OpenForms["frmBox"].Focus();
                Close();
            }
            else
            {
                return;
            }
        }

        private void decideReferenceTable()
        {
            //testerTableThisMonth = cmbModel.Text + "202306";
            //tableThisMonth = testerTableThisMonth;
            //testerTableLastMonth = cmbModel.Text + "202305";
            //tableLastMonth = testerTableLastMonth;

            testerTableThisMonth = cmbModel.Text.ToLower() + DateTime.Today.ToString("yyyyMM");
            tableThisMonth = testerTableThisMonth;
            testerTableLastMonth = cmbModel.Text.ToLower() + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
            tableLastMonth = testerTableLastMonth;

            testerTableThisMonth1 = cmbModel.Text.ToLower() + "_b3a502757" + DateTime.Today.ToString("yyyyMM");
            tableThisMonth1 = testerTableThisMonth1;
            testerTableLastMonth1 = cmbModel.Text.ToLower() + "_b3a502757" + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
    (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
            tableLastMonth1 = testerTableLastMonth1;

            //tableAssyThisMonth = "la20_523ab" + DateTime.Today.ToString("yyyyMM");
            //tableAssyLastMonth = "la20_523ab" + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
            //    (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
        }
        private string checkDataTableWithRealTable(DataTable dt1)
        {
            string serial;
            string result = String.Empty;
            if (formReturnMode) return result;
            if (cmbModel.Text == "BFB_0025")
            {
                string model = "BFB0025";
                productTable = "product_serial_0025";
                string sql = "select serialno, boxid FROM " + productTable;

                DataTable dt2 = new DataTable();
                TfSQL tf = new TfSQL();
                tf.sqlDataAdapterFillDatatable(sql, ref dt2);

                for (int i = 0; i < dt1.Rows.Count; i++)
                {

                    serial = dt1.Rows[i]["serialno"].ToString();
                    DataRow[] dr = dt2.Select("serialno = '" + serial + "'");
                    if (dr.Length >= 1)
                    {
                        string boxid = dr[0]["boxId"].ToString();
                        result += (i + 1 + ": " + serial + " / " + boxid + Environment.NewLine);
                    }
                }

                if (result == String.Empty)
                {
                    return String.Empty;
                }
                else
                {
                    return result;
                }

            }
            else
            {
                string[] model = cmbModel.Text.Split('_');
                productTable = "product_serial_" + model[1];
                string sql = "select serialno, boxid FROM " + productTable;

                DataTable dt2 = new DataTable();
                TfSQL tf = new TfSQL();
                tf.sqlDataAdapterFillDatatable(sql, ref dt2);

                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    if (cmbModel.Text == "LA20_517CB")
                    {
                        serial = VBS.Mid(dt1.Rows[i]["serialno"].ToString(), 2, 21);
                    }
                    else serial = dt1.Rows[i]["serialno"].ToString();
                    DataRow[] dr = dt2.Select("serialno = '" + serial + "'");
                    if (dr.Length >= 1)
                    {
                        string boxid = dr[0]["boxId"].ToString();
                        result += (i + 1 + ": " + serial + " / " + boxid + Environment.NewLine);
                    }
                }

                if (result == String.Empty)
                {
                    return String.Empty;
                }
                else
                {
                    return result;
                }
            }

        }
        private void printBarcode(string dir, string id, string m_model_long, DataGridView dgv1, ref DataGridView dgv2, ref TextBox txt, string shipKind)
        {
            TfPrint tf = new TfPrint();
            tf.createBoxidFiles(dir, id, m_model_long, dgv1, ref dgv2, ref txt, shipKind);
        }
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x112;
            const long SC_CLOSE = 0xF060L;
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & 0xFFF0L) == SC_CLOSE) { return; }
            base.WndProc(ref m);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (txtBoxId.Text == "")
            {
                MessageBox.Show("Please input the carton number!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtCarton.Focus();
                return;
            }
            DataTable dt1 = new DataTable();
            dt1 = (DataTable)dgvInline.DataSource;
            ExcelClass xl = new ExcelClass();
            xl.Export(dgvInline, txtBoxId.Text, "Box " + txtBoxId.Text);
        }

        private void cmbModel_SelectedIndexChanged(object sender, EventArgs e)
        {
            //txtCarton.Enabled = true;
            //{
            //    limit1 = 500;
            //    txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
            //    txtProductSerial.Enabled = true;
            //    txtProductSerial.Focus();
            //}
        }

        private void txtCarton_TextChanged(object sender, EventArgs e)
        {
            if (lblFrmName.Text != "VIEW")
            {
                if (cmbModel.Text == "BFB_0025")
                {
                    string[] box = cmbModel.Text.Split('_');
                    txtBoxId.Text = box[1] + "-" + DateTime.Today.ToString("yyMMdd") + "-" + txtCarton.Text;
                }
                else
                {
                    string[] box = cmbModel.Text.Split('_');
                    txtBoxId.Text = box[1] + "-" + DateTime.Today.ToString("yyMMdd") + "-" + txtCarton.Text;
                }
            }
        }

        private void txtCarton_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && !Char.IsControl(e.KeyChar))
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(e.KeyChar.ToString(), "\\d+"))
                {
                    e.Handled = true;
                    MessageBox.Show("Please input only number!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        private void printDataTable(DataTable dt)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    System.Diagnostics.Debug.Print(dt.Rows[i][j].ToString());
                }
            }
        }
        private void printDataView(DataView dv)
        {
            foreach (DataRowView drv in dv)
            {
                System.Diagnostics.Debug.Print(drv["lot"].ToString() + " " +
                    drv["tjudge"].ToString() + " " + drv["inspectdate"].ToString());
            }
        }

        private void dgvInline_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || dtAllProcess.Rows.Count == 0)
            {
                return;
            }
            string serial = dgvInline.Rows[e.RowIndex].Cells[0].Value.ToString();
            ShowProcessJudge(serial);
        }

        private void txtProductSerial_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (cmbModel.Text == "")
                    {

                        MessageBox.Show("Please select model name", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cmbModel.Focus();
                        return;
                    }
                    else
                    {
                        txtCount.Clear();
                        lbFRICTIONAlarm.Text = "";
                        // lbFRICTIONAlarm_2.Text = "";
                        txtResultDetail.Clear();
                        txtProductSerial.Enabled = false;
                        string serial = txtProductSerial.Text;

                        decideReferenceTable();
                        TfSQL tf = new TfSQL();
                        string model = cmbModel.Text;
                        //DataTable dtFAN = new DataTable();
                        DataTable dtFRICTION = new DataTable();
                        //string sqlFAN = string.Format("select serno, inspectdate, tjudge from {1} where serno = '{0}' and process = 'FAN'  UNION ALL select serno, inspectdate, tjudge from {2} where serno = '{0}' and process = 'FAN' order by inspectdate", txtProductSerial.Text, tableThisMonth, tableLastMonth);
                        //System.Diagnostics.Debug.Print(System.Environment.NewLine + sqlFAN);

                        //tf.sqlDataAdapterFillDatatableOqc(sqlFAN, ref dtFAN);
                        string sqlFRICTION = string.Format("select serno, inspectdate, tjudge from {1} where serno = '{0}' and process = 'SVFRICTION'  UNION ALL select serno, inspectdate, tjudge from {2} where serno = '{0}' and process = 'SVFRICTION' order by inspectdate DESC", txtProductSerial.Text, tableThisMonth, tableLastMonth);
                        System.Diagnostics.Debug.Print(System.Environment.NewLine + sqlFRICTION);
                        tf.sqlDataAdapterFillDatatableOqc(sqlFRICTION, ref dtFRICTION);
                        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        if (dtFRICTION.Rows.Count > 0)
                        {
                            int numberng = 0;
                            frictionnorow = false;
                            for (int i = 0; i < dtFRICTION.Rows.Count; i++)
                            {
                                if (dtFRICTION.Rows[i]["tjudge"].ToString() == "1")
                                {
                                    numberng++;
                                }
                            }
                            if (numberng <= 1 && dtFRICTION.Rows[0]["tjudge"].ToString() == "0")
                            {
                                frictiongood = true;
                                //int noiFRICTION = 1;
                                //string countdtFRICTION = dtFRICTION.Rows.Count.ToString();
                                int maxRowIndex = dtFRICTION.Rows.Count;
                                List<string> showFRICTION = new List<string>();
                                for (int rowIndex = maxRowIndex - 1; rowIndex >= 0; rowIndex--)
                                {
                                    DataRow row = dtFRICTION.Rows[rowIndex];
                                    string value = row[2].ToString();

                                    if (value == "0")
                                        value = "OK";
                                    if (value == "1")
                                        value = "NG";

                                    int noiFRICTION = maxRowIndex - rowIndex;
                                    showFRICTION.Add("No " + noiFRICTION + ": " + value + "\r\n");
                                }
                                //foreach (DataRow row in dtFRICTION.Rows)
                                //{
                                //    string value = row[2].ToString();
                                //    if (value == "0")
                                //        value = "OK";
                                //    if (value == "1")
                                //        value = "NG";
                                //    if (noiFRICTION <= dtFRICTION.Rows.Count)
                                //    {
                                //        showFRICTION.Add("No " + noiFRICTION + ": " + value + "\r\n");
                                //        noiFRICTION++;
                                //    }
                                //}
                                lbFRICTIONAlarm.Text = "Data FRICTION Đã Kiểm " + maxRowIndex + " lần\r\n" + String.Join("", showFRICTION.ToArray());
                            }
                            else
                            {

                                frictiongood = false;
                                string countdt = dtFRICTION.Rows.Count.ToString();
                                List<string> show = new List<string>();
                                int maxRowIndex = Int32.Parse(countdt);
                                for (int rowIndex = Int32.Parse(countdt) - 1; rowIndex >= 0; rowIndex--)
                                {
                                    DataRow row = dtFRICTION.Rows[rowIndex];
                                    string value = row[2].ToString();
                                    if (value == "0")
                                        value = "OK";
                                    if (value == "1")
                                        value = "NG";
                                    int noiFRICTION = maxRowIndex - rowIndex;
                                    show.Add("No " + noiFRICTION + ": " + value + "\r\n");
                                }
                                lbFRICTIONAlarm.Text = "Data FRICTION Đã Kiểm " + countdt + " lần\r\n" + String.Join("", show.ToArray());
                            }
                        }
                        else
                        {
                            frictionnorow = true;
                        }
                        #region Data OQC
                        string sql2 = "select serno, tjudge, inspectdate, " +
                  "MAX(case inspect when 'QACURREN' then inspectdata else null end) as QACURRENT," +
                  "MAX(case inspect when 'QAFG' then inspectdata else null end) as QAFG," +
                  "MAX(case inspect when 'QASPEED' then inspectdata else null end) as QASPEED" +

                  " FROM" +
                  " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableThisMonth + "data" +
                  " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'QAFAN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('QACURREN','QAFG','QASPEED'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableThisMonth + " where serno = '" + serial + "' and process = 'QAFAN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                  " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                  " GROUP BY serno, tjudge, inspectdate" +

                  " UNION ALL " +

                  "select serno, tjudge, inspectdate, " +
                   "MAX(case inspect when 'QACURREN' then inspectdata else null end) as QACURRENT," +
                  "MAX(case inspect when 'QAFG' then inspectdata else null end) as QAFG," +
                  "MAX(case inspect when 'QASPEED' then inspectdata else null end) as QASPEED" +
                  " FROM" +
                  " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableLastMonth + "data" +
                  " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'QAFAN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('QACURREN','QAFG','QASPEED'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableLastMonth + " where serno = '" + serial + "' and process = 'QAFAN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                  " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                  " GROUP BY serno, tjudge, inspectdate";
                        System.Diagnostics.Debug.Print(System.Environment.NewLine + sql2);
                        DataTable dt2 = new DataTable();

                        tf.sqlDataAdapterFillDatatableOqc(sql2, ref dt2);
                        #endregion
                        #region Data INLINE FAN
                        string sql1 = "select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                  "MAX(case inspect when 'CURRENT' then inspectdata else null end) as CURRENT," +
                  "MAX(case inspect when 'FG' then inspectdata else null end) as FG," +
                  "MAX(case inspect when 'SPEED' then inspectdata else null end) as SPEED" +

                  " FROM" +
                  " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableThisMonth + "data" +
                  " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'FAN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('CURRENT','FG','SPEED'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableThisMonth + " where serno = '" + serial + "' and process = 'FAN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                  " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                  " GROUP BY serno, tjudge, inspectdate" +

                  " UNION ALL " +

                  "select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                   "MAX(case inspect when 'CURRENT' then inspectdata else null end) as CURRENT," +
                  "MAX(case inspect when 'FG' then inspectdata else null end) as FG," +
                  "MAX(case inspect when 'SPEED' then inspectdata else null end) as SPEED" +

                  " FROM" +
                  " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableLastMonth + "data" +
                  " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'FAN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('CURRENT','FG','SPEED'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableLastMonth + " where serno = '" + serial + "' and process = 'FAN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                  " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                  " GROUP BY serno, tjudge, inspectdate";
                        System.Diagnostics.Debug.Print(System.Environment.NewLine + sql1);
                        DataTable dt1 = new DataTable();
                        tf.sqlDataAdapterFillDatatableOqc(sql1, ref dt1);

                        string sql3 = "select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                                    "MAX(case inspect when 'SVFI' then inspectdata else null end) as SVFI" +
                                    " FROM" +
                                    " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableThisMonth + "data" +
                                    " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'FAN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SVFI'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableThisMonth + " where serno = '" + serial + "' and process = 'FAN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                                    " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                                    " GROUP BY serno, tjudge, inspectdate" +
                        " UNION ALL " +
                        "select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                         "MAX(case inspect when 'SVFI' then inspectdata else null end) as SVFI" +
                        " FROM" +
                        " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableLastMonth + "data" +
                        " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'FAN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SVFI'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableLastMonth + " where serno = '" + serial + "' and process = 'FAN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                        " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                        " GROUP BY serno, tjudge, inspectdate";
                        System.Diagnostics.Debug.Print(System.Environment.NewLine + sql3);
                        DataTable dt3 = new DataTable();
                        tf.sqlDataAdapterFillDatatableOqc(sql3, ref dt3);

                        #endregion
                        #region -- Get All Process Judge --
                        string queryProcess = string.Format("SELECT serno, lot, inspectdate, process,judge from "
                            + "(SELECT serno, lot, inspectdate, process, judge, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from "
                            + "(SELECT DISTINCT '{0}' as serno, serno lot,inspectdate, process,"
                            + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {1} "
                            + "WHERE serno in (SELECT DISTINCT lot FROM {1} WHERE process = 'EN' AND serno = '{0}') "
                            + "OR serno = '{0}' "
                            + "UNION ALL SELECT DISTINCT '{0}' as serno, serno lot,inspectdate ,process,"
                            + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {2} "
                            + "WHERE serno in (SELECT DISTINCT lot FROM {2} WHERE process = 'EN' AND serno = '{0}') "
                            + "OR serno = '{0}' ORDER BY process) tbl) tb where ROW_NUMBER = 1 and process in ('FAN','QAFAN', 'SVFRICTION')", serial, tableThisMonth, tableLastMonth);
                        System.Diagnostics.Debug.Print(System.Environment.NewLine + queryProcess);
                        if (dtAllProcess == null || dtAllProcess.Rows.Count == 0)
                        {
                            dtAllProcess = new DataTable();
                            tf.sqlDataAdapterFillDatatablePqm(queryProcess, ref dtAllProcess);
                        }
                        else
                        {
                            bool checkSerial = dtAllProcess.AsEnumerable().Any(x => x["serno"].ToString() == serial);
                            if (!checkSerial)
                            {
                                DataTable dtProcessCurrentSerial = new DataTable();
                                tf.sqlDataAdapterFillDatatablePqm(queryProcess, ref dtProcessCurrentSerial);
                                dtAllProcess.Merge(dtProcessCurrentSerial);
                            }
                        }
                        ShowProcessJudge(serial);
                        dtAllProcess.Clear();
                        #endregion
                        DataRow dr = dtOverall.NewRow();
                        // dr["model"] = model.Substring(0, 4);
                        //  string modeldt = String.Empty;

                        dr["model"] = "BFB_0025";
                        dr["serialno"] = serial;
                        dr["lot"] = VBS.Mid(serial, 1, 5);
                        dr["return"] = formReturnMode ? "R" : "N";

                        if (dt2.Rows.Count != 0)
                        {
                            //T-judge OQC
                            dr["qacurrent"] = dt2.Rows[0]["qacurrent"].ToString();
                            dr["qafg"] = dt2.Rows[0]["qafg"].ToString();
                            dr["qaspeed"] = dt2.Rows[0]["qaspeed"].ToString();
                            string linepass = String.Empty;
                            string buff = dt2.Rows[0]["tjudge"].ToString();
                            if (buff == "0") linepass = "PASS";
                            else if (buff == "1") linepass = "FAIL";
                            else linepass = "ERROR";

                            dr["tjudge"] = linepass;
                            dr["inspectdate"] = dt2.Rows[0]["inspectdate"].ToString();
                        }

                        if (dt1.Rows.Count != 0)
                        {
                            dr["current"] = dt1.Rows[0]["current"].ToString();
                            dr["fg"] = dt1.Rows[0]["fg"].ToString();
                            dr["speed"] = dt1.Rows[0]["speed"].ToString();
                            string judge_line = String.Empty;

                            string buff = dt1.Rows[0]["tjudge_line"].ToString();
                            if (buff == "0" && frictiongood) judge_line = "PASS";
                            else if (buff == "0"&& !frictiongood) judge_line = "FAIL";
                            else if (buff == "1" || !frictiongood) judge_line = "FAIL";
                            else judge_line = "ERROR";
                            dr["tjudge_line"] = judge_line;
                            dr["date_line"] = dt1.Rows[0]["date_line"].ToString();
                        }
                        if (dt3.Rows.Count != 0&& frictiongood)
                        {
                            dr["svfi"] = dt3.Rows[0]["svfi"].ToString();
                        }
                        dtOverall.Rows.Add(dr);
                        updateDataGridViews(dtOverall, ref dgvInline);
                        if (okCount >= limit1 && !formAddMode)
                        {
                            txtProductSerial.Enabled = false;
                        }
                        else
                        {
                            txtProductSerial.Enabled = true;
                            txtProductSerial.Focus();
                            txtProductSerial.SelectAll();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "CAUTION", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void ShowProcessJudge(string serial)
        {
            if (dtAllProcess.Rows.Count > 0)
            {
                var datastring = string.Empty;
                var datarows = dtAllProcess.AsEnumerable().Where(x => x["serno"].ToString() == serial).ToList();
                for (int i = 0; i < datarows.Count; i++)
                {
                    var process = datarows[i]["process"] ?? string.Empty;
                    var judge = datarows[i]["judge"] ?? string.Empty;
                    if (!string.IsNullOrEmpty(process.ToString()) && process.ToString() != "SVFRICTION")
                    {
                        datastring += string.Format("{0}: {1}\r\n", process.ToString(), judge.ToString());

                    }
                }
                var checkFail = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("judge").Contains("FAILURE"));
                var checkFAN = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("FAN"));
                var checkFANOQC = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("QAFAN"));
                var checkFRICTION = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("SVFRICTION"));

                //   var checkenQA = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NMT5"));
                if (!checkFAN)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    lbFRICTIONAlarm.BackColor = Color.Red;
                    datastring += "FAN: NO DATA\r\n";
                }
                if (!frictiongood && !frictionnorow)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    lbFRICTIONAlarm.BackColor = Color.Red;
                    datastring += "FRICTION: FAILURE\r\n";
                }
                else if (!checkFRICTION)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    lbFRICTIONAlarm.BackColor = Color.Red;
                    datastring += "FRICTION: NO DATA\r\n";
                }
                else if(frictiongood)
                {
                    datastring += "FRICTION: PASS\r\n";
                }
                if (!checkFANOQC)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    lbFRICTIONAlarm.BackColor = Color.Red;
                    datastring += "QAFAN: NO DATA\r\n";
                }

                if (checkFail)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    lbFRICTIONAlarm.BackColor = Color.Red;
                    txtResultDetail.Text = datastring;
                }
                if (!checkFail && checkFAN && checkFRICTION && checkFANOQC && frictiongood)
                {
                    txtCount.Text = "OK";
                    txtCount.BackColor = Color.SpringGreen;
                    txtResultDetail.BackColor = Color.SpringGreen;
                    lbFRICTIONAlarm.BackColor = Color.SpringGreen;
                    txtResultDetail.Text = datastring;
                }
                txtResultDetail.Text = datastring;
            }
        }
        [System.Runtime.InteropServices.DllImport("winmm.dll")]
        private static extern int mciSendString(String command,
           StringBuilder buffer, int bufferSize, IntPtr hwndCallback);

        private string aliasName = "MediaFile";

        private void soundAlarm()
        {
            string currentDir = System.Environment.CurrentDirectory;
            string fileName = currentDir + @"\warning.mp3";
            string cmd;

            if (sound)
            {
                cmd = "stop " + aliasName;
                mciSendString(cmd, null, 0, IntPtr.Zero);
                cmd = "close " + aliasName;
                mciSendString(cmd, null, 0, IntPtr.Zero);
                sound = false;
            }

            cmd = "open \"" + fileName + "\" type mpegvideo alias " + aliasName;
            if (mciSendString(cmd, null, 0, IntPtr.Zero) != 0) return;
            cmd = "play " + aliasName;
            mciSendString(cmd, null, 0, IntPtr.Zero);
            sound = true;
        }
    }
}
