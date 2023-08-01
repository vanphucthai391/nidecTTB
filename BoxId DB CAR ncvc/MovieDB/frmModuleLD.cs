using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Security.Permissions;
using System.Runtime.InteropServices;
using System.Linq;
using System.Collections.Generic;

namespace BoxIdDb
{
    public partial class frmModuleLD : Form
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
        int bin1, bin2, bin3, bin4;
        bool inputBoxModeOriginal;
        string productTable;
        string testerTableThisMonth;
        string testerTableLastMonth;
        string tableThisMonth;
        string tableLastMonth;
        //string tableAssyThisMonth, tableAssyLastMonth;
        DataTable dtOverall;
        DataTable dtAllProcess;
        int limit1 = 500;
        public int limit2 = 0;
        bool sound;
        public frmModuleLD()
        {
            InitializeComponent();
            dtAllProcess = new DataTable();
        }


        private void frmModuleLD_Load(object sender, EventArgs e)
        {
            txtCarton.Enabled = false;
            user = txtUser.Text;
            txtBin1.Enabled = false;
            txtBin2.Enabled = false;
            txtBin3.Enabled = false;
            txtBin4.Enabled = false;
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
                    limit1 = 500;
                }
            }
            updateDataGridViews(dtOverall, ref dgvInline);
            if (!formEditMode)
            {
                txtProductSerial.Enabled = false;
            }
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
                if (model == "LD20")
                    model = "ld20";
                productTable = "product_serial_" + model;
                switch (model)
                {
                    case "LD20":
                        cmbModel.Text = "LD20";
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
            if (editMode && user == "admin" || editMode && user == "User_9")
            {
                btnChangeLimit.Visible = true;
                txtLimit.Visible = true;
            }
            if (!editMode && user == "admin" || !editMode && user == "User_9")
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
                string[] model = cmbModel.Text.Split('-');
                productTable = "product_serial_" + model[0];
            }
        }
        private void defineAndReadDtOverall(ref DataTable dt)
        {
            string boxId = txtBoxId.Text;
            setProductTable();
            dt.Columns.Add("serialno", Type.GetType("System.String"));
            dt.Columns.Add("model", Type.GetType("System.String"));
            dt.Columns.Add("lot", Type.GetType("System.String"));
            //dt.Columns.Add("inspectdate", Type.GetType("System.DateTime")); //date test NMT OQC
            //dt.Columns.Add("sdf0_oqc", Type.GetType("System.String"));
            //dt.Columns.Add("scurave_oqc", Type.GetType("System.String"));
            //dt.Columns.Add("sgrmsave_oqc", Type.GetType("System.String"));
            //dt.Columns.Add("srtpctg2_oqc", Type.GetType("System.String"));
            //dt.Columns.Add("sbtpctg2_oqc", Type.GetType("System.String"));
            //dt.Columns.Add("tjudge", Type.GetType("System.String"));
            dt.Columns.Add("date_line", Type.GetType("System.DateTime")); //date test Inline
            dt.Columns.Add("sdf0", Type.GetType("System.String"));
            dt.Columns.Add("scurave", Type.GetType("System.String"));
            dt.Columns.Add("sgrmsave", Type.GetType("System.String"));
            dt.Columns.Add("srtpctg2", Type.GetType("System.String"));
            dt.Columns.Add("sbtpctg2", Type.GetType("System.String"));
            dt.Columns.Add("tjudge_line", Type.GetType("System.String"));
            dt.Columns.Add("bin", Type.GetType("System.String"));
            dt.Columns.Add("return", Type.GetType("System.String"));

            if (!formEditMode)
            {
                string sql;
                sql = "select serialno, model, lot, date_line, SDF0, SBTPCTG2, SCURAVE, SGRMSAVE, SRTPCTG2, tjudge_line,bin, return " +
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
            bin1 = getCountBin1(dt1);
            txtBin1.Text = bin1.ToString();
            bin2 = getCountBin2(dt1);
            txtBin2.Text = bin2.ToString();
            bin3 = getCountBin3(dt1);
            txtBin3.Text = bin3.ToString();
            bin4 = getCountBin4(dt1);
            txtBin4.Text = bin4.ToString();


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
            }
            else
            {
                btnRegisterBoxId.Enabled = false;
            }

        }
        private int getOkCount(DataTable dt)
        {
            if (dt.Rows.Count <= 0) return 0;
            //DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge", "tjudge_line" });
            //DataRow[] dr = distinct.Select("tjudge = 'PASS' and tjudge_line = 'PASS'");
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge_line" });
            DataRow[] dr = distinct.Select("tjudge_line = 'PASS'");
            int dist = dr.Length;
            return dist;
        }
        private int getCountBin1(DataTable dt)
        {
            if (dt.Rows.Count <= 0) return 0;
            //DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge", "tjudge_line" });
            //DataRow[] dr = distinct.Select("tjudge = 'PASS' and tjudge_line = 'PASS'");
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "bin" });
            DataRow[] dr = distinct.Select("bin = 'Bin 1'");
            int dist = dr.Length;
            return dist;
        }
        private int getCountBin2(DataTable dt)
        {
            if (dt.Rows.Count <= 0) return 0;
            //DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge", "tjudge_line" });
            //DataRow[] dr = distinct.Select("tjudge = 'PASS' and tjudge_line = 'PASS'");
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "bin" });
            DataRow[] dr = distinct.Select("bin = 'Bin 2'");
            int dist = dr.Length;
            return dist;
        }
        private int getCountBin3(DataTable dt)
        {
            if (dt.Rows.Count <= 0) return 0;
            //DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge", "tjudge_line" });
            //DataRow[] dr = distinct.Select("tjudge = 'PASS' and tjudge_line = 'PASS'");
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "bin" });
            DataRow[] dr = distinct.Select("bin = 'Bin 3'");
            int dist = dr.Length;
            return dist;
        }
        private int getCountBin4(DataTable dt)
        {
            if (dt.Rows.Count <= 0) return 0;
            //DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge", "tjudge_line" });
            //DataRow[] dr = distinct.Select("tjudge = 'PASS' and tjudge_line = 'PASS'");
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "bin" });
            DataRow[] dr = distinct.Select("bin = 'Bin 4'");
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
        private void colorViewForFailAndBlank(ref DataGridView dgv)
        {
            int row = dgv.Rows.Count;
            for (int i = 0; i < row; ++i)
            {
                #region ALARM OQC FAIL
                //Alarm OQC FAIL or NODATA
                //if (dgv["col_judge_oqc", i].Value.ToString() == "FAIL" || dgv["col_judge_oqc", i].Value.ToString() == "PLS NG" || String.IsNullOrEmpty(dgv["col_judge_oqc", i].Value.ToString()))
                //{
                //    dgv["col_date", i].Style.BackColor = Color.Red;
                //    dgv["col_sdf0_oqc", i].Style.BackColor = Color.Red;
                //    dgv["col_sbtpctg2_oqc", i].Style.BackColor = Color.Red;
                //    dgv["col_scurave_oqc", i].Style.BackColor = Color.Red;
                //    dgv["col_sgrmsave_oqc", i].Style.BackColor = Color.Red;
                //    dgv["col_srtpctg2_oqc", i].Style.BackColor = Color.Red;
                //    dgv["col_judge_oqc", i].Style.BackColor = Color.Red;

                //    if (dgv.Name == "dgvInline") tabControl1.SelectedIndex = 1;
                //    else tabControl1.SelectedIndex = 0;

                //    soundAlarm();
                //}
                //else
                //{
                //    dgv.Rows[i].InheritedStyle.BackColor = Color.FromKnownColor(KnownColor.Window);

                //    tabControl1.SelectedIndex = 0;
                //}
                #endregion
                //Alarm INLINE FAIL or NODATA
                if (dgv["col_judge_inline", i].Value.ToString() == "FAIL" || dgv["col_judge_inline", i].Value.ToString() == "PLS NG" || String.IsNullOrEmpty(dgv["col_judge_inline", i].Value.ToString()))
                {
                    dgv["col_date_line", i].Style.BackColor = Color.Red;
                    dgv["col_sdf0_inline", i].Style.BackColor = Color.Red;
                    dgv["col_sbtpctg2_inline", i].Style.BackColor = Color.Red;
                    dgv["col_scurave_inline", i].Style.BackColor = Color.Red;
                    dgv["col_sgrmsave_inline", i].Style.BackColor = Color.Red;
                    dgv["col_srtpctg2_inline", i].Style.BackColor = Color.Red;
                    dgv["col_judge_inline", i].Style.BackColor = Color.Red;
                    dgv["bin", i].Style.BackColor = Color.Red;
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


        private void txtProductSerial_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (cmbModel.Text == "" || cmbModelMaster.Text == "")
                    {

                        MessageBox.Show("Please select model name", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cmbModel.Focus();
                        return;
                    }
                    else
                    {
                        txtCount.Clear();
                        txtResultDetail.Clear();
                        txtProductSerial.Enabled = false;
                        string serial = txtProductSerial.Text;

                        decideReferenceTable();
                        string codemodelmaster = null;
                        if (cmbModelMaster.Text == "LD20_001")
                            codemodelmaster = "01";
                        else if (cmbModelMaster.Text == "BMD_0015")
                            codemodelmaster = "02";
                        else if (cmbModelMaster.Text == "BMD_0016")
                            codemodelmaster = "03";
                        else if (cmbModelMaster.Text == "BMD_0019")
                            codemodelmaster = "04";
                        else if (cmbModelMaster.Text == "BMD_0103")
                            codemodelmaster = "05";
                        else if (cmbModelMaster.Text == "BMD_0124")
                            codemodelmaster = "06";
                        else if (cmbModelMaster.Text == "BMD_0219")
                            codemodelmaster = "07";
                        else if (cmbModelMaster.Text == "BMD_0226")
                            codemodelmaster = "08";
                        else if (cmbModelMaster.Text == "BMD_0232")
                            codemodelmaster = "09";
                        else if (cmbModelMaster.Text == "TRIAL MODEL")
                            codemodelmaster = "00";

                        string codecheckmodelmaster = VBS.Left(txtProductSerial.Text, 2);
                        int lengcode = txtProductSerial.Text.Length;
                        if (codemodelmaster == codecheckmodelmaster && lengcode == 10)
                        {
                            TfSQL tf = new TfSQL();
                            string model = cmbModel.Text;
                            DataTable dtfct = new DataTable();
                            string sqlfct = string.Format("select serno, inspectdate, tjudge from {1} where serno = '{0}' and process = 'EN' UNION ALL select serno, inspectdate, tjudge from {2} where serno = '{0}' and process = 'EN' order by inspectdate", txtProductSerial.Text, tableThisMonth, tableLastMonth);
                            tf.sqlDataAdapterFillDatatableOqc(sqlfct, ref dtfct);
                            if (dtfct.Rows.Count <= 3)
                            {
                                int noi = 1;
                                string countdt = dtfct.Rows.Count.ToString();
                                List<string> show = new List<string>();

                                foreach (DataRow row in dtfct.Rows)
                                {
                                    string value = row[2].ToString();
                                    if (value == "0")
                                        value = "OK";
                                    if (value == "1")
                                        value = "NG";

                                    if (noi <= dtfct.Rows.Count)
                                    {
                                        show.Add("No " + noi + ": " + value + "\n");
                                        noi++;
                                    }
                                }

                                lbENAlarm.Text = "Data FCT Đã Kiểm " + countdt + " Lần \n" + String.Join("", show.ToArray());
                                lbENAlarm.BackColor = Color.SpringGreen;

                                #region Data OQC
                                //string sql2 = "select serno, tjudge, inspectdate, " +
                                //"MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0," +
                                //"MAX(case inspect when 'SCURAVE' then inspectdata else null end) as SCURAVE," +
                                //"MAX(case inspect when 'SGRMSAVE' then inspectdata else null end) as SGRMSAVE," +
                                //"MAX(case inspect when 'SRTPCTG2' then inspectdata else null end) as SRTPCTG2," +
                                //"MAX(case inspect when 'SBTPCTG2' then inspectdata else null end) as SBTPCTG2" +
                                //" FROM" +
                                //" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableThisMonth + "data" +
                                //" WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'NMT5' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SDF0','SCURAVE','SGRMSAVE','SRTPCTG2','SBTPCTG2'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableThisMonth + " where serno = '" + serial + "' and process = 'NMT5' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                                //" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                                //" GROUP BY serno, tjudge, inspectdate" +

                                //" UNION ALL " +

                                //"select serno, tjudge, inspectdate, " +
                                //"MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0," +
                                //"MAX(case inspect when 'SCURAVE' then inspectdata else null end) as SCURAVE," +
                                //"MAX(case inspect when 'SGRMSAVE' then inspectdata else null end) as SGRMSAVE," +
                                //"MAX(case inspect when 'SRTPCTG2' then inspectdata else null end) as SRTPCTG2," +
                                //"MAX(case inspect when 'SBTPCTG2' then inspectdata else null end) as SBTPCTG2" +
                                //" FROM" +
                                //" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableLastMonth + "data" +
                                //" WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'NMT5' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SDF0','SCURAVE','SGRMSAVE','SRTPCTG2','SBTPCTG2'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableLastMonth + " where serno = '" + serial + "' and process = 'NMT5' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                                //" group by d.serno, d.tjudge, c.inspectdates, c.inspect, c.inspectdata, c.judge) e " +
                                //" GROUP BY serno, tjudge, inspectdate";

                                //System.Diagnostics.Debug.Print(System.Environment.NewLine + sql2);
                                //DataTable dt2 = new DataTable();
                                //TfSQL tf = new TfSQL();
                                //tf.sqlDataAdapterFillDatatableOqc(sql2, ref dt2);
                                #endregion

                                #region Data INLINE
                                string sql1 = "select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                          "MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0," +
                          "MAX(case inspect when 'SCURAVE' then inspectdata else null end) as SCURAVE," +
                          "MAX(case inspect when 'SGRMSAVE' then inspectdata else null end) as SGRMSAVE," +
                          "MAX(case inspect when 'SRTPCTG2' then inspectdata else null end) as SRTPCTG2," +
                          "MAX(case inspect when 'SBTPCTG2' then inspectdata else null end) as SBTPCTG2" +
                          " FROM" +
                          " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableThisMonth + "data" +
                          " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'EN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SDF0','SCURAVE','SGRMSAVE','SRTPCTG2','SBTPCTG2'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableThisMonth + " where serno = '" + serial + "' and process = 'EN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                          " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                          " GROUP BY serno, tjudge, inspectdate" +

                          " UNION ALL " +

                          "select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                          "MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0," +
                          "MAX(case inspect when 'SCURAVE' then inspectdata else null end) as SCURAVE," +
                          "MAX(case inspect when 'SGRMSAVE' then inspectdata else null end) as SGRMSAVE," +
                          "MAX(case inspect when 'SRTPCTG2' then inspectdata else null end) as SRTPCTG2," +
                          "MAX(case inspect when 'SBTPCTG2' then inspectdata else null end) as SBTPCTG2" +
                          " FROM" +
                          " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableLastMonth + "data" +
                          " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'EN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SDF0','SCURAVE','SGRMSAVE','SRTPCTG2','SBTPCTG2'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableLastMonth + " where serno = '" + serial + "' and process = 'EN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                          " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                          " GROUP BY serno, tjudge, inspectdate";
                                System.Diagnostics.Debug.Print(System.Environment.NewLine + sql1);
                                DataTable dt1 = new DataTable();

                                tf.sqlDataAdapterFillDatatableOqc(sql1, ref dt1);
                                #endregion
                                #region DATA INLINE OLD
                                //string sql1 = "select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                                // "MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0," +
                                //"MAX(case inspect when 'SCURAVE' then inspectdata else null end) as SCURAVE," +
                                //"MAX(case inspect when 'SGRMSAVE' then inspectdata else null end) as SGRMSAVE," +
                                //"MAX(case inspect when 'SRTPCTG2' then inspectdata else null end) as SRTPCTG2," +
                                //"MAX(case inspect when 'SBTPCTG2' then inspectdata else null end) as SBTPCTG2" +
                                //" FROM" +
                                //" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge FROM(select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE FROM (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag FROM (select * from " + tableThisMonth + "data" +
                                //" WHERE serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'EN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SDF0','SCURAVE','SGRMSAVE','SRTPCTG2','SBTPCTG2'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + tableThisMonth + " where serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'EN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and process = 'EN' order by inspectdate desc LIMIT 1) d" +
                                //" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                                //" GROUP BY serno, tjudge, inspectdate";

                                //" UNION ALL " +

                                //"select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                                //  "MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0," +
                                //"MAX(case inspect when 'SCURAVE' then inspectdata else null end) as SCURAVE," +
                                //"MAX(case inspect when 'SGRMSAVE' then inspectdata else null end) as SGRMSAVE," +
                                //"MAX(case inspect when 'SRTPCTG2' then inspectdata else null end) as SRTPCTG2," +
                                //"MAX(case inspect when 'SBTPCTG2' then inspectdata else null end) as SBTPCTG2" +
                                //" FROM" +
                                //" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge FROM(select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE FROM (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag FROM (select * from " + tableLastMonth + "data" +
                                //" WHERE serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'EN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SDF0','SCURAVE','SGRMSAVE','SRTPCTG2','SBTPCTG2'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + tableLastMonth + " where serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'EN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and process = 'EN' order by inspectdate desc LIMIT 1) d" +
                                //" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                                //" GROUP BY serno, tjudge, inspectdate" +

                                //" UNION ALL " +

                                //"select serno, tjudge as tjudge_line, inspectdate as date_line, " +
                                //"MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0," +
                                //"MAX(case inspect when 'SCURAVE' then inspectdata else null end) as SCURAVE," +
                                //"MAX(case inspect when 'SGRMSAVE' then inspectdata else null end) as SGRMSAVE," +
                                //"MAX(case inspect when 'SRTPCTG2' then inspectdata else null end) as SRTPCTG2," +
                                //"MAX(case inspect when 'SBTPCTG2' then inspectdata else null end) as SBTPCTG2" +
                                //" FROM" +
                                //" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge FROM(select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE FROM (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag FROM (select * from " + tableLastMonth + "data" +
                                //" WHERE serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'EN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('SDF0','SCURAVE','SGRMSAVE','SRTPCTG2','SBTPCTG2'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + tableLastMonth + " where serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'EN' and serno = '" + serial + "') order by serno) tbl where row_number =1) and process = 'EN' order by inspectdate desc LIMIT 1) d" +
                                //" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                                //" GROUP BY serno, tjudge, inspectdate";
                                //System.Diagnostics.Debug.Print(System.Environment.NewLine + sql1);
                                //DataTable dt1 = new DataTable();

                                //tf.sqlDataAdapterFillDatatablePqm(sql1, ref dt1);
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
                                    + "OR serno = '{0}' ORDER BY process) tbl) tb where ROW_NUMBER = 1 ", serial, tableThisMonth, tableLastMonth);
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
                                string barcode = txtProductSerial.Text;
                                if (barcode.StartsWith("00"))
                                {
                                    dr["model"] = "BMD-0273";
                                }
                                else if (barcode.StartsWith("01"))
                                {
                                    dr["model"] = "LD20-001";
                                }
                                else if (barcode.StartsWith("02"))
                                {
                                    dr["model"] = "BMD-0015";
                                }
                                else if (barcode.StartsWith("03"))
                                {
                                    dr["model"] = "BMD-0016";
                                }
                                else if (barcode.StartsWith("04"))
                                {
                                    dr["model"] = "BMD-0019";
                                }
                                else if (barcode.StartsWith("05"))
                                {
                                    dr["model"] = "BMD-0103";
                                }
                                else if (barcode.StartsWith("06"))
                                {
                                    dr["model"] = "BMD-0124";
                                }
                                else if (barcode.StartsWith("07"))
                                {
                                    dr["model"] = "BMD-0219";
                                }
                                else if (barcode.StartsWith("08"))
                                {
                                    dr["model"] = "BMD-0226";
                                }
                                else if (barcode.StartsWith("09"))
                                {
                                    dr["model"] = "BMD-0232";
                                }
                                dr["serialno"] = serial;
                                dr["lot"] = VBS.Mid(serial, 3, 4);
                                dr["return"] = formReturnMode ? "R" : "N";

                                //if (dt2.Rows.Count != 0)
                                //{
                                //    //T-judge OQC
                                //    string linepass = String.Empty;
                                //    string buff = dt2.Rows[0]["tjudge"].ToString();
                                //    if (buff == "0") linepass = "PASS";
                                //    else if (buff == "1") linepass = "FAIL";
                                //    else linepass = "ERROR";

                                //    dr["tjudge"] = linepass;
                                //    dr["inspectdate"] = dt2.Rows[0]["inspectdate"].ToString();
                                //}

                                if (dt1.Rows.Count != 0)
                                {
                                    dr["sdf0"] = dt1.Rows[0]["sdf0"].ToString();
                                    dr["scurave"] = dt1.Rows[0]["scurave"].ToString();
                                    dr["sgrmsave"] = dt1.Rows[0]["sgrmsave"].ToString();
                                    dr["srtpctg2"] = dt1.Rows[0]["srtpctg2"].ToString();
                                    dr["sbtpctg2"] = dt1.Rows[0]["sbtpctg2"].ToString();
                                    // dr["bin"] = dt1.Rows[0]["bin"].ToString();
                                    //T-judge LINE
                                    string judge_line = String.Empty;
                                    string buff = dt1.Rows[0]["tjudge_line"].ToString();
                                    if (buff == "0") judge_line = "PASS";
                                    else if (buff == "1") judge_line = "FAIL";
                                    else judge_line = "ERROR";

                                    dr["tjudge_line"] = judge_line;
                                    dr["date_line"] = dt1.Rows[0]["date_line"].ToString();

                                    string bindata = String.Empty;
                                    string bin = dt1.Rows[0]["sdf0"].ToString();
                                    double BinF0 = double.Parse(bin);
                                    if (170.00 <= BinF0 && BinF0 <= 175.049)
                                    {
                                        bindata = "Bin 1";
                                    }
                                    else if (175.05 <= BinF0 && BinF0 <= 180.049)
                                    {
                                        bindata = "Bin 2";

                                    }
                                    else if (180.05 <= BinF0 && BinF0 <= 185.049)
                                    {
                                        bindata = "Bin 3";
                                    }
                                    else if (185.05 <= BinF0 && BinF0 <= 190.00)
                                    {
                                        bindata = "Bin 4";
                                    }
                                    #region old
                                    //#region BIN 1
                                    //if (bin == "170.0") bindata = "Bin 1";
                                    //if (bin == "170") bindata = "Bin 1";
                                    //if (bin == "170.1") bindata = "Bin 1";
                                    //if (bin == "170.2") bindata = "Bin 1";
                                    //if (bin == "170.3") bindata = "Bin 1";
                                    //if (bin == "170.4") bindata = "Bin 1";
                                    //if (bin == "170.5") bindata = "Bin 1";
                                    //if (bin == "170.6") bindata = "Bin 1";
                                    //if (bin == "170.7") bindata = "Bin 1";
                                    //if (bin == "170.8") bindata = "Bin 1";
                                    //if (bin == "170.9") bindata = "Bin 1";
                                    //if (bin == "171.0") bindata = "Bin 1";
                                    //if (bin == "171") bindata = "Bin 1";
                                    //if (bin == "171.1") bindata = "Bin 1";
                                    //if (bin == "171.2") bindata = "Bin 1";
                                    //if (bin == "171.3") bindata = "Bin 1";
                                    //if (bin == "171.4") bindata = "Bin 1";
                                    //if (bin == "171.5") bindata = "Bin 1";
                                    //if (bin == "171.6") bindata = "Bin 1";
                                    //if (bin == "171.7") bindata = "Bin 1";
                                    //if (bin == "171.8") bindata = "Bin 1";
                                    //if (bin == "171.9") bindata = "Bin 1";
                                    //if (bin == "172.0") bindata = "Bin 1";
                                    //if (bin == "172") bindata = "Bin 1";
                                    //if (bin == "172.1") bindata = "Bin 1";
                                    //if (bin == "172.2") bindata = "Bin 1";
                                    //if (bin == "172.3") bindata = "Bin 1";
                                    //if (bin == "172.4") bindata = "Bin 1";
                                    //if (bin == "172.5") bindata = "Bin 1";
                                    //if (bin == "172.6") bindata = "Bin 1";
                                    //if (bin == "172.7") bindata = "Bin 1";
                                    //if (bin == "172.8") bindata = "Bin 1";
                                    //if (bin == "172.9") bindata = "Bin 1";
                                    //if (bin == "173.0") bindata = "Bin 1";
                                    //if (bin == "173") bindata = "Bin 1";
                                    //if (bin == "173.1") bindata = "Bin 1";
                                    //if (bin == "173.2") bindata = "Bin 1";
                                    //if (bin == "173.3") bindata = "Bin 1";
                                    //if (bin == "173.4") bindata = "Bin 1";
                                    //if (bin == "173.5") bindata = "Bin 1";
                                    //if (bin == "173.6") bindata = "Bin 1";
                                    //if (bin == "173.7") bindata = "Bin 1";
                                    //if (bin == "173.8") bindata = "Bin 1";
                                    //if (bin == "173.9") bindata = "Bin 1";
                                    //if (bin == "174.0") bindata = "Bin 1";
                                    //if (bin == "174") bindata = "Bin 1";
                                    //if (bin == "174.1") bindata = "Bin 1";
                                    //if (bin == "174.2") bindata = "Bin 1";
                                    //if (bin == "174.3") bindata = "Bin 1";
                                    //if (bin == "174.4") bindata = "Bin 1";
                                    //if (bin == "174.5") bindata = "Bin 1";
                                    //if (bin == "174.6") bindata = "Bin 1";
                                    //if (bin == "174.7") bindata = "Bin 1";
                                    //if (bin == "174.8") bindata = "Bin 1";
                                    //if (bin == "174.9") bindata = "Bin 1";
                                    //if (bin == "175.0") bindata = "Bin 1";
                                    //if (bin == "175") bindata = "Bin 1";
                                    //#endregion
                                    //#region BIN 2
                                    //if (bin == "175.1") bindata = "Bin 2";
                                    //if (bin == "175.2") bindata = "Bin 2";
                                    //if (bin == "175.3") bindata = "Bin 2";
                                    //if (bin == "175.4") bindata = "Bin 2";
                                    //if (bin == "175.5") bindata = "Bin 2";
                                    //if (bin == "175.6") bindata = "Bin 2";
                                    //if (bin == "175.7") bindata = "Bin 2";
                                    //if (bin == "175.8") bindata = "Bin 2";
                                    //if (bin == "175.9") bindata = "Bin 2";
                                    //if (bin == "176.0") bindata = "Bin 2";
                                    //if (bin == "176") bindata = "Bin 2";
                                    //if (bin == "176.1") bindata = "Bin 2";
                                    //if (bin == "176.2") bindata = "Bin 2";
                                    //if (bin == "176.3") bindata = "Bin 2";
                                    //if (bin == "176.4") bindata = "Bin 2";
                                    //if (bin == "176.5") bindata = "Bin 2";
                                    //if (bin == "176.6") bindata = "Bin 2";
                                    //if (bin == "176.7") bindata = "Bin 2";
                                    //if (bin == "176.8") bindata = "Bin 2";
                                    //if (bin == "176.9") bindata = "Bin 2";
                                    //if (bin == "177.0") bindata = "Bin 2";
                                    //if (bin == "177") bindata = "Bin 2";
                                    //if (bin == "177.1") bindata = "Bin 2";
                                    //if (bin == "177.2") bindata = "Bin 2";
                                    //if (bin == "177.3") bindata = "Bin 2";
                                    //if (bin == "177.4") bindata = "Bin 2";
                                    //if (bin == "177.5") bindata = "Bin 2";
                                    //if (bin == "177.6") bindata = "Bin 2";
                                    //if (bin == "177.7") bindata = "Bin 2";
                                    //if (bin == "177.8") bindata = "Bin 2";
                                    //if (bin == "177.9") bindata = "Bin 2";
                                    //if (bin == "178.0") bindata = "Bin 2";
                                    //if (bin == "178") bindata = "Bin 2";
                                    //if (bin == "178.1") bindata = "Bin 2";
                                    //if (bin == "178.2") bindata = "Bin 2";
                                    //if (bin == "178.3") bindata = "Bin 2";
                                    //if (bin == "178.4") bindata = "Bin 2";
                                    //if (bin == "178.5") bindata = "Bin 2";
                                    //if (bin == "178.6") bindata = "Bin 2";
                                    //if (bin == "178.7") bindata = "Bin 2";
                                    //if (bin == "178.8") bindata = "Bin 2";
                                    //if (bin == "178.9") bindata = "Bin 2";
                                    //if (bin == "179.0") bindata = "Bin 2";
                                    //if (bin == "179") bindata = "Bin 2";
                                    //if (bin == "179.1") bindata = "Bin 2";
                                    //if (bin == "179.2") bindata = "Bin 2";
                                    //if (bin == "179.3") bindata = "Bin 2";
                                    //if (bin == "179.4") bindata = "Bin 2";
                                    //if (bin == "179.5") bindata = "Bin 2";
                                    //if (bin == "179.6") bindata = "Bin 2";
                                    //if (bin == "179.7") bindata = "Bin 2";
                                    //if (bin == "179.8") bindata = "Bin 2";
                                    //if (bin == "179.9") bindata = "Bin 2";
                                    //if (bin == "180.0") bindata = "Bin 2";
                                    //if (bin == "180") bindata = "Bin 2";
                                    //#endregion
                                    //#region BIN 3

                                    //if (bin == "180.1") bindata = "Bin 3";
                                    //if (bin == "180.2") bindata = "Bin 3";
                                    //if (bin == "180.3") bindata = "Bin 3";
                                    //if (bin == "180.4") bindata = "Bin 3";
                                    //if (bin == "180.5") bindata = "Bin 3";
                                    //if (bin == "180.6") bindata = "Bin 3";
                                    //if (bin == "180.7") bindata = "Bin 3";
                                    //if (bin == "180.8") bindata = "Bin 3";
                                    //if (bin == "180.9") bindata = "Bin 3";
                                    //if (bin == "181.0") bindata = "Bin 3";
                                    //if (bin == "181") bindata = "Bin 3";
                                    //if (bin == "181.1") bindata = "Bin 3";
                                    //if (bin == "181.2") bindata = "Bin 3";
                                    //if (bin == "181.3") bindata = "Bin 3";
                                    //if (bin == "181.4") bindata = "Bin 3";
                                    //if (bin == "181.5") bindata = "Bin 3";
                                    //if (bin == "181.6") bindata = "Bin 3";
                                    //if (bin == "181.7") bindata = "Bin 3";
                                    //if (bin == "181.8") bindata = "Bin 3";
                                    //if (bin == "181.9") bindata = "Bin 3";
                                    //if (bin == "182.0") bindata = "Bin 3";
                                    //if (bin == "182") bindata = "Bin 3";
                                    //if (bin == "182.1") bindata = "Bin 3";
                                    //if (bin == "182.2") bindata = "Bin 3";
                                    //if (bin == "182.3") bindata = "Bin 3";
                                    //if (bin == "182.4") bindata = "Bin 3";
                                    //if (bin == "182.5") bindata = "Bin 3";
                                    //if (bin == "182.6") bindata = "Bin 3";
                                    //if (bin == "182.7") bindata = "Bin 3";
                                    //if (bin == "182.8") bindata = "Bin 3";
                                    //if (bin == "182.9") bindata = "Bin 3";
                                    //if (bin == "183.0") bindata = "Bin 3";
                                    //if (bin == "183") bindata = "Bin 3";
                                    //if (bin == "183.1") bindata = "Bin 3";
                                    //if (bin == "183.2") bindata = "Bin 3";
                                    //if (bin == "183.3") bindata = "Bin 3";
                                    //if (bin == "183.4") bindata = "Bin 3";
                                    //if (bin == "183.5") bindata = "Bin 3";
                                    //if (bin == "183.6") bindata = "Bin 3";
                                    //if (bin == "183.7") bindata = "Bin 3";
                                    //if (bin == "183.8") bindata = "Bin 3";
                                    //if (bin == "183.9") bindata = "Bin 3";
                                    //if (bin == "184.0") bindata = "Bin 3";
                                    //if (bin == "184") bindata = "Bin 3";
                                    //if (bin == "184.1") bindata = "Bin 3";
                                    //if (bin == "184.2") bindata = "Bin 3";
                                    //if (bin == "184.3") bindata = "Bin 3";
                                    //if (bin == "184.4") bindata = "Bin 3";
                                    //if (bin == "184.5") bindata = "Bin 3";
                                    //if (bin == "184.6") bindata = "Bin 3";
                                    //if (bin == "184.7") bindata = "Bin 3";
                                    //if (bin == "184.8") bindata = "Bin 3";
                                    //if (bin == "184.9") bindata = "Bin 3";
                                    //if (bin == "185.0") bindata = "Bin 3";
                                    //if (bin == "185") bindata = "Bin 3";
                                    //#endregion
                                    //#region BIN 4
                                    //if (bin == "185.1") bindata = "Bin 4";
                                    //if (bin == "185.2") bindata = "Bin 4";
                                    //if (bin == "185.3") bindata = "Bin 4";
                                    //if (bin == "185.4") bindata = "Bin 4";
                                    //if (bin == "185.5") bindata = "Bin 4";
                                    //if (bin == "185.6") bindata = "Bin 4";
                                    //if (bin == "185.7") bindata = "Bin 4";
                                    //if (bin == "185.8") bindata = "Bin 4";
                                    //if (bin == "185.9") bindata = "Bin 4";
                                    //if (bin == "186.0") bindata = "Bin 4";
                                    //if (bin == "186") bindata = "Bin 4";
                                    //if (bin == "186.1") bindata = "Bin 4";
                                    //if (bin == "186.2") bindata = "Bin 4";
                                    //if (bin == "186.3") bindata = "Bin 4";
                                    //if (bin == "186.4") bindata = "Bin 4";
                                    //if (bin == "186.5") bindata = "Bin 4";
                                    //if (bin == "186.6") bindata = "Bin 4";
                                    //if (bin == "186.7") bindata = "Bin 4";
                                    //if (bin == "186.8") bindata = "Bin 4";
                                    //if (bin == "186.9") bindata = "Bin 4";
                                    //if (bin == "187.0") bindata = "Bin 4";
                                    //if (bin == "187") bindata = "Bin 4";
                                    //if (bin == "187.1") bindata = "Bin 4";
                                    //if (bin == "187.2") bindata = "Bin 4";
                                    //if (bin == "187.3") bindata = "Bin 4";
                                    //if (bin == "187.4") bindata = "Bin 4";
                                    //if (bin == "187.5") bindata = "Bin 4";
                                    //if (bin == "187.6") bindata = "Bin 4";
                                    //if (bin == "187.7") bindata = "Bin 4";
                                    //if (bin == "187.8") bindata = "Bin 4";
                                    //if (bin == "187.9") bindata = "Bin 4";
                                    //if (bin == "188.0") bindata = "Bin 4";
                                    //if (bin == "188") bindata = "Bin 4";
                                    //if (bin == "188.1") bindata = "Bin 4";
                                    //if (bin == "188.2") bindata = "Bin 4";
                                    //if (bin == "188.3") bindata = "Bin 4";
                                    //if (bin == "188.4") bindata = "Bin 4";
                                    //if (bin == "188.5") bindata = "Bin 4";
                                    //if (bin == "188.6") bindata = "Bin 4";
                                    //if (bin == "188.7") bindata = "Bin 4";
                                    //if (bin == "188.8") bindata = "Bin 4";
                                    //if (bin == "188.9") bindata = "Bin 4";
                                    //if (bin == "189.0") bindata = "Bin 4";
                                    //if (bin == "189") bindata = "Bin 4";
                                    //if (bin == "189.1") bindata = "Bin 4";
                                    //if (bin == "189.2") bindata = "Bin 4";
                                    //if (bin == "189.3") bindata = "Bin 4";
                                    //if (bin == "189.4") bindata = "Bin 4";
                                    //if (bin == "189.5") bindata = "Bin 4";
                                    //if (bin == "189.6") bindata = "Bin 4";
                                    //if (bin == "189.7") bindata = "Bin 4";
                                    //if (bin == "189.8") bindata = "Bin 4";
                                    //if (bin == "189.9") bindata = "Bin 4";
                                    //if (bin == "190.0") bindata = "Bin 4";
                                    //if (bin == "190") bindata = "Bin 4";
                                    //#endregion
                                    #endregion
                                    dr["bin"] = bindata;

                                }

                                //if (dt2.Rows.Count != 0)
                                //{
                                //    dr["sdf0_oqc"] = dt2.Rows[0]["SDF0"].ToString();
                                //    dr["scurave_oqc"] = dt2.Rows[0]["SCURAVE"].ToString();
                                //    dr["sgrmsave_oqc"] = dt2.Rows[0]["SGRMSAVE"].ToString();
                                //    dr["srtpctg2_oqc"] = dt2.Rows[0]["SRTPCTG2"].ToString();
                                //    dr["sbtpctg2_oqc"] = dt2.Rows[0]["SBTPCTG2"].ToString();
                                //}
                                //if (txtCount.Text == "OK")
                                //{
                                //    dtOverall.Rows.Add(dr);
                                //    updateDataGridViews(dtOverall, ref dgvInline);
                                //}
                                dtOverall.Rows.Add(dr);
                                updateDataGridViews(dtOverall, ref dgvInline);
                            }
                            if (dtfct.Rows.Count > 3)
                            {
                                int noi = 1;
                                string countdt = dtfct.Rows.Count.ToString();
                                List<string> show = new List<string>();

                                foreach (DataRow row in dtfct.Rows)
                                {
                                    string value = row[2].ToString();
                                    if (value == "0")
                                        value = "OK";
                                    if (value == "1")
                                        value = "NG";

                                    if (noi <= dtfct.Rows.Count)
                                    {
                                        show.Add("No " + noi + ": " + value + "\n");
                                        noi++;
                                    }
                                }

                                lbENAlarm.Text = "Data FCT Đã Kiểm " + countdt + " Lần \n" + String.Join("", show.ToArray());
                                lbENAlarm.BackColor = Color.Red;
                                txtCount.Text = "NG";
                                txtCount.BackColor = Color.Red;
                                return;
                            }

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
                            if (txtCount.Text == "NG")
                            {
                                dgvInline.Rows.RemoveAt(dgvInline.Rows.Count - 1);
                                txtOkCount.Text = (okCount - 1).ToString() + "/" + limit1.ToString();
                            }

                            //  SearchDataF0();
                        }
                        else
                        {
                            txtCount.Text = "NG";
                            txtCount.BackColor = Color.Red;
                            txtResultDetail.Text = "Model name fail!!!";
                            txtResultDetail.BackColor = Color.Red;
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
                    if (!string.IsNullOrEmpty(process.ToString()))
                    {
                        datastring += string.Format("{0}: {1}\r\n", process.ToString(), judge.ToString());

                    }
                }
                var checkFail = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("judge").Contains("FAILURE"));
                var checkcase = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("CASE"));
                var checken = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("EN"));
                var checkheightcase = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("HEIGHTCASE"));
                var checkenQA = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NMT5"));
                if (!checkcase)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "CASE: NO DATA\r\n";
                }
                if (!checkheightcase)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "HEIGHTCASE: NO DATA\r\n";

                }
                if (!checken)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "EN: NO DATA\r\n";
                }
                if (!checkenQA)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "NMT5: NO DATA\r\n";
                }

                if (checkFail)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    txtResultDetail.Text = datastring;
                }
                if (!checkFail && checken && checkcase && checkheightcase && checkenQA)
                {
                    txtCount.Text = "OK";
                    txtCount.BackColor = Color.SpringGreen;
                    txtResultDetail.BackColor = Color.SpringGreen;
                    txtResultDetail.Text = datastring;
                }
                txtResultDetail.Text = datastring;
            }
        }
        #region SEARCH DATA F0
        private void SearchDataF0()
        {
            string sql2 = "select MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0 " +
                     " FROM" +
                     " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableThisMonth + "data" +
                     " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'EN' and serno = '" + txtProductSerial.Text + "') order by serno) tbl where row_number =1) and inspect in ('SDF0'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableThisMonth + " where serno = '" + txtProductSerial.Text + "' and process = 'EN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                     " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                     " GROUP BY serno, tjudge, inspectdate" +
                     " UNION ALL " +
                     "select MAX(case inspect when 'SDF0' then inspectdata else null end) as SDF0 " +
                     " FROM" +
                     " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableLastMonth + "data" +
                     " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'EN' and serno = '" + txtProductSerial.Text + "') order by serno) tbl where row_number =1) and inspect in ('SDF0'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableLastMonth + " where serno = '" + txtProductSerial.Text + "' and process = 'EN' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                     " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                     " GROUP BY serno, tjudge, inspectdate";
            TfSQL con = new TfSQL();
            con.getCompleteData(sql2, ref lbF0);
            double ValueF0 = double.Parse(lbF0.Text);
            int Bin1 = int.Parse(txtBin1.Text);
            int Bin2 = int.Parse(txtBin2.Text);
            int Bin3 = int.Parse(txtBin3.Text);
            int Bin4 = int.Parse(txtBin4.Text);
            if (170 <= ValueF0 && ValueF0 <= 175.0)
            {
                bin1 = Bin1 + 1;
                lbBin.Text = "BIN 1";
                lbBin.BackColor = Color.Pink;
            }
            else if (175.1 <= ValueF0 && ValueF0 <= 180.0)
            {
                bin2 = Bin2 + 1;
                lbBin.Text = "BIN 2";
                lbBin.BackColor = Color.DimGray;
            }
            else if (180.1 <= ValueF0 && ValueF0 <= 185.0)
            {
                bin3 = Bin3 + 1;
                lbBin.Text = "BIN 3";
                lbBin.BackColor = Color.DodgerBlue;
            }
            else if (185.1 <= ValueF0 && ValueF0 <= 190.0)
            {

                bin4 = Bin4 + 1;
                lbBin.Text = "BIN 4";
                lbBin.BackColor = Color.MediumOrchid;
            }
            txtBin1.Text = bin1.ToString();
            txtBin2.Text = bin2.ToString();
            txtBin3.Text = bin3.ToString();
            txtBin4.Text = bin4.ToString();

        }
        #endregion
        private void decideReferenceTable()
        {
            testerTableThisMonth = cmbModel.Text + DateTime.Today.ToString("yyyyMM");
            tableThisMonth = testerTableThisMonth;
            testerTableLastMonth = cmbModel.Text + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
            tableLastMonth = testerTableLastMonth;
            //tableAssyThisMonth = "la20_523ab" + DateTime.Today.ToString("yyyyMM");
            //tableAssyLastMonth = "la20_523ab" + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
            //    (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
        }

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
                bool res1;
                if (cmbModel.Text == "LD20")
                    res1 = tf.sqlMultipleInsertDeus(dt);
                else
                    res1 = tf.sqlMultipleInsertDeus(dt);

                if (res1)
                {
                    string shipKind = dtOverall.Rows[0]["return"].ToString();
                    string prt_model = cmbModel.Text.Substring(0, 4);
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
        private string checkDataTableWithRealTable(DataTable dt1)
        {
            string serial;
            string result = String.Empty;
            if (formReturnMode) return result;
            if (cmbModel.Text == "LD20")
            {
                string model = "ld20";
                productTable = "product_serial_" + model;
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
                txtCount.Clear();
                txtCount.BackColor = Color.LightGray;
                txtResultDetail.Clear();
                txtResultDetail.BackColor = Color.LightGray;
                lbENAlarm.Text = null;
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
                productTable = "product_serial_" + model[1];

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
                if (cmbModel.Text == "BMD_0015")
                    res2 = tf.sqlMultipleInsertBMD0015(dt);
                if (cmbModel.Text == "BMD_001")
                    res2 = tf.sqlMultipleInsertBMD0016(dt);
                else
                    res2 = tf.sqlMultipleInsertDeus(dt);

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
                    int res = tf.sqlDeleteBoxidld(boxid);

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
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x112;
            const long SC_CLOSE = 0xF060L;
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & 0xFFF0L) == SC_CLOSE) { return; }
            base.WndProc(ref m);
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
            txtCarton.Enabled = true;
            {
                limit1 = 500;
                txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
                txtProductSerial.Enabled = true;
                txtProductSerial.Focus();
            }
        }

        private void txtCarton_TextChanged(object sender, EventArgs e)
        {
            if (lblFrmName.Text != "VIEW")
            {
                if (cmbModel.Text == "LD20")
                {
                    string box = cmbModel.Text;
                    txtBoxId.Text = box + "-" + DateTime.Today.ToString("yyMMdd") + "-" + txtCarton.Text;
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbModelMaster.Enabled = false;
            txtCarton.Enabled = true;

            limit1 = 500;
            txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
            txtProductSerial.Enabled = true;
            txtProductSerial.Focus();

            // txtProductSerial.Focus();
        }

        private void btnEditCode_Click(object sender, EventArgs e)
        {
            cmbModelMaster.Enabled = true;
            cmbModelMaster.Focus();
        }

        private void btnResetCode_Click(object sender, EventArgs e)
        {
            cmbModelMaster.Enabled = false;
            txtProductSerial.Enabled = true;
            txtProductSerial.Focus();
            txtCount.Text = "";
            txtResultDetail.Text = "";
            txtCount.BackColor = Color.WhiteSmoke;
            txtResultDetail.BackColor = Color.WhiteSmoke;
            txtProductSerial.SelectAll();
        }

        private void printDataView(DataView dv)
        {
            foreach (DataRowView drv in dv)
            {
                System.Diagnostics.Debug.Print(drv["lot"].ToString() + " " +
                    drv["tjudge"].ToString() + " " + drv["inspectdate"].ToString());
            }
        }
        private void SearchData()
        {
            TfSQL tf = new TfSQL();

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
    }
}
