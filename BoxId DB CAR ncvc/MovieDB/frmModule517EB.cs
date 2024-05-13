using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Security.Permissions;
using System.Runtime.InteropServices;
using System.Linq;


namespace BoxIdDb
{
    public partial class frmModule517EB : Form
    {
        //eƒtƒH[ƒ€frmBox‚ÖƒCƒxƒ“ƒg”­¶‚ð˜A—iƒfƒŒƒQ[ƒgj
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        // ƒvƒŠƒ“ƒg—pƒeƒLƒXƒgƒtƒ@ƒCƒ‹‚Ì•Û‘¶—pƒtƒHƒ‹ƒ_‚ðAŠî–{Ý’èƒtƒ@ƒCƒ‹‚ÅÝ’è‚·‚é
        string appconfig = @"\\192.168.193.1\barcode$\BoxId Printer vc5\info.ini";
        string directory = @"C:\Users\takusuke.fujii\Desktop\Auto Print\\";

        //‚»‚Ì‘¼A”ñƒ[ƒJƒ‹•Ï”
        bool formEditMode;
        bool formReturnMode;
        bool formAddMode;
        string user;
        // string m_model;
        string testerTableThisMonth;
        string testerTableLastMonth;
        string tableThisMonth;
        string tableLastMonth;
        string testerThisMonth;
        string testerLastMonth;
        string tablethis;
        string tablelast;
        string m_lot;
        int okCount;
        string productTable;
        bool inputBoxModeOriginal;
        string inLineTableThisMonth;
        string inLineTableLastMonth;
        string oqcTableThisMonth;
        string oqcTableLastMonth;
        DataTable dtOverall;
        DataTable dtAllProcess;
        //DataTable dtTester;
        int limit1 = 60;
        public int limit2 = 0;
        bool sound;
        public frmModule517EB()
        {
            InitializeComponent();
        }

        // ƒ[ƒhŽž‚Ìˆ—
        private void frmModule_Load(object sender, EventArgs e)
        {
            //txtCarton.Enabled = false;
            user = txtUser.Text;
            txtLimit.Text = limit2.ToString();
            directory = readIni("TARGET DIRECTORY", "DIR", appconfig);
            this.Left = 250;
            this.Top = 20;
            dtOverall = new DataTable();
            defineAndReadDtOverall(ref dtOverall);
            //dtTester = new DataTable();
            //defineAndReaddtTester(ref dtTester);

            // ‚k‚h‚l‚h‚s‚Ì§Œä‚ðŒã‚Å’¼‚·•K—v‚ ‚è
            if (!formEditMode)
            {
                // ƒf[ƒ^ƒe[ƒuƒ‹‚Ìæ“ªs‚ÌƒVƒŠƒAƒ‹‚©‚çA‚k‚h‚l‚h‚s‚ð”»’f‚·‚é
                if (dtOverall.Rows.Count >= 0)
                {
                    limit1 = 60;
                }
            }

            // ƒOƒŠƒbƒgƒrƒ…[‚ÌXV
            updateDataGridViews(dtOverall, ref dgvInline);

            // ƒVƒŠƒAƒ‹—pƒeƒLƒXƒgƒ{ƒbƒNƒX‚Ì§Œä‚ðŒã‚Å’¼‚·•K—v‚ ‚è
            if (!formEditMode)
            {
                txtProductSerial.Enabled = false;
            }
        }

        // Ý’èƒeƒLƒXƒgƒtƒ@ƒCƒ‹‚Ì“Ç‚Ýž‚Ý
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
        // Windows API ‚ðƒCƒ“ƒ|[ƒg
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filepath);

        // ƒTƒuƒvƒƒV[ƒWƒƒFeƒtƒH[ƒ€‚ÅŒÄ‚Ño‚µAeƒtƒH[ƒ€‚Ìî•ñ‚ðAƒeƒLƒXƒgƒ{ƒbƒNƒX‚ÖŠi”[‚µ‚Äˆø‚«Œp‚®
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
                productTable = "product_serial_" + model;
                switch (model)
                {
                    case "517EB":
                        cmbModel.Text = "LA20_517EB";
                        limit1 = 60;
                        break;
                }
                //limit1 = 60;
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
                string[] model = cmbModel.Text.Split('_');
                productTable = "product_serial_" + model[1];
            }
            //else
            //{
            //    MessageBox.Show("Please choose model!", "Warring", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    cmbModel.Focus();
            //}
        }
        // ƒTƒuƒvƒƒV[ƒWƒƒF‚c‚a‚©‚ç‚Ì‚c‚s‚n‚u‚d‚q‚`‚k‚k‚Ö‚Ì“Ç‚Ýž‚Ý
        private void defineAndReadDtOverall(ref DataTable dt)
        {
            string boxId = txtBoxId.Text;
            setProductTable();
            dt.Columns.Add("serialno", Type.GetType("System.String"));
            dt.Columns.Add("model", Type.GetType("System.String"));
            dt.Columns.Add("lot", Type.GetType("System.String"));
            dt.Columns.Add("inspectdate", Type.GetType("System.DateTime")); //date test NMT
            dt.Columns.Add("cir_ccw", Type.GetType("System.String"));
            dt.Columns.Add("cg_ccw", Type.GetType("System.String"));
            dt.Columns.Add("cnr_ccw", Type.GetType("System.String"));
            dt.Columns.Add("tjudge", Type.GetType("System.String"));
            dt.Columns.Add("date_line", Type.GetType("System.DateTime")); //date test NO41
            dt.Columns.Add("aio_ccw", Type.GetType("System.String"));
            dt.Columns.Add("ano_ccw", Type.GetType("System.String"));
            dt.Columns.Add("air_ccw", Type.GetType("System.String"));
            dt.Columns.Add("anr_ccw", Type.GetType("System.String"));
            dt.Columns.Add("ais_ccw", Type.GetType("System.String"));
            dt.Columns.Add("tjudge_line", Type.GetType("System.String"));
            dt.Columns.Add("return", Type.GetType("System.String"));

            if (!formEditMode)
            {
                //string sql;
                //sql = "select serialno, model, lot, cir_ccw, cg_ccw, cnr_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, judge, return " + "FROM product_serial_517eb WHERE boxid='" + boxId + "'";
                string sql;
                sql = "select serialno, model, lot, inspectdate, cir_ccw, cg_ccw, cnr_ccw, tjudge, date_line, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return " +
                    "FROM " + productTable + " WHERE boxid='" + boxId + "'";
                //TfSQL tf = new TfSQL();
                //System.Diagnostics.Debug.Print(sql);
                //tf.sqlDataAdapterFillDatatable(sql, ref dt);
                TfSQL tf = new TfSQL();
                System.Diagnostics.Debug.Print(sql);
                tf.sqlDataAdapterFillDatatable(sql, ref dt);
            }
        }

        // ƒTƒuƒvƒƒV[ƒWƒƒFƒf[ƒ^ƒOƒŠƒbƒgƒrƒ…[‚ÌXV
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
            }
            else
            {
                btnRegisterBoxId.Enabled = false;
            }
        }

        // ƒTƒuƒvƒƒV[ƒWƒƒFƒVƒŠƒAƒ‹”Ô†d•¡‚È‚µ‚Ì‚o‚`‚r‚rŒÂ”‚ðŽæ“¾‚·‚é
        private int getOkCount(DataTable dt)
        {
            if (dt.Rows.Count <= 0) return 0;
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge", "tjudge_line" });
            DataRow[] dr = distinct.Select("tjudge = 'PASS' and tjudge_line = 'PASS'");
            int dist = dr.Length;
            return dist;
        }

        // ƒTƒuƒvƒƒV[ƒWƒƒFƒƒCƒ“ƒf[ƒ^ƒOƒŠƒbƒgƒrƒ…[‚Öƒf[ƒ^ƒe[ƒuƒ‹‚ðŠi”[A‚¨‚æ‚ÑWŒvƒOƒŠƒbƒhƒrƒ…[‚Ìì¬
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

        // ƒTƒuƒTƒuƒvƒƒV[ƒWƒƒFWŒv—p‚Ìƒf[ƒ^ƒe[ƒuƒ‹‚ðAƒf[ƒ^ƒOƒŠƒbƒhƒrƒ…[‚ÉŠi”[
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

                // Œ”‚Ì‘½‚¢ƒRƒ“ƒtƒBƒO‚ðA‚±‚Ì” ‚ÌƒƒCƒ“ƒ‚ƒfƒ‹‚Æ‚·‚é
                m_lot = a > b ? A : B;

                // Œ”‚Ì­‚È‚¢‚Ù‚¤‚ÌƒƒCƒ“ƒ‚ƒfƒ‹•¶Žš‚ðŽæ“¾‚µAƒZƒ‹”Ô’n‚ð“Á’è‚µ‚Äƒ}[ƒN‚·‚é
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

        // ƒTƒuƒvƒƒV[ƒWƒƒFƒeƒXƒgŒ‹‰Ê‚ª‚e‚`‚h‚k‚Ü‚½‚ÍƒŒƒR[ƒh‚È‚µ‚ÌƒVƒŠƒAƒ‹‚ðƒ}[ƒLƒ“ƒO‚·‚é
        private void colorViewForFailAndBlank(ref DataGridView dgv)
        {
            int row = dgv.Rows.Count;
            for (int i = 0; i < row; ++i)
            {
                if (dgv["col_judge_oqc", i].Value.ToString() == "FAIL" || dgv["col_judge_oqc", i].Value.ToString() == "PLS NG" || String.IsNullOrEmpty(dgv["col_judge_oqc", i].Value.ToString())|| String.IsNullOrEmpty(dgv["col_cg_ccw", i].Value.ToString()) || String.IsNullOrEmpty(dgv["col_cir_ccw", i].Value.ToString()) || String.IsNullOrEmpty(dgv["col_cnr_ccw", i].Value.ToString()))
                {
                    dgv["col_date", i].Style.BackColor = Color.Red;
                    dgv["col_cg_ccw", i].Style.BackColor = Color.Red;
                    dgv["col_cir_ccw", i].Style.BackColor = Color.Red;
                    dgv["col_cnr_ccw", i].Style.BackColor = Color.Red;
                    dgv["col_judge_oqc", i].Style.BackColor = Color.Red;

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
                    dgv["col_aio_ccw", i].Style.BackColor = Color.Red;
                    dgv["col_air_ccw", i].Style.BackColor = Color.Red;
                    dgv["col_ais_ccw", i].Style.BackColor = Color.Red;
                    dgv["col_ano_ccw", i].Style.BackColor = Color.Red;
                    dgv["col_anr_ccw", i].Style.BackColor = Color.Red;
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

        // ƒTƒuƒvƒƒV[ƒWƒƒFd•¡ƒŒƒR[ƒhA‚Ü‚½‚Í‚PƒZƒ‹‚Qd“ü—Í‚ðƒ}[ƒLƒ“ƒO‚·‚é
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
                    if (cmbModel.Text == "")
                    {
                        MessageBox.Show("Please select model name", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cmbModel.Focus();
                        return;
                    }
                    txtCount.Clear();
                    txtResultDetail.Clear();
                    txtProductSerial.Enabled = false;
                    string serial = txtProductSerial.Text;
                    decideReferenceTable();

                    if (serial != String.Empty)
                    {
                        string model = cmbModel.Text;
                        #region Data OQC
                        string sql2 = "select serno, tjudge, inspectdate, " +
                        "MAX(case inspect when 'CG_CCW' then inspectdata else null end) as CG_CCW," +
                        "MAX(case inspect when 'CIR_CCW' then inspectdata else null end) as CIR_CCW," +
                        "MAX(case inspect when 'CNR_CCW' then inspectdata else null end) as CNR_CCW" +
                        " FROM" +
                        " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + oqcTableThisMonth + "data" +
                        " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + oqcTableThisMonth + " where (process = 'NMT1' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('CG_CCW','CIR_CCW','CNR_CCW'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + oqcTableThisMonth + " where serno = '" + serial + "' and process = 'NMT1' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                        " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                        " GROUP BY serno, tjudge, inspectdate" +

                        " UNION ALL " +

                        "select serno, tjudge, inspectdate, " +
                        "MAX(case inspect when 'CG_CCW' then inspectdata else null end) as CG_CCW," +
                        "MAX(case inspect when 'CIR_CCW' then inspectdata else null end) as CIR_CCW," +
                        "MAX(case inspect when 'CNR_CCW' then inspectdata else null end) as CNR_CCW" +
                        " FROM" +
                        " (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + oqcTableLastMonth + "data" +
                        " WHERE serno = (SELECT serno from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + oqcTableLastMonth + " where (process = 'NMT1' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('CG_CCW','CIR_CCW','CNR_CCW'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + oqcTableLastMonth + " where serno = '" + serial + "' and process = 'NMT1' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
                        " group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                        " GROUP BY serno, tjudge, inspectdate order by inspectdate desc";

                        System.Diagnostics.Debug.Print(System.Environment.NewLine + sql2);
                        DataTable dt2 = new DataTable();
                        TfSQL tf = new TfSQL();
                        tf.sqlDataAdapterFillDatatableOqc(sql2, ref dt2);
                        #endregion
                        //Data INLINE
                        #region Data INLINE
                        string sql1 = "select serno, " +
                        "MAX(case inspect when 'AIO_CCW' then inspectdata else null end) as AIO_CCW," +
                        "MAX(case inspect when 'ANO_CCW' then inspectdata else null end) as ANO_CCW," +
                        "MAX(case inspect when 'AIR_CCW' then inspectdata else null end) as AIR_CCW," +
                        "MAX(case inspect when 'ANR_CCW' then inspectdata else null end) as ANR_CCW," +
                        "MAX(case inspect when 'AIS_CCW' then inspectdata else null end) as AIS_CCW," +
                        " inspectdate, judge FROM" +
                        " (select d.serno, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + tableThisMonth +
                        " WHERE serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'NO53' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('AIO_CCW','AIR_CCW','AIS_CCW','ANO_CCW','ANR_CCW') AND JUDGE = '0')" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno from " + testerTableThisMonth + " where serno = '" + serial + "')d" +
                        " group by d.serno, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                        " GROUP BY serno,inspectdate, judge" +

                        " UNION ALL" +

                        " select serno, " +
                        "MAX(case inspect when 'AIO_CCW' then inspectdata else null end) as AIO_CCW," +
                        "MAX(case inspect when 'ANO_CCW' then inspectdata else null end) as ANO_CCW," +
                        "MAX(case inspect when 'AIR_CCW' then inspectdata else null end) as AIR_CCW," +
                        "MAX(case inspect when 'ANR_CCW' then inspectdata else null end) as ANR_CCW," +
                        "MAX(case inspect when 'AIS_CCW' then inspectdata else null end) as AIS_CCW," +
                        " inspectdate, judge FROM" +
                        " (select d.serno, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + tableLastMonth +
                        " WHERE serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableThisMonth + " where (process = 'NO53' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('AIO_CCW','AIR_CCW','AIS_CCW','ANO_CCW','ANR_CCW') AND JUDGE = '0')" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno from " + testerTableThisMonth + " where serno = '" + serial + "')d" +
                        " group by d.serno, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                        " GROUP BY serno,inspectdate, judge" +

                        " UNION ALL" +

                        " select serno, " +
                        "MAX(case inspect when 'AIO_CCW' then inspectdata else null end) as AIO_CCW," +
                        "MAX(case inspect when 'ANO_CCW' then inspectdata else null end) as ANO_CCW," +
                        "MAX(case inspect when 'AIR_CCW' then inspectdata else null end) as AIR_CCW," +
                        "MAX(case inspect when 'ANR_CCW' then inspectdata else null end) as ANR_CCW," +
                        "MAX(case inspect when 'AIS_CCW' then inspectdata else null end) as AIS_CCW," +
                        " inspectdate, judge FROM" +
                        " (select d.serno, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + tableLastMonth +
                        " WHERE serno = (SELECT lot from(select lot, serno,process, inspectdate, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from " + testerTableLastMonth + " where (process = 'NO53' and serno = '" + serial + "') order by serno) tbl where row_number =1) and inspect in ('AIO_CCW','AIR_CCW','AIS_CCW','ANO_CCW','ANR_CCW') AND JUDGE = '0')" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno from " + testerTableLastMonth + " where serno = '" + serial + "')d" +
                        " group by d.serno, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
                        " GROUP BY serno,inspectdate, judge";
                        System.Diagnostics.Debug.Print(System.Environment.NewLine + sql1);
                        DataTable dt1 = new DataTable();

                        tf.sqlDataAdapterFillDatatablePqm(sql1, ref dt1);

                        #endregion


                        #region -- Get All Process Judge --


                        string lotthis = string.Format("SELECT lot FROM (SELECT lot, inspectdate, ROW_NUMBER() OVER(PARTITION BY lot ORDER BY inspectdate DESC) from {1} where serno = '{0}' and process = 'NO53' ORDER BY lot)tb where ROW_NUMBER = 1", serial, testerTableThisMonth);
                        string lotlast = string.Format("SELECT lot FROM (SELECT lot, inspectdate, ROW_NUMBER() OVER(PARTITION BY lot ORDER BY inspectdate DESC) from {1} where serno = '{0}' and process = 'NO53' ORDER BY lot)tb where ROW_NUMBER = 1", serial, testerTableLastMonth);
                        string queryProcess = string.Format("SELECT serno, lot, inspectdate, process,judge from "
                         + "(SELECT serno, lot, inspectdate, process, judge, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from "
                         + "(SELECT ({3}) as serno, serno lot, inspectdate, process,"
                         + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {4} "
                         + "WHERE serno in (SELECT DISTINCT lot FROM {4} WHERE process = 'NO53' AND serno = ({3})) "
                         + "OR serno = ({3})"
                         + "UNION ALL SELECT ({3}) as serno, serno lot, inspectdate, process,"
                         + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {5} "
                         + "WHERE serno in (SELECT DISTINCT lot FROM {5} WHERE process = 'NO53' AND serno = ({3})) "
                         + "OR serno = ({3})"
                         + "UNION ALL SELECT ({6}) as serno, serno lot, inspectdate, process,"
                         + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {5} "
                         + "WHERE serno in (SELECT DISTINCT lot FROM {5} WHERE process = 'NO53' AND serno = ({6})) "
                         + "OR serno = ({6})"
                         + "UNION ALL SELECT '{0}' as serno, serno lot, inspectdate, process, "
                         + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {1} "
                         + "WHERE serno in (SELECT DISTINCT lot FROM {1} WHERE process = 'NO53' AND serno = '{0}') "
                         + "OR serno = '{0}' "
                         + "UNION ALL SELECT '{0}' as serno, serno lot, inspectdate, process, "
                         + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {2} "
                         + "WHERE serno in (SELECT DISTINCT lot FROM {2} WHERE process = 'NO53' AND serno = '{0}') "
                         + "OR serno = '{0}' ORDER BY process) tbl) tb where ROW_NUMBER = 1", serial, testerTableThisMonth, testerTableLastMonth, lotthis, tablethis, tablelast, lotlast);
                        //string lotthis = string.Format("SELECT lot FROM (SELECT lot, inspectdate, ROW_NUMBER() OVER(PARTITION BY lot ORDER BY inspectdate DESC) from {1} where serno = '{0}' and process = 'NO53' ORDER BY lot)tb where ROW_NUMBER = 1", serial, testerTableThisMonth);
                        //string lotlast = string.Format("SELECT lot FROM (SELECT lot, inspectdate, ROW_NUMBER() OVER(PARTITION BY lot ORDER BY inspectdate DESC) from {1} where serno = '{0}' and process = 'NO53' ORDER BY lot)tb where ROW_NUMBER = 1", serial, testerTableLastMonth);
                        //string queryProcess = string.Format("SELECT serno, lot, inspectdate, process,judge from "
                        // + "(SELECT serno, lot, inspectdate, process, judge, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from "
                        // + "(SELECT ({3}) as serno, serno lot,inspectdate, process,"
                        // + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {4} "
                        // + "WHERE serno in (SELECT DISTINCT lot FROM {4} WHERE process = 'NO53' AND serno = ({3})) "
                        // + "OR serno = ({3})"
                        // + "UNION ALL SELECT ({6}) as serno, serno lot, inspectdate, process,"
                        // + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {5} "
                        // + "WHERE serno in (SELECT DISTINCT lot FROM {5} WHERE process = 'NO53' AND serno = ({6})) "
                        // + "OR serno = ({6})"
                        // + "UNION ALL SELECT '{0}' as serno, serno lot, inspectdate, process, "
                        // + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {1} "
                        // + "WHERE serno in (SELECT DISTINCT lot FROM {1} WHERE process = 'NO53' AND serno = '{0}') "
                        // + "OR serno = '{0}' "
                        // + "UNION ALL SELECT '{0}' as serno, serno lot, inspectdate, process, "
                        // + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {2} "
                        // + "WHERE serno in (SELECT DISTINCT lot FROM {2} WHERE process = 'NO53' AND serno = '{0}') "
                        // + "OR serno = '{0}' ORDER BY process) tbl) tb where ROW_NUMBER = 1", serial, testerTableThisMonth, testerTableLastMonth, lotthis, tablethis, tablelast, lotlast);
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
                        dr["model"] = model.Substring(0, 4) + "V" + model.Substring(4);

                        switch (model)
                        {
                            case "LA20_517CC":
                            case "LA20_517CC1":
                            case "LA20_517CC3":
                            case "LA20_517DB":
                                dr["serialno"] = serial;
                                dr["lot"] = VBS.Mid(serial, 13, 3).Length < 3 ? "Error" : VBS.Mid(serial, 13, 3);
                                break;
                            case "LA20_517CC2":
                            case "LA20_517DC":
                                dr["serialno"] = serial;
                                dr["lot"] = VBS.Mid(serial, 8, 3).Length < 3 ? "Error" : VBS.Mid(serial, 8, 3);
                                break;
                            case "LA20_517CB":
                                dr["serialno"] = serial;
                                dr["lot"] = VBS.Mid(serial, 9, 3).Length < 3 ? "Error" : VBS.Mid(serial, 9, 3);
                                break;
                            case "LA20_517EB":
                                dr["serialno"] = serial;
                                dr["lot"] = VBS.Mid(serial, 12, 6).Length < 3 ? "Error" : VBS.Mid(serial, 12, 6);
                                break;
                            default:
                                dr["serialno"] = serial;
                                dr["lot"] = VBS.Mid(serial, 9, 3).Length < 3 ? "Error" : VBS.Mid(serial, 9, 3);
                                break;
                        }

                        dr["return"] = formReturnMode ? "R" : "N";

                        if (dt2.Rows.Count != 0)
                        {
                            //T-judge OQC
                            string linepass = String.Empty;
                            string buff = dt2.Rows[0]["tjudge"].ToString();
                            if (buff == "0") linepass = "PASS";
                            else if (buff == "1") linepass = "FAIL";
                            else linepass = "ERROR";

                            dr["tjudge"] = linepass;
                            dr["inspectdate"] = dt2.Rows[0]["inspectdate"].ToString();
                        }
                        // else dr["tjudge"] = "FAIL";
                        if (dt1.Rows.Count != 0)
                        {
                            dr["aio_ccw"] = dt1.Rows[0]["aio_ccw"].ToString();
                            dr["ano_ccw"] = dt1.Rows[0]["ano_ccw"].ToString();
                            dr["air_ccw"] = dt1.Rows[0]["air_ccw"].ToString();
                            dr["anr_ccw"] = dt1.Rows[0]["anr_ccw"].ToString();
                            dr["ais_ccw"] = dt1.Rows[0]["ais_ccw"].ToString();
                            dr["date_line"] = dt1.Rows[0]["inspectdate"].ToString();
                            string buff = dt1.Rows[0]["judge"].ToString();
                            if (buff == "0") dr["tjudge_line"] = "PASS";
                            else if (buff == "1") dr["tjudge_line"] = "FAIL";
                        }

                        if (dt2.Rows.Count != 0)
                        {
                            dr["cg_ccw"] = dt2.Rows[0]["cg_ccw"].ToString();
                            dr["cir_ccw"] = dt2.Rows[0]["cir_ccw"].ToString();
                            dr["cnr_ccw"] = dt2.Rows[0]["cnr_ccw"].ToString();
                        }
                        if (txtCount.Text == "OK")
                        {
                            dtOverall.Rows.Add(dr);

                            // ƒf[ƒ^ƒOƒŠƒbƒgƒrƒ…[‚ÌXV
                            updateDataGridViews(dtOverall, ref dgvInline);
                        }
                    }
                    // “ü—Í—pƒeƒLƒXƒgƒ{ƒbƒNƒX‚ð•ÒW‰Â”\‚Ö–ß‚µA˜A‘±‚µ‚ÄƒXƒLƒƒƒ“‚Å‚«‚é‚æ‚¤AƒeƒLƒXƒg‚ð‘I‘ðó‘Ô‚É‚·‚é
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
                    //if (txtCount.Text == "NG")
                    //{
                    //    dgvInline.Rows.RemoveAt(dgvInline.Rows.Count - 1);
                    //    txtOkCount.Text = (okCount - 1).ToString() + "/" + limit1.ToString();

                    //}
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
                var datarows = dtAllProcess.Rows;
                for (int i = 0; i < datarows.Count; i++)
                {
                    var process = datarows[i]["process"] ?? string.Empty;
                    var judge = datarows[i]["judge"] ?? string.Empty;
                    if (!string.IsNullOrEmpty(process.ToString()))
                    {
                        datastring += string.Format("{0}: {1}\r\n", process.ToString(), judge.ToString());
                    }
                }
                txtResultDetail.Text = datastring;
                var checkFail = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("judge").Contains("FAILURE"));
                var checknmt1 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NMT1"));
                var checkno41 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO41"));
                var checkno43 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO43"));
                var checkno44 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO44"));
                var checkno47 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO47"));
                var checkno48 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO48"));
                var checkno53 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO53"));
                if (!checknmt1)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "NMT1: NO DATA\r\n";
                }

                if (!checkno41)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "NO41: NO DATA\r\n";
                }
                if (!checkno43)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "NO43: NO DATA\r\n";
                }
                if (!checkno44)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "NO44: NO DATA\r\n";

                }
                if (!checkno47)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "NO47: NO DATA\r\n";
                }
                if (!checkno48)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "NO48: NO DATA\r\n";
                }
                if (!checkno53)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    datastring += "NO53: NO DATA\r\n";
                }
                if (checkFail)
                {
                    txtResultDetail.BackColor = Color.Red;
                    txtCount.Text = "NG";
                    txtCount.BackColor = Color.Red;
                    txtResultDetail.Text = datastring;
                }
                if (!checkFail && checknmt1 && checkno41 && checkno43 && checkno44 && checkno47 && checkno48 && checkno53)
                {
                    txtCount.Text = "OK";
                    txtCount.BackColor = Color.SpringGreen;
                    txtResultDetail.BackColor = Color.SpringGreen;
                    txtResultDetail.Text = datastring;
                }
                txtResultDetail.Text = datastring;
            }
        }

        // ƒTƒuƒvƒƒV[ƒWƒƒFƒVƒŠƒAƒ‹‚©‚çAŽQÆ‚·‚×‚«‚o‚p‚lƒe[ƒuƒ‹–¼‚ð“Á’è‚·‚é
        private void decideReferenceTable()
        {
            string m = cmbModel.Text;
            //string m1 = cmbModel.Text;
            //m1 = m.Remove(m.Length - 1);
            oqcTableThisMonth = m + DateTime.Today.ToString("yyyyMM");
            oqcTableLastMonth = m + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
            inLineTableThisMonth = m + DateTime.Today.ToString("yyyyMM");
            inLineTableLastMonth = m + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
            testerTableThisMonth = cmbModel.Text + DateTime.Today.ToString("yyyyMM");
            tableThisMonth = testerTableThisMonth;
            testerTableLastMonth = cmbModel.Text + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
            tableLastMonth = testerTableLastMonth;
            string model = VBS.Mid(cmbModel.Text, 6, 4);
            string model_c = VBS.Left(cmbModel.Text, 9);

            switch (model)
            {
                case "517C":
                    testerTableThisMonth = cmbModel.Text + DateTime.Today.ToString("yyyyMM");
                    tableThisMonth = model_c + DateTime.Today.ToString("yyyyMM") + "data";
                    testerTableLastMonth = cmbModel.Text + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
                    tableLastMonth = model_c + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12") + "data";
                    break;
                case "517D":
                    testerTableThisMonth = cmbModel.Text + DateTime.Today.ToString("yyyyMM");
                    tableThisMonth = model_c + DateTime.Today.ToString("yyyyMM") + "data";
                    testerTableLastMonth = cmbModel.Text + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
                    tableLastMonth = model_c + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12") + "data";
                    break;
                default:
                    testerTableThisMonth = cmbModel.Text + DateTime.Today.ToString("yyyyMM");
                    testerThisMonth = cmbModel.Text + DateTime.Today.ToString("yyyyMM") + "data";
                    tableThisMonth = model_c + DateTime.Today.ToString("yyyyMM") + "data";
                    tablethis = model_c + DateTime.Today.ToString("yyyyMM");
                    tablelast = model_c + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
                    testerTableLastMonth = cmbModel.Text + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
                    testerLastMonth = cmbModel.Text + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12") + "data";
                    tableLastMonth = model_c + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12") + "data";
                    break;

            }
        }

        // ƒrƒ…[ƒ‚[ƒh‚ÅÄˆóü‚ðs‚¤
        private void btnPrint_Click(object sender, EventArgs e)
        {
            string boxId = txtBoxId.Text;
            string model = cmbModel.Text.Substring(0, 4) + "V" + cmbModel.Text.Substring(4);
            string shipKind = dtOverall.Rows[0]["return"].ToString();
            printBarcode(directory, boxId, model, dgvDateCode, ref dgvDateCode2, ref txtBoxIdPrint, shipKind);
        }

        // ŠeŽíŠm”FŒãAƒ{ƒbƒNƒX‚h‚c‚Ì”­sAƒVƒŠƒAƒ‹‚Ì“o˜^Aƒo[ƒR[ƒhƒ‰ƒxƒ‹‚ÌƒvƒŠƒ“ƒg‚ðs‚¤
        private void btnRegisterBoxId_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtCarton.Text))
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

                    //æ‚¸‚ÍADataTble‚Éƒ{ƒbƒNƒX‚h‚c‚ð“o˜^
                    DataTable dt = dtOverall.Copy();
                    dt.Columns.Add("boxid", Type.GetType("System.String"));
                    dt.Columns.Add("carton", Type.GetType("System.String"));
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        dt.Rows[i]["boxid"] = boxId;
                        dt.Rows[i]["carton"] = txtCarton.Text;
                    }

                    //DataTable‚©‚ç–{”Ôƒe[ƒuƒ‹‚ÖˆêŠ‡“o˜^
                    TfSQL tf = new TfSQL();
                    bool res1 = tf.sqlMultipleInsert517EB(dt);

                    if (res1)
                    {

                        string shipKind = dtOverall.Rows[0]["return"].ToString();
                        string prt_model = cmbModel.Text.Substring(0, 4) + "V" + cmbModel.Text.Substring(4);
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
                        //ˆê’U“o˜^‚µ‚½‚a‚n‚w‚h‚c‚ðÁ‹Ž‚·‚é
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
        }

        private string checkDataTableWithRealTable(DataTable dt1)
        {
            string serial;
            string result = String.Empty;
            if (formReturnMode) return result;
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

        // ƒTƒuƒvƒƒV[ƒWƒƒFƒo[ƒR[ƒh‚ðƒvƒŠƒ“ƒg‚·‚éi–{ƒo[ƒWƒ‡ƒ“‚ÍA‚a‚n‚w‚h‚c–¼‚ÌƒeƒLƒXƒgƒtƒ@ƒCƒ‹‚ð¶¬‚·‚éj
        private void printBarcode(string dir, string id, string m_model_long, DataGridView dgv1, ref DataGridView dgv2, ref TextBox txt, string shipKind)
        {
            TfPrint tf = new TfPrint();
            tf.createBoxidFiles(dir, id, m_model_long, dgv1, ref dgv2, ref txt, shipKind);
        }

        // ˆêŽžƒe[ƒuƒ‹‚Ì‘I‘ð‚³‚ê‚½•¡”ƒŒƒR[ƒh‚ðAˆêŠ‡Á‹Ž‚³‚¹‚é
        private void btnDeleteSelection_Click(object sender, EventArgs e)
        {
            DataGridView dgv = new DataGridView();

            if (tabControl1.SelectedTab == tabControl1.TabPages["tabInline"])
                dgv = dgvInline;

            // ƒZƒ‹‚Ì‘I‘ð”ÍˆÍ‚ª‚Q—ñˆÈã‚Ìê‡‚ÍAƒƒbƒZ[ƒW‚Ì•\Ž¦‚Ì‚Ý‚ÅƒvƒƒV[ƒWƒƒ‚ð”²‚¯‚é
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
            }

        }

        // ‚Pƒ‰ƒxƒ‹‚ ‚½‚è‚ÌƒVƒŠƒAƒ‹”‚ð•ÏX‚·‚éiŠÇ—Œ ŒÀƒ†[ƒU[‚Ì‚Ýj
        private void btnChangeLimit_Click(object sender, EventArgs e)
        {
            // ƒtƒH[ƒ€‚Si‚Pƒ‰ƒxƒ‹‚ ‚½‚èƒVƒŠƒAƒ‹”•ÏXj‚ðAƒfƒŒƒQ[ƒgƒCƒxƒ“ƒg‚ð•t‰Á‚µ‚ÄŠJ‚­
            bool bl = TfGeneral.checkOpenFormExists("frmCapacity");
            if (bl)
            {
                MessageBox.Show("Please close or complete another form.", "Warning",
                MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            }
            else
            {
                frmCapacity f4 = new frmCapacity();
                //ŽqƒCƒxƒ“ƒg‚ðƒLƒƒƒbƒ`‚µ‚ÄAƒf[ƒ^ƒOƒŠƒbƒh‚ðXV‚·‚é
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

        // “o˜^Ï‚Ý‚Ìƒ{ƒbƒNƒX‚h‚c‚ÖAƒ‚ƒWƒ…[ƒ‹‚ð’Ç‰ÁiŠÇ—ƒ†[ƒU[‚Ì‚Ýj
        private void btnAddSerial_Click(object sender, EventArgs e)
        {
            // ’Ç‰Áƒ‚[ƒh‚Å‚È‚¢ê‡‚ÍA’Ç‰Áƒ‚[ƒh‚Ì•\Ž¦‚ÖØ‚è‘Ö‚¦‚é
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
            // Šù‚É’Ç‰Áƒ‚[ƒh‚Ìê‡‚ÍA‚c‚a‚Ö‚Ì“o˜^‚ðs‚¤
            else
            {
                // ‚c‚d‚k‚d‚s‚d ‚r‚p‚k•¶‚ð”­s‚µAƒf[ƒ^ƒx[ƒX‚©‚çíœ‚·‚é
                string boxId = txtBoxId.Text;
                string sql = "delete from product_serial_517eb where boxid = '" + boxId + "'";
                System.Diagnostics.Debug.Print(sql);
                TfSQL tf = new TfSQL();
                bool res1 = tf.sqlExecuteNonQuery(sql, false);

                // DataTble‚Éƒ{ƒbƒNƒX‚h‚c‚ð’Ç‰Á‚µA–{”Ôƒe[ƒuƒ‹‚ÖˆêŠ‡“o˜^
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

                bool res2 = tf.sqlMultipleInsertOverall(dt);

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

                // ’Ç‰Áƒ‚[ƒh‚ðI—¹‚µA‰{——ƒ‚[ƒh‚Ì•\Ž¦‚ÖØ‚è‘Ö‚¦‚é
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

        // “o˜^Ï‚Ý‚Ìƒ{ƒbƒNƒX‚h‚c‚ÌAƒ‚ƒWƒ…[ƒ‹‚ðíœiŠÇ—ƒ†[ƒU[‚Ì‚Ýj
        private void btnDeleteSerial_Click(object sender, EventArgs e)
        {
            // ƒZƒ‹‚Ì‘I‘ð”ÍˆÍ‚ª‚Q—ñˆÈã‚Ìê‡‚ÍAƒƒbƒZ[ƒW‚Ì•\Ž¦‚Ì‚Ý‚ÅƒvƒƒV[ƒWƒƒ‚ð”²‚¯‚é
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
                // ‚c‚d‚k‚d‚s‚d ‚r‚p‚k•¶‚ð”­s‚µAƒf[ƒ^ƒx[ƒX‚©‚çíœ‚·‚é
                string boxId = txtBoxId.Text;
                string whereSer = string.Empty;
                foreach (DataGridViewCell cell in dgvInline.SelectedCells)
                {
                    whereSer += "'" + cell.Value.ToString() + "', ";
                }
                string sql = "delete from product_serial_517eb where boxid = '" + boxId + "' and  serialno in (" + VBS.Left(whereSer, whereSer.Length - 2) + ")";
                System.Diagnostics.Debug.Print(sql);
                TfSQL tf = new TfSQL();
                int res = tf.sqlExecuteNonQueryInt(sql, false);

                if (res >= 1)
                {
                    // ƒf[ƒ^ƒOƒŠƒbƒhƒrƒ…[‚©‚çíœ‚·‚é
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

        // “o˜^Ï‚Ý‚Ìƒ{ƒbƒNƒX‚h‚c‚¨‚æ‚ÑŠY“–ƒ‚ƒWƒ…[ƒ‹‚ÌíœiŠÇ—ƒ†[ƒU[‚Ì‚Ýj
        private void btnCancelBoxid_Click(object sender, EventArgs e)
        {
            // –{“–‚Éíœ‚µ‚Ä‚æ‚¢‚©A‚Qd‚ÅŠm”F‚·‚éB
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
                    int res = tf.sqlDeleteBoxid517EB(boxid);

                    dtOverall.Clear();
                    // ƒf[ƒ^ƒOƒŠƒbƒgƒrƒ…[‚ÌXV
                    updateDataGridViews(dtOverall, ref dgvInline);

                    //eƒtƒH[ƒ€frmBox‚Ìƒf[ƒ^ƒOƒŠƒbƒgƒrƒ…[‚ðXV‚·‚é‚½‚ßAƒfƒŒƒQ[ƒgƒCƒxƒ“ƒg‚ð”­¶‚³‚¹‚é
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

        // ƒLƒƒƒ“ƒZƒ‹Žž‚ÉAƒf[ƒ^ƒe[ƒuƒ‹‚ÌƒŒƒR[ƒh‚Ì•ÛŽ‚ª‚Å‚«‚È‚¢Ž|AŒx‚·‚é
        private void btnCancel_Click(object sender, EventArgs e)
        {
            // frmCapacity i‚a‚n‚w‚ ‚½‚èƒVƒŠƒAƒ‹ŒÂ”j‚ð•Â‚¶‚Ä‚¢‚È‚¢ê‡‚ÍAæ‚É•Â‚¶‚é‚æ‚¤’Ê’m‚·‚é
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

            // ƒf[ƒ^ƒe[ƒuƒ‹‚ÌƒŒƒR[ƒhŒ”‚ª‚È‚¢ê‡A‚Ü‚½‚Í•ÒWƒ‚[ƒh‚Ìê‡‚ÍA‚»‚Ì‚Ü‚Ü•Â‚¶‚é                        
            if (dtOverall.Rows.Count == 0 || !formEditMode)
            {
                Application.OpenForms["frmBox"].Focus();
                Close();
                return;
            }

            // ƒf[ƒ^ƒe[ƒuƒ‹‚ÌƒŒƒR[ƒhŒ”‚ª‚ ‚éê‡AˆêŽž“I‚É•ÛŽ‚³‚ê‚Ä‚¢‚éƒŒƒR[ƒh‚ªÁ–Å‚·‚éŽ|AŒx‚·‚é
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

        // •Â‚¶‚éƒ{ƒ^ƒ“‚âƒVƒ‡[ƒgƒJƒbƒg‚Å‚ÌI—¹‚ð‹–‚³‚È‚¢
        [SecurityPermission(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref Message m)
        {
            const int WM_SYSCOMMAND = 0x112;
            const long SC_CLOSE = 0xF060L;
            if (m.Msg == WM_SYSCOMMAND && (m.WParam.ToInt64() & 0xFFF0L) == SC_CLOSE) { return; }
            base.WndProc(ref m);
        }

        //MP3ƒtƒ@ƒCƒ‹i¡‰ñ‚ÍŒx‰¹j‚ðÄ¶‚·‚é
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

        // ƒf[ƒ^‚ðƒGƒNƒZƒ‹‚ÖƒGƒNƒXƒ|[ƒg
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

        private void txtCarton_TextChanged(object sender, EventArgs e)
        {
            if (lblFrmName.Text != "VIEW")
            {
                string[] box = cmbModel.Text.Split('_');
                txtBoxId.Text = box[1] + "-" + DateTime.Today.ToString("yyMMdd") + "-" + txtCarton.Text;
            }
        }

        private void txtCompSerial_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtProductSerial.Focus();
                txtProductSerial.SelectAll();
            }
        }

        private void CmbModel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (VBS.Mid(cmbModel.Text, 6, 4) == "517E")
            {
                limit1 = 60;
                //  limit1 = 60;
                txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
                txtProductSerial.Enabled = true;
                txtProductSerial.Focus();
            }
            else if (VBS.Mid(cmbModel.Text, 6, 4) == "517F")
            {
                limit1 = 65;
                txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
                txtProductSerial.Enabled = true;
                txtProductSerial.Focus();
            }
            else
            {
                limit1 = 80;
                txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
                txtProductSerial.Enabled = true;
                txtProductSerial.Focus();
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

        // ƒTƒuƒvƒƒV[ƒWƒƒFƒf[ƒ^ƒe[ƒuƒ‹‚Ì’†g‚ðƒ`ƒFƒbƒN‚·‚éA–{ƒAƒvƒŠƒP[ƒVƒ‡ƒ“‚É‘Î‚µ‚ÄA’¼Ú‚ÍŠÖŒW‚È‚¢
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

        // ƒTƒuƒvƒƒV[ƒWƒƒFƒf[ƒ^ƒrƒ…[‚Ì’†g‚ðƒ`ƒFƒbƒN‚·‚éA–{ƒAƒvƒŠƒP[ƒVƒ‡ƒ“‚É‘Î‚µ‚ÄA’¼Ú‚ÍŠÖŒW‚È‚¢
        private void printDataView(DataView dv)
        {
            foreach (DataRowView drv in dv)
            {
                System.Diagnostics.Debug.Print(drv["lot"].ToString() + " " +
                    drv["tjudge"].ToString() + " " + drv["inspectdate"].ToString());
            }
        }

        //private void dgvInline_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    if (e.RowIndex < 0 || e.ColumnIndex < 0 || dtAllProcess.Rows.Count == 0)
        //    {
        //        return;
        //    }
        //    string serial = dgvInline.Rows[e.RowIndex].Cells[0].Value.ToString();
        //    ShowProcessJudge(serial);
        //}
    }
}