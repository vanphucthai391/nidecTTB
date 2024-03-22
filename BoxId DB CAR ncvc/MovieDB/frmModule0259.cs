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
    public partial class frmModule0259 : Form
    {
        //親フォームfrmBoxへイベント発生を連絡（デレゲート）
        public delegate void RefreshEventHandler(object sender, EventArgs e);
        public event RefreshEventHandler RefreshEvent;

        // プリント用テキストファイルの保存用フォルダを、基本設定ファイルで設定する
        string appconfig = @"\\192.168.193.1\barcode$\BoxId Printer vc5\info.ini";
        string directory = @"C:\Users\takusuke.fujii\Desktop\Auto Print\\";

        //その他、非ローカル変数
        bool formEditMode;
        bool formReturnMode;
        bool formAddMode;
        string user;
        string m_model;
        string tablethis;
        string tablelast;
        string m_lot;
        int okCount;
        bool inputBoxModeOriginal;
        string testerTableThisMonth;
        string testerTableLastMonth;
        string tableThisMonth;
        string tableLastMonth;
        DataTable dtOverall;
        DataTable dtAllProcess;

        //DataTable dtTester;
        int limit1 =108;
        public int limit2 = 0;
        bool sound;

        // コンストラクタ
        public frmModule0259()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void frmModule_Load(object sender, EventArgs e)
        {
            cmbModel.SelectedIndex = 0;
            txtCarton.Enabled = false;
            user = txtUser.Text;
            txtLimit.Text = limit2.ToString();
            directory =  /*@"C:\Users\mt-qc20\Desktop\print\";*/ readIni("TARGET DIRECTORY", "DIR", appconfig);
            this.Left = 250;
            this.Top = 20;
            dtOverall = new DataTable();
            defineAndReadDtOverall(ref dtOverall);
            if (!formEditMode)
            {
                if (dtOverall.Rows.Count >= 0)
                {
                    limit1 = 108;
                }
            }
            updateDataGridViews(dtOverall, ref dgvInline);
            if (!formEditMode)
            {
                txtProductSerial.Enabled = false;
            }
            txtCarton.Enabled = true;
            txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();
            txtProductSerial.Enabled = true;
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
        // Windows API をインポート
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filepath);

        // サブプロシージャ：親フォームで呼び出し、親フォームの情報を、テキストボックスへ格納して引き継ぐ
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
                cmbModel.Text = "BMA_0259";
                limit1 = 108;
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

        // サブプロシージャ：ＤＢからのＤＴＯＶＥＲＡＬＬへの読み込み
        private void defineAndReadDtOverall(ref DataTable dt)
        {
            string boxId = txtBoxId.Text;
            dt.Columns.Add("serialno", Type.GetType("System.String"));
            dt.Columns.Add("model", Type.GetType("System.String"));
            dt.Columns.Add("lot", Type.GetType("System.String"));
            dt.Columns.Add("inspectdate", Type.GetType("System.DateTime"));
            dt.Columns.Add("cio_cw", Type.GetType("System.String"));
            dt.Columns.Add("cg_cw", Type.GetType("System.String"));
            dt.Columns.Add("cno_cw", Type.GetType("System.String"));
            dt.Columns.Add("tjudge", Type.GetType("System.String"));
            dt.Columns.Add("date_line", Type.GetType("System.DateTime"));
            dt.Columns.Add("aio_cw", Type.GetType("System.String"));
            dt.Columns.Add("ano_cw", Type.GetType("System.String"));
            dt.Columns.Add("air_cw", Type.GetType("System.String"));
            dt.Columns.Add("anr_cw", Type.GetType("System.String"));
            dt.Columns.Add("ais_cw", Type.GetType("System.String"));
            dt.Columns.Add("tjudge_line", Type.GetType("System.String"));
            dt.Columns.Add("return", Type.GetType("System.String"));
            if (!formEditMode)
            {
                string sql = "select serialno, model, lot, inspectdate, cio_cw, cg_cw, cno_cw, tjudge, date_line, aio_cw, ano_cw, air_cw, anr_cw, ais_cw, tjudge_line, return " +
"FROM product_serial_0259 WHERE boxid='" + boxId + "'";
                TfSQL tf = new TfSQL();
                System.Diagnostics.Debug.Print(sql);
                tf.sqlDataAdapterFillDatatable(sql, ref dt);
            }
        }

        // サブプロシージャ：データグリットビューの更新
        private void updateDataGridViews(DataTable dt1, ref DataGridView dgv1)
        {
            // 入力用ボックスの有効・無効を保持し、何れの場合も一時的に無効にする
            inputBoxModeOriginal = txtProductSerial.Enabled;
            txtProductSerial.Enabled = false;

            // データグリットビューへＤＴＡＡＴＡＢＬＥを格納
            updateDataGridViewsSub(dt1, ref dgv1);

            // テスト結果がＦＡＩＬまたはレコードなしのシリアルをマーキングする 
            colorViewForFailAndBlank(ref dgv1);
            // colorViewForFailAndBlank(ref dgv2);

            // 重複レコード、および１セル２重入力をマーキングする
            colorViewForDuplicateSerial(ref dgv1);
            // colorViewForDuplicateSerial(ref dgv2);

            // ２つ以上のコンフィグが混在する場合に警告し、データグリットビューをマークする

            //colorMixedLot(dt1, ref dgv1);

            //行ヘッダーに行番号を表示する（インライン）
            for (int i = 0; i < dgv1.Rows.Count; i++)
                dgv1.Rows[i].HeaderCell.Value = (i + 1).ToString();

            //行ヘッダーの幅を自動調節する（インライン）
            dgv1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            // 一番下の行を表示する（インライン）
            if (dgv1.Rows.Count >= 1)
                dgv1.FirstDisplayedScrollingRowIndex = dgv1.Rows.Count - 1;

            // 入力用ボックスの有効・無効を元へ戻す
            txtProductSerial.Enabled = inputBoxModeOriginal;

            // 現在の一時登録件数を変数へ保持する
            okCount = getOkCount(dt1);
            txtOkCount.Text = okCount.ToString() + "/" + limit1.ToString();

            // スキャン件数が既にＬＩＭＩＴに達している場合は、入力用ボックスを無効にする
            if (okCount == limit1)
            {
                txtProductSerial.Enabled = false;
            }
            else
            {
                txtProductSerial.Enabled = true;
            }

            // グリットレコード件数と、okCount数が一致している場合に、登録ボタンを有効にする
            if (okCount == limit1 && dgv1.Rows.Count == limit1)
            {
                btnRegisterBoxId.Enabled = true;
            }
            else
            {
                btnRegisterBoxId.Enabled = false;
            }
        }

        // サブプロシージャ：シリアル番号重複なしのＰＡＳＳ個数を取得する
        private int getOkCount(DataTable dt)
        {
            if (dt.Rows.Count <= 0) return 0;
            DataTable distinct = dt.DefaultView.ToTable(true, new string[] { "serialno", "tjudge", "tjudge_line" });
            DataRow[] dr = distinct.Select("tjudge = 'PASS' and tjudge_line = 'PASS'");
            int dist = dr.Length;
            return dist;
        }

        // サブプロシージャ：メインデータグリットビューへデータテーブルを格納、および集計グリッドビューの作成
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
                //Alarm OQC FAIL or NODATA
                if (dgv["col_judge_oqc", i].Value.ToString() == "FAIL" || dgv["col_judge_oqc", i].Value.ToString() == "PLS NG" || String.IsNullOrEmpty(dgv["col_judge_oqc", i].Value.ToString()))
                {
                    dgv["col_cg_cw", i].Style.BackColor = Color.Red;
                    dgv["col_cio_cw", i].Style.BackColor = Color.Red;
                    dgv["col_cno_cw", i].Style.BackColor = Color.Red;
                    dgv["col_date", i].Style.BackColor = Color.Red;
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
                    dgv["col_judge_inline", i].Style.BackColor = Color.Red;
                    dgv["col_aio_cw", i].Style.BackColor = Color.Red;
                    dgv["col_air_cw", i].Style.BackColor = Color.Red;
                    dgv["col_ais_cw", i].Style.BackColor = Color.Red;
                    dgv["col_ano_cw", i].Style.BackColor = Color.Red;
                    dgv["col_anr_cw", i].Style.BackColor = Color.Red;

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

        // サブプロシージャ：重複レコード、または１セル２重入力をマーキングする
        private void colorViewForDuplicateSerial(ref DataGridView dgv)
        {
            DataTable dt = ((DataTable)dgv.DataSource).Copy();
            if (dt.Rows.Count <= 0) return;

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                string serial;
                //if (cmbModel.Text == "LA20_517CB")
                //{
                //    serial = VBS.Mid(dgv["col_serial_no", i].Value.ToString(), 2, 21);
                //}
                //else 
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

        // シリアルがスキャンされた時の処理
        private void txtProductSerial_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (cmbModel.Text == "")
                {

                    MessageBox.Show("Please select model name", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cmbModel.Focus();
                    return;
                }

                // 入力用テキストボックスを編集不可にして、処理中のスキャンをブロックする
                txtProductSerial.Enabled = false;
                string serial = txtProductSerial.Text;

                decideReferenceTable();

                if (serial != String.Empty)
                {
                    string model = cmbModel.Text;

                    #region Data OQC
                    string sql2 = "select serno, tjudge, inspectdate, " +
"MAX(case inspect when 'CG_CW' then inspectdata else null end) as CG_CW," +
"MAX(case inspect when 'CIO_CW' then inspectdata else null end) as CIO_CW," +
"MAX(case inspect when 'CNO_CW' then inspectdata else null end) as CNO_CW" +
" FROM" +
" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableThisMonth + "data" +
" WHERE serno = (select serno from " + testerTableThisMonth + " where process = 'NMT2' and serno = '" + serial + "' LIMIT 1) and inspect in ('CG_CW','CIO_CW','CNO_CW'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableThisMonth + " where serno = '" + serial + "' and process = 'NMT2' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
" GROUP BY serno, tjudge, inspectdate" +

" UNION ALL " +

"select serno, tjudge, inspectdate, " +
"MAX(case inspect when 'CG_CW' then inspectdata else null end) as CG_CW," +
"MAX(case inspect when 'CIO_CW' then inspectdata else null end) as CIO_CW," +
"MAX(case inspect when 'CNO_CW' then inspectdata else null end) as CNO_CW" +
" FROM" +
" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge from (select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE from (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag from (select * from " + testerTableLastMonth + "data" +
" WHERE serno = (select serno from " + testerTableLastMonth + " where process = 'NMT2' and serno = '" + serial + "' LIMIT 1) and inspect in ('CG_CW','CIO_CW','CNO_CW'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + testerTableLastMonth + " where serno = '" + serial + "' and process = 'NMT2' and tjudge = '0' order by inspectdate desc LIMIT 1) d" +
" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
" GROUP BY serno, tjudge, inspectdate";
                    System.Diagnostics.Debug.Print(System.Environment.NewLine + sql2);
                    DataTable dt2 = new DataTable();
                    TfSQL tf = new TfSQL();
                    tf.sqlDataAdapterFillDatatableOqc(sql2, ref dt2);
                    #endregion
                    #region Data INLINE
                    string sql1 = "select serno, tjudge as tjudge_line, inspectdate as date_line, " +
"MAX(case inspect when 'AIO_CW' then inspectdata else null end) as AIO_CW," +
"MAX(case inspect when 'ANO_CW' then inspectdata else null end) as ANO_CW," +
"MAX(case inspect when 'AIR_CW' then inspectdata else null end) as AIR_CW," +
"MAX(case inspect when 'ANR_CW' then inspectdata else null end) as ANR_CW," +
"MAX(case inspect when 'AIS_CW' then inspectdata else null end) as AIS_CW" +
" FROM" +
" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge FROM(select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE FROM (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag FROM (select * from " + tableThisMonth + "data" +
" WHERE serno = (select lot from " + testerTableThisMonth + " where process = 'NO53' and serno = '" + serial + "' LIMIT 1) and inspect in ('AIO_CW','AIR_CW','AIS_CW','ANO_CW','ANR_CW'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + tableThisMonth + " where serno = (select lot from " + testerTableThisMonth + " where process = 'NO53' and serno = '" + serial + "' LIMIT 1) and process = 'NO41' order by inspectdate desc LIMIT 1) d" +
" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
" GROUP BY serno, tjudge, inspectdate" +

" UNION ALL " +

"select serno, tjudge as tjudge_line, inspectdate as date_line, " +
"MAX(case inspect when 'AIO_CW' then inspectdata else null end) as AIO_CW," +
"MAX(case inspect when 'ANO_CW' then inspectdata else null end) as ANO_CW," +
"MAX(case inspect when 'AIR_CW' then inspectdata else null end) as AIR_CW," +
"MAX(case inspect when 'ANR_CW' then inspectdata else null end) as ANR_CW," +
"MAX(case inspect when 'AIS_CW' then inspectdata else null end) as AIS_CW" +
" FROM" +
" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge FROM(select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE FROM (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag FROM (select * from " + tableLastMonth + "data" +
" WHERE serno = (select lot from " + testerTableThisMonth + " where process = 'NO53' and serno = '" + serial + "' LIMIT 1) and inspect in ('AIO_CW','AIR_CW','AIS_CW','ANO_CW','ANR_CW'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + tableLastMonth + " where serno = (select lot from " + testerTableThisMonth + " where process = 'NO53' and serno = '" + serial + "' LIMIT 1) and process = 'NO41' order by inspectdate desc LIMIT 1) d" +
" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
" GROUP BY serno, tjudge, inspectdate" +

" UNION ALL " +

"select serno, tjudge as tjudge_line, inspectdate as date_line, " +
"MAX(case inspect when 'AIO_CW' then inspectdata else null end) as AIO_CW," +
"MAX(case inspect when 'ANO_CW' then inspectdata else null end) as ANO_CW," +
"MAX(case inspect when 'AIR_CW' then inspectdata else null end) as AIR_CW," +
"MAX(case inspect when 'ANR_CW' then inspectdata else null end) as ANR_CW," +
"MAX(case inspect when 'AIS_CW' then inspectdata else null end) as AIS_CW" +
" FROM" +
" (select d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge FROM(select SERNO, INSPECTDATE, INSPECT, INSPECTDATA, JUDGE FROM (select SERNO, INSPECT, INSPECTDATA, JUDGE, max(inspectdate) as inspectdate, row_number() OVER(PARTITION BY inspect ORDER BY max(inspectdate) desc) as flag FROM (select * from " + tableLastMonth + "data" +
" WHERE serno = (select lot from " + testerTableLastMonth + " where process = 'NO53' and serno = '" + serial + "' LIMIT 1) and inspect in ('AIO_CW','AIR_CW','AIS_CW','ANO_CW','ANR_CW'))" + "a group by SERNO, INSPECTDATE , INSPECT , INSPECTDATA , JUDGE ) b where flag = 1) c," + "(select serno, tjudge from " + tableLastMonth + " where serno = (select lot from " + testerTableLastMonth + " where process = 'NO53' and serno = '" + serial + "' LIMIT 1) and process = 'NO41' order by inspectdate desc LIMIT 1) d" +
" group by d.serno, d.tjudge, c.inspectdate, c.inspect, c.inspectdata, c.judge) e " +
" GROUP BY serno, tjudge, inspectdate";
                    #endregion

                    System.Diagnostics.Debug.Print(System.Environment.NewLine + sql1);
                    DataTable dt1 = new DataTable();

                    tf.sqlDataAdapterFillDatatablePqm(sql1, ref dt1);
                    #region -- Get All Process Judge --
                    string queryProcess = "";
                    string lotthis = string.Format("SELECT lot FROM (SELECT lot, inspectdate, ROW_NUMBER() OVER(PARTITION BY lot ORDER BY inspectdate DESC) from {1} where serno = '{0}' and process = 'NO53' ORDER BY lot)tb where ROW_NUMBER = 1", serial, testerTableThisMonth);
                    string lotlast = string.Format("SELECT lot FROM (SELECT lot, inspectdate, ROW_NUMBER() OVER(PARTITION BY lot ORDER BY inspectdate DESC) from {1} where serno = '{0}' and process = 'NO53' ORDER BY lot)tb where ROW_NUMBER = 1", serial, testerTableLastMonth);
                    queryProcess = string.Format("SELECT serno, lot, inspectdate, process,judge from "
                    + "(SELECT serno, lot, inspectdate, process, judge, ROW_NUMBER() OVER(PARTITION BY process ORDER BY inspectdate DESC) from "
                    + "(SELECT ({3}) as serno, serno lot,inspectdate, process,"
                    + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {4} "
                    + "WHERE serno in (SELECT DISTINCT lot FROM {4} WHERE process = 'NO53' AND serno = ({3})) "
                    + "OR serno = ({3})"
                    + "UNION ALL SELECT ({6}) as serno, serno lot, inspectdate, process,"
                    + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {5} "
                    + "WHERE serno in (SELECT DISTINCT lot FROM {5} WHERE process = 'NO53' AND serno = ({6})) "
                    + "OR serno = ({6})"
                    + "UNION ALL SELECT ({3}) as serno, serno lot, inspectdate, process,"
                    + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {5} "
                    + "WHERE serno in (SELECT DISTINCT lot FROM {1} WHERE process = 'NO53' AND serno = ({3})) "
                    + "OR serno = ({3})"
                    + "UNION ALL SELECT '{0}' as serno, serno lot, inspectdate, process, "
                    + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {1} "
                    + "WHERE serno in (SELECT DISTINCT lot FROM {1} WHERE process = 'NO53' AND serno = '{0}') "
                    + "OR serno = '{0}' "
                    + "UNION ALL SELECT '{0}' as serno, serno lot, inspectdate, process, "
                    + "(CASE WHEN tjudge = '0' THEN 'PASS' ELSE 'FAILURE' END) AS judge FROM {2} "
                    + "WHERE serno in (SELECT DISTINCT lot FROM {2} WHERE process = 'NO53' AND serno = '{0}') "
                    + "OR serno = '{0}' ORDER BY process) tbl) tb where ROW_NUMBER = 1 and process in ('NMT2','NO41','NO43','NO53','NO56')", serial, testerTableThisMonth, testerTableLastMonth, lotthis, tablethis, tablelast, lotlast);
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
                    dr["model"] = "BMA_0259";
                    dr["serialno"] = serial;
                    dr["lot"] = VBS.Mid(serial, 14, 3).Length < 3 ? "Error" : VBS.Mid(serial, 14, 3);
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

                    if (dt1.Rows.Count != 0)
                    {
                        dr["aio_cw"] = dt1.Rows[0]["aio_cw"].ToString();
                        dr["ano_cw"] = dt1.Rows[0]["ano_cw"].ToString();
                        dr["air_cw"] = dt1.Rows[0]["air_cw"].ToString();
                        dr["anr_cw"] = dt1.Rows[0]["anr_cw"].ToString();
                        dr["ais_cw"] = dt1.Rows[0]["ais_cw"].ToString();
                        //T-judge LINE
                        string judge_line = String.Empty;
                        string buff = dt1.Rows[0]["tjudge_line"].ToString();
                        if (buff == "0") judge_line = "PASS";
                        else if (buff == "1") judge_line = "FAIL";
                        else judge_line = "ERROR";

                        dr["tjudge_line"] = judge_line;
                        dr["date_line"] = dt1.Rows[0]["date_line"].ToString();
                    }

                    if (dt2.Rows.Count != 0)
                    {
                        dr["cg_cw"] = dt2.Rows[0]["cg_cw"].ToString();
                        dr["cio_cw"] = dt2.Rows[0]["cio_cw"].ToString();
                        dr["cno_cw"] = dt2.Rows[0]["cno_cw"].ToString();
                    }

                    dtOverall.Rows.Add(dr);

                    // データグリットビューの更新
                    updateDataGridViews(dtOverall, ref dgvInline);
                }
                // 入力用テキストボックスを編集可能へ戻し、連続してスキャンできるよう、テキストを選択状態にする
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
                if (cmbModel.Text == "BMA_0051")
                {
                    var checkFail = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("judge").Contains("FAILURE"));
                    var checknmt2 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NMT2"));
                    var checkno41 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO41"));
                    var checkno43 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO43"));
                    var checkno53 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO53"));
                    var checkno56 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO56"));
                    var checkno44 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO44"));

                    if (!checknmt2)
                    {
                        txtResultDetail.BackColor = Color.Red;
                        txtCount.Text = "NG";
                        txtCount.BackColor = Color.Red;
                        datastring += "NMT2: NO DATA\r\n";
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
                    if (!checkno53)
                    {
                        txtResultDetail.BackColor = Color.Red;
                        txtCount.Text = "NG";
                        txtCount.BackColor = Color.Red;
                        datastring += "NO53: NO DATA\r\n";

                    }
                    if (!checkno56)
                    {
                        txtResultDetail.BackColor = Color.Red;
                        txtCount.Text = "NG";
                        txtCount.BackColor = Color.Red;
                        datastring += "NO56: NO DATA\r\n";
                    }
                    //if (!checkrivet)
                    //{
                    //    txtResultDetail.BackColor = Color.Red;
                    //    txtCount.Text = "NG";
                    //    txtCount.BackColor = Color.Red;
                    //    datastring += "RIVET: NO DATA\r\n";
                    //}

                    if (checkFail)
                    {
                        txtResultDetail.BackColor = Color.Red;
                        txtCount.Text = "NG";
                        txtCount.BackColor = Color.Red;
                        txtResultDetail.Text = datastring;
                    }
                    if (!checkFail && checkno41 && checkno43 && checkno44 && checkno53 && checkno56 && checknmt2)
                    {
                        txtCount.Text = "OK";
                        txtCount.BackColor = Color.SpringGreen;
                        txtResultDetail.BackColor = Color.SpringGreen;
                        txtResultDetail.Text = datastring;
                    }
                    txtResultDetail.Text = datastring;
                }
                else
                {

                    var checkFail = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("judge").Contains("FAILURE"));
                    var checknmt2 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NMT2"));
                    var checkno41 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO41"));
                    var checkno43 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO43"));
                    var checkno53 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO53"));
                    var checkno56 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO56"));
                    //  var checkno44 = dtAllProcess.AsEnumerable().Any(x => x.Field<string>("process").Contains("NO44"));

                    if (!checknmt2)
                    {
                        txtResultDetail.BackColor = Color.Red;
                        txtCount.Text = "NG";
                        txtCount.BackColor = Color.Red;
                        datastring += "NMT2: NO DATA\r\n";
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
                    //if (!checkno44)
                    //{
                    //    txtResultDetail.BackColor = Color.Red;
                    //    txtCount.Text = "NG";
                    //    txtCount.BackColor = Color.Red;
                    //    datastring += "NO44: NO DATA\r\n";
                    //}
                    if (!checkno53)
                    {
                        txtResultDetail.BackColor = Color.Red;
                        txtCount.Text = "NG";
                        txtCount.BackColor = Color.Red;
                        datastring += "NO53: NO DATA\r\n";

                    }
                    if (!checkno56)
                    {
                        txtResultDetail.BackColor = Color.Red;
                        txtCount.Text = "NG";
                        txtCount.BackColor = Color.Red;
                        datastring += "NO56: NO DATA\r\n";
                    }
                    //if (!checkrivet)
                    //{
                    //    txtResultDetail.BackColor = Color.Red;
                    //    txtCount.Text = "NG";
                    //    txtCount.BackColor = Color.Red;
                    //    datastring += "RIVET: NO DATA\r\n";
                    //}

                    if (checkFail)
                    {
                        txtResultDetail.BackColor = Color.Red;
                        txtCount.Text = "NG";
                        txtCount.BackColor = Color.Red;
                        txtResultDetail.Text = datastring;
                    }
                    if (!checkFail && checkno41 && checkno43 && checkno53 && checkno56 && checknmt2)
                    {
                        txtCount.Text = "OK";
                        txtCount.BackColor = Color.SpringGreen;
                        txtResultDetail.BackColor = Color.SpringGreen;
                        txtResultDetail.Text = datastring;
                    }
                    txtResultDetail.Text = datastring;
                }
            }
        }
        private void decideReferenceTable()
        {
            string modelsub = "0259";
            string model_sub = "BMA0_0259";
            string model_c = "BMA0_0259";
            switch (modelsub)
            {
                case "0259":
                    testerTableThisMonth = "BMA0_00259" + DateTime.Today.ToString("yyyyMM");
                    tableThisMonth = model_sub + DateTime.Today.ToString("yyyyMM");
                    testerTableLastMonth = "BMA0_00259" + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
                    tableLastMonth = model_sub + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
                    tablethis = model_c + DateTime.Today.ToString("yyyyMM");
                    tablelast = model_c + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                        (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
                    break;
            }
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
                bool res1 = tf.sqlMultipleInsert0259(dt);
                if (res1)
                {
                    // バーコードを印字（念のためメインモデルを今一度取得した後）
                    //m_model = dtOverall.Rows[0]["model"].ToString();
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
                    //一旦登録したＢＯＸＩＤを消去する
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
            string sql = "select serialno, boxid, model FROM product_serial_0259 where model='" + cmbModel.Text + "'";
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

            // セルの選択範囲が２列以上の場合は、メッセージの表示のみでプロシージャを抜ける
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
                string sql = "delete from product_serial_0259 where boxid = '" + boxId + "'";
                System.Diagnostics.Debug.Print(sql);
                TfSQL tf = new TfSQL();
                bool res1 = tf.sqlExecuteNonQuery(sql, false);

                // DataTbleにボックスＩＤを追加し、本番テーブルへ一括登録
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

                // 追加モードを終了し、閲覧モードの表示へ切り替える
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
                // ＤＥＬＥＴＥ ＳＱＬ文を発行し、データベースから削除する
                string boxId = txtBoxId.Text;
                string whereSer = string.Empty;
                foreach (DataGridViewCell cell in dgvInline.SelectedCells)
                {
                    whereSer += "'" + cell.Value.ToString() + "', ";
                }
                string sql = "delete from product_serial_0259 where boxid = '" + boxId + "' and  serialno in (" + VBS.Left(whereSer, whereSer.Length - 2) + ")";
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
                    int res = tf.sqlDeleteBoxid0259(boxid);
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
    }
}