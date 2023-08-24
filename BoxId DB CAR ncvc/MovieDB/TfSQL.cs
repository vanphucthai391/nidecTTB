using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using Npgsql;

namespace BoxIdDb
{
    public class TfSQL
    {
        NpgsqlConnection connection;
        static string conStringBoxidDb = @"Server=192.168.193.4;Port=5432;User Id=pqm;Password=dbuser;Database=boxidcardb; CommandTimeout=100; Timeout=100;";
        static string conStringPQMDb = @"Server=192.168.193.4;Port=5432;User Id=pqm;Password=dbuser;Database=pqmdb; CommandTimeout=100; Timeout=100;";
        static string conStringOQC = @"Server=192.168.193.4;Port=5432;User Id=pqm;Password=dbuser;Database=ncardb; CommandTimeout=100; Timeout=100;";
        static string conStringERP = @"Server=192.168.193.4;Port=5432;User Id=pqm;Password=dbuser;Database=ERP_DB; CommandTimeout=100; Timeout=100;";


        public void getComboBoxData(string sql, ref ComboBox cmb)
        {
            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter();
            NpgsqlCommand command;
            DataSet ds = new DataSet();
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                command = new NpgsqlCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    cmb.Items.Add(row[0].ToString());
                }
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
            }
        }

        public void getComboBoxDataViaCsv(string sql, ref ComboBox cmb)
        {
            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter();
            NpgsqlCommand command;
            DataSet ds = new DataSet();
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                command = new NpgsqlCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                connection.Close();

                string text = ds.Tables[0].Rows[0][0].ToString();
                if (text == String.Empty) return;

                cmb.Items.Clear();
                string[] words = text.Split(',');
                foreach (string s in words)
                {
                    string t = s.Trim();
                    cmb.Items.Add(t);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
            }
        }
        public void getAutoCompleteData(string sql, ref TextBox txt)
        {
            txt.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txt.AutoCompleteSource = AutoCompleteSource.CustomSource;
            AutoCompleteStringCollection DataCollection = new AutoCompleteStringCollection();

            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter();
            NpgsqlCommand command;
            DataSet ds = new DataSet();
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                command = new NpgsqlCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                connection.Close();
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    DataCollection.Add(row[0].ToString());
                }
                txt.AutoCompleteCustomSource = DataCollection;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
            }
        }

        public double sqlExecuteScalarDouble(string sql)
        {
            double response;
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                response = Convert.ToDouble(command.ExecuteScalar());// 1 unique value. Ex: MAX, MIN, SUM, COUNT,... 
                connection.Close();
                return response;
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL executeschalar moethod failed." + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return 100;
            }
        }

        public string sqlExecuteScalarString(string sql)
        {
            string response;
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                response = Convert.ToString(command.ExecuteScalar());
                connection.Close();
                return response;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                System.Diagnostics.Debug.Print(ex.Message);
                connection.Close();
                return String.Empty;
            }
        }

        public string sqlScalarString(string sql)
        {
            string response;
            try
            {
                connection = new NpgsqlConnection(conStringPQMDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                response = Convert.ToString(command.ExecuteScalar());
                connection.Close();
                return response;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                System.Diagnostics.Debug.Print(ex.Message);
                connection.Close();
                return String.Empty;
            }
        }

        public bool sqlExecuteScalarBool(string sql)
        {
            bool response;
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                response = (bool)command.ExecuteScalar();
                connection.Close();
                return response;
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL executeschalar moethod failed." + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        public long sqlExecuteScalarLong(string sql)
        {
            long response;
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                response = (long)command.ExecuteScalar();
                connection.Close();
                return response;
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL executeschalar moethod failed." + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return 0;
            }
        }

        public bool sqlExecuteNonQuery(string sql, bool result_message_show)
        {
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                int response = command.ExecuteNonQuery();
                if (response >= 1)
                {
                    if (result_message_show) { MessageBox.Show("Successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    connection.Close();
                    return true;
                }
                else
                {
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        public int sqlExecuteNonQueryInt(string sql, bool result_message_show)
        {
            try
            {
                connection = new NpgsqlConnection(conStringBoxidDb);
                connection.Open();
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);
                int response = command.ExecuteNonQuery();
                if (response >= 1)
                {
                    if (result_message_show) { MessageBox.Show("Successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    connection.Close();
                    return response;
                }
                else
                {
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    return 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return 0;
            }
        }

        // ＦＩＬＬ ＦＲＯＭ ＢＯＸＩＤ ＤＢ
        public void sqlDataAdapterFillDatatable(string sql, ref DataTable dt)
        {
            NpgsqlConnection connection = new NpgsqlConnection(conStringBoxidDb);
            NpgsqlCommand command = new NpgsqlCommand();

            using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter())
            {
                command.CommandText = sql;
                command.Connection = connection;
                adapter.SelectCommand = command;
                adapter.Fill(dt);
            }
        }

        public void sqlDataAdapterFillDatatablePqm(string sql, ref DataTable dt)
        {
            NpgsqlConnection connection = new NpgsqlConnection(conStringPQMDb);
            NpgsqlCommand command = new NpgsqlCommand();

            using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter())
            {
                command.CommandText = sql;
                command.Connection = connection;
                adapter.SelectCommand = command;
                adapter.Fill(dt);
            }
        }

        public void sqlDataAdapterFillDatatableOqc(string sql, ref DataTable dt)
        {
            NpgsqlConnection connection = new NpgsqlConnection(conStringPQMDb);
            NpgsqlCommand command = new NpgsqlCommand();

            using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter())
            {
                command.CommandText = sql;
                command.Connection = connection;
                adapter.SelectCommand = command;
                adapter.Fill(dt);
            }
        }
        public void sqlDataAdapterFillDatatableERP(string sql, ref DataTable dt)
        {
            NpgsqlConnection connection = new NpgsqlConnection(conStringERP);
            NpgsqlCommand command = new NpgsqlCommand();

            using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter())
            {
                command.CommandText = sql;
                command.Connection = connection;
                adapter.SelectCommand = command;
                adapter.Fill(dt);
            }
        }
        public void sqlDataAdapterFillDatatableOqc517EB(string sql, ref DataTable dt)
        {
            NpgsqlConnection connection = new NpgsqlConnection(conStringPQMDb);
            NpgsqlCommand command = new NpgsqlCommand();

            using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter())
            {
                command.CommandText = sql;
                command.Connection = connection;
                adapter.SelectCommand = command;
                adapter.Fill(dt);
            }
        }

        // ボックスＩＤ・シリアルの一括登録
        public bool sqlMultipleInsertOverall(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                //string sql = "INSERT INTO product_serial_rtcd(boxid, serialno, model, carton, lot, cio_ccw, cg_ccw, cno_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cio_ccw, :cg_ccw, :cno_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge_line, :return, :inspectdate, :tjudge, :date_line)";
                string sql = "INSERT INTO product_serial_rtcd1(boxid, serialno, model, carton, lot, cio_ccw, cg_ccw, cno_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cio_ccw, :cg_ccw, :cno_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge_line, :return, :inspectdate, :tjudge, :date_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                //OQC
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("cio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cg_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cno_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("aio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ano_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("air_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("anr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ais_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    //OQC
                    command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["cio_ccw"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["cg_ccw"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["cno_ccw"].ToString();
                    //INLINE
                    command.Parameters[10].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[11].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[12].Value = dt.Rows[i]["aio_ccw"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["ano_ccw"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["air_ccw"].ToString();
                    command.Parameters[15].Value = dt.Rows[i]["anr_ccw"].ToString();
                    command.Parameters[16].Value = dt.Rows[i]["ais_ccw"].ToString();
                    command.Parameters[17].Value = dt.Rows[i]["return"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        public bool sqlMultipleInsert517EB(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_517eb(boxid, serialno, model, carton, lot, cir_ccw, cg_ccw, cnr_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge, return, inspectdate, date_line,tjudge_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cir_ccw, :cg_ccw, :cnr_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge, :return, :inspectdate, :date_line, :tjudge_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cnr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cir_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cg_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("aio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ano_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("air_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("anr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ais_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.Timestamp));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.Timestamp));
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    command.Parameters[5].Value = dt.Rows[i]["cir_ccw"].ToString();
                    command.Parameters[6].Value = dt.Rows[i]["cg_ccw"].ToString();
                    command.Parameters[7].Value = dt.Rows[i]["cnr_ccw"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["aio_ccw"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["ano_ccw"].ToString();
                    command.Parameters[10].Value = dt.Rows[i]["air_ccw"].ToString();
                    command.Parameters[11].Value = dt.Rows[i]["anr_ccw"].ToString();
                    command.Parameters[12].Value = dt.Rows[i]["ais_ccw"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["return"].ToString();
                    command.Parameters[15].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[16].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[17].Value = dt.Rows[i]["tjudge_line"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }
        public bool sqlMultipleInsert523AB(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_523ab(boxid, serialno, model, carton, lot, cio_ccw, cg_ccw, cno_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cio_ccw, :cg_ccw, :cno_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge_line, :return, :inspectdate, :tjudge, :date_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                //OQC
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("cio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cg_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cno_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("aio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ano_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("air_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("anr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ais_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    //OQC
                    command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["cio_ccw"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["cg_ccw"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["cno_ccw"].ToString();
                    //INLINE
                    command.Parameters[10].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[11].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[12].Value = dt.Rows[i]["aio_ccw"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["ano_ccw"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["air_ccw"].ToString();
                    command.Parameters[15].Value = dt.Rows[i]["anr_ccw"].ToString();
                    command.Parameters[16].Value = dt.Rows[i]["ais_ccw"].ToString();
                    command.Parameters[17].Value = dt.Rows[i]["return"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        public bool sqlMultipleInsert523AC(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_523ac(boxid, serialno, model, carton, lot, cio_ccw, cg_ccw, cno_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cio_ccw, :cg_ccw, :cno_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge_line, :return, :inspectdate, :tjudge, :date_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                //OQC
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("cio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cg_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cno_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("aio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ano_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("air_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("anr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ais_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    //OQC
                    command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["cio_ccw"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["cg_ccw"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["cno_ccw"].ToString();
                    //INLINE
                    command.Parameters[10].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[11].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[12].Value = dt.Rows[i]["aio_ccw"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["ano_ccw"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["air_ccw"].ToString();
                    command.Parameters[15].Value = dt.Rows[i]["anr_ccw"].ToString();
                    command.Parameters[16].Value = dt.Rows[i]["ais_ccw"].ToString();
                    command.Parameters[17].Value = dt.Rows[i]["return"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        public bool sqlMultipleInsert517FB(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_517fb(boxid, serialno, model, carton, lot, cio_ccw, cg_ccw, cno_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cio_ccw, :cg_ccw, :cno_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge_line, :return, :inspectdate, :tjudge, :date_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                //OQC
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("cio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cg_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cno_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("aio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ano_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("air_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("anr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ais_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    //OQC
                    command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["cio_ccw"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["cg_ccw"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["cno_ccw"].ToString();
                    //INLINE
                    command.Parameters[10].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[11].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[12].Value = dt.Rows[i]["aio_ccw"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["ano_ccw"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["air_ccw"].ToString();
                    command.Parameters[15].Value = dt.Rows[i]["anr_ccw"].ToString();
                    command.Parameters[16].Value = dt.Rows[i]["ais_ccw"].ToString();
                    command.Parameters[17].Value = dt.Rows[i]["return"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }
        public bool sqlMultipleInsert0148(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_0148(boxid, serialno, model, carton, lot, cio_ccw, cg_ccw, cno_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return, terminal, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cio_ccw, :cg_ccw, :cno_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge_line, :return, :terminal, :inspectdate, :tjudge, :date_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                //OQC
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("cio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cg_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cno_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("aio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ano_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("air_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("anr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ais_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("terminal", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    //OQC
                    command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["cio_ccw"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["cg_ccw"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["cno_ccw"].ToString();
                    //INLINE
                    command.Parameters[10].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[11].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[12].Value = dt.Rows[i]["aio_ccw"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["ano_ccw"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["air_ccw"].ToString();
                    command.Parameters[15].Value = dt.Rows[i]["anr_ccw"].ToString();
                    command.Parameters[16].Value = dt.Rows[i]["ais_ccw"].ToString();
                    command.Parameters[17].Value = dt.Rows[i]["return"].ToString();
                    command.Parameters[18].Value = dt.Rows[i]["terminal"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }
        public bool sqlMultipleInsertBMD0015(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_bmd0015(boxid, serialno, model, carton, lot, cio_ccw, cg_ccw, cno_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cio_ccw, :cg_ccw, :cno_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge_line, :return, :inspectdate, :tjudge, :date_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                //OQC
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("cio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cg_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cno_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("aio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ano_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("air_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("anr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ais_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    //OQC
                    command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["cio_ccw"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["cg_ccw"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["cno_ccw"].ToString();
                    //INLINE
                    command.Parameters[10].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[11].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[12].Value = dt.Rows[i]["aio_ccw"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["ano_ccw"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["air_ccw"].ToString();
                    command.Parameters[15].Value = dt.Rows[i]["anr_ccw"].ToString();
                    command.Parameters[16].Value = dt.Rows[i]["ais_ccw"].ToString();
                    command.Parameters[17].Value = dt.Rows[i]["return"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }
        public bool sqlMultipleInsertBMD0016(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_bmd0016(boxid, serialno, model, carton, lot, cio_ccw, cg_ccw, cno_ccw, aio_ccw, ano_ccw, air_ccw, anr_ccw, ais_ccw, tjudge_line, return, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :cio_ccw, :cg_ccw, :cno_ccw, :aio_ccw, :ano_ccw, :air_ccw, :anr_ccw, :ais_ccw, :tjudge_line, :return, :inspectdate, :tjudge, :date_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                //OQC
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("cio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cg_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("cno_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("aio_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ano_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("air_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("anr_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("ais_ccw", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    //OQC
                    command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["cio_ccw"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["cg_ccw"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["cno_ccw"].ToString();
                    //INLINE
                    command.Parameters[10].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[11].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[12].Value = dt.Rows[i]["aio_ccw"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["ano_ccw"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["air_ccw"].ToString();
                    command.Parameters[15].Value = dt.Rows[i]["anr_ccw"].ToString();
                    command.Parameters[16].Value = dt.Rows[i]["ais_ccw"].ToString();
                    command.Parameters[17].Value = dt.Rows[i]["return"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }

        // ボックスＩＤ・シリアルの一括登録
        public int sqlDeleteBoxid(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            //string sql2 = "delete from product_serial_rtcd where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_rtcd1 where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }
        public int sqlDeleteBoxidld(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_ld20 where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }
        public int sqlDeleteBoxid517EB(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_517eb where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }

        public int sqlDeleteBoxid523(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_523 where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }
        public int sqlDeleteBoxid_0025(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_0025 where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }
        public int sqlDeleteBoxid_0148(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_0148 where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }
        public int sqlDeleteBoxid_517fb(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_517fb where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }
        public int sqlDeleteBoxid_523ab(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_523ab where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }
        public int sqlDeleteBoxid_523ac(string boxid)
        {
            int res1 = 0;
            int res2 = 0;
            string sql1 = "delete from box_id_rt where boxid = '" + boxid + "'";
            string sql2 = "delete from product_serial_523ac where boxid = '" + boxid + "'";

            System.Diagnostics.Debug.Print(sql1);
            System.Diagnostics.Debug.Print(sql2);

            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                NpgsqlCommand command1 = new NpgsqlCommand(sql1, connection);
                NpgsqlCommand command2 = new NpgsqlCommand(sql2, connection);
                res1 = command1.ExecuteNonQuery();
                res2 = command2.ExecuteNonQuery();

                transaction.Commit();
                connection.Close();
                return res2;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return -1;
            }
        }
        public void getCompleteData(string sql, ref Label txt)
        {
            NpgsqlDataAdapter adapter = new NpgsqlDataAdapter();
            NpgsqlCommand command;
            DataSet ds = new DataSet();
            try
            {
                connection = new NpgsqlConnection(conStringPQMDb);
                connection.Open();
                command = new NpgsqlCommand(sql, connection);
                adapter.SelectCommand = command;
                adapter.Fill(ds);
                adapter.Dispose();
                command.Dispose();
                connection.Close();
                if (ds.Tables[0] != null && ds.Tables[0].Rows.Count > 0)
                {
                    txt.Text = ds.Tables[0].Rows[0][0].ToString();
                    // ds.Clear();
                }

                //    foreach (DataRow row in ds.Tables[0].Rows)
                //{
                //    txt.Text = row[0].ToString();
                //}
                //txt.AutoCompleteCustomSource = DataCollection;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
            }
        }
        public bool sqlMultipleInsertDeus(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_ld20(boxid, serialno, model, carton, lot, sdf0, scurave, sgrmsave, srtpctg2, sbtpctg2, tjudge_line, date_line, bin, return) VALUES (:boxid, :serialno, :model, :carton, :lot, :sdf0, :scurave, :sgrmsave, :srtpctg2, :sbtpctg2, :tjudge_line, :date_line,:bin,:return)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
               
                //OQC
                //command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                //command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                //command.Parameters.Add(new NpgsqlParameter("sdf0oqc", NpgsqlTypes.NpgsqlDbType.Varchar));
                //command.Parameters.Add(new NpgsqlParameter("scuroqc", NpgsqlTypes.NpgsqlDbType.Varchar));
                //command.Parameters.Add(new NpgsqlParameter("sgrmsoqc", NpgsqlTypes.NpgsqlDbType.Varchar));
                //command.Parameters.Add(new NpgsqlParameter("srtpcoqc", NpgsqlTypes.NpgsqlDbType.Varchar));
                //command.Parameters.Add(new NpgsqlParameter("sbtpcoqc", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("sdf0", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("scurave", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("sgrmsave", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("srtpctg2", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("sbtpctg2", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("bin", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                   
                    ////OQC
                    //command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    //command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    //command.Parameters[7].Value = dt.Rows[i]["sdf0oqc"].ToString();
                    //command.Parameters[8].Value = dt.Rows[i]["scuroqc"].ToString();
                    //command.Parameters[9].Value = dt.Rows[i]["sgrmsoqc"].ToString();
                    //command.Parameters[8].Value = dt.Rows[i]["srtpcoqc"].ToString();
                    //command.Parameters[9].Value = dt.Rows[i]["sbtpcoqc"].ToString();
                    //INLINE
                    command.Parameters[5].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["sdf0"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["scurave"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["sgrmsave"].ToString();
                    command.Parameters[10].Value = dt.Rows[i]["srtpctg2"].ToString();
                    command.Parameters[11].Value = dt.Rows[i]["sbtpctg2"].ToString();
                    command.Parameters[12].Value = dt.Rows[i]["bin"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["return"].ToString();



                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }
        public bool sqlMultipleInsertBFB0025(DataTable dt)
        {
            int res1;
            bool res2 = false;
            connection = new NpgsqlConnection(conStringBoxidDb);
            connection.Open();
            NpgsqlTransaction transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[0][i].ToString());
            }

            try
            {
                string sql = "INSERT INTO product_serial_0025(boxid, serialno, model, carton, lot, qacurrent, qafg, qaspeed, current, fg, speed, svfi, pcbbarcode, tjudge_line, return, inspectdate, tjudge, date_line) VALUES (:boxid, :serialno, :model, :carton, :lot, :qacurrent, :qafg, :qaspeed, :current, :fg, :speed,:svfi,:pcbbarcode, :tjudge_line, :return, :inspectdate, :tjudge, :date_line)";
                NpgsqlCommand command = new NpgsqlCommand(sql, connection);

                command.Parameters.Add(new NpgsqlParameter("boxid", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("serialno", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("model", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("carton", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("lot", NpgsqlTypes.NpgsqlDbType.Varchar));
                //OQC
                command.Parameters.Add(new NpgsqlParameter("tjudge", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("inspectdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("qacurrent", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("qafg", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("qaspeed", NpgsqlTypes.NpgsqlDbType.Varchar));
                //INLINE
                command.Parameters.Add(new NpgsqlParameter("tjudge_line", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("date_line", NpgsqlTypes.NpgsqlDbType.TimestampTZ));
                command.Parameters.Add(new NpgsqlParameter("current", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("fg", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("speed", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("svfi", NpgsqlTypes.NpgsqlDbType.Varchar));
                command.Parameters.Add(new NpgsqlParameter("pcbbarcode", NpgsqlTypes.NpgsqlDbType.Varchar));

                command.Parameters.Add(new NpgsqlParameter("return", NpgsqlTypes.NpgsqlDbType.Varchar));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    command.Parameters[0].Value = dt.Rows[i]["boxid"].ToString();
                    command.Parameters[1].Value = dt.Rows[i]["serialno"].ToString();
                    command.Parameters[2].Value = dt.Rows[i]["model"].ToString();
                    command.Parameters[3].Value = dt.Rows[i]["carton"].ToString();
                    command.Parameters[4].Value = dt.Rows[i]["lot"].ToString();
                    //OQC
                    command.Parameters[5].Value = dt.Rows[i]["tjudge"].ToString();
                    command.Parameters[6].Value = DateTime.Parse(dt.Rows[i]["inspectdate"].ToString());
                    command.Parameters[7].Value = dt.Rows[i]["qacurrent"].ToString();
                    command.Parameters[8].Value = dt.Rows[i]["qafg"].ToString();
                    command.Parameters[9].Value = dt.Rows[i]["qaspeed"].ToString();
                    //INLINE
                    command.Parameters[10].Value = dt.Rows[i]["tjudge_line"].ToString();
                    command.Parameters[11].Value = DateTime.Parse(dt.Rows[i]["date_line"].ToString());
                    command.Parameters[12].Value = dt.Rows[i]["current"].ToString();
                    command.Parameters[13].Value = dt.Rows[i]["fg"].ToString();
                    command.Parameters[14].Value = dt.Rows[i]["speed"].ToString();
                    command.Parameters[15].Value = dt.Rows[i]["svfi"].ToString();
                    command.Parameters[16].Value = dt.Rows[i]["pcbbarcode"].ToString();
                    command.Parameters[17].Value = dt.Rows[i]["return"].ToString();

                    System.Diagnostics.Debug.Print(command.ToString());
                    res1 = command.ExecuteNonQuery();
                    if (res1 == -1) res2 = true;
                }

                if (!res2)
                {
                    transaction.Commit();
                    connection.Close();
                    return true;
                }
                else
                {
                    transaction.Rollback();
                    MessageBox.Show("Not successful!", "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    connection.Close();
                    transaction.Rollback();
                    return false;
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show("Not successful!" + "\r\n" + ex.Message
                                , "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                connection.Close();
                return false;
            }
        }
    }
}
