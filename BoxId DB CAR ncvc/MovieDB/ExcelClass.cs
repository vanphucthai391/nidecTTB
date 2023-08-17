using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace BoxIdDb
{
    public class ExcelClass
    {

        public void ExportToExcel(DataTable dt)
        {
            try
            {
                if (dt == null || dt.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                Excel.Application excelApp = new Excel.Application();
                excelApp.Workbooks.Add();
                Excel.Worksheet ws = excelApp.ActiveSheet;

                // column headings
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ws.Cells[1, (i + 1)] = dt.Columns[i].ColumnName;
                }

                // rows
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ws.Cells[(i + 2), (j + 1)] = dt.Rows[i][j];
                    }
                }
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Database Responce", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void Export(DataGridView dgv, string sheetName, string title)
        {
            string[] model = sheetName.Split('-');
            //Tạo các đối tượng Excel

            Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbooks oBooks;

            Microsoft.Office.Interop.Excel.Sheets oSheets;

            Microsoft.Office.Interop.Excel.Workbook oBook;

            Microsoft.Office.Interop.Excel.Worksheet oSheet;

            //Tạo mới một Excel WorkBook 

            oExcel.Visible = true;

            oExcel.DisplayAlerts = false;

            oExcel.Application.SheetsInNewWorkbook = 1;

            oBooks = oExcel.Workbooks;

            oBook = (Microsoft.Office.Interop.Excel.Workbook)(oExcel.Workbooks.Add(Type.Missing));

            oSheets = oBook.Worksheets;

            oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(1);

            oSheet.Name = sheetName;

            // Tạo phần đầu nếu muốn
            Microsoft.Office.Interop.Excel.Range head = oSheet.get_Range("A1", "P2");
            if (model[0] == "517EB")
            {
                head = oSheet.get_Range("A1", "M2");
            }
            else if (model[0] == "0025")
                head = oSheet.get_Range("A1", "O2");
            else head = oSheet.get_Range("A1", "P2");

            head.MergeCells = true;

            head.Value2 = title;

            head.Font.Bold = true;

            head.Font.Name = "Tahoma";

            head.Font.Size = "18";

            //head.WrapText = true;

            head.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            head.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Tạo tiêu đề cột 
            Microsoft.Office.Interop.Excel.Range rowHead = oSheet.get_Range("A3", "P3");

            Microsoft.Office.Interop.Excel.Range cl1 = oSheet.get_Range("A3", "A3");
            cl1.Value2 = "SerialNo";
            oSheet.Columns[1].NumberFormat = "Text";
            cl1.ColumnWidth = 25.0;

            Microsoft.Office.Interop.Excel.Range cl2 = oSheet.get_Range("B3", "B3");
            cl2.Value2 = "Model";
            cl2.ColumnWidth = 13.0;

            Microsoft.Office.Interop.Excel.Range cl3 = oSheet.get_Range("C3", "C3");
            cl3.Value2 = "Lot";
            cl3.ColumnWidth = 6.0;

            if (model[0] == "517EB")
            {   //OQC
                Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value2 = "Datetest NMT";
                cl4.ColumnWidth = 20.0;
                Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value2 = "CIR_CCW-Motor COMP Rated current";
                cl5.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value2 = "CG_CCW-Motor COMP Vibration";
                cl6.ColumnWidth = 10.0;

                Microsoft.Office.Interop.Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value2 = "CNR_CCW-Motor COMP Rated speed";
                cl7.ColumnWidth = 10.0;

                Microsoft.Office.Interop.Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                cl8.Value2 = "Judge OQC";
                cl8.ColumnWidth = 7.0;
                //INLINE
                Microsoft.Office.Interop.Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                cl9.Value2 = "Datetest NO41";
                cl9.ColumnWidth = 20.0;

                Microsoft.Office.Interop.Excel.Range cl10 = oSheet.get_Range("J3", "J3");
                cl10.Value2 = "AIO_CCW-No load current";
                cl10.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl11 = oSheet.get_Range("K3", "K3");
                cl11.Value2 = "ANO_CCW-No load speed";
                cl11.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl12 = oSheet.get_Range("L3", "L3");
                cl12.Value2 = "AIR_CCW-Rated load current";
                cl12.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl13 = oSheet.get_Range("M3", "M3");
                cl13.Value2 = "ANR_CCW-Rated load speed";
                cl13.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl14 = oSheet.get_Range("N3", "N3");
                cl14.Value2 = "AIS_CCW-Stall current";
                cl14.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl15 = oSheet.get_Range("O3", "O3");
                cl15.Value2 = "Judge LINE";
                cl15.ColumnWidth = 7.0;

                Microsoft.Office.Interop.Excel.Range cl16 = oSheet.get_Range("P3", "P3");
                cl16.Value2 = "Return";
                cl16.ColumnWidth = 7.0;

                rowHead = oSheet.get_Range("A3", "P3");
            }
            else if (model[0] == "0148")
            {   //OQC
                Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value2 = "Datetest NMT";
                cl4.ColumnWidth = 20.0;
                Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value2 = "CIR_CCW-Motor COMP Rated current";
                cl5.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value2 = "CG_CCW-Motor COMP Vibration";
                cl6.ColumnWidth = 10.0;

                Microsoft.Office.Interop.Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value2 = "CNR_CCW-Motor COMP Rated speed";
                cl7.ColumnWidth = 10.0;

                Microsoft.Office.Interop.Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                cl8.Value2 = "Judge OQC";
                cl8.ColumnWidth = 7.0;
                //INLINE
                Microsoft.Office.Interop.Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                cl9.Value2 = "Datetest NO41";
                cl9.ColumnWidth = 20.0;

                Microsoft.Office.Interop.Excel.Range cl10 = oSheet.get_Range("J3", "J3");
                cl10.Value2 = "AIO_CCW-No load current";
                cl10.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl11 = oSheet.get_Range("K3", "K3");
                cl11.Value2 = "ANO_CCW-No load speed";
                cl11.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl12 = oSheet.get_Range("L3", "L3");
                cl12.Value2 = "AIR_CCW-Rated load current";
                cl12.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl13 = oSheet.get_Range("M3", "M3");
                cl13.Value2 = "ANR_CCW-Rated load speed";
                cl13.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl14 = oSheet.get_Range("N3", "N3");
                cl14.Value2 = "AIS_CCW-Stall current";
                cl14.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl15 = oSheet.get_Range("O3", "O3");
                cl15.Value2 = "Judge LINE";
                cl15.ColumnWidth = 7.0;

                Microsoft.Office.Interop.Excel.Range cl16 = oSheet.get_Range("P3", "P3");
                cl16.Value2 = "Return";
                cl16.ColumnWidth = 7.0;

                Microsoft.Office.Interop.Excel.Range cl17 = oSheet.get_Range("Q3", "Q3");
                cl17.Value2 = "Terminal";
                cl17.ColumnWidth = 20.0;

                rowHead = oSheet.get_Range("A3", "Q3");
            }
            else if (model[0] == "0025")
            {   //OQC
                Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value2 = "Datetest QA";
                cl4.ColumnWidth = 20.0;
                Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value2 = "QA CURRENT";
                cl5.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value2 = "QA FG";
                cl6.ColumnWidth = 10.0;

                Microsoft.Office.Interop.Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value2 = "QA SPEED";
                cl7.ColumnWidth = 10.0;

                Microsoft.Office.Interop.Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                cl8.Value2 = "Judge QA";
                cl8.ColumnWidth = 7.0;
                //INLINE
                Microsoft.Office.Interop.Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                cl9.Value2 = "Datetest LINE";
                cl9.ColumnWidth = 20.0;

                Microsoft.Office.Interop.Excel.Range cl10 = oSheet.get_Range("J3", "J3");
                cl10.Value2 = "CURRENT";
                cl10.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl11 = oSheet.get_Range("K3", "K3");
                cl11.Value2 = "FG";
                cl11.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl12 = oSheet.get_Range("L3", "L3");
                cl12.Value2 = "SPEED";
                cl12.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl13 = oSheet.get_Range("M3", "M3");
                cl13.Value2 = "Judge INLINE";
                cl13.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl14 = oSheet.get_Range("N3", "N3");
                cl14.Value2 = "SVFI";
                cl14.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl15 = oSheet.get_Range("O3", "O3");
                cl15.Value2 = "Return";
                cl15.ColumnWidth = 13.0;
                rowHead = oSheet.get_Range("A3", "O3");
            }
            else if (model[0] == "LD20")
            {   //OQC
                Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value2 = "Date test";
                cl4.ColumnWidth = 20.0;
                Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value2 = "SwdF0(G)[Hz]";
                cl5.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value2 = "Sgl Current(ave)[mA]";
                cl6.ColumnWidth = 10.0;

                Microsoft.Office.Interop.Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value2 = "Sgl Grms(ave)[G]";
                cl7.ColumnWidth = 10.0;

                Microsoft.Office.Interop.Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                cl8.Value2 = "Sgl RiseTime(PctG2)[ms]";
                cl8.ColumnWidth = 7.0;
                //INLINE
                Microsoft.Office.Interop.Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                cl9.Value2 = "Sgl BrakeTime(PctG2)[ms]";
                cl9.ColumnWidth = 20.0;

                Microsoft.Office.Interop.Excel.Range cl10 = oSheet.get_Range("J3", "J3");
                cl10.Value2 = "Judge LINE";
                cl10.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl11 = oSheet.get_Range("K3", "K3");
                cl11.Value2 = "Bin Data";
                cl11.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl12 = oSheet.get_Range("L3", "L3");
                cl12.Value2 = "Return";
                cl12.ColumnWidth = 13.0;

                //Microsoft.Office.Interop.Excel.Range cl13 = oSheet.get_Range("M3", "M3");
                //cl13.Value2 = "ANR_CCW-Rated load speed";
                //cl13.ColumnWidth = 13.0;

                //Microsoft.Office.Interop.Excel.Range cl14 = oSheet.get_Range("N3", "N3");
                //cl14.Value2 = "AIS_CCW-Stall current";
                //cl14.ColumnWidth = 13.0;

                //Microsoft.Office.Interop.Excel.Range cl15 = oSheet.get_Range("O3", "O3");
                //cl15.Value2 = "Judge LINE";
                //cl15.ColumnWidth = 7.0;

                //Microsoft.Office.Interop.Excel.Range cl16 = oSheet.get_Range("P3", "P3");
                //cl16.Value2 = "Return";
                //cl16.ColumnWidth = 7.0;

                rowHead = oSheet.get_Range("A3", "P3");
            }
            else
            {
                //OQC
                Microsoft.Office.Interop.Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value2 = "Datetest NMT";
                cl4.ColumnWidth = 20.0;

                Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value2 = "CIO_CCW-Motor COMP No load current";
                cl5.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value2 = "CG_CCW-Motor COMP Vibration";
                cl6.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value2 = "CNO_CCW-Motor COMP No load speed";
                cl7.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                cl8.Value2 = "Judge OQC";
                cl8.ColumnWidth = 7.0;

                //INLINE
                Microsoft.Office.Interop.Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                cl9.Value2 = "Datetest NO41";
                cl9.ColumnWidth = 20.0;

                Microsoft.Office.Interop.Excel.Range cl10 = oSheet.get_Range("J3", "J3");
                cl10.Value2 = "AIO_CCW-No load current";
                cl10.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl11 = oSheet.get_Range("K3", "K3");
                cl11.Value2 = "ANO_CCW-No load speed";
                cl11.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl12 = oSheet.get_Range("L3", "L3");
                cl12.Value2 = "AIR_CCW-Rated load current";
                cl12.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl13 = oSheet.get_Range("M3", "M3");
                cl13.Value2 = "ANR_CCW-Rated load speed";
                cl13.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl14 = oSheet.get_Range("N3", "N3");
                cl14.Value2 = "AIS_CCW-Stall current";
                cl14.ColumnWidth = 13.0;

                Microsoft.Office.Interop.Excel.Range cl15 = oSheet.get_Range("O3", "O3");
                cl15.Value2 = "Judge LINE";
                cl15.ColumnWidth = 7.0;

                Microsoft.Office.Interop.Excel.Range cl16 = oSheet.get_Range("P3", "P3");
                cl16.Value2 = "Return";
                cl16.ColumnWidth = 7.0;

                rowHead = oSheet.get_Range("A3", "P3");
            }
            #region old
            //else
            //{
            //    Microsoft.Office.Interop.Excel.Range cl5 = oSheet.get_Range("D3", "D3");
            //    cl5.Value2 = "Current [mA]";
            //    cl5.ColumnWidth = 13.0;

            //    Microsoft.Office.Interop.Excel.Range cl6 = oSheet.get_Range("E3", "E3");
            //    cl6.Value2 = "Vibration /9.8[G]";
            //    cl6.ColumnWidth = 10.0;

            //    Microsoft.Office.Interop.Excel.Range cl7 = oSheet.get_Range("F3", "F3");
            //    cl7.Value2 = "Vibration [m/s2]";
            //    cl7.ColumnWidth = 10.0;

            //    Microsoft.Office.Interop.Excel.Range cl8 = oSheet.get_Range("G3", "G3");
            //    cl8.Value2 = "Vibration * 10";
            //    cl8.ColumnWidth = 13.0;

            //    Microsoft.Office.Interop.Excel.Range cl9 = oSheet.get_Range("H3", "H3");
            //    cl9.Value2 = "Frequency [Hz]";
            //    cl9.ColumnWidth = 13.0;

            //    Microsoft.Office.Interop.Excel.Range cl10 = oSheet.get_Range("I3", "I3");
            //    cl10.Value2 = "AIO_CW-No load current";
            //    cl10.ColumnWidth = 13.0;

            //    Microsoft.Office.Interop.Excel.Range cl11 = oSheet.get_Range("J3", "J3");
            //    cl11.Value2 = "ANO_CW-No load speed";
            //    cl11.ColumnWidth = 13.0;

            //    Microsoft.Office.Interop.Excel.Range cl12 = oSheet.get_Range("K3", "K3");
            //    cl12.Value2 = "AIR_CW-Rated load current";
            //    cl12.ColumnWidth = 13.0;

            //    Microsoft.Office.Interop.Excel.Range cl13 = oSheet.get_Range("L3", "L3");
            //    cl13.Value2 = "ANR_CW-Rated load speed";
            //    cl13.ColumnWidth = 13.0;

            //    Microsoft.Office.Interop.Excel.Range cl14 = oSheet.get_Range("M3", "M3");
            //    cl14.Value2 = "AIS_CW-Stall current";
            //    cl14.ColumnWidth = 13.0;

            //    Microsoft.Office.Interop.Excel.Range cl15 = oSheet.get_Range("N3", "N3");
            //    cl15.Value2 = "Judge";
            //    cl15.ColumnWidth = 7.0;

            //    Microsoft.Office.Interop.Excel.Range cl16 = oSheet.get_Range("O3", "O3");
            //    cl16.Value2 = "Return";
            //    cl16.ColumnWidth = 7.0;

            //    rowHead = oSheet.get_Range("A3", "O3");
            //}
            #endregion

            rowHead.WrapText = true;
            rowHead.Font.Bold = true;

            // Kẻ viền

            rowHead.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

            // Thiết lập màu nền

            rowHead.Interior.ColorIndex = 15;

            // Tạo mẳng đối tượng để lưu trữ toàn bộ dữ liệu trong DataTable,

            // vì dữ liệu được được gán vào các Cell trong Excel phải thông qua object thuần.

            //object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];

            //Chuyển dữ liệu từ DataTable vào mảng đối tượng

            //for (int r = 0; r < dt.Rows.Count; r++)

            //{

            //    DataRow dr = dt.Rows[r];

            //    for (int c = 0; c < dt.Columns.Count; c++)

            //    {
            //        arr[r, c] = dr[c];
            //    }
            //}

            //Thiết lập vùng điền dữ liệu

            int rowStart = 4;

            int columnStart = 1;

            int rowEnd = rowStart + dgv.Rows.Count - 1;
            int columnEnd;

            if (model[0] == "0025")
            {
                columnEnd = 15;
            }
            else
            {
                columnEnd = 16;
            }
            //if (model[0] == "517EB")
            //{
            //    columnEnd = 16;
            //}
            //else columnEnd = 13;

            // Ô bắt đầu điền dữ liệu

            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowStart, columnStart];

            // Ô kết thúc điền dữ liệu

            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnEnd];

            // Lấy về vùng điền dữ liệu

            Microsoft.Office.Interop.Excel.Range range = oSheet.get_Range(c1, c2);
            range.WrapText = true;

            //Điền dữ liệu vào vùng đã thiết lập
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                if (model[0] == "517EB")
                {
                    oExcel.Cells[i + rowStart, 1] = "'" + dgv["col_serial_no", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 2] = dgv["col_model", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 3] = dgv["col_lot", i].Value.ToString();
                    //OQC
                    oExcel.Cells[i + rowStart, 4] = dgv["col_date", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 5] = dgv["col_cir_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 6] = dgv["col_cg_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 7] = dgv["col_cnr_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 8] = dgv["col_judge_oqc", i].Value.ToString();
                    //INLINE
                    oExcel.Cells[i + rowStart, 9] = dgv["col_date_line", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 10] = dgv["col_aio_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 11] = dgv["col_ano_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 12] = dgv["col_air_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 13] = dgv["col_anr_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 14] = dgv["col_ais_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 15] = dgv["col_judge_oqc", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 16] = dgv["col_return", i].Value.ToString();
                }
                else if (model[0] == "0025")
                {
                    oExcel.Cells[i + rowStart, 1] = "'" + dgv["col_serial_no", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 2] = dgv["col_model", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 3] = dgv["col_lot", i].Value.ToString();
                    //OQC
                    oExcel.Cells[i + rowStart, 4] = dgv["col_date", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 5] = dgv["col_qacurrent", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 6] = dgv["col_qafg", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 7] = dgv["col_qaspeed", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 8] = dgv["col_judge_qa", i].Value.ToString();
                    //INLINE
                    oExcel.Cells[i + rowStart, 9] = dgv["col_date_line", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 10] = dgv["col_current", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 11] = dgv["col_fg", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 12] = dgv["col_speed", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 13] = dgv["col_judge_inline", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 14] = dgv["col_svfi", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 15] = dgv["col_return", i].Value.ToString();
                }
                else if (model[0] == "0148")
                {
                    oExcel.Cells[i + rowStart, 1] = "'" + dgv["col_serial_no", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 2] = dgv["col_model", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 3] = dgv["col_lot", i].Value.ToString();
                    //OQC
                    oExcel.Cells[i + rowStart, 4] = dgv["col_date", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 5] = dgv["col_cir_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 6] = dgv["col_cg_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 7] = dgv["col_cnr_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 8] = dgv["col_judge_oqc", i].Value.ToString();
                    //INLINE
                    oExcel.Cells[i + rowStart, 9] = dgv["col_date_line", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 10] = dgv["col_aio_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 11] = dgv["col_ano_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 12] = dgv["col_air_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 13] = dgv["col_anr_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 14] = dgv["col_ais_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 15] = dgv["col_judge_oqc", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 16] = dgv["col_return", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 17] = dgv["col_terminal", i].Value.ToString();
                }
                else if (model[0] == "LD20")
                {
                    oExcel.Cells[i + rowStart, 1] = "'" + dgv["col_serial_no", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 2] = dgv["col_model", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 3] = dgv["col_lot", i].Value.ToString();
                    //Inline
                    oExcel.Cells[i + rowStart, 4] = dgv["col_date_line", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 5] = dgv["col_sdf0_inline", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 6] = dgv["col_scurave_inline", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 7] = dgv["col_sgrmsave_inline", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 8] = dgv["col_srtpctg2_inline", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 9] = dgv["col_sbtpctg2_inline", i].Value.ToString();
                    //INLINE
                    oExcel.Cells[i + rowStart, 10] = dgv["col_judge_inline", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 11] = dgv["bin", i].Value.ToString();
                    //oExcel.Cells[i + rowStart, 11] = dgv["col_ano_ccw", i].Value.ToString();
                    //oExcel.Cells[i + rowStart, 12] = dgv["col_air_ccw", i].Value.ToString();
                    //oExcel.Cells[i + rowStart, 13] = dgv["col_anr_ccw", i].Value.ToString();
                    //oExcel.Cells[i + rowStart, 14] = dgv["col_ais_ccw", i].Value.ToString();
                    //oExcel.Cells[i + rowStart, 15] = dgv["col_judge_oqc", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 16] = dgv["col_return", i].Value.ToString();
                }
                else //if (model[0] == "517DB" || model[0] == "517DC")
                {
                    oExcel.Cells[i + rowStart, 1] = "'" + dgv["col_serial_no", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 2] = dgv["col_model", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 3] = dgv["col_lot", i].Value.ToString();
                    //OQC
                    oExcel.Cells[i + rowStart, 4] = dgv["col_date", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 5] = dgv["col_cio_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 6] = dgv["col_cg_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 7] = dgv["col_cno_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 8] = dgv["col_judge_oqc", i].Value.ToString();
                    //ININE
                    oExcel.Cells[i + rowStart, 9] = dgv["col_date_line", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 10] = dgv["col_aio_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 11] = dgv["col_ano_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 12] = dgv["col_air_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 13] = dgv["col_anr_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 14] = dgv["col_ais_ccw", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 15] = dgv["col_judge_oqc", i].Value.ToString();
                    oExcel.Cells[i + rowStart, 16] = dgv["col_return", i].Value.ToString();
                }

                //else
                //{
                //    oExcel.Cells[i + rowStart, 1] = dgv["col_serial_no", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 2] = dgv["col_model", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 3] = dgv["col_lot", i].Value.ToString();

                //    oExcel.Cells[i + rowStart, 4] = dgv["col_current", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 5] = dgv["col_vib_g", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 6] = dgv["col_vib_ms2", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 7] = dgv["col_vibration", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 8] = dgv["col_freq", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 9] = dgv["col_aio_cw", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 10] = dgv["col_ano_cw", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 11] = dgv["col_air_cw", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 12] = dgv["col_anr_cw", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 13] = dgv["col_ais_cw", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 14] = dgv["col_judge", i].Value.ToString();
                //    oExcel.Cells[i + rowStart, 15] = dgv["col_return", i].Value.ToString();
                //}
            }

            // Kẻ viền

            range.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

            rowHead.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            rowHead.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Căn giữa cột STT

            Microsoft.Office.Interop.Excel.Range c3 = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[rowEnd, columnStart];

            Microsoft.Office.Interop.Excel.Range c4 = oSheet.get_Range(c1, c3);

            oSheet.get_Range(c3, c4).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheets);
        }
    }
}
