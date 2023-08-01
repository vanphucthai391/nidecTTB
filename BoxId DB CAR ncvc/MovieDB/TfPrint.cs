using System;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Data;

namespace BoxIdDb
{
    public class TfPrint
    {
        public void createBoxidFiles(string dir, string id, string model, 
            DataGridView dgv1, ref DataGridView dgv2, ref TextBox txt, string shipKind)
        {
            //boxId 名のテキストファイルを生成
            string file = dir + id + ".txt"; 
            FileInfo fi = new FileInfo(file);
            using (FileStream fs = fi.Create())
            {
                fs.Close();
            }

            //boxId 名のビットマップファイルを生成（①まとめ用）
            Bitmap bmp1 = new Bitmap(80 * 7 - 30, 50 * 5);
            Graphics gr = Graphics.FromImage(bmp1);
            gr.Clear(Color.White);

            //テキストボックスのビットマップを生成、その１（②サブ－１）
            txt.Text = "Box: " + id;// + "   Model: " + model;
            txt.Width = 80 * 7 - 30;
            Bitmap bmp2_1 = new Bitmap(txt.Width, txt.Height);
            txt.DrawToBitmap(bmp2_1, new Rectangle(0, 0, txt.Width, txt.Height));
            gr.DrawImage(bmp2_1, 0, 10);

            //テキストボックスのビットマップを生成、その２（②サブ－２）
            txt.Text = "Model: " + model;
            Bitmap bmp2_2 = new Bitmap(txt.Width, txt.Height);
            txt.DrawToBitmap(bmp2_2, new Rectangle(0, 0, txt.Width, txt.Height));
            gr.DrawImage(bmp2_2, 0, 10 + txt.Height);

            //データグリットビュー用のビットマップを生成（③サブ）
            adjustDummyDatagridview(dgv1, ref dgv2);
            Bitmap bmp3 = new Bitmap(80 * 7 - 30, 50 * 5);
            dgv2.DrawToBitmap(bmp3, new Rectangle(0, 60, dgv2.Width, 50 * 6));

            txt.Text = "Judge: PASS";
            Bitmap bmp2_3 = new Bitmap(txt.Width, txt.Height);
            txt.DrawToBitmap(bmp2_3, new Rectangle(0, 0, txt.Width, txt.Height));
            gr.DrawImage(bmp2_3, 0, txt.Height + 140);

            //ビットマップのコピー＆ペースト
            for (int i = 0; i < 2; i++)
            {
                Rectangle destRect = new Rectangle(0, txt.Height + 60 * i, 80 * 7, 50 * 6);
                Rectangle srcRect = new Rectangle(80 * 7 * i, 0, 80 * 7, 50 * 6);
                gr.DrawImage(bmp3, destRect, srcRect, GraphicsUnit.Pixel);
            }

            //８ビット形式に変換し、ファイルに保存する
            Bitmap bmp4 = TfBitmap.CopyToBpp(bmp1, 8);
            file = dir + id + ".bmp";
            bmp4.Save(file, ImageFormat.Bmp);
        }

        //ダミーデータグリットビューの幅調整
        private int adjustDummyDatagridview(DataGridView dgv1, ref DataGridView dgv2)
        {
            DataTable dt = ((DataTable)dgv1.DataSource).Copy();
            dgv2.DataSource = dt;

            int k = dgv2.Columns.Count;
            dgv2.Width = 115 * k;
            dgv2.Height = 31 * k;
            for (int i = 0; i < k; i++)
            {
                dgv2.Columns[i].Width = 90;
            }
            return k;
        }

        // 外箱用のプリント（ファイル経由でなく、直接プリンターにコマンドを送信する）
        public void printBigBarcode(string apn,string qpn, string vpn, string desc, string qty, 
            string carton, string stage, string shaft, string overlay, string rtn, DataTable dt)
        {
            long res;
            int x, y, BarHeight;
            string printerName = "SEWOO Label Printer";

            /* 1. LK_OpenPrinter() */
            if (LKBPRINT.LK_OpenPrinter(printerName) != LKBPRINT.LK_SUCCESS) { return; }

            /* 2. LK_SetupPrinter() */
            res = LKBPRINT.LK_SetupPrinter("101", 	// 10~104 (Unit is mm) label Width // small 70, big 101
                "201", 		// 5~350 (Unit is mm) Label Lengsh  //  small 30, big 201
                0,				// 0=Label with Gap, 1=Label with Black Mark, 2=Label with Continuous.
                "3",			// if(MediaType==0) <GapHeight> else <BlackMarkHeight>. (Unit is mm) // small 3, big 7
                "0",			// if(MediaType==0) <not used> else <distance from BlackMark to perforation>. (Unit is mm)
                8,				// 0 ~ 15
                6,				// 2 ~ 6 (Unit is Inch)
                1				// 1 ~ 9999 copies
                );

            System.Diagnostics.Debug.Print(res.ToString());
            if (res != LKBPRINT.LK_SUCCESS) { LKBPRINT.LK_ClosePrinter(); return; }

            /* 3-1. page 1 test */
            LKBPRINT.LK_StartPage();
            BarHeight = 5 * 8;	// 12mm

            x = 15 * 8;
            y = (2 + 67) * 8;
            LKBPRINT.LK_PrintWindowsFont(x, y, 0, 40, 1, 0, 0, "Arial", "Nidec Copal(Vietnam) Co.,Ltd.");

            x = 15 * 8;
            y = (10 + 69) * 8;
            LKBPRINT.LK_PrintDeviceFont(x, y, 0, 4, 1, 1, 0, "APN:");
            x = 35 * 8;
            y = (10 + 69) * 8;
            LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, apn);

            x = 15 * 8;
            y = (20 + 69) * 8;
            LKBPRINT.LK_PrintDeviceFont(x, y, 0, 4, 1, 1, 0, "QPN:");
            x = 35 * 8;
            y = (20 + 69) * 8;
            LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, qpn);

            x = 15 * 8;
            y = (30 + 69) * 8;
            LKBPRINT.LK_PrintDeviceFont(x, y, 0, 4, 1, 1, 0, "Vender PN:");
            x = 35 * 8;
            y = (30 + 69) * 8;
            LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, vpn);

            x = 15 * 8;
            y = (40 + 69) * 8;
            LKBPRINT.LK_PrintDeviceFont(x, y, 0, 4, 1, 1, 0, "Desc:");
            x = 35 * 8;
            y = (40 + 69) * 8;
            LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, desc);

            x = 15 * 8;
            y = (50 + 69) * 8;
            LKBPRINT.LK_PrintDeviceFont(x, y, 0, 4, 1, 1, 0, "QTY:");
            x = 35 * 8;
            y = (50 + 69) * 8;
            LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, qty);

            x = 15 * 8;
            y = (63 + 69) * 8;
            LKBPRINT.LK_PrintDeviceFont(x, y, 0, 4, 1, 1, 0, "Country of origin:  Vietnam");

            x = 15 * 8;
            y = (70 + 69) * 8;
            LKBPRINT.LK_PrintDeviceFont(x, y, 0, 4, 1, 1, 0, "L/C:");

            for (int i = 0; i < dt.Rows.Count -1; i++)
            {
                if (i == 0)
                {
                    x = 15 * 8;
                    y = (75 + 69) * 8;
                    LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, dt.Rows[i]["lot"].ToString());                
                }
                else if (i == 1)
                {
                    x = 45 * 8;
                    y = (75 + 69) * 8;
                    LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, dt.Rows[i]["lot"].ToString());                
                }
                else if (i == 2)
                {
                    x = 75 * 8;
                    y = (75 + 69) * 8;
                    LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, dt.Rows[i]["lot"].ToString());
                }
                else if (i == 3)
                {
                    x = 15 * 8;
                    y = (85 + 69) * 8;
                    LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, dt.Rows[i]["lot"].ToString());         
                }
                else if (i == 4)
                {
                    x = 45 * 8;
                    y = (85 + 69) * 8;
                    LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, dt.Rows[i]["lot"].ToString());            
                }
                else if (i == 5)
                {
                    x = 75 * 8;
                    y = (85 + 69) * 8;
                    LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, dt.Rows[i]["lot"].ToString());
                }
            }

            x = 15 * 8;
            y = (100 + 69) * 8;
            LKBPRINT.LK_PrintDeviceFont(x, y, 0, 4, 1, 1, 0, "Carton ID:");
            x = 35 * 8;
            y = (100 + 69) * 8;
            LKBPRINT.LK_PrintBarCode(x, y, 0, "1A", 2, 4, BarHeight, 1, carton);

            x = 10 * 8;
            y = (114 + 69) * 8;
            LKBPRINT.LK_PrintWindowsFont(x, y, 0, 40, 1, 0, 0, "Arial", "Stage: " + stage);
            x = 45 * 8;
            y = (114 + 69) * 8;
            string shipKind = (rtn == "N" ? "New" : (rtn == "R" ? "Re-Screen" : "Error"));
            LKBPRINT.LK_PrintWindowsFont(x, y, 0, 40, 1, 0, 0, "Arial", "Ship Kind: " + shipKind);

            x = 10 * 8;
            y = (122 + 69) * 8;
            string shaftShow = (shaft == "null" ? string.Empty : "Shaft: " + shaft);
            LKBPRINT.LK_PrintWindowsFont(x, y, 0, 40, 1, 0, 0, "Arial", shaftShow);
            x = 45 * 8;
            y = (122 + 69) * 8;
            string overlayShow = (overlay == "null" ? string.Empty : "Overlay : " + overlay);
            LKBPRINT.LK_PrintWindowsFont(x, y, 0, 40, 1, 0, 0, "Arial", overlayShow);
            
            LKBPRINT.LK_EndPage();

            /* 4. LK_ClosePrinter() */
            LKBPRINT.LK_ClosePrinter();
        }
    }
}
