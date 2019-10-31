using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DSLG_OCR.CSV;
using DSLG_OCR.Common;
using DSLG_OCR.OCR;

namespace DSLG_OCR
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 環境設定情報取得
            DSLG_OCR.Config.getConfig c = new Config.getConfig();

            // 配車データパス
            lblHdataPath.Text = global.cnfHaishaPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // MDB最適化
            mdbCompact();

            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 配車データ存在確認
            if (!System.IO.File.Exists(global.cnfHaishaPath))
            {
                string msg = "配車データ " + global.cnfHaishaPath + "が存在しません。" + Environment.NewLine + "メインメニューより配車データの登録を行ってください。";
                MessageBox.Show(msg, "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            this.Hide();

            // 日付入力
            DSLG_OCR.OCR.frmDate frm = new OCR.frmDate();
            frm.ShowDialog();

            if (frm.MyProperty == global.flgOn)
            {
                // OCR認識処理
                getOcrData(this, Properties.Settings.Default.imgInPath, Properties.Settings.Default.imgOutPath, Properties.Settings.Default.ngPath, Properties.Settings.Default.frmPath, false);

                // 照合
                dataVerifi();
            }

            frm.Dispose();

            this.Show();
        }

        private void getOcrData(Form frm, string inPath, string outPath, string ngPath, string fmtPath, bool bFax)
        {
            OcrPV6.Class1 ocr = new OcrPV6.Class1(frm, inPath, outPath, ngPath, fmtPath, false);
            clsMakeCsvfile cs = new clsMakeCsvfile(this);
            int m = cs.getCSVFile();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Config.frmConfig frm = new Config.frmConfig();
            frm.ShowDialog();
            this.Show();

            // 環境設定情報取得
            DSLG_OCR.Config.getConfig c = new Config.getConfig();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 照合
            dataVerifi();
        }

        private void dataVerifi()
        {
            // 件数確認
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            adp.Fill(dts.伝票番号);
            int g = dts.伝票番号.Count(a => a.照合ステータス == global.STATUS_UNVERI);

            //if (g == 0)
            //{
            //    MessageBox.Show("照合する物品受領書データがありません。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}

            // 照合処理
            clsMakeCsvfile c = new clsMakeCsvfile(this);
            //c.getHaishaCsv();             // 配車データロード
            c.importHaishaCsv();            // 配車CSVインポート
            int v = c.findDenOverlap();     // 伝票№重複チェック
            int p = c.findPastData();       // 過去データ照合（重複チェック）
            int j = c.findHaishaData();     // 配車データ照合
            c.pastDataUpdate();             // 照合済みデータで過去データ更新
            int ok = c.haishaDataUpdate();  // 配車データ更新
            int un = c.unmDataUpdate();     // 未照合伝票テーブル更新
            c.ngToUnmData();                // NGデータ未照合伝票テーブル更新

            // 終了メッセージ
            StringBuilder sb = new StringBuilder();
            sb.Append("物品受領書の伝票番号の配車データ照合処理が終了しました。 ");
            sb.Append(DateTime.Now.ToShortDateString()).Append(" ");
            sb.Append(DateTime.Now.ToLongTimeString());
            sb.Append(Environment.NewLine + Environment.NewLine);
            sb.Append("伝票件数：" + g.ToString() + "件").Append(Environment.NewLine);
            sb.Append("照合完了：" + ok.ToString() + "件").Append(Environment.NewLine);
            sb.Append("照合未完了").Append(Environment.NewLine);
            sb.Append(">伝票番号重複：" + v.ToString() + "件").Append(Environment.NewLine);
            sb.Append(">過去データ登録済：" + p.ToString() + "件").Append(Environment.NewLine);
            sb.Append(">配車データ未登録：" + j.ToString() + "件").Append(Environment.NewLine + Environment.NewLine);
            sb.Append("照合未完了の伝票は修正確認画面で確認してください。");

            MessageBox.Show(sb.ToString(), "照合結果", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // ログ出力
            Utility.logOutput(sb.ToString(), "OCR照合");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string rtn = openHaishaData();

            if (rtn != string.Empty)
            {
                lblHdataPath.Text = rtn;
            }
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     配車データ選択 </summary>
        /// <returns>
        ///     パスを含むファイル名</returns>
        /// ---------------------------------------------------------------
        private string openHaishaData()
        {
            string result = string.Empty;

            //ダイアログボックスの初期設定
            openFileDialog1.Title = "配車データ選択";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "CSVファイル(*.csv)|*.csv|テキストファイル(*.txt)|*.txt";

            //ダイアログボックスの表示
            DialogResult ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return result;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "Excelアサイン確認書取り込み", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return result;
            }

            // 環境設定テーブル更新
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.環境設定TableAdapter adp = new DSLGDataSetTableAdapters.環境設定TableAdapter();
            adp.Fill(dts.環境設定);

            var s = dts.環境設定.Single(a => a.ID == global.configKEY);
            s.配車データパス = openFileDialog1.FileName;
            s.更新年月日 = DateTime.Now;
            adp.Update(dts.環境設定);
            
            // 環境設定情報取得
            DSLG_OCR.Config.getConfig c = new Config.getConfig();

            // ファイルのパスを返す
            return openFileDialog1.FileName;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmUnDenpyo frm = new frmUnDenpyo();
            frm.ShowDialog();
            this.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmDenNum frm = new frmDenNum();
            frm.ShowDialog();
            this.Show();
        }


        /// ---------------------------------------------------------------------
        /// <summary>
        ///     MDBファイルを最適化する </summary>
        /// ---------------------------------------------------------------------
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();
                string OldDb = Properties.Settings.Default.mdbOlePath;
                string NewDb = Properties.Settings.Default.mdbPathTemp;

                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.mdbPath + global.MDBBACK);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.mdbPath + global.MDBFILE, Properties.Settings.Default.mdbPath + global.MDBBACK);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.mdbPath + global.MDBTEMP, Properties.Settings.Default.mdbPath + global.MDBFILE);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }
        
    }
}
