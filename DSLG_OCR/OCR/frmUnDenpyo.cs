using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using DSLG_OCR.Common;
using DSLG_OCR.CSV;

namespace DSLG_OCR.OCR
{
    public partial class frmUnDenpyo : Form
    {
        public frmUnDenpyo()
        {
            InitializeComponent();
        }

        DSLGDataSet dts = new DSLGDataSet();
        DSLGDataSetTableAdapters.修正確認データTableAdapter adp = new DSLGDataSetTableAdapters.修正確認データTableAdapter();
        
        // データインデックス
        int cI = 0;

        private void frmUnDenpyo_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データ読み込み
            adp.Fill(dts.修正確認データ);

            // データテーブル件数カウント
            if (!dts.修正確認データ.Any())
            {
                MessageBox.Show("対象となる未照合データがありません", "未照合データ確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                Environment.Exit(0);
            }

            // 画面表示ズーム倍率初期化
            global.miMdlZoomRate = 0;
            
            // データ表示
            cI = 0;
            dataShow(cI);
        }

        private void dataShow(int iX)
        {
            // フォーム初期化
            formInitialize(iX);

            // データ表示
            DSLGDataSet.修正確認データRow r = (DSLGDataSet.修正確認データRow)dts.修正確認データ.Rows[iX];
            txtDenNum.Text = r.伝票番号.ToString();

            switch (r.照合ステータス)
            {
                case global.STATUS_DENOVERLAP:
                    lblErrMsg.Text = "データ内で伝票№が重複しています";
                    break;

                case global.STATUS_PASTOVERLAP:
                    lblErrMsg.Text = "過去データと重複しています";
                    break;

                case global.STATUS_UNFIND:
                    lblErrMsg.Text = "配車データ未登録です";
                    break;

                case global.STATUS_NG:
                    lblErrMsg.Text = "ＮＧ：ＯＣＲ認識で既定の書式と認識されませんでした。";
                    break;

                default:
                    break;
            }

            // 画像表示
            ShowImage(r.画像名);
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        /// ------------------------------------------------------------------------------
        public void ShowImage(string tempImgName)
        {
            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                lblNoImage.Visible = false;
                global.pblImagePath = string.Empty;
                return;
            }

            //画像ファイルがあるとき表示
            if (File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;

                //画像ロード
                Leadtools.Codecs.RasterCodecs.Startup();
                Leadtools.Codecs.RasterCodecs cs = new Leadtools.Codecs.RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                Leadtools.RasterPaintProperties prop = new Leadtools.RasterPaintProperties();
                prop = Leadtools.RasterPaintProperties.Default;
                prop.PaintDisplayMode = Leadtools.RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(tempImgName, 0, Leadtools.Codecs.CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    //leadImg.ScaleFactor *= global.ZOOM_RATE;
                    leadImg.ScaleFactor *= Properties.Settings.Default.zoomRate;
                }
                else
                {
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = Leadtools.WinForms.RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                Leadtools.ImageProcessing.GrayscaleCommand grayScaleCommand = new Leadtools.ImageProcessing.GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                // 画像の右端に合わせてスクロール
                leadImg.ScrollPosition = new Point(leadImg.Width, 0);

                cs.Dispose();
                Leadtools.Codecs.RasterCodecs.Shutdown();
                //global.pblImagePath = tempImgName;
            }
            else
            {
                //画像ファイルがないとき
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;

                leadImg.Visible = false;
                //global.pblImagePath = string.Empty;
            }
        }

        /// ----------------------------------------------------------------------
        /// <summary>
        ///     表示中のデータを更新します </summary>
        /// <param name="iX">
        ///     レコードインデックス</param>
        /// ----------------------------------------------------------------------
        private void curDataUpDate(int iX)
        {
            DSLGDataSetTableAdapters.未照合伝票TableAdapter cAdp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            cAdp.Fill(dts.未照合伝票);

            DSLGDataSet.修正確認データRow r = (DSLGDataSet.修正確認データRow)dts.修正確認データ.Rows[iX];

            var s = dts.未照合伝票.Single(a => a.ID == r.ID);
            s.伝票番号 = Utility.StrtoInt(txtDenNum.Text);
            cAdp.Update(dts.未照合伝票);

            // 修正確認データ再読み込み
            adp.Fill(dts.修正確認データ);
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            curDataUpDate(cI);

            //レコードの移動
            cI = 0;
            dataShow(cI);
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            curDataUpDate(cI);

            //レコードの移動
            if (cI > 0)
            {
                cI--;
                dataShow(cI);
            }   
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            curDataUpDate(cI);

            //レコードの移動
            if (cI + 1 < dts.修正確認データ.Rows.Count)
            {
                cI++;
                dataShow(cI);
            }   
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            curDataUpDate(cI);

            //レコードの移動
            cI = dts.修正確認データ.Rows.Count - 1;
            dataShow(cI);
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //カレントデータの更新
            curDataUpDate(cI);

            //レコードの移動
            cI = hScrollBar1.Value;
            dataShow(cI);
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     フォーム表示初期化 </summary>
        /// <param name="cIx">
        ///     カレントレコードインデックス</param>
        ///------------------------------------------------------------------------------------
        private void formInitialize(int cIx)
        {
            txtDenNum.MaxLength = 6;

            // テキストボックス表示色設定

            lblNoImage.Visible = false;
            
            // スクロールバー設定
            hScrollBar1.Enabled = true;
            hScrollBar1.Minimum = 0;
            hScrollBar1.Maximum = dts.修正確認データ.Count - 1;
            hScrollBar1.Value = cIx;
            hScrollBar1.LargeChange = 1;
            hScrollBar1.SmallChange = 1;

            //移動ボタン制御
            btnFirst.Enabled = true;
            btnNext.Enabled = true;
            btnBefore.Enabled = true;
            btnEnd.Enabled = true;

            //最初のレコード
            if (cIx == 0)
            {
                btnBefore.Enabled = false;
                btnFirst.Enabled = false;
            }

            //最終レコード
            if ((cIx + 1) == dts.修正確認データ.Count)
            {
                btnNext.Enabled = false;
                btnEnd.Enabled = false;
            }

            //データ数表示
            lblPage.Text = " (" + (cI + 1).ToString() + "/" + dts.修正確認データ.Rows.Count.ToString() + ")";
        }

        private void txtDenNum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmUnDenpyo_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dts.修正確認データ.Any())
            {
                curDataUpDate(cI);
            }

            this.Dispose();
        }

        private void frmUnDenpyo_Shown(object sender, EventArgs e)
        {
            button1.Focus();
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void leadImg_ScrollPositionChanged(object sender, EventArgs e)
        {
            //int sp = leadImg.ScrollPosition.X;
            //MessageBox.Show(sp.ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // 表示中データ更新
            curDataUpDate(cI);

            // 消込画面表示
            frmMnkeshikomi frm = new frmMnkeshikomi(txtDenNum.Text);
            frm.ShowDialog();

            if (frm.mDate == string.Empty)
            {
                frm.Dispose();
                return;
            }

            // 手動消込（未照合ＯＫ）
            manualKeshikomi(cI, frm.mDenNum, frm.mDate, frm.mMaker);

            // 伝票番号を過去データに追加
            clsMakeCsvfile c = new clsMakeCsvfile(this);
            c.addPastData(frm.mDenNum);

            frm.Dispose();

            // 未照合データ再読み込み
            adp.Fill(dts.修正確認データ);

            // データテーブル件数カウント
            if (!dts.修正確認データ.Any())
            {
                MessageBox.Show("該当する未照合データがありません", "未照合データ確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                this.Close();
            }
            else
            {
                // データ表示
                if (cI > dts.修正確認データ.Rows.Count - 1)
                {
                    cI = dts.修正確認データ.Rows.Count - 1;
                }

                dataShow(cI);
            }
        }

        /// ------------------------------------------------------------
        /// <summary>
        ///     手動消込 </summary>
        /// <param name="iX">
        ///     レコードインデックス</param>
        /// ------------------------------------------------------------
        private void manualKeshikomi(int iX, int mDen, string mDate, string mMaker)
        {
            // データセット
            DSLGDataSetTableAdapters.未照合伝票TableAdapter adp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            adp.Fill(dts.未照合伝票);

            DSLGDataSet.修正確認データRow r = (DSLGDataSet.修正確認データRow)dts.修正確認データ.Rows[iX];

            // 画像を未照合OKフォルダへ移動する
            // 画像名を変更します 
            string newImgNm = global.cnfUnmOkImgPath + mDate.Replace("/", "") + mMaker + "_" + r.伝票番号.ToString() + ".tif";
            System.IO.File.Move(r.画像名, newImgNm);

            // 伝票番号テーブルの照合ステータス更新
            DSLGDataSet.未照合伝票Row d = dts.未照合伝票.Single(a => a.ID == r.ID);
            d.伝票番号 = mDen;
            d.日付 = DateTime.Parse(mDate);
            d.メーカー名 = mMaker;
            d.照合ステータス = global.STATUS_UNVERIOK;
            d.更新年月日 = DateTime.Now;
            d.画像名 = newImgNm;

            // データ更新
            adp.Update(dts.未照合伝票);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の物品受領書の伝票№を削除します。よろしいですか？", "消込確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 伝票番号データ削除
            dataDelete(cI);

            // 未照合データ再読み込み
            adp.Fill(dts.修正確認データ);

            // データテーブル件数カウント
            if (!dts.修正確認データ.Any())
            {
                MessageBox.Show("全ての未照合データが削除されました", "未照合データ確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                this.Close();
            }
            else
            {
                // データ表示
                if (cI > (dts.修正確認データ.Rows.Count - 1))
                {
                    cI = dts.修正確認データ.Rows.Count - 1;
                }

                dataShow(cI);
            }
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     伝票番号データ削除 </summary>
        /// <param name="iX">
        ///     レコードインデックス</param>
        /// ---------------------------------------------------------------
        private void dataDelete(int iX)
        {
            // データセット
            DSLGDataSetTableAdapters.未照合伝票TableAdapter adp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            adp.Fill(dts.未照合伝票);

            DSLGDataSet.修正確認データRow r = (DSLGDataSet.修正確認データRow)dts.修正確認データ.Rows[iX];

            // 画像を削除します
            System.IO.File.Delete(r.画像名);

            // 伝票番号データを削除します
            DSLGDataSet.未照合伝票Row d = dts.未照合伝票.Single(a => a.ID == r.ID);
            d.Delete();

            // データ更新
            adp.Update(dts.未照合伝票);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("未照合伝票の再照合処理を行います。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 表示中データ更新
            curDataUpDate(cI);

            // 再照合
            getUnVerifi();
        }

        /// --------------------------------------------------------------
        /// <summary>
        ///     再照合処理 </summary>
        /// --------------------------------------------------------------
        private void getUnVerifi()
        {
            // 再照合
            int n = unVerifi();

            string msg = string.Empty;

            if (n == 0)
            {
                msg = "照合された伝票はありませんでした";
            }
            else
            {
                msg = n.ToString() + "件の伝票が照合されました";
            }

            // 確認メッセージ
            MessageBox.Show(msg, "再照合結果", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // ログ出力
            msg += " " + (DateTime.Now.ToShortDateString()) + " " + DateTime.Now.ToLongTimeString();
            Utility.logOutput(msg, "再照合");

            // 修正確認データ再読み込み
            adp.Fill(dts.修正確認データ);

            // データテーブル件数カウント
            if (!dts.修正確認データ.Any())
            {
                MessageBox.Show("該当する未照合データがありません", "未照合データ確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                this.Close();
            }
            else
            {
                // データ表示
                if (cI > (dts.修正確認データ.Rows.Count - 1))
                {
                    cI = dts.修正確認データ.Rows.Count - 1;
                }

                dataShow(cI);
            }
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     未照合伝票再照合 </summary>
        /// <returns>
        ///     照合件数</returns>
        /// ---------------------------------------------------------------
        private int unVerifi()
        {
            // カーソルを待機にする
            this.Cursor = Cursors.WaitCursor;

            clsMakeCsvfile c = new clsMakeCsvfile(this);
            //c.getHaishaCsv();       // 配車データロード
            c.importHaishaCsv();    // 配車CSVインポート

            c.unStatusToUnveri();   // 未照合伝票の照合ステータスを未処理に書き換える（※未照合OK以外）
            c.findDenOverlapUn();   // 伝票№重複チェック
            c.findPastDataUn();     // 過去データ照合（重複チェック）
            c.findHaishaDataUn();   // 配車データ照合
            c.pastDataUpdateUn();   // 再照合済みデータで過去データ更新
            int n = c.haishaDataUpdateUn(); // 配車データ更新

            // カーソルを戻す
            this.Cursor = Cursors.Default;

            return n;
        }
    }
}
