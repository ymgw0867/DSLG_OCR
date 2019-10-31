using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DSLG_OCR.Common;

namespace DSLG_OCR.CSV
{
    public partial class frmDenNum : Form
    {
        public frmDenNum()
        {
            InitializeComponent();
        }

        private void frmDenNum_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビュー
            GridViewSetting(dg1);
            //GridViewSetting2(dg2);

            // 画像表示
            lblNoImage.Visible = false;

            // 画面表示ズーム倍率初期化
            global.miMdlZoomRate = 0;

            // 日付検索
            dateTimePicker1.Checked = false;

            // 配車データ読み込み
            adp.Fill(dts.配車);

            // 未照合伝票データ読み込み
            mAdp.Fill(dts.未照合伝票);

            // コンボボックス
            cmbMakerLoad();
            cmbKekkaLoad();
            cmbDenpyo();

            // 件数
            lblCnt.Visible = false;

            // 取消リンク
            linkLabel2.Visible = false;

            // 画像拡大・縮小ボタン
            btnMinus.Enabled = false;
            btnPlus.Enabled = false;

            // 画像印刷ボタン
            btnImgPrn.Enabled = false;
        }

        string colDate = "col1";
        string colMaker = "col2";
        string colDenNum = "col3";
        string colStatus = "col4";
        string colImg = "col5";
        string colID = "col6";

        DSLGDataSet dts = new DSLGDataSet();
        DSLGDataSetTableAdapters.配車TableAdapter adp = new DSLGDataSetTableAdapters.配車TableAdapter();
        DSLGDataSetTableAdapters.未照合伝票TableAdapter mAdp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();

        const string MIOK = "未照合ＯＫ";
        const string SHOK = "照合済";

        /// --------------------------------------------------------
        /// <summary>
        ///     メーカーコンボボックス　</summary>
        /// --------------------------------------------------------
        private void cmbMakerLoad()
        {
            comboBox1.Items.Add("全て");

            var s = dts.配車.Select(a => a.メーカー名).Distinct();

            foreach (var t in s)
            {
                comboBox1.Items.Add(t);
            }

            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.SelectedIndex = 0;
        }

        /// --------------------------------------------------------
        /// <summary>
        ///     結果コンボボックス　</summary>
        /// --------------------------------------------------------
        private void cmbKekkaLoad()
        {
            comboBox2.Items.Add("全て");
            comboBox2.Items.Add("配車・未照合");
            comboBox2.Items.Add("配車・照合済");
            comboBox2.Items.Add("未照合ＯＫ");
            comboBox2.Items.Add("伝票番号重複");
            comboBox2.Items.Add("過去伝票登録あり");
            comboBox2.Items.Add("配車データ未登録");
            comboBox2.Items.Add("ＮＧ");
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.SelectedIndex = 0;
        }

        /// --------------------------------------------------------
        /// <summary>
        ///     伝票コンボボックス　</summary>
        /// --------------------------------------------------------
        private void cmbDenpyo()
        {
            comboBox3.Items.Add("全て");
            comboBox3.Items.Add("配車");
            comboBox3.Items.Add("ＯＣＲ読取伝票");
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.SelectedIndex = 0;
        }

        /// ---------------------------------------------------------------------------
        /// <summary>
        ///     配車データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// ---------------------------------------------------------------------------
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更するe

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 22;
                tempDGV.RowTemplate.Height = 22;

                // 全体の高さ
                //tempDGV.Height = 637;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // カラム定義
                tempDGV.Columns.Add(colDate, "日付");
                tempDGV.Columns.Add(colMaker, "メーカー");
                tempDGV.Columns.Add(colDenNum, "伝票番号");
                tempDGV.Columns.Add(colStatus, "結果");
                tempDGV.Columns.Add(colImg, "画像名");
                tempDGV.Columns.Add(colID, "ID");

                tempDGV.Columns[colImg].Visible = false;
                tempDGV.Columns[colID].Visible = false;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                // 表示位置
                tempDGV.Columns[colDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colDenNum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colStatus].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;

                ////ソート機能制限
                //for (int i = 0; i < tempDGV.Columns.Count; i++)
                //{
                //    // Alignment
                //    if (i == 0 || i == 3)
                //    {
                //        tempDGV.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                //    }

                //    // ソート機能制限
                //    //tempDGV.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //}

                // 各列幅指定
                tempDGV.Columns[colDate].Width = 100;
                tempDGV.Columns[colDenNum].Width = 80;
                tempDGV.Columns[colStatus].Width = 120;
                tempDGV.Columns[colMaker].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                // 編集可否
                tempDGV.ReadOnly = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // データ検索・表示
            dataSerch();
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     データ検索・表示 </summary>
        /// ---------------------------------------------------------------
        private void dataSerch()
        {
            string dt = string.Empty;

            if (dateTimePicker1.Checked)
            {
                dt = dateTimePicker1.Value.ToShortDateString();
            }

            // 
            int cmb3 = 0;
            if (comboBox3.SelectedIndex > 0)
            {
                cmb3 = comboBox3.SelectedIndex;
            }

            // メーカー指定
            string cmb = string.Empty;
            if (comboBox1.SelectedIndex > 0)
            {
                cmb = comboBox1.Text;
            }

            // 結果
            string cmb2 = string.Empty;
            if (comboBox2.SelectedIndex > 0)
            {
                cmb2 = (comboBox2.SelectedIndex - 1).ToString();
            }

            // 配車データ表示
            showHaishaData(dg1, cmb3, dt, cmb, Utility.StrtoInt(txtDen.Text), cmb2);
        }

        /// ---------------------------------------------------------------------------
        /// <summary>
        ///     配車データ表示 </summary>
        /// <param name="g">
        ///     DataGridViewオブジェクト </param>
        /// <param name="dt">
        ///     指定日付</param>
        /// <param name="mk">
        ///     指定メーカー</param>
        /// <param name="dNum">
        ///     指定伝票番号</param>
        /// ---------------------------------------------------------------------------
        private void showHaishaData(DataGridView g, int sKbn, string dt, string mk, int dNum, string kekka)
        {
            try
            {
                // カーソルを待機状態にする
                this.Cursor = Cursors.WaitCursor;

                // データ抽出
                var s = dts.配車
                    .Select(a => new
                    {
                        照合区分 = 1,
                        a.日付,
                        a.メーカー名,
                        a.伝票番号,
                        a.照合ステータス,
                        a.画像名,
                        a.ID
                    }).Union(dts.未照合伝票.Select(a => new
                    {
                        照合区分 = 2,
                        a.日付,
                        a.メーカー名,
                        a.伝票番号,
                        a.照合ステータス,
                        a.画像名,
                        a.ID
                    }))
                    .OrderByDescending(a => a.日付)
                    .ThenBy(a => a.メーカー名)
                    .ThenBy(a => a.伝票番号);

                // 日付指定
                if (sKbn != 0)
                {
                    s = s.Where(a => a.照合区分 == sKbn)
                                            .OrderByDescending(a => a.日付)
                                            .ThenBy(a => a.メーカー名)
                                            .ThenBy(a => a.伝票番号);
                }

                // 日付指定
                if (dt != string.Empty)
                {
                    s = s.Where(a => a.日付 == DateTime.Parse(dt))
                                            .OrderByDescending(a => a.日付)
                                            .ThenBy(a => a.メーカー名)
                                            .ThenBy(a => a.伝票番号);
                }

                // メーカー指定
                if (mk != string.Empty)
                {
                    s = s.Where(a => a.メーカー名 == mk).OrderByDescending(a => a.日付)
                                           .ThenBy(a => a.メーカー名)
                                           .ThenBy(a => a.伝票番号);
                }

                // 結果指定
                if (kekka != string.Empty)
                {
                    s = s.Where(a => a.照合ステータス == Utility.StrtoInt(kekka))
                                            .OrderByDescending(a => a.日付)
                                            .ThenBy(a => a.メーカー名)
                                            .ThenBy(a => a.伝票番号);
                }

                // 伝票番号指定
                if (dNum != global.flgOff)
                {
                    s = s.Where(a => a.伝票番号 == dNum).OrderByDescending(a => a.日付)
                                           .ThenBy(a => a.メーカー名)
                                           .ThenBy(a => a.伝票番号);
                }

                g.Rows.Clear();
                int i = 0;

                foreach (var t in s)
                {
                    g.Rows.Add();
                    g[colDate, i].Value = t.日付.ToShortDateString();
                    g[colMaker, i].Value = t.メーカー名;
                    g[colDenNum, i].Value = t.伝票番号;

                    if (t.画像名 == null)
                    {
                        g[colImg, i].Value = string.Empty;
                    }
                    else
                    {
                        g[colImg, i].Value = t.画像名;
                    }

                    g[colID, i].Value = t.ID;

                    // 照合ステータス
                    switch (t.照合ステータス)
                    {
                        case global.STATUS_VERIFI:
                            g[colStatus, i].Value = SHOK;
                            g.Rows[i].DefaultCellStyle.ForeColor = Color.Black;
                            break;

                        case global.STATUS_UNVERIOK:
                            g[colStatus, i].Value = MIOK;
                            g.Rows[i].DefaultCellStyle.ForeColor = Color.Black;
                            break;

                        case global.STATUS_UNVERI: // 実際には未処理のステータスだが配車データの未処理ステータス(0)で使用
                            g[colStatus, i].Value = "未照合";
                            g.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                            break;

                        case global.STATUS_DENOVERLAP:
                            g[colStatus, i].Value = "伝票番号重複";
                            g.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                            break;

                        case global.STATUS_PASTOVERLAP:
                            g[colStatus, i].Value = "過去伝票番号あり";
                            g.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                            break;

                        case global.STATUS_UNFIND:
                            g[colStatus, i].Value = "配車データ未登録";
                            g.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                            break;

                        case global.STATUS_NG:
                            g[colStatus, i].Value = "ＮＧ";
                            g.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                            break;

                        default:
                            break;
                    }
                    i++;
                }

                g.CurrentCell = null;

                // 画像クリア
                leadImg.Visible = false;
                lblNoImage.Visible = false;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                // カーソルを元に戻す
                this.Cursor = Cursors.Default;
            }

            // 配車情報がないとき
            if (g.RowCount == 0)
            {
                lblCnt.Visible = false;
                MessageBox.Show("該当するデータが存在しませんでした", "検索結果", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                lblCnt.Visible = true;
                lblCnt.Text = "(" + g.RowCount.ToString("#,0") + "件)";
            }
        }

        private void dg1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //string imgPath = dg1[colImg, dg1.SelectedRows[0].Index].Value.ToString();
            //ShowImage(imgPath);
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
                lblNoImage.Visible = true;

                // 画像拡大・縮小ボタン
                btnMinus.Enabled = false;
                btnPlus.Enabled = false;
                btnImgPrn.Enabled = false;

                return;
            }

            //画像ファイルがあるとき表示
            if (System.IO.File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;
                btnImgPrn.Enabled = true;

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
                btnImgPrn.Enabled = false;
                //global.pblImagePath = string.Empty;
            }
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     検索条件リセット </summary>
        /// ---------------------------------------------------------
        private void serReset()
        {
            dateTimePicker1.Checked = false;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            txtDen.Text = string.Empty;
        }

        private void dg1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dg1[e.ColumnIndex, e.RowIndex].Value != null)
            {
                string imgPath = dg1[colImg, dg1.SelectedRows[0].Index].Value.ToString();
                ShowImage(imgPath);
            
                // 取消リンクをオフ
                linkLabel2.Visible = false;

                // 未照合ＯＫ伝票のとき取消可能とする
                if (dg1[colStatus, e.RowIndex].Value.ToString() == MIOK)
                {
                    linkLabel2.Visible = true;
                }
                else if (dg1[colStatus, e.RowIndex].Value.ToString() == SHOK)
                {
                    linkLabel2.Visible = true;
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("検索条件を初期状態に戻します。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 検索条件リセット
            serReset();
        }

        private void txtDen_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string msg = "伝票番号：" + dg1[colDenNum, dg1.SelectedRows[0].Index].Value + " の" + dg1[colStatus, dg1.SelectedRows[0].Index].Value + "を取り消します。" + Environment.NewLine + "よろしいですか";
            if (MessageBox.Show(msg, "取消確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question,MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // ＩＤ取得
            int sID = Utility.StrtoInt(dg1[colID, dg1.SelectedRows[0].Index].Value.ToString());

            // 伝票番号
            int sDen = Utility.StrtoInt(dg1[colDenNum, dg1.SelectedRows[0].Index].Value.ToString());
            
            // 日付
            DateTime sDate = DateTime.Parse(dg1[colDate, dg1.SelectedRows[0].Index].Value.ToString());

            // 画像名取得
            string sImgNm = dg1[colImg, dg1.SelectedRows[0].Index].Value.ToString();

            // 取消処理
            switch (dg1[colStatus, dg1.SelectedRows[0].Index].Value.ToString())
            {
                case MIOK:  // 未照合ＯＫ
                    unmOkCancel(sID, sImgNm);
                    break;

                case SHOK:  // 照合済み
                    veriCancel(sID, sDen, sDate, sImgNm);
                    break;

                default:
                    break;
            }

            // データ再表示
            dataSerch();
        }

        /// --------------------------------------------------------------------------
        /// <summary>
        ///     未照合ＯＫ取消 </summary>
        /// <param name="sID">
        ///     ID</param>
        /// <param name="sImgNm">
        ///     画像名</param>
        /// --------------------------------------------------------------------------
        private void unmOkCancel(int sID, string sImgNm)
        {
            mAdp.Fill(dts.未照合伝票);
            DSLGDataSet.未照合伝票Row r = dts.未照合伝票.Single(a => a.ID == sID);
                                    
            // 画像移動
            string newImgNm = global.cnfUnmImgPath + System.IO.Path.GetFileName(sImgNm);
            System.IO.File.Move(sImgNm, newImgNm);
            
            // 未照合伝票書き換え
            int sDen = r.伝票番号;
            r.メーカー名 = string.Empty;
            r.画像名 = newImgNm;
            r.照合ステータス = global.STATUS_UNFIND;
            r.更新年月日 = DateTime.Now;

            // データベース更新
            mAdp.Update(dts.未照合伝票);

            // 該当伝票を過去データから削除する
            clsMakeCsvfile c = new clsMakeCsvfile(this);
            c.pastDataCancel(sDen);
        }

        /// --------------------------------------------------------------------------
        /// <summary>
        ///     照合済取消 </summary>
        /// <param name="sID">
        ///     ID</param>
        /// <param name="sDen">
        ///     伝票番号</param>
        /// <param name="sDate">
        ///     日付</param>
        /// <param name="sImgNm">
        ///     画像名</param>
        /// --------------------------------------------------------------------------
        private void veriCancel(int sID, int sDen, DateTime sDate, string sImgNm)
        {
            // 配車データ書き換え
            adp.Fill(dts.配車);
            DSLGDataSet.配車Row r = dts.配車.Single(a => a.ID == sID);

            // 値書き換え
            r.画像名 = string.Empty;
            r.照合ステータス = global.flgOff;
            r.更新年月日 = DateTime.Now;

            CSV.clsMakeCsvfile c = new clsMakeCsvfile(this);

            // 未照合画像連番取得
            int unNum = c.getUnNumber(sDate) + 1;

            // 画像移動
            //C:\DSLG_OCR\TIF\20150510ABC商事_397377.tif
            //C:\DSLG_OCR\UNMIMG\20150512UN0024_395767.tif

            string newImgNm = global.cnfUnmImgPath + sDate.ToShortDateString().Replace("/", "") + global.UNMARK + unNum.ToString().PadLeft(4, '0') + "_" + sDen.ToString() + ".tif";  
            System.IO.File.Move(sImgNm, newImgNm);

            // 未処理連番テーブル更新
            c.setUnNumber(sDate, unNum);

            // 未照合伝票に新規登録
            mAdp.Fill(dts.未照合伝票);
            DSLGDataSet.未照合伝票Row m = dts.未照合伝票.New未照合伝票Row();
            m.伝票番号 = sDen;
            m.メーカー名 = string.Empty;
            m.日付 = sDate;
            m.画像名 = newImgNm;
            m.照合ステータス = global.STATUS_UNFIND;

            dts.未照合伝票.Add未照合伝票Row(m);
            
            // データベース更新
            adp.Update(dts.配車);
            mAdp.Update(dts.未照合伝票);

            // 該当伝票を過去データから削除する
            c.pastDataCancel(sDen);
        }

        private void btnImgPrn_Click(object sender, EventArgs e)
        {
            //印刷確認
            if (!leadImg.Visible) return;
            if (MessageBox.Show("この伝票画像を印刷します。よろしいですか？", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            //画像印刷
            cPrint prn = new cPrint();
            prn.Image(leadImg);
        }
    }
}
