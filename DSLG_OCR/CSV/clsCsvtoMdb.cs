using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DSLG_OCR.Common;

namespace DSLG_OCR.CSV
{
    class clsMakeCsvfile
    {
        public clsMakeCsvfile(Form owner)
        {
            _Owner = owner;
        }

        Form _Owner;

        // チェック用基幹コード配列
        string[] mainCode;

        // 配車クラス
        clsHaisha[] h;

        // 読込枚数
        int denCnt = 0;

        ///-------------------------------------------------------------------
        /// <summary>
        ///     配車データ(CSV)を取得して配列にセットする　</summary>
        ///-------------------------------------------------------------------
        public bool getHaishaCsv()
        {
            // 配車データパス
            string cFile = global.cnfHaishaPath;

            if (System.IO.File.Exists(cFile))
            {
                int i = 0;
                foreach (var t in System.IO.File.ReadAllLines(cFile, Encoding.Default))
                {
                    string [] hCsv = t.Split(',');
                    
                    Array.Resize(ref h, i + 1);
                    h[i] = new clsHaisha(); 
                    h[i].hDate = hCsv[0];
                    h[i].hMaker = hCsv[1];
                    h[i].hDenNum = hCsv[2];

                    i++;
                }
                return true;
            }
            else
            {
                // 基幹マスターが存在しないとき
                string msg = "配車データ：" + cFile + "が存在しません。" + Environment.NewLine + Environment.NewLine +
                             "データを選択後、再実行してください。";

                MessageBox.Show(msg, "配車データ未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     配車データ(CSV)を配車テーブルにインポートする　</summary>
        ///-------------------------------------------------------------------
        public bool importHaishaCsv()
        {
            // 配車テーブル読み込み
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.配車TableAdapter adpHt = new DSLGDataSetTableAdapters.配車TableAdapter();
            adpHt.Fill(dts.配車);

            // 配車データパス
            string cFile = global.cnfHaishaPath;
            
            // 日付
            DateTime dt = DateTime.Today;

            if (System.IO.File.Exists(cFile))
            {
                foreach (var t in System.IO.File.ReadAllLines(cFile, Encoding.Default))
                {
                    // カンマごとに分割し配列にセット
                    string[] hCsv = t.Split(',');

                    // ダブルコーテーションを除去
                    hCsv[0] = hCsv[0].Replace(@"""", string.Empty); // 日付
                    hCsv[1] = hCsv[1].Replace(@"""", string.Empty); // メーカー名
                    hCsv[2] = hCsv[2].Replace(@"""", string.Empty); // 伝票番号

                    string cDt = string.Empty;

                    // CSVデータの内容をチェック
                    if (!csvCheck(hCsv, out cDt))
                    {
                        continue;   // エラーのとき読み飛ばし
                    }

                    // 配車データに追加登録
                    if (DateTime.TryParse(cDt, out dt))
                    {
                        if (!dts.配車.Any(a => a.日付 == dt && a.メーカー名 == hCsv[1] &&
                                              a.伝票番号 == Utility.StrtoInt(hCsv[2])))
                        {
                            // 配車テーブルに追加登録
                            DSLGDataSet.配車Row hr = dts.配車.New配車Row();
                            hr.日付 = dt;
                            hr.メーカー名 = hCsv[1];
                            hr.伝票番号 = Utility.StrtoInt(hCsv[2]);
                            hr.照合ステータス = global.flgOff;
                            hr.画像名 = "";
                            hr.更新年月日 = DateTime.Now;
                            dts.配車.Add配車Row(hr);
                        }
                    }
                }

                // データベース更新
                adpHt.Update(dts.配車);

                return true;
            }
            else
            {
                // 配車CSVが存在しないとき
                string msg = "配車データ：" + cFile + "が存在しません。" + Environment.NewLine + Environment.NewLine +
                             "データを選択後、再実行してください。";

                MessageBox.Show(msg, "配車データ未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     CSV配車データの内容をチェック </summary>
        /// <param name="hCsv">
        ///     CSV配車データ配列</param>
        /// <param name="dt">
        ///     日付</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// ------------------------------------------------------------------------
        private bool csvCheck(string[] hCsv, out string dt)
        {
            dt = string.Empty;

            // 日付が８桁でない (yyyymmddになっていない）
            if (hCsv[0].Length != 8)
            {
                return false;
            }

            // 正しい日付か？
            DateTime cDt;
            string ymd = hCsv[0].Substring(0, 4) + "/" + hCsv[0].Substring(4, 2) + "/" + hCsv[0].Substring(6, 2);
            if (DateTime.TryParse(ymd, out cDt))
            {
                dt = cDt.ToShortDateString();
            }
            else
            {
                return false;
            }

            // 伝票番号が数字か？
            if (!Utility.NumericCheck(hCsv[2]))
            {
                return false;
            }
            
            return true;
        }


        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     物品受領書・伝票番号をMDBへ読み込み </summary> 
        /// -----------------------------------------------------------------------------
        public int getCSVFile()
        {
            // 対象CSVファイル数を取得
            string _inPath = Properties.Settings.Default.imgOutPath;
            denCnt = System.IO.Directory.GetFiles(_inPath, "*.csv").Count();

            // 読込件数
            int dCnt = 0;

            // ＣＳＶファイルがなければ終了
            if (denCnt == 0) return 0;

            // オーナーフォームを無効にする
            _Owner.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = _Owner;
            frmP.Show();

            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            
            // テーブルアダプタ
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            
            // テーブルアダプタに読み取りデータを読み込む
            adp.Fill(dts.伝票番号);

            try
            {
                // CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    // 件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "物品受領書画像ロード中　" + cCnt.ToString() + "/" + denCnt.ToString();
                    frmP.progressValue = cCnt * 100 / denCnt;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    int denNum = 0;
                    string imgName = string.Empty;

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // 伝票番号
                        if (stCSV[2] != string.Empty)
                        {
                            denNum = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[2], 6));
                        }
                        else if (stCSV[3] != string.Empty)
                        {
                            denNum = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[3], 6));
                        }
                        else
                        {
                            continue;
                        }

                        // 画像名
                        imgName = Utility.GetStringSubMax(stCSV[1].Trim(), 21);

                        // 読込件数
                        dCnt++;

                        // データセットに読み取りデータを追加する
                        dts.伝票番号.Add伝票番号Row(denNum, imgName, DateTime.Today, "", global.STATUS_UNVERI, DateTime.Now);              
                    }
                }

                // データベースへ反映
                adp.Update(dts);

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }

                // いったんオーナーをアクティブにする
                _Owner.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                _Owner.Enabled = true;

                // 戻り値
                return dCnt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "物品受領書伝票№読み込み処理", MessageBoxButtons.OK);
                return dCnt;
            }
            finally
            {
                
            }
        }

        /// ------------------------------------------------------------
        /// <summary>
        ///     伝票番号重複チェック </summary>
        /// <returns>
        ///     件数</returns>
        /// ------------------------------------------------------------
        public int findDenOverlap()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            adp.Fill(dts.伝票番号);

            // 結果件数
            int dNum = 0;

            // 未処理伝票が存在するか
            if (dts.伝票番号.Any(a => a.照合ステータス == global.STATUS_UNVERI))
            {
                // 重複している伝票番号を抽出
                var fff = dts.伝票番号.Where(a => a.照合ステータス == global.STATUS_UNVERI)
                                    .GroupBy(a => a.伝票番号)
                                    .Where(a => a.Count() > 1)
                                    .Select(a => a.Key)
                                    .ToArray();

                if (fff.Length > 0)
                {
                    // OCR認識日付取得
                    DateTime dt = getOcrDate();
                    string sDt = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0');

                    // 未照合画像連番取得
                    int unNum = getUnNumber(dt);

                    for (int i = 0; i < fff.Length; i++)
                    {
                        // 伝票番号テーブルの照合ステータス更新
                        foreach (var t in dts.伝票番号.Where(a => a.伝票番号 == Utility.StrtoInt(fff[i].ToString()) && a.照合ステータス == global.STATUS_UNVERI))
                        {
                            // 画像を未処理フォルダへ移動する
                            unNum++;
                            string newImgNm = global.cnfUnmImgPath + sDt + global.UNMARK + unNum.ToString().PadLeft(4, '0') + "_" + fff[i].ToString() + ".tif";
                            System.IO.File.Move(Properties.Settings.Default.imgOutPath + t.画像名, newImgNm);

                            // 伝票番号テーブル書き換え
                            t.日付 = DateTime.Parse(dt.ToShortDateString());
                            t.画像名 = newImgNm;
                            t.照合ステータス = global.STATUS_DENOVERLAP;
                            t.更新年月日 = DateTime.Now;

                            dNum++;
                        } 
                    }

                    // データベース更新
                    adp.Update(dts.伝票番号);

                    // 未処理連番テーブル更新
                    setUnNumber(dt, unNum);
                }
            }

            // 後片付け
            adp.Dispose();

            // 件数を返す
            return dNum;
        }

        /// ------------------------------------------------------------
        /// <summary>
        ///     未照合伝票・伝票番号重複チェック </summary>
        /// ------------------------------------------------------------
        public void findDenOverlapUn()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合伝票TableAdapter adp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            adp.Fill(dts.未照合伝票);
                        
            // 未処理伝票が対象
            if (dts.未照合伝票.Any(a => a.照合ステータス == global.STATUS_UNVERI))
            {
                // 重複している伝票番号を抽出
                var fff = dts.未照合伝票.Where(a => a.照合ステータス == global.STATUS_UNVERI)
                                        .GroupBy(a => a.伝票番号)
                                        .Where(a => a.Count() > 1)
                                        .Select(a => a.Key)
                                        .ToArray();

                if (fff.Length > 0)
                {
                    for (int i = 0; i < fff.Length; i++)
                    {
                        // 伝票番号テーブルの照合ステータス更新
                        foreach (var t in dts.未照合伝票.Where(a => a.伝票番号 == Utility.StrtoInt(fff[i].ToString()) && 
                                                                    a.照合ステータス == global.STATUS_UNVERI))
                        {
                            // 照合ステータス書き換え
                            t.照合ステータス = global.STATUS_DENOVERLAP;
                            t.更新年月日 = DateTime.Now;
                        }
                    }

                    // データベース更新
                    adp.Update(dts.未照合伝票);
                }
            }

            // 後片付け
            adp.Dispose();
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     OCR認識日付取得 </summary>
        /// <returns>
        ///     OCR認識日付</returns>
        /// ---------------------------------------------------------
        private DateTime getOcrDate()
        {
            DateTime result = DateTime.Today;

            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.環境設定TableAdapter adp = new DSLGDataSetTableAdapters.環境設定TableAdapter();
            adp.Fill(dts.環境設定);
            
            foreach (var r in dts.環境設定.Where(a => a.ID == global.configKEY))
	        {
                result = r.OCR認識日付;
	        }
            
            return result;
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     未照合画像日別連番取得 </summary>
        /// <param name="dt">
        ///     日付</param>
        /// <returns>
        ///     該当日付の連番</returns>
        /// ---------------------------------------------------------------
        public int getUnNumber(DateTime dt)
        {
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合連番TableAdapter adp = new DSLGDataSetTableAdapters.未照合連番TableAdapter();
            adp.Fill(dts.未照合連番);

            int result = 0;
            foreach (var t in dts.未照合連番.Where(a => a.日付 == dt))
            {
                result = t.連番; 
            } ;

            return result;
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     未照合連番テーブル更新 </summary>
        /// <param name="dt">
        ///     日付</param>
        /// <returns>
        ///     該当日付の連番</returns>
        /// ---------------------------------------------------------------
        public void setUnNumber(DateTime dt, int sNum)
        {
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合連番TableAdapter adp = new DSLGDataSetTableAdapters.未照合連番TableAdapter();
            adp.Fill(dts.未照合連番);

            if (dts.未照合連番.Any(a => a.日付 == dt))
            {
                // 更新
                DSLGDataSet.未照合連番Row r = dts.未照合連番.Single(a => a.日付 == dt);
                r.連番 = sNum;
                r.更新年月日 = DateTime.Now;
            }
            else
            {
                // 新規登録
                DSLGDataSet.未照合連番Row r = dts.未照合連番.New未照合連番Row();
                r.日付 = dt;
                r.連番 = sNum;
                r.更新年月日 = DateTime.Now;

                dts.未照合連番.Add未照合連番Row(r);
            }

            adp.Update(dts.未照合連番);
        }
        
        /// ------------------------------------------------------------
        /// <summary>
        ///     過去データ突合</summary>
        /// <returns>
        ///     件数</returns>
        /// ------------------------------------------------------------
        public int findPastData()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            DSLGDataSetTableAdapters.過去データTableAdapter padp = new DSLGDataSetTableAdapters.過去データTableAdapter();

            adp.Fill(dts.伝票番号);
            padp.Fill(dts.過去データ);

            // 結果件数
            int dNum = 0;

            // OCR認識日付取得
            DateTime dt = getOcrDate();
            string sDt = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0');

            // 未照合画像連番取得
            int unNum = getUnNumber(dt);
            int cNum = unNum;

            // 未照合伝票を順次読む
            if (dts.伝票番号.Any(a => a.照合ステータス == global.STATUS_UNVERI))
            {
                foreach (var t in dts.伝票番号.Where(a => a.照合ステータス == global.STATUS_UNVERI).OrderBy(a => a.ID))
                {
                    // 過去データを検索
                    if (dts.過去データ.Any(a => a.伝票番号 == t.伝票番号))
                    {
                        // 画像を未処理フォルダへ移動する
                        cNum++;
                        string newImgNm = global.cnfUnmImgPath + sDt + global.UNMARK + cNum.ToString().PadLeft(4, '0') + "_" + t.伝票番号.ToString() + ".tif";
                        System.IO.File.Move(Properties.Settings.Default.imgOutPath + t.画像名, newImgNm);

                        // 伝票番号テーブルの照合ステータス更新
                        DSLGDataSet.伝票番号Row d = dts.伝票番号.Single(a => a.ID == t.ID);
                        d.日付 = DateTime.Parse(dt.ToShortDateString());
                        d.照合ステータス = global.STATUS_PASTOVERLAP;
                        d.更新年月日 = DateTime.Now;
                        d.画像名 = newImgNm;

                        dNum++;
                    }
                }

                // 過去データとの重複があったとき
                if (unNum != cNum)
                {
                    // データ更新
                    adp.Update(dts.伝票番号);

                    // 未処理連番テーブル更新
                    setUnNumber(dt, cNum);
                }
            }

            // 後片付け
            adp.Dispose();
            padp.Dispose();

            // 件数を返す
            return dNum;
        }

        /// ------------------------------------------------------------
        /// <summary>
        ///     未照合伝票・過去データ突合</summary>
        ///     
        ///     対象ステータス：未処理、ＮＧ
        /// ------------------------------------------------------------
        public void findPastDataUn()
        {
            bool un = true;

            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合伝票TableAdapter adp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            DSLGDataSetTableAdapters.過去データTableAdapter padp = new DSLGDataSetTableAdapters.過去データTableAdapter();

            adp.Fill(dts.未照合伝票);
            padp.Fill(dts.過去データ);

            // 未照合伝票を順次読む
            if (dts.未照合伝票.Any(a => a.照合ステータス == global.STATUS_UNVERI || 
                                        a.照合ステータス == global.STATUS_NG))
            {
                foreach (var t in dts.未照合伝票.Where(a => a.照合ステータス == global.STATUS_UNVERI || 
                                                           a.照合ステータス == global.STATUS_NG).OrderBy(a => a.ID))
                {
                    // 有効伝票番号のとき
                    if (t.伝票番号 > 0)
                    {
                        // 過去データを検索
                        if (dts.過去データ.Any(a => a.伝票番号 == t.伝票番号))
                        {
                            // 照合ステータス更新
                            DSLGDataSet.未照合伝票Row d = dts.未照合伝票.Single(a => a.ID == t.ID);
                            d.照合ステータス = global.STATUS_PASTOVERLAP;
                            d.更新年月日 = DateTime.Now;

                            un = false;
                        }
                    }
                }

                // 過去データとの重複があったとき
                if (!un)
                {
                    // データ更新
                    adp.Update(dts.未照合伝票);
                }
            }

            // 後片付け
            adp.Dispose();
            padp.Dispose();
        }
        
        /// ------------------------------------------------------------
        /// <summary>
        ///     配車データ照合 </summary>
        /// ------------------------------------------------------------
        public int findHaishaData()
        {
            // 新画像名
            string newImgNm = string.Empty;

            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            DSLGDataSetTableAdapters.配車TableAdapter adpH = new DSLGDataSetTableAdapters.配車TableAdapter();
            adp.Fill(dts.伝票番号);
            adpH.Fill(dts.配車);

            // 未照合件数
            int dNum = 0;

            // OCR認識日付取得
            DateTime dt = getOcrDate();
            string sDt = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0');

            // 未照合画像連番取得
            int unNum = getUnNumber(dt);
            int cNum = unNum;

            // 伝票データを順次読む
            if (dts.伝票番号.Any(a => a.照合ステータス == global.STATUS_UNVERI))
            {
                foreach (var t in dts.伝票番号.Where(a => a.照合ステータス == global.STATUS_UNVERI).OrderBy(a => a.ID))
                {
                    bool md = false;
                    string hDate = string.Empty;
                    string hMaker = string.Empty;

                    // 配車データを検索 2016/08/03 照合条件に日付を追加
                    // 配車データを検索 2016/08/10 照合条件の日付はＯＣＲ認識日付に変更
                    foreach (var it in dts.配車.Where(a => a.伝票番号 == t.伝票番号 && a.日付 == DateTime.Parse(dt.ToShortDateString())))
                    {
                        md = true;
                        hDate = it.日付.ToShortDateString();      // 日付
                        hMaker = it.メーカー名;                   // メーカー名
                        break;
                    }

                    DSLGDataSet.伝票番号Row d = dts.伝票番号.Single(a => a.ID == t.ID);

                    if (md)
                    {
                        // 配車データに該当あり

                        // 画像を処理済フォルダへ移動する
                        newImgNm = global.cnfTifPath + hDate.Replace("/", "") + hMaker + "_" + t.伝票番号.ToString() + ".tif";

                        // 再度読み込んだとき等、登録済みのときは削除する
                        if (System.IO.File.Exists(newImgNm))
                        {
                            System.IO.File.Delete(newImgNm);
                        }

                        System.IO.File.Move(Properties.Settings.Default.imgOutPath + t.画像名, newImgNm);

                        // 伝票番号テーブル情報更新
                        d.日付 = DateTime.Parse(hDate);
                        d.メーカー名 = hMaker;
                        d.照合ステータス = global.STATUS_VERIFI;
                        d.画像名 = newImgNm;
                        d.更新年月日 = DateTime.Now;
                    }
                    else
                    {
                        // 画像を未処理フォルダへ移動する
                        cNum++;
                        newImgNm = global.cnfUnmImgPath + sDt + global.UNMARK + cNum.ToString().PadLeft(4, '0') + "_" + t.伝票番号.ToString() + ".tif";
                        System.IO.File.Move(Properties.Settings.Default.imgOutPath + t.画像名, newImgNm);

                        // 伝票番号テーブルの照合ステータス更新
                        d.日付 = DateTime.Parse(dt.ToShortDateString());
                        d.照合ステータス = global.STATUS_UNFIND;
                        d.画像名 = newImgNm;
                        d.更新年月日 = DateTime.Now;

                        dNum++;
                    }
                }

                // データ更新
                adp.Update(dts.伝票番号);

                // 配車データ未登録があったとき
                if (unNum != cNum)
                {
                    // 未処理連番テーブル更新
                    setUnNumber(dt, cNum);
                }
            }

            // 後片付け
            adp.Dispose();

            // 未処理件数を返す
            return dNum;
        }

        /// ------------------------------------------------------------
        /// <summary>
        ///     未照合伝票・配車データ照合 </summary>
        ///     
        ///     対象ステータス：未処理、ＮＧ
        ///     
        ///     2016/04/05 再照合条件に日付を追加
        /// ------------------------------------------------------------
        public void findHaishaDataUn()
        {
            // 新画像名
            string newImgNm = string.Empty;

            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合伝票TableAdapter adp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            DSLGDataSetTableAdapters.配車TableAdapter adpH = new DSLGDataSetTableAdapters.配車TableAdapter();
            adp.Fill(dts.未照合伝票);
            adpH.Fill(dts.配車);

            // 未照合伝票を順次読む
            if (dts.未照合伝票.Any(a => a.照合ステータス == global.STATUS_UNVERI || 
                                       a.照合ステータス == global.STATUS_NG))
            {
                foreach (var t in dts.未照合伝票.Where(a => a.照合ステータス == global.STATUS_UNVERI ||
                                                            a.照合ステータス == global.STATUS_NG).OrderBy(a => a.ID))
                {
                    bool md = false;
                    string hDate = string.Empty;
                    string hMaker = string.Empty;

                    // 有効な伝票番号のとき
                    if (t.伝票番号 > 0)
                    {
                        // 配車データを検索 2016/04/05 再照合条件に日付を追加
                        foreach (var it in dts.配車.Where(a => a.伝票番号 == t.伝票番号 && a.日付 == t.日付))
                        {
                            md = true;
                            hDate = it.日付.ToShortDateString();  // 日付
                            hMaker = it.メーカー名;              // メーカー名
                        }

                        DSLGDataSet.未照合伝票Row d = dts.未照合伝票.Single(a => a.ID == t.ID);

                        if (md)
                        {
                            // 配車データに該当あり

                            // パスを含む新画像名
                            newImgNm = global.cnfTifPath + hDate.Replace("/", "") + hMaker + "_" + t.伝票番号.ToString() + ".tif";

                            // 再度読み込んだとき等、登録済みのときは削除する
                            if (System.IO.File.Exists(newImgNm))
                            {
                                System.IO.File.Delete(newImgNm);
                            }

                            // 画像を処理済フォルダへ移動する
                            System.IO.File.Move(t.画像名, newImgNm);

                            // 伝票番号テーブル情報更新
                            d.日付 = DateTime.Parse(hDate);
                            d.メーカー名 = hMaker;
                            d.照合ステータス = global.STATUS_VERIFI;
                            d.画像名 = newImgNm;
                            d.更新年月日 = DateTime.Now;
                        }
                        else
                        {
                            // 伝票番号テーブルの照合ステータス更新
                            d.照合ステータス = global.STATUS_UNFIND;
                            d.更新年月日 = DateTime.Now;
                        }
                    }
                }

                // データ更新
                adp.Update(dts.未照合伝票);
            }

            // 後片付け
            adp.Dispose();
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVデータファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string sPath, string backPath, string [] arrayData)
        {
            //// ファイル名
            //string outFileName = sPath + ".csv";

            //// タイムスタンプ付加文字列
            //string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

            //// 出力ファイルが存在するとき
            //if (System.IO.File.Exists(outFileName))
            //{
            //    // 更新年月日を取得
            //    DateTime dt = System.IO.File.GetLastWriteTime(outFileName);

            //    // 更新年月日のタイムスタンプ付加文字列
            //    string moveFileName = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') +
            //                         dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') +
            //                         dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');

            //    // ファイルリネーム
            //    System.IO.File.Move(outFileName, sPath + "_" + moveFileName + ".csv");
            //}

            //// csvファイル出力
            //System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding(932));

            //// バックアップフォルダがなければ作成
            //if (!System.IO.Directory.Exists(Properties.Settings.Default.backupPath))
            //{
            //    System.IO.Directory.CreateDirectory(Properties.Settings.Default.backupPath);
            //}

            //// バックアップファイルを出力
            //// ファイル名文字列
            //string bkFileName = backPath + "_" + newFileName + ".csv";
            //System.IO.File.WriteAllLines(bkFileName, arrayData, System.Text.Encoding.GetEncoding(932));
        }

        /// ----------------------------------------------------------
        /// <summary>
        ///     照合済み伝票データを過去データに登録する </summary>
        /// <returns>
        ///     照合済み件数</returns>
        /// ----------------------------------------------------------
        public int pastDataUpdate()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            DSLGDataSetTableAdapters.過去データTableAdapter pAdp = new DSLGDataSetTableAdapters.過去データTableAdapter();

            adp.Fill(dts.伝票番号);
            pAdp.Fill(dts.過去データ);

            // 照合件数
            int dNum = 0;

            // 照合済みデータを抽出
            if (dts.伝票番号.Any(a => a.照合ステータス == global.STATUS_VERIFI))
            {
                foreach (var t in dts.伝票番号.Where(a => a.照合ステータス == global.STATUS_VERIFI))
                {
                    // 過去データに未登録の伝票番号を追加する
                    if (!dts.過去データ.Any(a => a.伝票番号 == t.伝票番号))
                    {
                        DSLGDataSet.過去データRow r = dts.過去データ.New過去データRow();
                        r.伝票番号 = t.伝票番号;
                        r.更新年月日 = DateTime.Now;
                        dts.過去データ.Add過去データRow(r);

                        dNum++;
                    }
                }

                pAdp.Update(dts.過去データ);
            }

            // 後片付け
            adp.Dispose();

            // 照合件数を返す
            return dNum;
        }

        /// ----------------------------------------------------------
        /// <summary>
        ///     指定伝票番号を過去データに登録する </summary>
        /// ----------------------------------------------------------
        public void addPastData(int sDen)
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.過去データTableAdapter pAdp = new DSLGDataSetTableAdapters.過去データTableAdapter();
            pAdp.Fill(dts.過去データ);

            // 過去データに未登録を確認して伝票番号を追加する
            if (!dts.過去データ.Any(a => a.伝票番号 == sDen))
            {
                DSLGDataSet.過去データRow r = dts.過去データ.New過去データRow();
                r.伝票番号 = sDen;
                r.更新年月日 = DateTime.Now;
                dts.過去データ.Add過去データRow(r);
                pAdp.Update(dts.過去データ);
            }

            // 後片付け
            pAdp.Dispose();
        }

        /// ----------------------------------------------------------
        /// <summary>
        ///     照合済み未照合伝票データを過去データに登録する </summary>
        /// <returns>
        ///     照合済み件数</returns>
        /// ----------------------------------------------------------
        public int pastDataUpdateUn()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合伝票TableAdapter adp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            DSLGDataSetTableAdapters.過去データTableAdapter pAdp = new DSLGDataSetTableAdapters.過去データTableAdapter();

            adp.Fill(dts.未照合伝票);
            pAdp.Fill(dts.過去データ);

            // 照合件数
            int dNum = 0;

            // 照合済みデータを抽出
            if (dts.未照合伝票.Any(a => a.照合ステータス == global.STATUS_VERIFI))
            {
                foreach (var t in dts.未照合伝票.Where(a => a.照合ステータス == global.STATUS_VERIFI))
                {
                    // 過去データに未登録の伝票番号を追加する
                    if (!dts.過去データ.Any(a => a.伝票番号 == t.伝票番号))
                    {
                        DSLGDataSet.過去データRow r = dts.過去データ.New過去データRow();
                        r.伝票番号 = t.伝票番号;
                        r.更新年月日 = DateTime.Now;
                        dts.過去データ.Add過去データRow(r);

                        dNum++;
                    }
                }

                pAdp.Update(dts.過去データ);
            }

            // 後片付け
            adp.Dispose();

            // 照合件数を返す
            return dNum;
        }

        /// ----------------------------------------------------------
        /// <summary>
        ///     過去データを削除する </summary>
        /// <param name="sDen">
        ///     伝票番号</param>
        /// ----------------------------------------------------------
        public void pastDataCancel(int sDen)
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.過去データTableAdapter pAdp = new DSLGDataSetTableAdapters.過去データTableAdapter();
            pAdp.Fill(dts.過去データ);
            
            // 照合済みデータを抽出
            if (dts.過去データ.Any(a => a.伝票番号 == sDen))
            {
                var t = dts.過去データ.Single(a => a.伝票番号 == sDen);
                t.Delete();
                pAdp.Update(dts.過去データ);
            }

            // 後片付け
            pAdp.Dispose();
        }

        /// ----------------------------------------------------------
        /// <summary>
        ///     照合済み伝票番号データで配車テーブルを更新する </summary>
        /// <returns>
        ///     照合済み件数</returns>
        /// ----------------------------------------------------------
        public int haishaDataUpdate()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            DSLGDataSetTableAdapters.配車TableAdapter hAdp = new DSLGDataSetTableAdapters.配車TableAdapter();

            adp.Fill(dts.伝票番号);
            hAdp.Fill(dts.配車);

            // 照合件数
            int dNum = 0;

            // 照合済みデータを抽出
            foreach (var t in dts.伝票番号.Where(a => a.照合ステータス == global.STATUS_VERIFI))
            {
                // 配車データの照合結果を更新する
                if (dts.配車.Any(a => a.伝票番号 == t.伝票番号 && a.日付 == t.日付))
                {
                    DSLGDataSet.配車Row r = dts.配車.Single(a => a.伝票番号 == t.伝票番号 && a.日付 == t.日付);
                    r.画像名 = t.画像名;
                    r.照合ステータス = t.照合ステータス;
                    r.更新年月日 = DateTime.Now;

                    dNum++;

                    // 伝票番号データを削除
                    t.Delete();
                }
            }

            // データベースを更新
            hAdp.Update(dts.配車);
            adp.Update(dts.伝票番号);

            // 後片付け
            adp.Dispose();
            hAdp.Dispose();

            // 照合件数を返す
            return dNum;
        }

        /// -----------------------------------------------------------------
        /// <summary>
        ///     照合済みの未照合伝票データで配車テーブルを更新する </summary>
        /// <returns>
        ///     照合済み件数</returns>
        /// -----------------------------------------------------------------
        public int haishaDataUpdateUn()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合伝票TableAdapter adp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            DSLGDataSetTableAdapters.配車TableAdapter hAdp = new DSLGDataSetTableAdapters.配車TableAdapter();

            adp.Fill(dts.未照合伝票);
            hAdp.Fill(dts.配車);

            // 照合件数
            int dNum = 0;

            // 照合済みデータを抽出
            foreach (var t in dts.未照合伝票.Where(a => a.照合ステータス == global.STATUS_VERIFI))
            {
                // 配車データの照合結果を更新する
                if (dts.配車.Any(a => a.伝票番号 == t.伝票番号))
                {
                    DSLGDataSet.配車Row r = dts.配車.Single(a => a.伝票番号 == t.伝票番号 && a.日付 == t.日付);
                    r.画像名 = t.画像名;
                    r.照合ステータス = t.照合ステータス;
                    r.更新年月日 = DateTime.Now;

                    dNum++;

                    // 未照合伝票データを削除
                    t.Delete();
                }
            }

            // データベースを更新
            hAdp.Update(dts.配車);
            adp.Update(dts.未照合伝票);

            // 後片付け
            adp.Dispose();
            hAdp.Dispose();

            // 照合件数を返す
            return dNum;
        }

        /// ----------------------------------------------------------
        /// <summary>
        ///     NGデータを出力する </summary>
        /// <returns>
        ///     出力件数を返す</returns>
        /// ----------------------------------------------------------
        public int ngOutput()
        {
            //// データセット
            //RELODataSet dts = new RELODataSet();
            //RELODataSetTableAdapters.読み取りデータTableAdapter adp = new RELODataSetTableAdapters.読み取りデータTableAdapter();
            //RELODataSetTableAdapters.NGTableAdapter adpng = new RELODataSetTableAdapters.NGTableAdapter();
            //adp.Fill(dts.読み取りデータ);
            //adpng.Fill(dts.NG);

            //if (dts.読み取りデータ.Any(a => a.NGフラグ == global.flgOn))
            //{
            //    foreach (var t in dts.読み取りデータ.Where(a => a.NGフラグ == global.flgOn).OrderBy(a => a.ID))
            //    {
            //        dts.NG.AddNGRow(t.画像名, denCnt, t.枚目, t.番号, DateTime.Now);

            //        // 画像をＮＧフォルダへコピー
            //        ngImageCopy(Properties.Settings.Default.dataPath + t.画像名);
            //    }

            //    adpng.Update(dts.NG);
            //}

            //// 戻り値
            //int rCnt = dts.読み取りデータ.Count(a => a.NGフラグ == global.flgOn);

            //// 後片付け
            //adp.Dispose();
            //adpng.Dispose();

            //// 書き込み件数を返す
            //return rCnt;
            return 0;
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     NGデータがある画像をＮＧフォルダへコピーします </summary>
        /// <param name="imgPath">
        ///     NG画像パス</param>
        /// -------------------------------------------------------------------------
        private void ngImageCopy(string imgPath)
        {
            // NGフォルダがなければ作成
            if (!System.IO.Directory.Exists(Properties.Settings.Default.ngPath))
            {
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.ngPath);
            }

            // NGPath
            string newPath = Properties.Settings.Default.ngPath + System.IO.Path.GetFileName(imgPath);

            if (!System.IO.File.Exists(newPath))
            {
                System.IO.File.Copy(imgPath, newPath, true);
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     伝票番号テーブル全行削除 </summary>
        ///--------------------------------------------------------------------
        public void dataDelete()
        {
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            adp.Fill(dts.伝票番号);

            for (int i = 0; i < dts.伝票番号.Rows.Count; i++)
            {
                dts.伝票番号.Rows[i].Delete();
            }

            adp.Update(dts.伝票番号);
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     画像ファイル削除 </summary>
        /// -------------------------------------------------------------------
        public void imageDelete()
        {
            foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.imgOutPath, "*.tif"))
            {
                System.IO.File.Delete(file);
            }
        }

        /// ---------------------------------------------------------------------------------------
        /// <summary>
        ///     「未照合OK」「ＮＧ」以外の未照合伝票の照合ステータスを未処理に書き換える </summary>
        /// ---------------------------------------------------------------------------------------
        public void unStatusToUnveri()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合伝票TableAdapter adp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            adp.Fill(dts.未照合伝票);

            foreach (var t in dts.未照合伝票.Where(a => a.照合ステータス != global.STATUS_UNVERIOK && 
                                                       a.照合ステータス != global.STATUS_NG))
            {
                t.照合ステータス = global.STATUS_UNVERI;                
            }

            adp.Update(dts.未照合伝票);                        
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     未照合データを未照合伝票データに登録し伝票番号データから削除する </summary>
        /// <returns>
        ///     件数</returns>
        /// -------------------------------------------------------------------------------
        public int unmDataUpdate()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.伝票番号TableAdapter adp = new DSLGDataSetTableAdapters.伝票番号TableAdapter();
            DSLGDataSetTableAdapters.未照合伝票TableAdapter pAdp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();

            adp.Fill(dts.伝票番号);
            pAdp.Fill(dts.未照合伝票);

            // 未照合件数
            int dNum = 0;

            // 未照合データを抽出
            foreach (var t in dts.伝票番号.Where(a => a.照合ステータス != global.STATUS_VERIFI && a.照合ステータス != global.STATUS_UNVERI))
            {
                // 未照合伝票テーブルに追加する
                DSLGDataSet.未照合伝票Row r = dts.未照合伝票.New未照合伝票Row();
                r.伝票番号 = t.伝票番号;
                r.画像名 = t.画像名;
                r.日付 = t.日付;
                r.メーカー名 = t.メーカー名;
                r.照合ステータス = t.照合ステータス;
                r.更新年月日 = DateTime.Now;
                dts.未照合伝票.Add未照合伝票Row(r);

                dNum++;

                // 伝票番号データ削除
                t.Delete();
            }

            // データベースを更新
            pAdp.Update(dts.未照合伝票);
            adp.Update(dts.伝票番号);

            // 後片付け
            adp.Dispose();
            pAdp.Dispose();

            // 未照合件数を返す
            return dNum;
        }
        
        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     NG画像を未照合データを未照合伝票データに登録する </summary>
        /// <returns>
        ///     件数</returns>
        /// -------------------------------------------------------------------------------
        public int ngToUnmData()
        {
            // データセット
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.未照合伝票TableAdapter pAdp = new DSLGDataSetTableAdapters.未照合伝票TableAdapter();
            pAdp.Fill(dts.未照合伝票);
            
            // OCR認識日付取得
            DateTime dt = getOcrDate();
            string sDt = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0');

            // 未照合追加件数
            int dNum = 0;

            // ＮＧ画像を抽出
            foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.ngPath, "*.tif"))
            {
                // 未照合伝票テーブル登録済みの画像はネグる 2015/12/15
                if (dts.未照合伝票.Any(a => a.画像名 == file))
                {
                    continue;
                }
                
                // 未照合伝票テーブルに追加する
                DSLGDataSet.未照合伝票Row r = dts.未照合伝票.New未照合伝票Row();
                r.伝票番号 = 0;
                r.画像名 = file;
                r.日付 = DateTime.Parse(sDt.Substring(0, 4) + "/" + sDt.Substring(4, 2) + "/" + sDt.Substring(6, 2));
                r.メーカー名 = "";
                r.照合ステータス = global.STATUS_NG;
                r.更新年月日 = DateTime.Now;
                dts.未照合伝票.Add未照合伝票Row(r);

                dNum++;
            }

            // データベースを更新
            pAdp.Update(dts.未照合伝票);

            // 後片付け
            pAdp.Dispose();

            // 未照合件数を返す
            return dNum;
        }

    }
}
