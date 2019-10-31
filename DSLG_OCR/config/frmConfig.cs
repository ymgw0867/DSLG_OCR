using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DSLG_OCR.Common;
using System.Data.OleDb;

namespace DSLG_OCR.Config
{
    public partial class frmConfig : Form
    {
        public frmConfig()
        {
            InitializeComponent();

            adp.Fill(db.環境設定);

            var s = db.環境設定.Single(a => a.ID == global.configKEY);

            if (s.Is照合済みフォルダパスNull())
            {
                lblPath.Text = string.Empty;
            }
            else
            {
                lblPath.Text = s.照合済みフォルダパス;
            }
            
            if (s.Is未照合画像フォルダパスNull())
            {
                lblUnmImgPath.Text = string.Empty;
            }
            else
            {
                lblUnmImgPath.Text = s.未照合画像フォルダパス;
            }
            
            if (s.Is未照合OKフォルダパスNull())
            {
                lblUnmOkPath.Text = string.Empty;
            }
            else
            {
                lblUnmOkPath.Text = s.未照合OKフォルダパス;
            }


            //if (s.Isログ作成パスNull())
            //{
            //    lblLogPath.Text = string.Empty;
            //}
            //else
            //{
            //    lblLogPath.Text = s.ログ作成パス;
            //}

            //if (s.Isログファイル名Null())
            //{
            //    txtLogName.Text = string.Empty;
            //}
            //else
            //{
            //    txtLogName.Text = s.ログファイル名;
            //}
        }

        DSLGDataSet db = new DSLGDataSet();
        DSLGDataSetTableAdapters.環境設定TableAdapter adp = new DSLGDataSetTableAdapters.環境設定TableAdapter();

        private void frmConfig_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);
        }

        /// <summary>
        /// フォルダダイアログ選択
        /// </summary>
        /// <returns>フォルダー名</returns>
        private string userFolderSelect()
        {
            string fName = string.Empty;

            //出力フォルダの選択ダイアログの表示
            // FolderBrowserDialog の新しいインスタンスを生成する (デザイナから追加している場合は必要ない)
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            // ダイアログの説明を設定する
            folderBrowserDialog1.Description = "フォルダを選択してください";

            // ルートになる特殊フォルダを設定する (初期値 SpecialFolder.Desktop)
            folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.Desktop;

            // 初期選択するパスを設定する
            folderBrowserDialog1.SelectedPath = @"C:\DSLG_OCR";

            // [新しいフォルダ] ボタンを表示する (初期値 true)
            folderBrowserDialog1.ShowNewFolderButton = true;

            // ダイアログを表示し、戻り値が [OK] の場合は、選択したディレクトリを表示する
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                fName = folderBrowserDialog1.SelectedPath + @"\";
            }
            else
            {
                // 不要になった時点で破棄する
                folderBrowserDialog1.Dispose();
                return fName;
            }

            // 不要になった時点で破棄する
            folderBrowserDialog1.Dispose();

            return fName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            string sPath = userFolderSelect();
            if (sPath != string.Empty)
            {
                lblPath.Text = sPath;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // データ更新
            DataUpdate();
        }

        private void DataUpdate()
        {
            if (MessageBox.Show("データを更新してよろしいですか","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;

            // エラーチェック
            //if (!errCheck(lblLogPath, txtLogName, "処理結果ログ出力先パス", "処理結果ログファイル名")) return;

            DSLGDataSet.環境設定Row r = db.環境設定.Single(a => a.ID == global.configKEY);

            r.照合済みフォルダパス = lblPath.Text;
            r.未照合画像フォルダパス = lblUnmImgPath.Text;
            r.未照合OKフォルダパス = lblUnmOkPath.Text;            
            r.更新年月日 = DateTime.Now;

            // データ更新
            adp.Update(r);
 
            // 終了
            this.Close();
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェック </summary>
        /// <param name="lbl">
        ///     ラベルオブジェクト</param>
        /// <param name="txt">
        ///     テキストボックスオブジェクト</param>
        /// <param name="lblName">
        ///     ラベル摘要</param>
        /// <param name="txtName">
        ///     テキストボックス摘要</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// ------------------------------------------------------------------------------------
        private bool errCheck(Label lbl, TextBox txt, string lblName, string txtName)
        {
            // パス
            if (lbl.Text.Trim() == string.Empty)
            {
                MessageBox.Show(lblName + "が指定されていません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                lbl.Focus();
                return false;
            }

            // ファイル名
            if (txt.Text == string.Empty)
            {
                MessageBox.Show(txtName + "が指定されていません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt.Focus();
                return false;
            }

            // ファイル名の禁止文字「\ / ? : * " > < |.」
            if (txt.Text.Contains(@"\") || txt.Text.Contains("/") || txt.Text.Contains("?") ||
                txt.Text.Contains(":") || txt.Text.Contains("*") || txt.Text.Contains(">") ||
                txt.Text.Contains(@"""") || txt.Text.Contains("|") || txt.Text.Contains("."))
            {
                MessageBox.Show("ファイル名に使用できない文字が含まれています", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txt.Focus();
                return false;
            }
            
            return true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmConfig_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            string sPath = userFolderSelect();
            if (sPath != string.Empty)
            {
                lblUnmImgPath.Text = sPath;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            string sPath = userFolderSelect();
            if (sPath != string.Empty)
            {
                lblUnmOkPath.Text = sPath;
            }
        }
    }
}
