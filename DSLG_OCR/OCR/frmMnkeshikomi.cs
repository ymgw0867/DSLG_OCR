using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DSLG_OCR.Common;

namespace DSLG_OCR.OCR
{
    public partial class frmMnkeshikomi : Form
    {
        public frmMnkeshikomi(string denNum)
        {
            InitializeComponent();

            _denNum = denNum;
        }

        // 伝票番号
        string _denNum = string.Empty;

        private void button2_Click(object sender, EventArgs e)
        {
            mDate = string.Empty;
            mMaker = string.Empty;

            this.Close();
        }

        private void frmMnkeshikomi_Load(object sender, EventArgs e)
        {
            // メーカーコンボボックスに配車テーブルのメーカー名をセットします
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.配車TableAdapter adp = new DSLGDataSetTableAdapters.配車TableAdapter();
            adp.Fill(dts.配車);
            
            var s = dts.配車.Select(a => new 
                {
                    a.メーカー名
                }).Distinct();

            foreach (var t in s)
            {
                comboBox1.Items.Add(t.メーカー名); 
            }

            // 伝票番号
            txtDenNum.Text = _denNum;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 手動消込処理
            getKeshikomi();
        }

        /// --------------------------------------------------------------
        /// <summary>
        ///     手動消込処理 </summary>
        /// --------------------------------------------------------------
        private void getKeshikomi()
        {
            if (comboBox1.Text.Trim() == string.Empty)
            {
                MessageBox.Show("メーカー名を選択または入力してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("表示中の物品受領書の伝票№の手動消込を行い「未照合ＯＫ」扱いとします。" + Environment.NewLine + "よろしいですか？", "消込確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            mDenNum = Utility.StrtoInt(txtDenNum.Text);
            mDate = dateTimePicker1.Value.ToShortDateString();
            mMaker = comboBox1.Text;

            this.Close();
        }

        public int mDenNum { get; set; }
        public string mDate { get; set; }
        public string mMaker { get; set; }

        private void txtDenNum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
