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
    public partial class frmDate : Form
    {
        public frmDate()
        {
            InitializeComponent();
            dateTimePicker1.Value = DateTime.Today;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MyProperty = global.flgOff;
            this.Close();
        }

        private void frmDate_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        public int MyProperty { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("物品受領書のOCR認識処理を実行します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            //  OCR認識日付を更新
            cnfUpdate(dateTimePicker1.Value.ToShortDateString());

            // フラグ
            MyProperty = global.flgOn;
            this.Close();
        }

        /// -------------------------------------------------------------
        /// <summary>
        ///     OCR認識日付を更新 </summary>
        /// <param name="dt">
        ///     日付</param>
        /// -------------------------------------------------------------
        private void cnfUpdate(string dt)
        {
            DSLGDataSet dts = new DSLGDataSet();
            DSLGDataSetTableAdapters.環境設定TableAdapter adp = new DSLGDataSetTableAdapters.環境設定TableAdapter();
            adp.Fill(dts.環境設定);

            DSLGDataSet.環境設定Row r = dts.環境設定.Single(a => a.ID == global.configKEY);
            r.OCR認識日付 = DateTime.Parse(dt);
            adp.Update(dts);
        }
    }
}
