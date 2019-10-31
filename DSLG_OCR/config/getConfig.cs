using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DSLG_OCR.Common;

namespace DSLG_OCR.Config
{
    public class getConfig
    {
        DSLGDataSet db = new DSLGDataSet();
        DSLGDataSetTableAdapters.環境設定TableAdapter adp = new DSLGDataSetTableAdapters.環境設定TableAdapter();

        public getConfig()
        {
            try
            {
                adp.Fill(db.環境設定);
                DSLGDataSet.環境設定Row r = db.環境設定.Single(a => a.ID == global.configKEY);

                global.cnfTifPath = r.照合済みフォルダパス;
                global.cnfUnmImgPath = r.未照合画像フォルダパス;
                global.cnfUnmOkImgPath = r.未照合OKフォルダパス;

                if (r.Is配車データパスNull())
                {
                    global.cnfHaishaPath = string.Empty;
                }
                else
                {
                    global.cnfHaishaPath = r.配車データパス;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定情報取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
            }
        }
    }
}
