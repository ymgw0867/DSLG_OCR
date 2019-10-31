using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DSLG_OCR.Common
{
    class global
    {
        public static string pblImagePath;

        #region 画像表示倍率（%）・座標
        public static float miMdlZoomRate = 0;      // 現在の表示倍率
        public static float miMdlZoomRate_TATE = 0; // 現在のタテ表示倍率
        public static float miMdlZoomRate_YOKO = 0; // 現在のヨコ表示倍率
        public static float ZOOM_RATE = 0.23f;      // 標準倍率
        public static float ZOOM_RATE_TATE = 0.25f; // タテ標準倍率
        public static float ZOOM_RATE_YOKO = 0.28f; // ヨコ標準倍率
        public static float ZOOM_MAX = 2.00f;       // 最大倍率
        public static float ZOOM_MIN = 0.05f;       // 最小倍率
        public static float ZOOM_STEP = 0.02f;      // ステップ倍率
        public static float ZOOM_NOW;               // 現在の倍率

        public static int RECTD_NOW;                // 現在の座標
        public static int RECTS_NOW;                // 現在の座標
        public static int RECT_STEP = 20;           // ステップ座標
        #endregion
        
        #region ローカルMDB関連定数
        public const string MDBFILE = "DSLG.mdb";         // MDBファイル名
        public const string MDBTEMP = "DSLG_Temp.mdb";    // 最適化一時ファイル名
        public const string MDBBACK = "DSLG_Back.mdb";    // 最適化後バックアップファイル名
        #endregion

        #region フラグオン・オフ定数
        public const int flgOn = 1;            //フラグ有り(1)
        public const int flgOff = 0;           //フラグなし(0)
        public const string FLGON = "1";
        public const string FLGOFF = "0";
        #endregion

        public static int pblDenNum;            // データ数

        public const int configKEY = 1;        // 環境設定データキー

        //ＯＣＲ処理ＣＳＶデータの検証要素
        public const int CSVLENGTH = 197;          //データフィールド数 2011/06/11
        public const int CSVFILENAMELENGTH = 21;   //ファイル名の文字数 2011/06/11  
 
        #region 環境設定項目
        public static string cnfTifPath;            // 照合済みフォルダパス
        public static string cnfUnmImgPath;         // 未照合画像フォルダパス
        public static string cnfUnmOkImgPath;       // 未照合OK画像フォルダパス
        public static string cnfHaishaPath;         // 配車データパス名
        public static string cnfPastPath;           // 重複データ出力先パス
        public static string cnfPastName;           // 重複データファイル名
        public static string cnfLostPath;           // 未登録データ出力先パス
        public static string cnfLostName;           // 未登録データファイル名
        public static string cnfLogPath;            // ログデータ出力先パス
        public static string cnfLogName;            // ログデータファイル名
        #endregion

        #region 伝票番号テーブル照合ステータス
        public const int STATUS_UNVERI = 0;         // 未処理
        public const int STATUS_VERIFI = 1;         // 照合済み
        public const int STATUS_UNVERIOK = 2;       // 未照合OK
        public const int STATUS_DENOVERLAP = 3;     // データ内伝票№重複
        public const int STATUS_PASTOVERLAP = 4;    // 過去データ重複
        public const int STATUS_UNFIND = 5;         // 配車データ未登録
        public const int STATUS_NG = 6;             // NG
        #endregion

        public const string UNMARK = "UN";  // 未照合画像マーク
    }
}
