namespace DSLG_OCR.OCR
{
    partial class frmUnDenpyo
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmUnDenpyo));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.leadImg = new Leadtools.WinForms.RasterImageViewer();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDenNum = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.hScrollBar1 = new System.Windows.Forms.HScrollBar();
            this.btnEnd = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnBefore = new System.Windows.Forms.Button();
            this.btnFirst = new System.Windows.Forms.Button();
            this.btnPlus = new System.Windows.Forms.Button();
            this.btnMinus = new System.Windows.Forms.Button();
            this.lblNoImage = new System.Windows.Forms.Label();
            this.lblPage = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblErrMsg = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Location = new System.Drawing.Point(0, 1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(1182, 531);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // leadImg
            // 
            this.leadImg.Cursor = System.Windows.Forms.Cursors.Hand;
            this.leadImg.Location = new System.Drawing.Point(0, 1);
            this.leadImg.Name = "leadImg";
            this.leadImg.Size = new System.Drawing.Size(1182, 531);
            this.leadImg.TabIndex = 1;
            this.leadImg.ScrollPositionChanged += new System.EventHandler(this.leadImg_ScrollPositionChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(914, 587);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 24);
            this.label1.TabIndex = 2;
            this.label1.Text = "伝票№：";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtDenNum
            // 
            this.txtDenNum.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtDenNum.Location = new System.Drawing.Point(995, 569);
            this.txtDenNum.Name = "txtDenNum";
            this.txtDenNum.Size = new System.Drawing.Size(180, 55);
            this.txtDenNum.TabIndex = 4;
            this.txtDenNum.Text = "123456";
            this.txtDenNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtDenNum.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDenNum_KeyPress);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(1066, 634);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(109, 32);
            this.button1.TabIndex = 0;
            this.button1.Text = "終了(&E)";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button3.Location = new System.Drawing.Point(687, 634);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(126, 32);
            this.button3.TabIndex = 1;
            this.button3.Text = "再照合(&H)";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button4.Location = new System.Drawing.Point(819, 634);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(126, 32);
            this.button4.TabIndex = 2;
            this.button4.Text = "未照合ＯＫ(&M)";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button5.Location = new System.Drawing.Point(951, 634);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(109, 32);
            this.button5.TabIndex = 3;
            this.button5.Text = "削除(&D)";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // hScrollBar1
            // 
            this.hScrollBar1.Location = new System.Drawing.Point(63, 535);
            this.hScrollBar1.Name = "hScrollBar1";
            this.hScrollBar1.Size = new System.Drawing.Size(904, 25);
            this.hScrollBar1.TabIndex = 7;
            this.hScrollBar1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.hScrollBar1_Scroll);
            // 
            // btnEnd
            // 
            this.btnEnd.Font = new System.Drawing.Font("メイリオ", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnEnd.Image = ((System.Drawing.Image)(resources.GetObject("btnEnd.Image")));
            this.btnEnd.Location = new System.Drawing.Point(1045, 536);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(26, 27);
            this.btnEnd.TabIndex = 11;
            this.btnEnd.TabStop = false;
            this.btnEnd.UseVisualStyleBackColor = true;
            this.btnEnd.Click += new System.EventHandler(this.btnEnd_Click);
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("メイリオ", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.Location = new System.Drawing.Point(1020, 536);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(26, 27);
            this.btnNext.TabIndex = 10;
            this.btnNext.TabStop = false;
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnBefore
            // 
            this.btnBefore.Font = new System.Drawing.Font("メイリオ", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnBefore.Image = ((System.Drawing.Image)(resources.GetObject("btnBefore.Image")));
            this.btnBefore.Location = new System.Drawing.Point(995, 536);
            this.btnBefore.Name = "btnBefore";
            this.btnBefore.Size = new System.Drawing.Size(26, 27);
            this.btnBefore.TabIndex = 9;
            this.btnBefore.TabStop = false;
            this.btnBefore.UseVisualStyleBackColor = true;
            this.btnBefore.Click += new System.EventHandler(this.btnBefore_Click);
            // 
            // btnFirst
            // 
            this.btnFirst.Font = new System.Drawing.Font("メイリオ", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnFirst.Image = ((System.Drawing.Image)(resources.GetObject("btnFirst.Image")));
            this.btnFirst.Location = new System.Drawing.Point(970, 536);
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.Size = new System.Drawing.Size(26, 27);
            this.btnFirst.TabIndex = 8;
            this.btnFirst.TabStop = false;
            this.btnFirst.UseVisualStyleBackColor = true;
            this.btnFirst.Click += new System.EventHandler(this.btnFirst_Click);
            // 
            // btnPlus
            // 
            this.btnPlus.Image = ((System.Drawing.Image)(resources.GetObject("btnPlus.Image")));
            this.btnPlus.Location = new System.Drawing.Point(7, 536);
            this.btnPlus.Name = "btnPlus";
            this.btnPlus.Size = new System.Drawing.Size(27, 27);
            this.btnPlus.TabIndex = 5;
            this.btnPlus.TabStop = false;
            this.btnPlus.UseVisualStyleBackColor = true;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click);
            // 
            // btnMinus
            // 
            this.btnMinus.Image = ((System.Drawing.Image)(resources.GetObject("btnMinus.Image")));
            this.btnMinus.Location = new System.Drawing.Point(33, 536);
            this.btnMinus.Name = "btnMinus";
            this.btnMinus.Size = new System.Drawing.Size(27, 27);
            this.btnMinus.TabIndex = 6;
            this.btnMinus.TabStop = false;
            this.btnMinus.UseVisualStyleBackColor = true;
            this.btnMinus.Click += new System.EventHandler(this.btnMinus_Click);
            // 
            // lblNoImage
            // 
            this.lblNoImage.Font = new System.Drawing.Font("メイリオ", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNoImage.ForeColor = System.Drawing.SystemColors.AppWorkspace;
            this.lblNoImage.Location = new System.Drawing.Point(418, 219);
            this.lblNoImage.Name = "lblNoImage";
            this.lblNoImage.Size = new System.Drawing.Size(322, 42);
            this.lblNoImage.TabIndex = 120;
            this.lblNoImage.Text = "画像はありません";
            this.lblNoImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblPage
            // 
            this.lblPage.Location = new System.Drawing.Point(1097, 535);
            this.lblPage.Name = "lblPage";
            this.lblPage.Size = new System.Drawing.Size(85, 25);
            this.lblPage.TabIndex = 121;
            this.lblPage.Text = "(1000/1000)";
            this.lblPage.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(26, 587);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 24);
            this.label2.TabIndex = 123;
            this.label2.Text = "照合結果：";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.lblErrMsg);
            this.panel1.Location = new System.Drawing.Point(125, 572);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(715, 55);
            this.panel1.TabIndex = 124;
            // 
            // lblErrMsg
            // 
            this.lblErrMsg.AutoSize = true;
            this.lblErrMsg.Font = new System.Drawing.Font("Meiryo UI", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblErrMsg.ForeColor = System.Drawing.Color.Red;
            this.lblErrMsg.Location = new System.Drawing.Point(48, 13);
            this.lblErrMsg.Name = "lblErrMsg";
            this.lblErrMsg.Size = new System.Drawing.Size(77, 28);
            this.lblErrMsg.TabIndex = 0;
            this.lblErrMsg.Text = "label2";
            // 
            // frmUnDenpyo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1186, 674);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblPage);
            this.Controls.Add(this.lblNoImage);
            this.Controls.Add(this.hScrollBar1);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.btnBefore);
            this.Controls.Add(this.btnFirst);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtDenNum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.leadImg);
            this.Controls.Add(this.pictureBox1);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MinimizeBox = false;
            this.Name = "frmUnDenpyo";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "未照合伝票確認";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmUnDenpyo_FormClosing);
            this.Load += new System.EventHandler(this.frmUnDenpyo_Load);
            this.Shown += new System.EventHandler(this.frmUnDenpyo_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private Leadtools.WinForms.RasterImageViewer leadImg;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDenNum;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.HScrollBar hScrollBar1;
        private System.Windows.Forms.Button btnEnd;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnBefore;
        private System.Windows.Forms.Button btnFirst;
        private System.Windows.Forms.Button btnPlus;
        private System.Windows.Forms.Button btnMinus;
        private System.Windows.Forms.Label lblNoImage;
        private System.Windows.Forms.Label lblPage;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblErrMsg;
    }
}