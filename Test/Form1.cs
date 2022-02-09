using System;
using System.Drawing;
using System.Diagnostics;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using System.Drawing.Printing;
using iTextSharp;
using O2S.Components.PDFRender4NET;
using O2S.Components.PDFRender4NET.Printing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using iTextSharp.text.pdf;

namespace Test
{
    /// <summary>
    /// Form1 ��ժҪ˵����
    /// </summary>
    public class Form1 : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.TextBox textBox9;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private Timer timer1;
        private Button button3;
        private Button button4;
        private Label label7;
        private TextBox textBox11;
        private Button button5;
        private Button button6;
        private IContainer components;


        #region ����֤�Ķ����ӿ�
        /// <summary>
        /// ����֤�Ķ����ӿ�
        /// </summary>
        /// <returns></returns>
        [DllImport("kernel32", EntryPoint = "GetModuleFileNameA")]
        static extern int GetModuleFileName(int hModule, string lpFileName, int nSize);

        //��ʼ���豸(����:1-16,USB:1001-1016)
        [DllImport("termb.dll")]
        static extern int InitComm(int X);
        //������Ϻ󣬹ر��豸
        [DllImport("termb.dll")]
        static extern int CloseComm();
        //֤����֤������ֵ����Ҫ�ж�
        [DllImport("termb.dll")]
        static extern int Authenticate();
        //��������Ϣ,iActive=1;�����ɹ�����Ƭ��Ϣ�����zp.bmp�ļ���;��׷�ӵ�ַ,iActive=3;
        [DllImport("termb.dll")]
        static extern int Read_Content(int Active);

        //�����ɹ���������·�����ȡ��Ӧ�����֤��Ϣ��
        //		const	int ERR_SUCCESS			= 1;//�ɹ�
        //		const	int ERR_FAIL		    	= 0;//ʧ��
        //		const	int ERR_SAVESERIALNO		= -1;//�����к�ʧ�� δ��Ȩ����
        //		const	int ERR_CANCELSERIALNO		= -1;//���к�ȡ��  δ��Ȩ����
        //		const	int ERR_OPENECOMM		= -2;//�򿪴���
        //		const	int ERR_CLOSECOMM		= -3;//�رմ���
        //		const	int ERR_SAMSTATUS		= -4;//ȡsam״̬ʧ��
        //		const	int ERR_SAMID		    	= -5;//ȡsamIDʧ��
        //		const	int ERR_FINDCARD		= -6;//�ҿ�����
        //		const	int ERR_SELECTCARD		= -7;//ѡ������
        //		const	int ERR_BASEINFO		= -8;//��ȡ������Ϣ����
        //		const	int ERR_APPINFO			= -9;//��ȡ������Ϣ����
        //		const	int ERR_MNGINFO			= -10;//��ȡMNG��Ϣ����
        //����
        [DllImport("termb.dll")]
        static extern int GetPeopleName(StringBuilder lpBuffer, uint strLen);
        //��ַ
        [DllImport("termb.dll")]
        static extern int GetPeopleAddress(StringBuilder lpBuffer, uint strLen);
        //���֤����
        [DllImport("termb.dll")]
        static extern int GetPeopleIDCode(StringBuilder lpBuffer, uint strLen);
        //��������
        [DllImport("termb.dll")]
        static extern int GetPeopleBirthday(StringBuilder lpBuffer, uint strLen);
        //����
        [DllImport("termb.dll")]
        static extern int GetPeopleNation(StringBuilder lpBuffer, uint strLen);
        //�Ա�
        [DllImport("termb.dll")]
        static extern int GetPeopleSex(StringBuilder lpBuffer, uint strLen);
        //��֤����
        [DllImport("termb.dll")]
        static extern int GetDepartment(StringBuilder lpBuffer, uint strLen);
        //֤����ʼ����
        [DllImport("termb.dll")]
        static extern int GetStartDate(StringBuilder lpBuffer, uint strLen);
        //֤����������
        [DllImport("termb.dll")]
        static extern int GetEndDate(StringBuilder lpBuffer, uint strLen);
        //��Ƭ
        [DllImport("termb.dll")]
        static extern int GetPhotoBMP(StringBuilder lpBuffer, uint strLen);
        //׷�ӵ�ַ
        [DllImport("termb.dll")]
        static extern int GetAppAddress(uint index, StringBuilder lpBuffer, uint strLen);

        [DllImport("termb.dll")]
        static extern int GetSAMIDToStr1(StringBuilder lpBuffer, uint strLen);

        //�ڴ��л�ȡ�ṹ����Ϣ
        [DllImport("termb.dll")]
        static extern int GetIdCardTxtInfo(ref CIdCardInfo info);

        //������Ƭ���ṹ��-��ʱ�䣩
        [DllImport("termb.dll")]
        static extern bool SaveCardData2Bmp(CIdCardInfo pInf, StringBuilder szBmpFile, uint iFace);

        //������Ƭ���ṹ��-��ʱ�䣩
        [DllImport("termb.dll")]
        static extern bool SaveCardData2BmpByTime(CIdCardInfo info, StringBuilder szBmpFile, uint mode, StringBuilder szTime);

        //������Ƭ�����ı�Ϊ��-��ʱ�䣩
        [DllImport("termb.dll")]
        static extern bool SaveCardData2BmpByTxt(StringBuilder szTxt, StringBuilder szBase64Wlt, StringBuilder szBmpFile, uint mode);

        //������Ƭ�����ı�Ϊ��-��ʱ�䣩
        [DllImport("termb.dll")]
        static extern bool SaveCardData2BmpByTxtByTime(StringBuilder szTxt, StringBuilder szBase64Wlt, StringBuilder szBmpFile, uint mode, StringBuilder szTime);
        #endregion

        private static int intPort = 0;//�˿ں�
        private static int iRet = 0;
        public static string folder = AppDomain.CurrentDomain.BaseDirectory + "temp";

        public Form1()
        {
            //
            // Windows ���������֧���������
            //
            InitializeComponent();

            ReadPrintName();
            //
            // TODO: �� InitializeComponent ���ú�����κι��캯������
            //
        }

        /// <summary>
        /// ������������ʹ�õ���Դ��
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows ������������ɵĴ���
        /// <summary>
        /// �����֧������ķ��� - ��Ҫʹ�ô���༭���޸�
        /// �˷��������ݡ�
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "������";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 62);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "��ַ��";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(31, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "���֤���룺";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(31, 116);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 3;
            this.label4.Text = "�������ڣ�";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(31, 140);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 12);
            this.label5.TabIndex = 4;
            this.label5.Text = "���壺";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(31, 167);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(41, 12);
            this.label6.TabIndex = 5;
            this.label6.Text = "�Ա�";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(440, 208);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(88, 32);
            this.button1.TabIndex = 6;
            this.button1.Text = "����";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(31, 270);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(65, 12);
            this.label8.TabIndex = 10;
            this.label8.Text = "׷�ӵ�ַ��";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(31, 244);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(89, 12);
            this.label9.TabIndex = 9;
            this.label9.Text = "֤���������ڣ�";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(31, 218);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(89, 12);
            this.label10.TabIndex = 8;
            this.label10.Text = "֤����ʼ���ڣ�";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(31, 193);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(65, 12);
            this.label11.TabIndex = 7;
            this.label11.Text = "��֤���أ�";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(120, 34);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(312, 21);
            this.textBox1.TabIndex = 12;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(120, 60);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(312, 21);
            this.textBox2.TabIndex = 12;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(120, 87);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(312, 21);
            this.textBox3.TabIndex = 12;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(120, 114);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(312, 21);
            this.textBox4.TabIndex = 12;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(120, 138);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(312, 21);
            this.textBox5.TabIndex = 12;
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(120, 165);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(312, 21);
            this.textBox6.TabIndex = 12;
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(120, 191);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(312, 21);
            this.textBox7.TabIndex = 12;
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(120, 216);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(312, 21);
            this.textBox8.TabIndex = 12;
            // 
            // textBox9
            // 
            this.textBox9.Location = new System.Drawing.Point(120, 242);
            this.textBox9.Name = "textBox9";
            this.textBox9.Size = new System.Drawing.Size(312, 21);
            this.textBox9.TabIndex = 12;
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(120, 268);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(312, 21);
            this.textBox10.TabIndex = 12;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(534, 209);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(88, 32);
            this.button2.TabIndex = 6;
            this.button2.Text = "���";
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(440, 32);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(104, 136);
            this.pictureBox1.TabIndex = 13;
            this.pictureBox1.TabStop = false;
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 3000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(440, 257);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(88, 32);
            this.button3.TabIndex = 14;
            this.button3.Text = "��PDF-100";
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(440, 305);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(88, 32);
            this.button4.TabIndex = 15;
            this.button4.Text = "��ӡ-100";
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(31, 315);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(77, 12);
            this.label7.TabIndex = 16;
            this.label7.Text = "��ӡ�����ƣ�";
            // 
            // textBox11
            // 
            this.textBox11.Location = new System.Drawing.Point(120, 312);
            this.textBox11.Name = "textBox11";
            this.textBox11.Size = new System.Drawing.Size(312, 21);
            this.textBox11.TabIndex = 17;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(534, 305);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(88, 32);
            this.button5.TabIndex = 18;
            this.button5.Text = "��ӡ-78";
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(534, 257);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(88, 32);
            this.button6.TabIndex = 19;
            this.button6.Text = "��PDF-78";
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(626, 340);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.textBox11);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox7);
            this.Controls.Add(this.textBox8);
            this.Controls.Add(this.textBox9);
            this.Controls.Add(this.textBox10);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button2);
            this.Name = "Form1";
            this.Text = "�������֤������ӡ����";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        /// <summary>
        /// Ӧ�ó��������ڵ㡣
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new Form1());
        }

        //������ini�ļ���TextBox11
        public void ReadPrintName()
        {
            IniFiles inifile = new IniFiles(AppDomain.CurrentDomain.BaseDirectory + @"PrintName.ini");
            textBox11.Text = inifile.ReadString("PrintSetting", "Name", "");
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            try
            {
                string str = "";
                str = str.PadLeft(100);
                StringBuilder str1 = new StringBuilder(str);

                if (intPort == 0)
                {
                    for (int i = 1; i < 16; i++)
                    {
                        iRet = InitComm(i);
                        if (iRet == 1)
                        {
                            intPort = i;
                            break;
                        }
                    }
                }
                else
                {
                    iRet = InitComm(intPort);
                }

                if (intPort == 0)
                {
                    for (int j = 1001; j < 1016; j++)
                    {
                        iRet = InitComm(j);
                        if (iRet == 1)
                        {
                            intPort = j;
                            break;
                        }
                    }
                }

                if (intPort == 0)
                {
                    MessageBox.Show("δ�ҵ����֤�Ķ������������֤�Ķ����Ƿ�����Ա�����ȷ���ӻ������֤�Ķ����ĵ�Դ�����Ƿ�����", "��ʾ��Ϣ");
                    return;
                }

                str1.Length = 0;
                str1.Append(str);
                iRet = GetSAMIDToStr1(str1, 100);

                iRet = Authenticate();

                iRet = Read_Content(1);

                if (iRet != 1)
                {   //\u000D�ǻ��б�־
                    MessageBox.Show("����ʧ�ܣ�������2�ֿ��ܣ�\u000D1�����֤δ�źá�\u000D2�������֤���𻵻��Ǽ�֤��", "��ʾ��Ϣ");
                    return;
                }

                CIdCardInfo info = new CIdCardInfo();
                iRet = GetIdCardTxtInfo(ref info);

                StringBuilder frontpath = new StringBuilder($@"{folder}\{info.m_szName}-����.bmp");
                StringBuilder backpath = new StringBuilder($@"{folder}\{info.m_szName}-����.bmp");
                bool isFrontSave = false, isBackSave = false;

                //��������
                string tempName = "";
                if (info.m_szName.Length == 2)
                    tempName = info.m_szName[0] + " " + info.m_szName[1];
                else
                    tempName = info.m_szName;

                //����
                FileStream fsReadWlt = new FileStream("xp.wlt", FileMode.Open, FileAccess.Read);
                info.m_szWltData = ToBase64String(fsReadWlt);

                StringBuilder txt = new StringBuilder
                (
                    //��ʽΪ������|�Ա�|����|��������|סַ|���֤��|ǩ������|��ʼ����|�������ڡ�
                    $"{tempName}|{info.m_szSex}|{info.m_szNation}|{info.m_szBirthday}|{info.m_szAddress}|{info.m_szID}|{info.m_szDepartment}|{info.m_szStartDate}|{info.m_szEndDate}"
                );
                StringBuilder wltdata = new StringBuilder(info.m_szWltData);
                StringBuilder timedata = new StringBuilder();
                isFrontSave = SaveCardData2BmpByTxtByTime(txt, wltdata, frontpath, 1, timedata);
                isBackSave = SaveCardData2BmpByTxtByTime(txt, wltdata, backpath, 2, timedata);

                if (isFrontSave && isBackSave)
                {
                    CombinImage(frontpath.ToString(), backpath.ToString(), $@"{folder}\{info.m_szName}-100-������", 100);
                    CombinImage(frontpath.ToString(), backpath.ToString(), $@"{folder}\{info.m_szName}-78-������", 78);
                }
                else
                    MessageBox.Show(this, "�����ջ����յ���ʧ��", "������Ϣ��", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);

                iRet = GetPeopleName(str1, 100);
                textBox1.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetPeopleAddress(str1, 100);
                textBox2.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetPeopleIDCode(str1, 100);
                textBox3.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetPeopleBirthday(str1, 100);
                textBox4.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetPeopleNation(str1, 100);
                textBox5.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetPeopleSex(str1, 100);
                textBox6.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetDepartment(str1, 100);
                textBox7.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetStartDate(str1, 100);
                textBox8.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetEndDate(str1, 100);
                textBox9.Text = str1.ToString().Trim();

                str1.Length = 0;
                str1.Append(str);
                iRet = GetAppAddress(0, str1, 100);
                textBox10.Text = str1.ToString().Trim();


                FileStream fsReadPic = new FileStream("zp.bmp", FileMode.Open, FileAccess.Read);
                Bitmap bmTemp = new Bitmap(fsReadPic);
                pictureBox1.Image = bmTemp;
                fsReadPic.Close();

                CloseComm();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "������Ϣ��", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            try
            {
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                pictureBox1.Image = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "������Ϣ��", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }
        }

        private void button3_Click(object sender, System.EventArgs e)
        {

            string destfilename = $"{textBox1.Text}-100-������";
            System.Diagnostics.Process.Start("explorer.exe", Path.Combine(folder, destfilename + ".pdf"));//�Զ����ļ���
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string destfilename = $"{textBox1.Text}-78-������";
            System.Diagnostics.Process.Start("explorer.exe", Path.Combine(folder, destfilename + ".pdf"));//�Զ����ļ���
        }

        private void button4_Click(object sender, System.EventArgs e)
        {
            string destfilename = $"{textBox1.Text}-100-������";
            IniFiles inifile = new IniFiles(AppDomain.CurrentDomain.BaseDirectory + @"PrintName.ini");
            inifile.WriteString("PrintSetting", "Name", textBox11.Text);

            PdfPrint print = new PdfPrint(Path.Combine(folder, destfilename + ".pdf"));
            print.PrintX(textBox11.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string destfilename = $"{textBox1.Text}-78-������";
            IniFiles inifile = new IniFiles(AppDomain.CurrentDomain.BaseDirectory + @"PrintName.ini");
            inifile.WriteString("PrintSetting", "Name", textBox11.Text);

            PdfPrint pdf = new PdfPrint(Path.Combine(folder, destfilename + ".pdf"));
            pdf.PrintX(textBox11.Text);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
        }


        //�ϲ�ͼƬ
        private void CombinImage(string image1, string image2, string destfilename, int mode)
        {
            Image img1 = Image.FromFile(image1);
            Bitmap map1 = new Bitmap(img1);
            Image img2 = Image.FromFile(image2);
            Bitmap map2 = new Bitmap(img2);

            var width = Math.Max(img1.Width, img2.Width);
            var height = img1.Height * 3;

            Bitmap bitmap = new Bitmap(width, height);
            Graphics g1 = Graphics.FromImage(bitmap);

            if (mode == 100)
            {
                g1.FillRectangle(Brushes.White, new Rectangle(0, 0, width, height));
                g1.DrawImage(map1, 0, 50, img1.Width, img2.Height);
                g1.DrawImage(map2, 0, img2.Width + 10, img2.Width, img2.Height);
            }
            else
            {
                g1.FillRectangle(Brushes.White, new Rectangle(0, 0, width, height));
                g1.DrawImage(map1, 0, 50, img1.Width * (float)0.78, img2.Height * (float)0.78);
                g1.DrawImage(map2, 0, img2.Width + 10, img2.Width * (float)0.78, img2.Height * (float)0.78);
            }

            map1.Dispose();
            map2.Dispose();

            Image img = bitmap;
            string jpgfile = destfilename + ".jpg";
            img.Save(jpgfile);
            img.Dispose();

            ConvertJPG2PDF(jpgfile, destfilename + ".pdf");//ת��
            File.Delete(jpgfile);//ɾ��JPGͼƬ
        }

        //JPGת��PDF
        private void ConvertJPG2PDF(string jpgfile, string pdffile)
        {
            var document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 25, 25, 25, 25);
            using (var stream = new FileStream(pdffile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                iTextSharp.text.pdf.PdfWriter.GetInstance(document, stream);
                document.Open();
                using (var imageStream = new FileStream(jpgfile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var image = iTextSharp.text.Image.GetInstance(imageStream);
                    if (image.Height > iTextSharp.text.PageSize.A4.Height - 25)
                    {
                        image.ScaleToFit(iTextSharp.text.PageSize.A4.Width - 25, iTextSharp.text.PageSize.A4.Height - 25);
                    }
                    else if (image.Width > iTextSharp.text.PageSize.A4.Width - 25)
                    {
                        image.ScaleToFit(iTextSharp.text.PageSize.A4.Width - 25, iTextSharp.text.PageSize.A4.Height - 25);
                    }
                    image.Alignment = iTextSharp.text.Image.ALIGN_MIDDLE;
                    document.Add(image);
                }

                document.Close();
            }
        }

        private const int BufferSize = 1024 * 8;
        //����Wlt�ļ�
        public static string ToBase64String(Stream s)
        {
            byte[] buff = null;
            StringBuilder rtnvalue = new StringBuilder();

            using (System.IO.BinaryReader br = new System.IO.BinaryReader(s))
            {
                do
                {
                    buff = br.ReadBytes(BufferSize);
                    rtnvalue.Append(Convert.ToBase64String(buff));

                } while (buff.Length != 0);

                br.Close();
            }

            return rtnvalue.ToString(); ;
        }

        // ����������ͼƬ
        public Image ZoomPicture(Image SourceImage, int TargetWidth, int TargetHeight)
        {
            int IntWidth; //�µ�ͼƬ��
            int IntHeight; //�µ�ͼƬ��
            try
            {
                System.Drawing.Imaging.ImageFormat format = SourceImage.RawFormat;
                System.Drawing.Bitmap SaveImage = new System.Drawing.Bitmap(TargetWidth, TargetHeight);
                Graphics g = Graphics.FromImage(SaveImage);
                g.Clear(Color.White);

                //��������ͼƬ�Ĵ�С http://www.cnblogs.com/roucheng/

                if (SourceImage.Width > TargetWidth && SourceImage.Height <= TargetHeight)//��ȱ�Ŀ��ͼƬ��ȴ󣬳��ȱ�Ŀ��ͼƬ����С
                {
                    IntWidth = TargetWidth;
                    IntHeight = (IntWidth * SourceImage.Height) / SourceImage.Width;
                }
                else if (SourceImage.Width <= TargetWidth && SourceImage.Height > TargetHeight)//��ȱ�Ŀ��ͼƬ���С�����ȱ�Ŀ��ͼƬ���ȴ�
                {
                    IntHeight = TargetHeight;
                    IntWidth = (IntHeight * SourceImage.Width) / SourceImage.Height;
                }
                else if (SourceImage.Width <= TargetWidth && SourceImage.Height <= TargetHeight) //�����Ŀ��ͼƬ����С
                {
                    IntHeight = SourceImage.Width;
                    IntWidth = SourceImage.Height;
                }
                else//�����Ŀ��ͼƬ�ĳ�����
                {
                    IntWidth = TargetWidth;
                    IntHeight = (IntWidth * SourceImage.Height) / SourceImage.Width;
                    if (IntHeight > TargetHeight)//���¼���
                    {
                        IntHeight = TargetHeight;
                        IntWidth = (IntHeight * SourceImage.Width) / SourceImage.Height;
                    }
                }

                g.DrawImage(SourceImage, (TargetWidth - IntWidth) / 2, (TargetHeight - IntHeight) / 2, IntWidth, IntHeight);
                SourceImage.Dispose();

                return SaveImage;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "������Ϣ��", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            }

            return null;
        }


    }

    public struct CIdCardInfo
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 31)]
        public string m_szName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 6)]
        public string m_szSex;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 11)]
        public string m_szNation;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 9)]
        public string m_szBirthday;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 71)]
        public string m_szAddress;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 19)]
        public string m_szID;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 31)]                                                                                                                            
        public string m_szDepartment;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 9)]
        public string m_szStartDate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 9)]
        public string m_szEndDate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 37)]
        public string m_szReserve;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 1024)]
        public string m_szWltData;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string m_szTxtData;
        [MarshalAs(UnmanagedType.I2, SizeConst = 100)]
        public short m_iFingerLen;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 1024)]
        public string m_szFingerData;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 100)]
        public string m_szSamid;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 38862)]
        public string m_szBmp;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string m_szContent;
    }
}
 