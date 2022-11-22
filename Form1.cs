using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ChromeDriverUpdater;
using System.Runtime.InteropServices;
using WindowsInput;

namespace WP_AutoPost
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Thread thread1;
        private ChromeDriverService chromeDriverService;
        private ChromeOptions chromeOptions;
        private ChromeDriver driver;
        private TextBox texttitle = new TextBox();
        private TextBox textbody = new TextBox();
        private List<string> script = new List<string>();
        private string id = string.Empty;
        private string pw = string.Empty;
        private DateTime now = DateTime.Now;
        System.Text.RegularExpressions.Regex cntStr = new System.Text.RegularExpressions.Regex(" ");

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(
          uint dwFlags,
          int dx,
          int dy,
          uint cButtons,
          uint dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern int GetCursorPos(out Form1.POINTAPI pt);

        [DllImport("user32.dll")]
        private static extern int SetCursorPos(int x, int y);

        private void Mouse_Left_Click() => new InputSimulator().Mouse.LeftButtonClick();
        private void Mouse_Wheel_Sroll_Down(int bb) => new InputSimulator().Mouse.VerticalScroll(bb);
        private void Mouse_Wheel_Sroll_Up(int dd) => new InputSimulator().Mouse.VerticalScroll(dd);

        private void Form1_Load(object sender, EventArgs e)
        {
            //radioButton1.Checked = true;
            //textBox4.Visible = false;
            //if (radioButton1.Checked == true)
            //{
            //    label7.Text = "티스토리 접속 링크";
            //    label10.Text = "티스토리 글작성 링크";
            //    label12.Visible = false;
            //    dateTimePicker1.Visible = false;
            //}
            Init();
            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 프로그램 시작 설정 중")));
            ChUpdate();
            StartPosition = FormStartPosition.Manual;
            Location = new Point(0, 0);
        }

        public struct POINTAPI
        {
            public int x;
            public int y;
        }

        public void ChUpdate()
        {
            try
            {
                if (System.IO.File.Exists(Application.StartupPath + "\\chromedriver.exe"))
                {
                    new ChromeDriverUpdater.ChromeDriverUpdater().Update(Application.StartupPath + "\\chromedriver.exe");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("오류 내용 : " + ex);
            }
        }

        private void Init()
        {
            foreach (Process process in ((IEnumerable<Process>)Process.GetProcesses()).Where<Process>((Func<Process, bool>)(pr => pr.ProcessName == "chromedriver")))
                process.Kill();
            //textBox1.Text = Properties.Settings.Default.id;
            //textBox2.Text = Properties.Settings.Default.pw;
            //numericUpDown1.Value = Properties.Settings.Default.bcnt;
            //textBox3.Text = Properties.Settings.Default.tconn;
            //textBox5.Text = Properties.Settings.Default.tpost;
            //textBox4.Text = Properties.Settings.Default.wconn;
            //textBox6.Text = Properties.Settings.Default.wpost;
            //numericUpDown2.Value = Properties.Settings.Default.sec;
            //dateTimePicker1.Value = Properties.Settings.Default.date;
            //numericUpDown3.Value = Properties.Settings.Default.delay;
            switch (Properties.Settings.Default.radio)
            {
                case "1":
                    radioButton1.Checked = true;
                    break;
                case "2":
                    radioButton2.Checked = true;
                    break;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                radioButton2.Checked = false;
                label7.Text = "티스토리 접속 링크";
                textBox3.Visible = true;
                textBox4.Visible = false;
                label12.Visible = false;
                dateTimePicker1.Visible = false;
                Properties.Settings.Default.radio = "1";
                Properties.Settings.Default.Save();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                radioButton1.Checked = false;
                label7.Text = "워드프레스 접속 링크";
                textBox3.Visible = false;
                textBox4.Visible = true;
                radioButton1.Checked = false;
                label12.Visible = true;
                dateTimePicker1.Visible = true;
                Properties.Settings.Default.radio = "2";
                Properties.Settings.Default.Save();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount < 1)
            {
                MessageBox.Show("검색 가능한 키워드가 없습니다.");
                return;
            }
            this.DoThread();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();
            openFileDialog1.Filter = "txt file(*.txt)|*.txt";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (dataGridView1.Rows.Count > 1)
                    dataGridView1.Rows.Clear();
                DataGridInsert(openFileDialog1.FileName, dataGridView1);
            }
        }

        private void DataGridInsert(string file, DataGridView dgv)
        {
            using (StreamReader sr = new StreamReader(file, Encoding.UTF8))
            {
                while (!sr.EndOfStream)
                {
                    try
                    {
                        string line = sr.ReadLine();
                        if (line != " ")
                        {
                            int col = dgv.Columns.Count;
                            if (line.ToString().IndexOf(",") > -1)
                            {
                                string[] dt = line.Split(new char[] { ',' });

                                this.Invoke((Action)(() => dgv.Rows.Add()));
                                for (int co = 0; co < col; co++)
                                {
                                    this.Invoke((Action)(() => dgv[co, dgv.Rows.Count - 2].Value = dt[co].Trim()));
                                }
                                this.Invoke((Action)(() => dgv.Update()));

                            }
                            else
                            {
                                string dt = line;
                                this.Invoke((Action)(() => dgv.Rows.Add()));
                                this.Invoke((Action)(() => dgv[col - 1, dgv.Rows.Count - 2].Value = dt.Trim()));
                                this.Invoke((Action)(() => dgv.Update()));
                            }
                        }
                    }
                    catch { continue; }
                }
                sr.Close();
            }
        }

        private void DataTextInsert(string file, TextBox text)
        {
            using (StreamReader sr = new StreamReader(file, Encoding.UTF8))
            {
                while (!sr.EndOfStream)
                {
                    try
                    {
                        string line = sr.ReadLine();
                        if (line != " ")
                        {
                            if (text.Text.Length < 1)
                            {
                                this.Invoke((Action)(() => text.Text = line));
                            }
                            else
                            {
                                this.Invoke((Action)(() => text.Text = text.Text + ":" + line));
                            }                            
                        }
                    }
                    catch { continue; }
                }
                sr.Close();
            }

            //if (file.IndexOf("body") > -1)
            //{
            //    if (text.Text.IndexOf(".") > -1)
            //    {
            //        text.Text = text.Text.Replace(".", "\n");
            //    }

            //    if (text.Text.IndexOf("니다.") > -1 | text.Text.IndexOf("니다") > -1)
            //    {
            //        text.Text = text.Text.Replace("니다", "니다.\n");
            //    }

            //    if (text.Text.IndexOf("데요.") > -1 | text.Text.IndexOf("데요") > -1)
            //    {
            //        text.Text = text.Text.Replace("데요", "데요.\n");
            //    }

            //    if (text.Text.IndexOf("어요.") > -1 | text.Text.IndexOf("어요") > -1)
            //    {
            //        text.Text = text.Text.Replace("어요", "어요.\n");
            //    }

            //    if (text.Text.IndexOf("해요.") > -1 | text.Text.IndexOf("해요") > -1)
            //    {
            //        text.Text = text.Text.Replace("해요", "해요.\n");
            //    }
            //    if (text.Text.IndexOf("했죠.") > -1 | text.Text.IndexOf("했죠") > -1)
            //    {
            //        text.Text = text.Text.Replace("했죠", "했죠.\n");
            //    }
            //}
            Thread.Sleep(100);
        }

        private void DataListinsert(object Data, ListBox lst)
        {
            try
            {
                this.Invoke((Action)(() => lst.Items.Add(Data)));
                this.Invoke((Action)(() => lst.Update()));
            }
            catch { }
        }

        private void DataListDelete(DataGridView dgv)
        {
            try
            {
                this.Invoke((Action)(() => dgv.Rows.RemoveAt(0)));
                this.Invoke((Action)(() => dgv.Update()));
            }
            catch { }
        }

        private void DataListDelete2(ListBox lst, int index)
        {
            try
            {
                this.Invoke((Action)(() => lst.Items.RemoveAt(index)));
                this.Invoke((Action)(() => lst.Update()));
            }
            catch { }
        }

        private void DoThread()
        {
            this.thread1 = new Thread(new ThreadStart(this.Process1));
            this.thread1.SetApartmentState(ApartmentState.STA);
            this.thread1.Start();
        }

        private void Process1()
        {
            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 프로그램 시작")));
            this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));

            bool flags = false;
            bool login = false;
            int index1 = 0;

            string sNid = string.Empty;
            string Snpw = string.Empty;
            string[] strkey = null;
            string key = string.Empty;
            int keycnt = 0;
            string platform = string.Empty;
            string link = string.Empty;
            int spcnt = 1;
            DateTime regdate = DateTime.Now;            
            int Delay = 1;
            string restitle = string.Empty;
            string resbody = string.Empty;
            string ttt = string.Empty;
            string ttb = string.Empty;
            string transUrl = "https://papago.naver.com/";

            while (index1 != dataGridView1.RowCount)
            {
                this.Invoke((Action)(() => this.listBox1.Items.Clear()));
                this.Invoke((Action)(() => this.listBox2.Items.Clear()));

                try
                {
                    this.Invoke((Action)(() => sNid = dataGridView1[1, index1].Value.ToString()));
                    this.Invoke((Action)(() => Snpw = dataGridView1[2, index1].Value.ToString()));
                    if (dataGridView1[3, index1].Value.ToString().IndexOf(",") > -1)
                    {
                        this.Invoke((Action)(() => strkey = dataGridView1[3, index1].Value.ToString().Split(new char[] { ',' },StringSplitOptions.None)));
                    }
                    else
                    {
                        this.Invoke((Action)(() => key = dataGridView1[3, index1].Value.ToString()));
                    }
                    this.Invoke((Action)(() => platform = dataGridView1[4, index1].Value.ToString()));
                    this.Invoke((Action)(() => keycnt = Convert.ToInt32(dataGridView1[5, index1].Value.ToString())));
                    this.Invoke((Action)(() => link = dataGridView1[6, index1].Value.ToString()));
                    this.Invoke((Action)(() => spcnt = Convert.ToInt32(dataGridView1[7, index1].Value.ToString())));
                    this.Invoke((Action)(() => regdate = Convert.ToDateTime(dataGridView1[8, index1].Value.ToString())));
                    this.Invoke((Action)(() => Delay = Convert.ToInt32(dataGridView1[9, index1].Value.ToString())));
                    if (strkey != null)
                    {
                        for (int kcnt = 0; kcnt < strkey.Count(); kcnt++)
                        {
                            this.Invoke((Action)(() => key = strkey[kcnt].ToString()));
                            try
                            {
                                this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss ") + key + " 키워드 검색 시작")));
                                this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                DateTime fromd = DateTime.Now.AddYears(-1);
                                DateTime tod = DateTime.Now.AddMonths(-1);
                                string kurl = "https://s.search.naver.com/p/blog/search.naver?where=blog&sm=tab_pge&api_type=1&query=" + key + "&sm=tab_opt&nso=so:r,p:from" + fromd.ToString("yyyy''MM''dd") + "to" + tod.ToString("yyyy''MM''dd");
                                Thread.Sleep(1000);

                                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                using (WebClient client = new WebClient())
                                {
                                    client.Encoding = Encoding.UTF8;
                                    string source = client.DownloadString(kurl);
                                    doc.LoadHtml(source);
                                }

                                string[] result = doc.ParsedText.Split(new string[] { "sp_blog" }, StringSplitOptions.None);
                                Thread.Sleep(100);

                                for (int index2 = 1; index2 <= keycnt; index2++)
                                {
                                    flags = false;
                                    string burl = indexParse(result[index2], "data-url=\\\"", "\\\" aria-pressed");
                                    Thread.Sleep(100);
                                    if (burl.IndexOf("naver.com") > -1)
                                    {
                                        if (listBox1.Items.Count > 1)
                                        {
                                            for (int index3 = 0; index3 < listBox1.Items.Count; index3++)
                                            {
                                                if (burl == listBox1.Items[index3].ToString())
                                                {
                                                    flags = true;
                                                    break;
                                                }
                                            }
                                            if (!flags)
                                            {
                                                this.Invoke((Action)(() => DataListinsert(burl, listBox1)));
                                                Thread.Sleep(100);
                                            }
                                        }
                                        else
                                        {
                                            this.Invoke((Action)(() => DataListinsert(burl, listBox1)));
                                            Thread.Sleep(100);
                                        }
                                    }
                                }
                                this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss ") + key + " 키워드 검색 완료")));
                                this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                flags = true;
                            }
                            catch { flags = false; }
                        }
                    }
                    else
                    {
                        try
                        {
                            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss ") + key + " 키워드 검색 시작")));
                            this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                            DateTime fromd = DateTime.Now.AddYears(-1);
                            DateTime tod = DateTime.Now.AddMonths(-1);

                            //string kurl = "https://s.search.naver.com/p/blog/search.naver?where=blog&sm=tab_pge&api_type=1&query=" + key + "&sm=tab_opt&nso=so:r,p:from" + fromd.ToString("yyyy''MM''dd") + "to" + tod.ToString("yyyy''MM''dd");
                            string kurl = "https://rss.blog.naver.com/" + key + ".xml";
                            Thread.Sleep(1000);

                            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                            using (WebClient client = new WebClient())
                            {
                                client.Encoding = Encoding.UTF8;
                                string source = client.DownloadString(kurl);
                                doc.LoadHtml(source);
                            }

                            //string[] result = doc.ParsedText.Split(new string[] { "sp_blog" }, StringSplitOptions.None);
                            string[] result = doc.ParsedText.Split(new string[] { "<item>" }, StringSplitOptions.None);
                            Thread.Sleep(100);

                            for (int index2 = 1; index2 <= keycnt; index2++)
                            {
                                flags = false;
                                string burl = indexParse(result[index2], "<link>", "</link>");
                                Thread.Sleep(100);
                                if (burl.IndexOf("naver.com") > -1)
                                {
                                    if (listBox1.Items.Count > 1)
                                    {
                                        for (int index3 = 0; index3 < listBox1.Items.Count; index3++)
                                        {
                                            if (burl == listBox1.Items[index3].ToString())
                                            {
                                                flags = true;
                                                break;
                                            }
                                        }
                                        if (!flags)
                                        {
                                            DataListinsert(burl, listBox1);
                                            Thread.Sleep(100);
                                        }
                                    }
                                    else
                                    {
                                        DataListinsert(burl, listBox1);
                                        Thread.Sleep(100);
                                    }
                                }
                            }
                            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss ") + key + " 키워드 검색 완료")));
                            this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                            flags = true;
                        }
                        catch { flags = false; }
                    }
                    if (!flags)
                    {
                        this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 조회할 링크가 없습니다.")));
                        this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                        //this.Invoke((Action)(() => DataListDelete(dataGridView1)));
                    }
                    else
                    {
                        this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 검색 시작")));
                        this.File_info(this.listBox3.Items[0].ToString());
                        flags = false;
                        login = false;
                        int index3 = 0;
                        int index5 = 0;
                        try
                        {
                            this.chromeDriverService = ChromeDriverService.CreateDefaultService();
                            this.chromeOptions = new ChromeOptions();
                            this.chromeOptions.AddArgument("--disable-extensions");
                            this.chromeOptions.AddArgument("--disable-notifications");
                            this.chromeOptions.AddArgument("window-size=1250,1050");
                            this.chromeOptions.AddArgument("window-position=680,0");
                            this.chromeOptions.AddArgument("--incognito");
                            this.chromeOptions.AddExcludedArgument("enable-automation");
                            this.chromeOptions.AddArgument("disable-infobars");
                            this.chromeDriverService.HideCommandPromptWindow = true;
                            this.driver = new ChromeDriver(this.chromeDriverService, this.chromeOptions);
                            this.driver.Manage().Cookies.DeleteAllCookies();
                            if (link.Substring(link.Length - 1, 1) == "/")
                            {
                                this.driver.Navigate().GoToUrl(link + "wp-admin");
                            }
                            else
                            {
                                this.driver.Navigate().GoToUrl(link + "/wp-admin");
                            }
                            
                            Thread.Sleep(Delay * 1000);
                            while (index3 != listBox1.Items.Count)
                            {
                                string linkurl = listBox1.Items[index3].ToString();
                                //this.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
                                //this.driver.Navigate().GoToUrl();
                                //Thread.Sleep(Delay * 1000);

                                try
                                {
                                    //this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 구글 검색 결과 확인")));
                                    //this.File_info(this.listBox3.Items[0].ToString());
                                    //IWebElement res = this.driver.FindElement(By.CssSelector("#topstuff > div.mnr-c > div > p:nth-child(1)"));
                                    //if (res.Text.IndexOf("일치하는 검색결과가 없습니다") > -1)
                                    //{
                                    this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 검색 링크 저장")));
                                    this.File_info(this.listBox3.Items[0].ToString());
                                    this.Invoke((Action)(() => this.DataListinsert(linkurl, listBox2)));
                                    Thread.Sleep(1000);

                                    this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 작성 가능 원고 추출 시작")));
                                    this.File_info(this.listBox3.Items[0].ToString());
                                    flags = false;

                                    this.Invoke((Action)(() => this.texttitle.Text = ""));
                                    this.Invoke((Action)(() => this.textbody.Text = ""));
                                    string burl = listBox2.Items[index5].ToString();
                                    try
                                    {
                                        HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();
                                        using (WebClient client2 = new WebClient())
                                        {
                                            client2.Encoding = Encoding.UTF8;
                                            string ssource = client2.DownloadString(burl);
                                            doc2.LoadHtml(ssource);
                                        }
                                        HtmlAgilityPack.HtmlNodeCollection ifnode = doc2.DocumentNode.SelectNodes("//iframe[@src]");
                                        string postview = string.Empty;
                                        foreach (var node in ifnode)
                                        {
                                            HtmlAgilityPack.HtmlAttribute attr = node.Attributes["src"];
                                            postview = attr.Value;
                                        }

                                        string vurl = "https://blog.naver.com" + postview;
                                        HtmlAgilityPack.HtmlDocument doc3 = new HtmlAgilityPack.HtmlDocument();
                                        using (WebClient client3 = new WebClient())
                                        {
                                            client3.Encoding = Encoding.UTF8;
                                            string ssource = client3.DownloadString(vurl);
                                            doc3.LoadHtml(ssource);
                                        }

                                        HtmlAgilityPack.HtmlNode setitle = doc3.DocumentNode.SelectSingleNode("//div[@class='se-module se-module-text se-title-text']");
                                        HtmlAgilityPack.HtmlNodeCollection setext = doc3.DocumentNode.SelectNodes("//div[@class='se-component se-text se-l-default']");

                                        string title = setitle.InnerText.Replace("\n", "").Trim();
                                        this.Invoke((Action)(() => this.texttitle.Text = title));

                                        foreach (var tex in setext)
                                        {
                                            try
                                            {
                                                string bodytx = tex.InnerText;
                                                string bodytxx = bodytx.Replace("\n", "").Trim();

                                                if (bodytxx != "")
                                                {
                                                    this.Invoke((Action)(() => this.textbody.Text = this.textbody.Text + bodytxx + ":"));
                                                }
                                                Thread.Sleep(100);
                                            }
                                            catch { continue; }
                                        }
                                        int returnStr = int.Parse(cntStr.Matches(textbody.Text, 0).Count.ToString());
                                        int textLentgh = textbody.Text.Trim().Length - returnStr;
                                        if (textLentgh >= spcnt)
                                        {
                                            //string[] filenm = burl.Split(new char[] { '/' }, StringSplitOptions.None);
                                            //string fnm = filenm[3] + "-" + filenm[4];
                                            string fnm = burl.Replace("https://", "").Replace("/", ".");
                                            ttt = "";
                                            ttb = "";
                                            try
                                            {
                                                if (link.IndexOf("https://") > -1)
                                                {
                                                    ttt = link.Replace("http://", "");
                                                    ttb = link.Replace("http://", "");
                                                }
                                                else
                                                {
                                                    ttt = link.Replace("http://", "");
                                                    ttb = link.Replace("http://", "");
                                                }
                                            }
                                            catch { }
                                            this.Invoke((Action)(() => this.File_save(texttitle.Text.Trim(), ttt, fnm + "_title")));
                                            this.Invoke((Action)(() => this.File_save(textbody.Text.Trim(), ttb, fnm + "_body")));
                                            Thread.Sleep(100);
                                            //index5++;
                                            flags = true;
                                        }
                                        else
                                        {
                                            this.Invoke((Action)(() => this.DataListDelete2(listBox2, index5)));
                                        }
                                        Thread.Sleep(100);
                                    }
                                    catch { }


                                    if (!flags)
                                    {
                                        this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 발행 할 수 있는 원고가 없습니다.")));
                                        this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                        //this.DoStop();
                                        Thread.Sleep(1000);
                                        //this.driver.Quit();
                                    }
                                    else
                                    {
                                        this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 원고 발행 시작")));
                                        this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                        flags = false;
                                        //int index6 = 0;
                                        this.Invoke((Action)(() => this.Script_info()));
                                        int woncount = this.script.Count();
                                        Thread.Sleep(100);

                                        if (platform == "티스토리")
                                        {
                                            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 티스토리 이동 및 로그인")));
                                            this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                            try
                                            {
                                                this.driver.Navigate().GoToUrl(link);
                                                this.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                                                Thread.Sleep(1000);

                                                while (true)
                                                {

                                                }
                                            }
                                            catch { }
                                        }
                                        else if (platform == "워드프레스")
                                        {
                                            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 워드프레스 이동 및 로그인")));
                                            this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                            try
                                            {                                                    
                                                if (!login)
                                                {
                                                    //this.driver.Navigate().GoToUrl(link + "/wp-admin");
                                                    //this.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                                                    Thread.Sleep(1000);

                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement wid = this.driver.FindElement(By.CssSelector("input[id*='user_login']"));
                                                            actions.MoveToElement(wid).Click().Perform();
                                                            Clipboard.SetText(sNid);
                                                            SendKeys.SendWait("^{v}");
                                                            Thread.Sleep(1000);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement wpw = this.driver.FindElement(By.CssSelector("input[id*='user_pass']"));
                                                            actions.MoveToElement(wpw).Click().Perform();
                                                            Clipboard.SetText(Snpw);
                                                            SendKeys.SendWait("^{v}");
                                                            Thread.Sleep(1000);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement lgsmit = this.driver.FindElement(By.CssSelector("input[id*='wp-submit']"));
                                                            actions.MoveToElement(lgsmit).Click().Perform();
                                                            this.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                                                            Thread.Sleep(1000);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    for (int idx1 = 1; idx1 < 3; idx1++)
                                                    {
                                                        Thread.Sleep(1000);
                                                        try
                                                        {
                                                            if (this.driver.FindElement(By.CssSelector("li[id*='wp-admin-bar-my-account']")).Displayed == true)
                                                            {
                                                                flags = true;
                                                                login = true;
                                                                break;
                                                            }
                                                        }
                                                        catch { }
                                                    }
                                                }
                                                else
                                                {
                                                    //flags = false;
                                                    //login = false;
                                                }
                                                Thread.Sleep(1000);

                                                if (!login)
                                                {
                                                    this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 워드프레스 로그인 실패")));
                                                    this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                                    this.Invoke((Action)(() => this.driver.Quit()));                                                        ;
                                                    Thread.Sleep(1000);
                                                }
                                                else
                                                {
                                                    this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 워드프레스 글 작성 시작")));
                                                    this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                                    flags = false;

                                                    this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss ") + (index5 + 1) + "번째 작성")));
                                                    this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                                        
                                                    if (link.Substring(link.Length - 1, 1) == "/")
                                                    {
                                                        this.driver.Navigate().GoToUrl(link + "wp-admin/post-new.php");
                                                }
                                                    else
                                                    {
                                                        this.driver.Navigate().GoToUrl(link + "/wp-admin/post-new.php");
                                                }
                                                    this.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
                                                    Thread.Sleep(1000);

                                                    try
                                                    {
                                                        if (this.driver.FindElement(By.CssSelector("div[class*='components-modal__content']")).Displayed == true)
                                                        {
                                                            while (true)
                                                            {
                                                                try
                                                                {
                                                                    Actions actions = new Actions((IWebDriver)this.driver);
                                                                    IWebElement combtn = this.driver.FindElement(By.CssSelector("button[class*='components-button has-icon']"));
                                                                    actions.MoveToElement(combtn).Click().Perform();
                                                                    Thread.Sleep(1000);
                                                                    break;
                                                                }
                                                                catch { }
                                                            }
                                                        }
                                                    }
                                                    catch { }


                                                    string postlink = listBox2.Items[index5].ToString().Replace("https://", "").Replace("/", ".");
                                                    string titlefd = Application.StartupPath + "\\script\\" + ttt + "\\" + now.ToString("yyyyMMdd") + "\\title";
                                                    string bodyfd = Application.StartupPath + "\\script\\" + ttb + "\\" + now.ToString("yyyyMMdd") + "\\body";

                                                    //this.texttitle.Text = "";
                                                    //this.textbody.Text = "";

                                                    //this.Invoke((Action)(() => this.DataTextInsert(titlefd + "\\" + postlink + "_title.txt", texttitle)));
                                                    //this.Invoke((Action)(() => this.DataTextInsert(bodyfd + "\\" + postlink + "_body.txt", textbody)));

                                                    if (this.driver.WindowHandles.Count > 1)
                                                    {
                                                        this.driver.SwitchTo().Window(driver.WindowHandles.Last());
                                                        //타이틀번역
                                                        try
                                                        {
                                                            //this.Invoke((Action)(() => restitle = this.Translate(texttitle.Text)));
                                                            //this.Invoke((Action)(() => restitle = this.indexParse(restitle, "\"translatedText\":\"", "\",\"engineType")));
                                                            //this.Invoke((Action)(() => Thread.Sleep(100)));
                                                            //this.Invoke((Action)(() => restitle = restitle.Replace(":", "\r\n")));
                                                            //this.Invoke((Action)(() => Thread.Sleep(100)));
                                                            //this.Invoke((Action)(() => texttitle.Text = restitle));
                                                            this.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(3);
                                                            this.driver.Navigate().GoToUrl(transUrl);
                                                            Thread.Sleep(1000);
                                                            while (true)
                                                            {
                                                                try
                                                                {
                                                                    Actions actions = new Actions((IWebDriver)this.driver);
                                                                    IWebElement src = this.driver.FindElement(By.CssSelector("#sourceEditArea"));
                                                                    actions.MoveToElement(src).Click().Perform();
                                                                    Thread.Sleep(1000);
                                                                    Clipboard.SetText(this.texttitle.Text);
                                                                    SendKeys.SendWait("^{v}");
                                                                    Thread.Sleep(3000);
                                                                    flags = true;
                                                                    break;
                                                                }
                                                                catch { }
                                                            }

                                                            if (!flags)
                                                            {

                                                            }
                                                            else
                                                            {
                                                                flags = false;
                                                                Actions actions = new Actions((IWebDriver)this.driver);
                                                                IWebElement copy = this.driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[1]/section/div/div[1]/div[2]/div/div[7]/span[2]/span/span/button"));
                                                                ////*[@id="root"]/div/div[1]/section/div/div[1]/div[2]/div/div[7]/span[2]/span/span/button
                                                                Clipboard.Clear();
                                                                actions.MoveToElement(copy).DoubleClick().Perform();
                                                                flags = true;

                                                                if (!flags)
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    this.Invoke((Action)(() => this.texttitle.Clear()));
                                                                    this.Invoke((Action)(() => this.texttitle.Text = Clipboard.GetText()));
                                                                    Thread.Sleep(1000);
                                                                }
                                                            }


                                                        }
                                                        catch { }
                                                        //본문번역
                                                        try
                                                        {
                                                            try
                                                            {
                                                                Actions actions = new Actions((IWebDriver)this.driver);
                                                                IWebElement cls = this.driver.FindElement(By.CssSelector("button[class*='btn_text_clse']"));
                                                                if (cls.Displayed == true)
                                                                {
                                                                    actions.MoveToElement(cls).Click().Perform();
                                                                    Thread.Sleep(100);
                                                                }
                                                            }
                                                            catch{ }
                                                            //this.Invoke((Action)(() => resbody = this.Translate(textbody.Text)));
                                                            //this.Invoke((Action)(() => resbody = this.indexParse(resbody, "\"translatedText\":\"", "\",\"engineType")));
                                                            //this.Invoke((Action)(() => Thread.Sleep(100)));
                                                            //this.Invoke((Action)(() => resbody = resbody.Replace(":", "\r\n")));
                                                            //this.Invoke((Action)(() => Thread.Sleep(100)));
                                                            //this.Invoke((Action)(() => textbody.Text = resbody));
                                                            while (true)
                                                            {
                                                                try
                                                                {
                                                                    Actions actions = new Actions((IWebDriver)this.driver);
                                                                    IWebElement src = this.driver.FindElement(By.CssSelector("#sourceEditArea"));
                                                                    //this.Invoke((Action)(() => src.Clear()));
                                                                    actions.MoveToElement(src).Click().Perform();
                                                                    Thread.Sleep(1000);
                                                                    Clipboard.SetText(this.textbody.Text);
                                                                    SendKeys.SendWait("^{v}");
                                                                    Thread.Sleep(5000);
                                                                    flags = true;
                                                                    break;
                                                                }
                                                                catch { }
                                                            }

                                                            if (!flags)
                                                            {

                                                            }
                                                            else
                                                            {
                                                                flags = false;

                                                                Form1.SetCursorPos(1333, 305);
                                                                Thread.Sleep(100);
                                                                this.Mouse_Wheel_Sroll_Up(10);
                                                                Thread.Sleep(3000);
                                                                Form1.SetCursorPos(1333, 305);
                                                                this.Mouse_Left_Click();
                                                                this.Mouse_Left_Click();
                                                                this.Mouse_Left_Click();
                                                                Thread.Sleep(100);
                                                                //this.Mouse_Wheel_Sroll_Down(-15);
                                                                //Thread.Sleep(3000);
                                                                //this.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(3);
                                                                //Actions actions = new Actions((IWebDriver)this.driver);
                                                                //IWebElement copy = this.driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[1]/section/div/div[1]/div[2]/div/div[7]/span[2]/span/span/button"));
                                                                //////*[@id="root"]/div/div[1]/section/div/div[1]/div[2]/div/div[7]/span[2]/span/span/button
                                                                //Clipboard.Clear();
                                                                //actions.MoveToElement(copy).DoubleClick().Perform();

                                                                SendKeys.SendWait("^{c}");
                                                                Thread.Sleep(1000);
                                                                flags = true;

                                                                if (!flags)
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    this.Invoke((Action)(() => this.textbody.Clear()));
                                                                    this.Invoke((Action)(() => this.textbody.Text = Clipboard.GetText()));
                                                                    this.Invoke((Action)(() => this.textbody.Text = textbody.Text.Replace(":", "\r\n")));
                                                                    Thread.Sleep(1000);
                                                                }
                                                            }
                                                        }
                                                        catch { }
                                                    }
                                                    else
                                                    {
                                                        var js = $"window.open('{transUrl}','_blank');";
                                                        ((IJavaScriptExecutor)driver).ExecuteScript(js);
                                                        Thread.Sleep(100);
                                                        this.driver.SwitchTo().Window(driver.WindowHandles.Last());
                                                        //타이틀번역
                                                        try
                                                        {
                                                            //this.Invoke((Action)(() => restitle = this.Translate(texttitle.Text)));
                                                            //this.Invoke((Action)(() => restitle = this.indexParse(restitle, "\"translatedText\":\"", "\",\"engineType")));
                                                            //this.Invoke((Action)(() => Thread.Sleep(100)));
                                                            //this.Invoke((Action)(() => restitle = restitle.Replace(":", "\r\n")));
                                                            //this.Invoke((Action)(() => Thread.Sleep(100)));
                                                            //this.Invoke((Action)(() => texttitle.Text = restitle));
                                                            this.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(3);
                                                            //this.driver.Navigate().GoToUrl("https://naver.papago.com");
                                                            Thread.Sleep(1000);
                                                            while (true)
                                                            {
                                                                try
                                                                {
                                                                    Actions actions = new Actions((IWebDriver)this.driver);
                                                                    IWebElement src = this.driver.FindElement(By.CssSelector("#sourceEditArea"));
                                                                    actions.MoveToElement(src).Click().Perform();
                                                                    Thread.Sleep(1000);
                                                                    Clipboard.SetText(this.texttitle.Text);
                                                                    SendKeys.SendWait("^{v}");
                                                                    Thread.Sleep(3000);
                                                                    flags = true;
                                                                    break;
                                                                }
                                                                catch { }
                                                            }

                                                            if (!flags)
                                                            {

                                                            }
                                                            else
                                                            {
                                                                flags = false;
                                                                Actions actions = new Actions((IWebDriver)this.driver);
                                                                IWebElement copy = this.driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[1]/section/div/div[1]/div[2]/div/div[7]/span[2]/span/span/button"));
                                                                ////*[@id="root"]/div/div[1]/section/div/div[1]/div[2]/div/div[7]/span[2]/span/span/button
                                                                Clipboard.Clear();
                                                                actions.MoveToElement(copy).DoubleClick().Perform();
                                                                Thread.Sleep(3000);
                                                                flags = true;
                                                                
                                                                if (!flags)
                                                                {

                                                                }  
                                                                else
                                                                {
                                                                    this.Invoke((Action)(() => this.texttitle.Clear()));
                                                                    this.Invoke((Action)(() => this.texttitle.Text = Clipboard.GetText()));
                                                                    Thread.Sleep(1000);
                                                                }
                                                            }


                                                        }
                                                        catch { }
                                                        //본문번역
                                                        try
                                                        {
                                                            try
                                                            {
                                                                Actions actions = new Actions((IWebDriver)this.driver);
                                                                IWebElement cls = this.driver.FindElement(By.CssSelector("button[class*='btn_text_clse']"));
                                                                if (cls.Displayed == true)
                                                                {
                                                                    actions.MoveToElement(cls).Click().Perform();
                                                                    Thread.Sleep(100);
                                                                }
                                                            }
                                                            catch { }
                                                            //this.Invoke((Action)(() => resbody = this.Translate(textbody.Text)));
                                                            //this.Invoke((Action)(() => resbody = this.indexParse(resbody, "\"translatedText\":\"", "\",\"engineType")));
                                                            //this.Invoke((Action)(() => Thread.Sleep(100)));
                                                            //this.Invoke((Action)(() => resbody = resbody.Replace(":", "\r\n")));
                                                            //this.Invoke((Action)(() => Thread.Sleep(100)));
                                                            //this.Invoke((Action)(() => textbody.Text = resbody));
                                                            while (true)
                                                            {
                                                                try
                                                                {
                                                                    Actions actions = new Actions((IWebDriver)this.driver);
                                                                    IWebElement src = this.driver.FindElement(By.CssSelector("#sourceEditArea"));
                                                                    //this.Invoke((Action)(() => src.Clear()));
                                                                    actions.MoveToElement(src).Click().Perform();
                                                                    Thread.Sleep(1000);
                                                                    Clipboard.SetText(this.textbody.Text);
                                                                    SendKeys.SendWait("^{v}");
                                                                    Thread.Sleep(5000);
                                                                    flags = true;
                                                                    break;
                                                                }
                                                                catch { }
                                                            }

                                                            if (!flags)
                                                            {

                                                            }
                                                            else
                                                            {
                                                                flags = false;
                                                                Form1.SetCursorPos(1333, 305);
                                                                Thread.Sleep(100);
                                                                this.Mouse_Wheel_Sroll_Up(10);
                                                                Thread.Sleep(3000);
                                                                Form1.SetCursorPos(1333, 305);
                                                                this.Mouse_Left_Click();
                                                                this.Mouse_Left_Click();
                                                                this.Mouse_Left_Click();
                                                                Thread.Sleep(100);
                                                                //this.Mouse_Wheel_Sroll_Down(-15);
                                                                //Thread.Sleep(3000);
                                                                //this.driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(3);
                                                                //Actions actions = new Actions((IWebDriver)this.driver);
                                                                //IWebElement copy = this.driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[1]/section/div/div[1]/div[2]/div/div[7]/span[2]/span/span/button"));
                                                                //////*[@id="root"]/div/div[1]/section/div/div[1]/div[2]/div/div[7]/span[2]/span/span/button
                                                                //Clipboard.Clear();
                                                                //actions.MoveToElement(copy).DoubleClick().Perform();
                                                              
                                                                SendKeys.SendWait("^{c}");
                                                                Thread.Sleep(1000);
                                                                flags = true;

                                                                if (!flags)
                                                                {

                                                                }
                                                                else
                                                                {
                                                                    this.Invoke((Action)(() => this.textbody.Clear()));
                                                                    this.Invoke((Action)(() => this.textbody.Text = Clipboard.GetText()));
                                                                    this.Invoke((Action)(() => this.textbody.Text = textbody.Text.Replace(":","\r\n")));
                                                                    Thread.Sleep(1000);
                                                                }
                                                            }
                                                        }
                                                        catch { }
                                                    }

                                                    this.driver.SwitchTo().Window(this.driver.WindowHandles.First());

                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement tit = this.driver.FindElement(By.CssSelector("h1[class*='wp-block wp-block-post-title block-editor-block-list__block editor-post-title editor-post-title__input rich-text']"));
                                                            actions.MoveToElement(tit).Click().Perform();
                                                            Clipboard.SetText(this.texttitle.Text);
                                                            SendKeys.SendWait("^{v}");
                                                            Thread.Sleep(1000);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement tbody1 = null;
                                                            IWebElement tbody2 = null;
                                                            try
                                                            {
                                                                tbody1 = this.driver.FindElement(By.CssSelector("p[class*='block-editor-default-block-appender__content']"));
                                                                actions.MoveToElement(tbody1).Click().Perform();
                                                                Clipboard.SetText(this.textbody.Text);
                                                                Thread.Sleep(1000);
                                                                SendKeys.SendWait("^{v}");
                                                                Thread.Sleep(1000);
                                                                break;                                                                                                                                        
                                                            }
                                                            catch { }
                                                            //block-editor-default-block-appender__content
                                                            //block-editor-rich-text__editable block-editor-block-list__block wp-block is-selected wp-block-paragraph rich-text                                                                
                                                            try
                                                            {
                                                                tbody2 = this.driver.FindElement(By.CssSelector("p[class*='block-editor-rich-text__editable block-editor-block-list__block wp-block is-selected wp-block-paragraph rich-text']"));
                                                                actions.MoveToElement(tbody2).Click().Perform();
                                                                Clipboard.SetText(this.textbody.Text);
                                                                Thread.Sleep(1000);
                                                                SendKeys.SendWait("^{v}");
                                                                Thread.Sleep(1000);
                                                                break;
                                                            }
                                                            catch { }
                                                        }
                                                        catch { }
                                                    }

                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement sidbar1 = this.driver.FindElement(By.CssSelector("button[class*='components-button edit-post-sidebar__panel-tab']"));
                                                            actions.MoveToElement(sidbar1).Click().Perform();
                                                            Thread.Sleep(1000);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement combodown = this.driver.FindElement(By.CssSelector("button[class*='components-button edit-post-post-schedule__toggle is-tertiary']"));
                                                            actions.MoveToElement(combodown).Click().Perform();
                                                            Thread.Sleep(1000);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    this.Invoke((Action)(() => regdate = regdate.AddHours(6 * (index5 + 1))));
                                                    string year = regdate.ToString("yyyy");
                                                    string mont = regdate.ToString("MM");
                                                    string day = regdate.ToString("dd");
                                                    string Hhch = regdate.ToString("tt");

                                                    //발행 년도
                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement yy = this.driver.FindElement(By.CssSelector("div[class*='components-datetime__time-field-year']"));
                                                            actions.MoveToElement(yy).Click().Perform();
                                                            //yy.Clear();
                                                            Thread.Sleep(1000);
                                                            SendKeys.SendWait("^{a}");
                                                            Thread.Sleep(100);
                                                            SendKeys.SendWait("{DEL}");
                                                            Thread.Sleep(100);
                                                            Clipboard.SetText(year);
                                                            SendKeys.SendWait("^{v}");
                                                            Thread.Sleep(100);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    //발행 월
                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement Mth = this.driver.FindElement(By.CssSelector("div[class*='components-datetime__time-field-month']"));
                                                            actions.MoveToElement(Mth).Click().Perform();
                                                            var opt = 0;
                                                            int op = 0;
                                                            try
                                                            {
                                                                opt = this.driver.FindElements(By.XPath("/html/body/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div/div[2]/div/div/div/div[2]/div[1]/fieldset[2]/div/div[1]/div/div/div/div/select/option")).Count();
                                                            }
                                                            catch { }
                                                            try
                                                            {
                                                                if (opt == 0)
                                                                {
                                                                    opt = this.driver.FindElements(By.XPath("/html/body/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div/div[2]/div/div/div/div[2]/div[1]/fieldset[2]/div/div[2]/div/div/div/div/select/option")).Count();
                                                                    op = 1;
                                                                }
                                                            }
                                                            catch { }
                                                            ////*[@id="inspector-select-control-11"]/option[1]
                                                            for (int idx1 = 1; idx1 <= opt; idx1++)
                                                            {
                                                                try
                                                                {
                                                                    var mmele = "";
                                                                    if (op == 0)
                                                                    {
                                                                        mmele = this.driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div/div[2]/div/div/div/div[2]/div[1]/fieldset[2]/div/div[1]/div/div/div/div/select/option[" + idx1 + "]")).Text;
                                                                    }
                                                                    else
                                                                    {
                                                                        mmele = this.driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div/div[2]/div/div/div/div[2]/div[1]/fieldset[2]/div/div[2]/div/div/div/div/select/option[" + idx1 + "]")).Text;
                                                                    }
                                                                    //
                                                                    if (mont.Substring(0, 1) == "0")
                                                                    {
                                                                        this.Invoke((Action)(() => mont = mont.Replace("0", "")));
                                                                    }
                                                                    var mmont = mont + "월";
                                                                    if (mmele == mmont)
                                                                    {
                                                                        Actions act = new Actions((IWebDriver)this.driver);
                                                                        IWebElement mt = null;
                                                                        if (op == 0)
                                                                        {
                                                                                mt = this.driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div/div[2]/div/div/div/div[2]/div[1]/fieldset[2]/div/div[1]/div/div/div/div/select/option[" + idx1 + "]"));
                                                                        }
                                                                        else
                                                                        {
                                                                            mt = this.driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/div/div[2]/div/div/div/div[2]/div[1]/fieldset[2]/div/div[2]/div/div/div/div/select/option[" + idx1 + "]"));
                                                                        }
                                                                        mt.Click();
                                                                        break;
                                                                    }
                                                                }
                                                                catch { }
                                                            }
                                                            Thread.Sleep(100);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    //발행 일
                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement dd = this.driver.FindElement(By.CssSelector("div[class*='components-datetime__time-field components-datetime__time-field-day']"));
                                                            actions.MoveToElement(dd).Click().Perform();
                                                            //this.Invoke((Action)(() => dd.Clear()));
                                                            SendKeys.SendWait("^{a}");
                                                            Thread.Sleep(100);
                                                            SendKeys.SendWait("{DEL}");
                                                            Thread.Sleep(100);
                                                            Clipboard.SetText(day);
                                                            SendKeys.SendWait("^{v}");
                                                            Thread.Sleep(100);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    //발행 시간
                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement hour = this.driver.FindElement(By.CssSelector("div[class*='components-datetime__time-field-hours-input']"));
                                                            actions.MoveToElement(hour).Click().Perform();
                                                            SendKeys.SendWait("^{a}");
                                                            Thread.Sleep(100);
                                                            SendKeys.SendWait("{DEL}");
                                                            Thread.Sleep(100);
                                                            Clipboard.SetText(regdate.ToString("hh"));
                                                            SendKeys.SendWait("^{v}");
                                                            Thread.Sleep(100);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    //발행 분
                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement sec = this.driver.FindElement(By.CssSelector("div[class*='components-datetime__time-field-minutes-input']"));
                                                            actions.MoveToElement(sec).Click().Perform();
                                                            SendKeys.SendWait("^{a}");
                                                            Thread.Sleep(100);
                                                            SendKeys.SendWait("{DEL}");
                                                            Thread.Sleep(100);
                                                            Clipboard.SetText(regdate.ToString("mm"));
                                                            SendKeys.SendWait("^{v}");
                                                            Thread.Sleep(100);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    //발행 오전/오후
                                                    try
                                                    {
                                                        if (this.driver.FindElement(By.CssSelector("button[class*='components-datetime__time-am-button is-secondary']")).Displayed == true)                                                                
                                                            while (true)
                                                            {
                                                                if (Hhch == "오전")
                                                                {
                                                                    try
                                                                    {
                                                                        Actions actions = new Actions((IWebDriver)this.driver);
                                                                        IWebElement tt = this.driver.FindElement(By.CssSelector("button[class*='components-datetime__time-am-button is-secondary']"));
                                                                        actions.MoveToElement(tt).Click().Perform();
                                                                        Thread.Sleep(100);
                                                                        break;
                                                                    }
                                                                    catch { }
                                                                }
                                                                else
                                                                {
                                                                    try
                                                                    {
                                                                        Actions actions = new Actions((IWebDriver)this.driver);
                                                                        IWebElement tt = this.driver.FindElement(By.CssSelector("button[class*='components-datetime__time-pm-button is-primary']"));
                                                                        actions.MoveToElement(tt).Click().Perform();
                                                                        Thread.Sleep(100);
                                                                        break;
                                                                    }
                                                                    catch { }
                                                                }
                                                            }
                                                    }
                                                    catch { }
                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement summit = this.driver.FindElement(By.CssSelector("button[class*='components-button editor-post-publish-panel__toggle editor-post-publish-button__button is-primary']"));
                                                            actions.MoveToElement(summit).Click().Perform();
                                                            Thread.Sleep(1000);
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    while (true)
                                                    {
                                                        try
                                                        {
                                                            Actions actions = new Actions((IWebDriver)this.driver);
                                                            IWebElement resummit = this.driver.FindElement(By.CssSelector("button[class*='components-button editor-post-publish-button editor-post-publish-button__button is-primary']"));
                                                            actions.MoveToElement(resummit).Click().Perform();
                                                            Thread.Sleep(3000);
                                                            flags = true;
                                                            break;
                                                        }
                                                        catch { }
                                                    }

                                                    if (!flags)
                                                    {
                                                        this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss ") + (index5 + 1) + "번째 글 작성 실패")));
                                                        this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                                        this.Invoke((Action)(() => this.DataListDelete2(listBox2, index5)));
                                                        this.Invoke((Action)(() => Thread.Sleep(Delay * 1000)));
                                                    }
                                                    else
                                                    {
                                                        this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss ") + (index5 + 1) + "번째 글 작성 성공")));
                                                        this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                                                        index5++;
                                                        this.Invoke((Action)(() => Thread.Sleep(Delay * 1000)));
                                                    }

                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                    //}
                                }
                                catch { }

                                this.Invoke((Action)(() => this.DataListDelete2(listBox1, index3)));
                                this.Invoke((Action)(() => Thread.Sleep(Delay * 1000)));
                                //this.Invoke((Action)(() => this.driver.Close()));
                                if (listBox1.Items.Count == index3)
                                {
                                    break;                                    
                                }
                            }
                            //flags = true;

                        }
                        catch { }

                        //if (!flags)
                        //{
                        //    this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 검색 가능한 링크가 없습니다.")));
                        //    this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                        //    this.Invoke((Action)(() => this.DoStop()));
                        //    this.Invoke((Action)(() => Thread.Sleep(1000)));
                        //    this.Invoke((Action)(() => this.driver.Quit()));
                        //}
                        //else
                        //{

                        //}
                    }
                }
                catch { }
                this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") +  " " + link.Replace("http://", "").Replace("/","")  +" 글 발행 완료.")));
                this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
                this.Invoke((Action)(() => this.DataListDelete(dataGridView1)));
                this.Invoke((Action)(() => this.driver.Quit()));
            }

            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 모든 글 작성 완료")));
            this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
            this.Invoke((Action)(() => this.File_save_list(listBox2)));
            this.Invoke((Action)(() => this.driver.Quit()));
            //this.Invoke((Action)(() => this.DoStop()));

            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 프로그램 종료")));
            this.Invoke((Action)(() => this.File_info(this.listBox3.Items[0].ToString())));
        }

        private void DoStop()
        {
            try
            {
                this.thread1.Interrupt();
                this.thread1.Abort();
            }
            catch { }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 프로그램 중지중")));
            this.DoStop();
            Thread.Sleep(3000);
            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 프로그램 중지되었습니다.")));
        }

        public string indexParse(string Data, string index1, string index2)
        {
            int id1 = Data.IndexOf(index1, 0) + index1.Length;
            int id2 = Data.IndexOf(index2);
            try
            {
                string str = Data.Substring(id1, id2 - id1);
                str = Regex.Replace(str, @"[^a-zA-Z0-9가-힣_\W]", "", RegexOptions.IgnoreCase | RegexOptions.Singleline);

                return str;
            }
            catch
            {
                string str = null;
                str = " ";

                return str;
            }
        }

        public bool File_info(string strMsg)
        {
            try
            {
                string strChkFolder = "";
                string strFilename = "";
                string strLocal = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\\"));

                strChkFolder = strLocal + "\\Log";
                if (!System.IO.Directory.Exists(strChkFolder))
                {
                    System.IO.Directory.CreateDirectory(strChkFolder);
                }

                strFilename = strChkFolder + "\\" + DateTime.Now.ToString("yyyyMMdd") + "_log.txt";

                System.IO.StreamWriter FileWriter = new System.IO.StreamWriter(strFilename, true);
                FileWriter.Write(strMsg + "\r\n");
                FileWriter.Flush();
                FileWriter.Close();
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool File_save_list(ListBox lst)
        {
            try
            {
                string strChkFolder = "";
                //string strFilename = "";
                string strLocal = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\\"));

                strChkFolder = Application.StartupPath + "\\script\\" + now.ToString("yyyyMMdd") + "\\won";
                if (!System.IO.Directory.Exists(strChkFolder))
                {
                    System.IO.Directory.CreateDirectory(strChkFolder);
                }

                using (StreamWriter streamWriter = new StreamWriter(strChkFolder + "\\" + DateTime.Now.ToString("yyyyMMdd") + "_script.txt"))
                {
                    for (int index10 = 0; index10 < lst.Items.Count; index10++)
                    {
                        string svdata = lst.Items[index10].ToString();

                        streamWriter.Write(svdata + "\r\n");
                        streamWriter.Flush();
                    }
                    streamWriter.Close();
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool File_save(string strMsg, string id, string filename)
        {
            try
            {
                string strChkFolder = "";
                string strFilename = "";
                string strLocal = Application.ExecutablePath.Substring(0, Application.ExecutablePath.LastIndexOf("\\"));

                strChkFolder = Application.StartupPath + "\\script\\" + id + "\\" + now.ToString("yyyyMMdd");
                if (!System.IO.Directory.Exists(strChkFolder))
                {
                    System.IO.Directory.CreateDirectory(strChkFolder);
                }

                if (filename.IndexOf("title") > -1)
                {
                    if (!System.IO.Directory.Exists(strChkFolder + "\\title"))
                    {
                        System.IO.Directory.CreateDirectory(strChkFolder + "\\title");
                        strChkFolder = strChkFolder + "\\title";
                        strFilename = strChkFolder + "\\" + filename + ".txt";
                    }
                    else
                    {
                        strChkFolder = strChkFolder + "\\title";
                        strFilename = strChkFolder + "\\" + filename + ".txt";
                    }
                }
                else
                {
                    if (!System.IO.Directory.Exists(strChkFolder + "\\body")) {
                        System.IO.Directory.CreateDirectory(strChkFolder + "\\body");
                        strChkFolder = strChkFolder + "\\body";
                        strFilename = strChkFolder + "\\" + filename + ".txt";
                    }
                    else
                    {
                        strChkFolder = strChkFolder + "\\body";
                        strFilename = strChkFolder + "\\" + filename + ".txt";
                    }
                }
                //strFilename = strChkFolder + "\\" + filename + ".txt";

                System.IO.StreamWriter FileWriter = new System.IO.StreamWriter(strFilename, true);
                FileWriter.Write(strMsg + "\r\n");
                FileWriter.Flush();
                FileWriter.Close();
            }
            catch
            {
                return false;
            }

            return true;
        }

        private void Script_info()
        {
            if (!System.IO.Directory.Exists(Application.StartupPath + "\\script\\" + DateTime.Now.ToString("yyyyMMdd") + "\\body"))
            {
                this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 원고 저장 폴더를 찾을 수 없습니다.")));
                this.File_info(this.listBox3.Items[0].ToString());
                return;
            }

            String FolderName = Application.StartupPath + "\\script\\" + DateTime.Now.ToString("yyyyMMdd") + "\\body";
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(FolderName);
            foreach (System.IO.FileInfo File in di.GetFiles())
            {
                if (File.Extension.ToLower().CompareTo(".txt") == 0)
                {
                    String FileNameOnly = File.Name.Substring(0, File.Name.Length - 4);
                    String FullFileName = File.FullName;

                    this.Invoke((Action)(() => script.Add(FullFileName)));
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.chromeDriverService = ChromeDriverService.CreateDefaultService();
            this.chromeOptions = new ChromeOptions();
            this.chromeOptions.AddArgument("--disable-extensions");
            this.chromeOptions.AddArgument("--disable-notifications");
            this.chromeOptions.AddArgument("window-size=1250,1050");
            this.chromeOptions.AddArgument("window-position=680,0");
            this.chromeOptions.AddArgument("--incognito");
            this.chromeOptions.AddExcludedArgument("enable-automation");
            this.chromeOptions.AddArgument("disable-infobars");
            this.chromeDriverService.HideCommandPromptWindow = true;
            this.driver = new ChromeDriver(this.chromeDriverService, this.chromeOptions);
            this.driver.Manage().Cookies.DeleteAllCookies();
            this.driver.Navigate().GoToUrl("https://naver.com");
            Thread.Sleep(1000);
            var js = $"window.open('{"https://popago.com"}','_blank');";
            ((IJavaScriptExecutor)driver).ExecuteScript(js);
            Thread.Sleep(1000);

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //if (textBox1.TextLength > 0 & e.KeyCode == System.Windows.Forms.Keys.Enter)
            //{
            //    Properties.Settings.Default.id = textBox1.Text;
            //    Properties.Settings.Default.Save();
            //}
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            //if (textBox2.TextLength > 0 & e.KeyCode == System.Windows.Forms.Keys.Enter)
            //{
            //    Properties.Settings.Default.pw = textBox2.Text;
            //    Properties.Settings.Default.Save();
            //}
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.bcnt = numericUpDown1.Value;
            //Properties.Settings.Default.Save();
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.sec = numericUpDown2.Value;
            //Properties.Settings.Default.Save();
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            //if (textBox3.TextLength > 0 & e.KeyCode == System.Windows.Forms.Keys.Enter)
            //{
            //    Properties.Settings.Default.tconn = textBox3.Text;
            //    Properties.Settings.Default.Save();
            //}
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            //if (textBox4.TextLength > 0 & e.KeyCode == System.Windows.Forms.Keys.Enter)
            //{
            //    Properties.Settings.Default.wconn = textBox4.Text;
            //    Properties.Settings.Default.Save();
            //}
        }

        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.date = dateTimePicker1.Value;
            //Properties.Settings.Default.Save();
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.delay = numericUpDown3.Value;
            //Properties.Settings.Default.Save();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            StringFormat drawFormat = new StringFormat();
            drawFormat.FormatFlags = StringFormatFlags.DirectionRightToLeft;
            using (Brush brush = new SolidBrush(Color.Red))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font
                    , brush, e.RowBounds.Location.X + 35, e.RowBounds.Location.Y + 4, drawFormat);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                if (textBox1.Text != "" & textBox1.TextLength > 1 & textBox2.Text != "" & textBox2.TextLength > 1
                    & textBox3.Text != "" & textBox3.TextLength > 1 & textBox5.Text != "" & textBox5.TextLength > 1)
                {
                    int cnt = this.dataGridView1.RowCount;
                    Invoke((Action)(() => this.dataGridView1.Rows.Add()));
                    Invoke((Action)(() => this.dataGridView1[1, cnt].Value = textBox1.Text));
                    Invoke((Action)(() => this.dataGridView1[2, cnt].Value = textBox2.Text));
                    Invoke((Action)(() => this.dataGridView1[3, cnt].Value = textBox5.Text));
                    Invoke((Action)(() => this.dataGridView1[4, cnt].Value = "티스토리"));
                    Invoke((Action)(() => this.dataGridView1[5, cnt].Value = numericUpDown1.Value));
                    Invoke((Action)(() => this.dataGridView1[6, cnt].Value = textBox3.Text));
                    Invoke((Action)(() => this.dataGridView1[7, cnt].Value = numericUpDown2.Value));
                    Invoke((Action)(() => this.dataGridView1[8, cnt].Value = dateTimePicker1.Value));
                    Invoke((Action)(() => this.dataGridView1[9, cnt].Value = numericUpDown3.Value));
                    Invoke((Action)(() => this.dataGridView1.Update()));
                }
                TextInit();
            }
            else if (radioButton2.Checked == true)
            {
                if (textBox1.Text != "" & textBox1.TextLength > 1 & textBox2.Text != "" & textBox2.TextLength > 1
                    & textBox4.Text != "" & textBox4.TextLength > 1 & textBox5.Text != "" & textBox5.TextLength > 1)
                {
                    int cnt = this.dataGridView1.RowCount;
                    Invoke((Action)(() => this.dataGridView1.Rows.Add()));
                    Invoke((Action)(() => this.dataGridView1[1, cnt].Value = textBox1.Text));
                    Invoke((Action)(() => this.dataGridView1[2, cnt].Value = textBox2.Text));
                    Invoke((Action)(() => this.dataGridView1[3, cnt].Value = textBox5.Text));
                    Invoke((Action)(() => this.dataGridView1[4, cnt].Value = "워드프레스"));
                    Invoke((Action)(() => this.dataGridView1[5, cnt].Value = numericUpDown1.Value));
                    Invoke((Action)(() => this.dataGridView1[6, cnt].Value = textBox4.Text));
                    Invoke((Action)(() => this.dataGridView1[7, cnt].Value = numericUpDown2.Value));
                    Invoke((Action)(() => this.dataGridView1[8, cnt].Value = dateTimePicker1.Value));
                    Invoke((Action)(() => this.dataGridView1[9, cnt].Value = numericUpDown3.Value));
                    Invoke((Action)(() => this.dataGridView1.Update()));
                }
                TextInit();
            }
        }

        private void TextInit()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            numericUpDown1.Value = 1;
            numericUpDown2.Value = 500;
            numericUpDown3.Value = 1;            
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int index1 = 0; index1 < dataGridView1.RowCount; index1++)
            {
                try
                {
                    if (this.dataGridView1[0, index1].Value.ToString() == "true")
                    {
                        Invoke((Action)(() => this.dataGridView1.Rows.RemoveAt(index1)));
                        Invoke((Action)(() => this.dataGridView1.Update()));
                    }
                }
                catch { }
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount < 1)
            {
                MessageBox.Show("검색 가능한 정보가 없습니다.");
                return;
            }

            this.DoThread();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 프로그램 중지중")));
            this.DoStop();
            Thread.Sleep(3000);
            this.Invoke((Action)(() => this.listBox3.Items.Insert(0, DateTime.Now.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss") + " 프로그램 중지되었습니다.")));
        }

        public string Translate(string TextBody)
        {
            try
            {
                string url = "https://openapi.naver.com/v1/papago/n2mt";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Headers.Add("X-Naver-Client-Id", "CW911KiVEpGj6D60KwB5");
                request.Headers.Add("X-Naver-Client-Secret", "WFSFjkQRUI");
                request.Method = "POST";
                string query = TextBody;
                byte[] byteDataParams = Encoding.UTF8.GetBytes("source=ko&target=en&text=" + query);
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = byteDataParams.Length;
                Stream st = request.GetRequestStream();
                st.Write(byteDataParams, 0, byteDataParams.Length);
                st.Close();
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream stream = response.GetResponseStream();
                StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                string text = reader.ReadToEnd();
                stream.Close();
                response.Close();
                reader.Close();
                return text;
            }
            catch { return null; }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                foreach (Process process in ((IEnumerable<Process>)Process.GetProcesses()).Where<Process>((Func<Process, bool>)(pr => pr.ProcessName == "chrome")))
                    process.Kill();
            }
            catch { }
            try
            {
                foreach (Process process in ((IEnumerable<Process>)Process.GetProcesses()).Where<Process>((Func<Process, bool>)(pr => pr.ProcessName == "chromedriver")))
                    process.Kill();
            }
            catch { }
        }
    }
}
