using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Autodesk.Revit.UI;
using System.Runtime.InteropServices;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Net;

namespace SinoTunnel
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public static string path
        {
            get;
            set;
        }

        public static string name
        {
            get;
            set;
        }

        public static string frame_path
        {
            get;
            set;
        }

        public static string counting_temp_path
        {
            get;
            set;
        }

        public static string saveas_FilePath
        {
            get;
            set;
        }

        public static bool bolt
        {
            get;
            set;
        }

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // height of ellipse
            int nHeightEllipse // width of ellipse
        );

        //拖曳用參數
        int mov;
        int movX;
        int movY;

        //外部事件處理:讀取其他的cs檔
        ExternalEvent externalEvent_attached_pipeline;
        Attached_pipeline handler_attached_pipeline = new Attached_pipeline();

        ExternalEvent externalEvent_contact_channel;
        Contact_channel handler_contact_channel = new Contact_channel();

        ExternalEvent externalEvent_envelope;
        Envelope handler_envelope = new Envelope();

        ExternalEvent externalEvent_inverted_arc;
        inverted_arc handler_inverted_arc = new inverted_arc();

        ExternalEvent externalEvent_place_adative_circle;
        place_adative_circle handler_place_adative_circle = new place_adative_circle();

        ExternalEvent externalEvent_track_bed;
        Track_bed handler_track_bed = new Track_bed();

        ExternalEvent externalEvent_U_shape_steel;
        U_shape_steel handler_U_shape_steel = new U_shape_steel();

        ExternalEvent externalEvent_section_for_test;
        section_for_test handler_section_for_test = new section_for_test();

        ExternalEvent externalEvent_testhao;
        testhao handler_testhao = new testhao();

        ExternalEvent externalEvent_rebar_in_circle;
        rebar_in_circle handler_rebar_in_circle = new rebar_in_circle();

        ExternalEvent externalEvent_shear_rebar;
        shear_rebar handler_shear_rebar = new shear_rebar();

        ExternalEvent externalEvent_Counting; //***
        Counting handler_Counting = new Counting(); //***


        //set color
        Color select_color = Color.FromArgb(31, 123, 221);
        Color option_color = Color.FromArgb(106, 181, 255);
        
        public Form1(UIDocument uIDocument)
        {
            InitializeComponent();
            FormBorderStyle = FormBorderStyle.None;
            Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));

            //建立外部事件
            externalEvent_attached_pipeline = ExternalEvent.Create(handler_attached_pipeline);
            externalEvent_contact_channel = ExternalEvent.Create(handler_contact_channel);
            externalEvent_envelope = ExternalEvent.Create(handler_envelope);
            externalEvent_inverted_arc = ExternalEvent.Create(handler_inverted_arc);
            externalEvent_place_adative_circle = ExternalEvent.Create(handler_place_adative_circle);
            externalEvent_track_bed = ExternalEvent.Create(handler_track_bed);
            externalEvent_U_shape_steel = ExternalEvent.Create(handler_U_shape_steel);
            externalEvent_section_for_test = ExternalEvent.Create(handler_section_for_test);
            externalEvent_testhao = ExternalEvent.Create(handler_testhao);
            externalEvent_rebar_in_circle = ExternalEvent.Create(handler_rebar_in_circle);
            externalEvent_shear_rebar = ExternalEvent.Create(handler_shear_rebar);
            externalEvent_Counting = ExternalEvent.Create(handler_Counting); //***
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            HttpClient client = client_login();
            if (client == null)
            {
                label5.Text = "- 認證失敗，10秒後強制關閉";
                this.Enabled = false;
                timer1.Start();
            }
            var result = client.GetAsync($"/user/me").Result;
            if (result.IsSuccessStatusCode)
            {
                string s = result.Content.ReadAsStringAsync().Result;
                label5.Text += " " + DecodeEncodedNonAsciiCharacters(s.Substring(8, s.Length - 10));
            }
            else
            {
                //MessageBox.Show("認證失敗，5秒後強制關閉");
                label5.Text = "- 認證失敗，10秒後強制關閉";
                this.Enabled = false;
                timer1.Start();
                //System.Threading.Thread.Sleep(10000);
                //this.Dispose();
            }
            //找出字體大小,並算出比例
            float dpiX, dpiY;
            Graphics graphics = this.CreateGraphics();
            dpiX = graphics.DpiX;
            dpiY = graphics.DpiY;
            int intPercent = (dpiX == 96) ? 100 : (dpiX == 120) ? 125 : 150;

            // 針對字體變更Form的大小
            this.Height = this.Height * intPercent / 100;
            this.Width = this.Width * intPercent / 100;
            this.Size = new System.Drawing.Size(this.gradientPanel1.Size.Width + 10 , this.gradientPanel1.Size.Height + this.gradientPanel2.Height + 10);

            Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
            BG_panel.Location = Point.Empty;
            BG_panel.Size = this.Size;
           
            load_button.ForeColor = Color.MidnightBlue;

            load_panel.Show();
            build_panel.Hide();
            result_panel.Hide();

            int start_position = gradientPanel2.Location.X + gradientPanel2.Size.Width;
            int start_position_y = gradientPanel1.Location.Y + gradientPanel1.Height;

            load_panel.Location = new Point(start_position, start_position_y);
            build_panel.Location = new Point(start_position, start_position_y);
            result_panel.Location = new Point(start_position, start_position_y);
        }
        

        private void load_button_Click(object sender, EventArgs e)
        {
            load_button.ForeColor = Color.MidnightBlue;
            build_button.ForeColor = Color.DarkGray;
            result_button.ForeColor = Color.DarkGray;

            load_panel.Show();
            build_panel.Hide();
            result_panel.Hide();
        }
        private void build_button_Click(object sender, EventArgs e)
        {
            load_button.ForeColor = Color.DarkGray;
            build_button.ForeColor = Color.MidnightBlue;
            result_button.ForeColor = Color.DarkGray;

            load_panel.Hide();
            build_panel.Show();
            result_panel.Hide();
        }
        private void result_button_Click(object sender, EventArgs e)
        {
            load_button.ForeColor = Color.DarkGray;
            build_button.ForeColor = Color.DarkGray;
            result_button.ForeColor = Color.MidnightBlue;

            load_panel.Hide();
            build_panel.Hide();
            result_panel.Show();
        }
        private void label1_Click(object sender, EventArgs e)
        {
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Application.Exit(); 
        }

        private void filepath_button_Click(object sender, EventArgs e)
        {
            filepath_textbox.Text = "";
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                filepath_textbox.Text = folderBrowserDialog.SelectedPath + @"\";
            }
            try
            {
                path = filepath_textbox.Text;
                string[] filename_list = Directory.GetFiles(path);
                filename_comboBox.Items.Clear(); // 清除原有檔案重新讀取
                foreach (string filename in filename_list)
                {
                    string realname = Path.GetFileName(filename);
                    if (realname.Contains(".xlsx"))
                        filename_comboBox.Items.Add(Path.GetFileName(filename));
                }
            }
            catch(Exception ex) { TaskDialog.Show("Error", "路徑輸入錯誤." + "\n\n" + ex.Message + "\n" + ex.ToString()); }
        }

        private void filename_comboBox_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                string lastChar = filepath_textbox.Text[filepath_textbox.Text.Length - 1].ToString();
                if (!lastChar.Equals("\\")) { path = filepath_textbox.Text + @"\"; }
                string[] filename_list = Directory.GetFiles(path);
                filename_comboBox.Items.Clear(); // 清除原有檔案重新讀取
                foreach (string filename in filename_list)
                {
                    string realname = Path.GetFileName(filename);
                    if (realname.Contains(".xlsx"))
                        filename_comboBox.Items.Add(Path.GetFileName(filename));
                }
            }
            catch (Exception ex) { TaskDialog.Show("Error", "路徑輸入錯誤." + "\n\n" + ex.Message + "\n" + ex.ToString()); }
        }

        private void inverted_arc_button_Click(object sender, EventArgs e)
        {
            externalEvent_inverted_arc.Raise();
        }
        private void track_bed_button_Click(object sender, EventArgs e)
        {
            externalEvent_track_bed.Raise();
        }

        private void envelope_button_Click(object sender, EventArgs e)
        {
            externalEvent_envelope.Raise();
        }

        private void circle_button_Click(object sender, EventArgs e)
        {
            if (bolt_checkbox.Checked)
                externalEvent_testhao.Raise();
            externalEvent_place_adative_circle.Raise();
        }

        private void attached_button_Click(object sender, EventArgs e)
        {
            externalEvent_attached_pipeline.Raise();
            externalEvent_U_shape_steel.Raise();
        }

        private void channel_button_Click(object sender, EventArgs e)
        {
            externalEvent_contact_channel.Raise();
        }

        private void gradientPanel1_MouseDown(object sender, MouseEventArgs e)
        {
            mov = 1;
            movX = e.X;
            movY = e.Y;
        }

        private void gradientPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            if(mov == 1)
            {
                this.Location = new Point(this.Left + e.X - this.movX, this.Top + e.Y - this.movY);
            }
        }

        private void gradientPanel1_MouseUp(object sender, MouseEventArgs e)
        {
            mov = 0;
        }

        private void counting_button_Click(object sender, EventArgs e) //***
        {
            externalEvent_Counting.Raise();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            externalEvent_section_for_test.Raise();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frame_path = "";
            OpenFileDialog openfiledialog = new OpenFileDialog();
            openfiledialog.ShowDialog();
            textBox1.Text = openfiledialog.FileName;

            frame_path = openfiledialog.FileName;
        }

        private void load_panel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void rebar_bottom_Click(object sender, EventArgs e)
        {
            externalEvent_shear_rebar.Raise();
            externalEvent_rebar_in_circle.Raise();
        }

        private void filename_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try { name = filename_comboBox.SelectedItem.ToString(); }
            catch (NullReferenceException) { }
        }

        private void bolt_checkbox_CheckedChanged(object sender, EventArgs e)
        {
            bolt = bolt_checkbox.Checked;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            counting_temp_path = "";
            OpenFileDialog openfiledialog = new OpenFileDialog();
            openfiledialog.ShowDialog();
            textBox3.Text = openfiledialog.FileName;

            counting_temp_path = openfiledialog.FileName;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Stream myStream;
            saveas_FilePath = "";
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "Microsoft Excel 活頁簿 (*.xlsx)|*.xlsx";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = saveFileDialog1.OpenFile()) != null)
                {
                    myStream.Close();
                }
            }
            textBox4.Text = saveFileDialog1.FileName.ToString();
            saveas_FilePath = saveFileDialog1.FileName.ToString();
        }

        public static HttpClient client_login()
        {
            try
            {
                HttpClient client = new HttpClient();
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                //client.BaseAddress = new Uri("http://127.0.0.1:8000/");
                client.BaseAddress = new Uri("https://bimdata.sinotech.com.tw/");
                client.DefaultRequestHeaders.Accept.Clear();
                var headerValue = new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json");
                client.DefaultRequestHeaders.Accept.Add(headerValue);
                client.DefaultRequestHeaders.ConnectionClose = true;
                Task.WaitAll(client.GetAsync($"/login/?USERNAME={Environment.UserName}&REVITAPI=SinoTunnel"));
                //Task.WaitAll(client.GetAsync($"/login/?USERNAME=11111&REVITAPI=SinoPipe"));
                return client;
            }
            catch (Exception)
            {
                return null;
            }
        }
        static string DecodeEncodedNonAsciiCharacters(string value)
        {
            return Regex.Replace(
                value,
                @"\\u(?<Value>[a-zA-Z0-9]{4})",
                m => {
                    return ((char)int.Parse(m.Groups["Value"].Value, NumberStyles.HexNumber)).ToString();
                });
        }

        int time_s = 10;
        private void timer1_Tick(object sender, EventArgs e)
        {
            time_s = time_s - 1;
            if (time_s == 0)
            {
                timer1.Stop();
                this.Dispose();
                this.Close();
            }
            label5.Text = "- 認證失敗，" + time_s.ToString() + "秒後強制關閉";
        }

        private void filepath_textbox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
