using OpenTK;
using OpenTK.Graphics.ES10;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tao.OpenGl;
using Tao.Platform.Windows;
using OpenTK.Graphics.OpenGL;
using GL = OpenTK.Graphics.OpenGL.GL;
using ClearBufferMask = OpenTK.Graphics.OpenGL.ClearBufferMask;
using MatrixMode = OpenTK.Graphics.OpenGL.MatrixMode;
using EnableCap = OpenTK.Graphics.OpenGL.EnableCap;
using BeginMode = OpenTK.Graphics.OpenGL.BeginMode;
using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.CacheProviders;
using GMap.NET.Internals;
using GMap.NET.ObjectModel;
using GMap.NET.Projections;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsPresentation;
using Excel = Microsoft.Office.Interop.Excel;
using AForge;
using AForge.Video;
using AForge.Video.DirectShow;
using Accord.Video.FFMPEG;
using System.Net;
using AForge.Video.VFW;

namespace turksat
{
    public partial class Frm_Home : Form
    {
        MJPEGStream streamvideo;
        public static List<string> data0 = new List<string>();
        public static List<string> data1 = new List<string>();
        public static List<string> data2 = new List<string>();
        public static List<string> data3 = new List<string>();
        public static List<string> data4 = new List<string>();
        public static List<string> data5 = new List<string>();
        public static List<string> data6 = new List<string>();
        public static List<string> data7 = new List<string>();
        public static List<string> data8 = new List<string>();
        public static List<string> data9 = new List<string>();
        public static List<string> data10 = new List<string>();
        public static List<string> data11 = new List<string>();
        public static List<string> data12 = new List<string>();
        public static List<string> data13 = new List<string>();
        public static List<string> data14 = new List<string>();
        public static List<string> data15 = new List<string>();
        public static List<string> data16 = new List<string>();
        public static List<string> data17 = new List<string>();
        public static List<string> data18 = new List<string>();
        public static int data_count = 0;
        public string[][] values = new string[18][];
        public Frm_Home()
        {
            //values[0][0]=data[0];//0,0
            //values[1][0] = data[1];
            //values[2][0] = data[2];
            //values[3][0] = data[3];
            //values[4][0] = data[4];
            ////yeni veriler geldiğinde 2.boyut değiştir.
            //values[0][1] = data[0];
            //values[0][2] = data[0];
            //values[0][3] = data[0];
            //values[0][4] = data[0];
            //values[0][5] = data[0];
            //values[0][6] = data[0];
            //values[0][7] = data[0];
            //values[0][8] = data[0];
            //int data_count = 0;
            ////array doldurmak için
            //for (int i = 0; i < 23; i++)
            //{
            //    values[i][data_count] = data[i];
            //    data_count++;
            //}

            ////yazdırmak için tüm eklenmiş veriyi gezer
            //for (int i = 0; i < 23; i++)
            //{
            //    for (int j = 0; j < data_count; j++)
            //    {
            //        //excell satırı= values[i][j];
            //    }
            //}
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;

            SerialPort.DataReceived += new SerialDataReceivedEventHandler(SerilPortOnReceiveData);

 //----------------------------------------------------KAMERA IP LİNKİ-----------------------------------------------------------------------------          
            streamvideo = new MJPEGStream("http://192.168.95.51:81/stream");
            streamvideo.NewFrame += GetNewFrame;
            //-------------------------------------------------------------------------------------------------------------------------------------------------
           /* txt_TakimNo.Text = "467887";
            txt_PilGerilimi.Text = "6.7";
            txt_UyduStatusu.Text = "Yükselişte";
            txt_VideoAktarimBilgisi.Text = "Record";
           */

            //OpenGlControl.InitializeContexts();
            //Gl.glClearColor(0.0f, 0.0f, 0.0f, 0.0f);
            //Gl.glMatrixMode(Gl.GL_PROJECTION);
            //Gl.glLoadIdentity();
            //Gl.glOrtho(0.0, 1.0, 0.0, 1.0, -1.0, 1.0);
            ////Gl.glOrtho(0.0, 1.0, 0.0, 1.0, 0.0, 1.0);
            //Gl.glMatrixMode(Gl.GL_MODELVIEW);
        }
//----------------------------------------kamera----------------------------------------------------------
        private void GetNewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap bmp = (Bitmap)eventArgs.Frame.Clone();
            picBoxVideo.Image = bmp;

        }

        private void btn_VideoAktarButonu_Click(object sender, EventArgs e)
        {
            streamvideo.Start();
        }
        private void btnVideoSonlandir_Click(object sender, EventArgs e)
        {
            //streamvideo.Stop();
            //picBoxVideo.Image = null;
            WebClient client = new WebClient();
            client.Credentials = new NetworkCredential("esp32", "esp32");
            client.UploadFile("ftp://192.168.43.23/Adsız.png", @"C:\Users\Aslı\Desktop\Adsız.png");
        }


 //-------------------------------------------------------------------------------------------------------

        //string realdata = "47697,67,5/7/2021/20/28/35/78372.06,2114.81,156.2,9.00,24.50,7.01,42.22,27.00,321.00,Düşüşte,-633,-5.30,956.50,-1.01,Evet";
        string realdata = string.Empty;
        string[] data;


        double yuk, pil, sck, bas;
        float pitch, roll, yaw;


        private void SerialPortBaglan()
        {

            //if (txtSifre.Text == "12345")
            //{
            //    SerialPort.Open();
            //    lbl_portBilgisi.Text = "BAĞLANDI";
            //    lbl_portBilgisi.ForeColor = Color.Green;
            //}

            //else
            //{
            //    SerialPort.Close();
            //    lbl_portBilgisi.Text = "BAĞLANMADI";
            //    lbl_portBilgisi.ForeColor = Color.Red;
            //}
            try
            {
                if (!SerialPort.IsOpen)
                {
                    //SerialPort.PortName = "COM6"; 
                    SerialPort.PortName = cmb_Ports.Text;
                    SerialPort.BaudRate = 9600;
                    SerialPort.DataBits = 8;
                    SerialPort.Parity = Parity.None;
                    SerialPort.StopBits = StopBits.One;

                    SerialPort.Open();
                    lbl_portBilgisi.Text = "BAĞLANDI";
                    lbl_portBilgisi.ForeColor = Color.Green;
                }
            }
            catch (Exception)
            {
                SerialPort.Close();
                lbl_portBilgisi.Text = "BAĞLANMADI";
                lbl_portBilgisi.ForeColor = Color.Red;
            }
        }

        DataTable tbl = new DataTable();
        private void SerilPortOnReceiveData(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(250);

            realdata = string.Empty;
            realdata = SerialPort.ReadExisting();

            //MessageBox.Show(realdata);
            //Console.WriteLine(realdata);

            //degeralyaz();

           

            data = realdata.Split(','); // alınan datalar sıralanır
            if ((data.Length == 19))
            {
                //data0.Add(data[0]);
                //data1.Add(data[1]);
                //data2.Add(data[2]);

                //data23.Add(data[18]);
               /* txt_PaketNumarasi.Text = data[0];
                txt_GonderimSaati.Text = data[1];
                txt_Basinc1.Text = data[2];
                txt_Basinc2.Text = data[3];
                txt_Yukseklik1.Text = data[4];
                txt_Yukseklik2.Text = data[5];
                txt_IrtifaFarki.Text = data[6];
                txt_inisHizi.Text = data[7];
                txt_Sicaklik.Text = data[8];
               
                txt_GPS1_Latitude.Text = data[9];
                txt_GPS1_Longitude.Text = data[10];
                txt_GPS1_Altitude.Text = data[11];
                txt_GPS2_Latitude.Text = data[12];
                txt_GPS2_Longitude.Text = data[13];
                txt_GPS2_Altitude.Text = data[14];
                
                txt_Pitch.Text = data[15];
                txt_Roll.Text = data[16];
                txt_Yaw.Text = data[17];
                txt_DonusSayisi.Text = data[18];
               */
               

                //pitch = float.Parse(data[12], System.Globalization.CultureInfo.InvariantCulture);
                //roll = float.Parse(data[13], System.Globalization.CultureInfo.InvariantCulture);
                //yaw = float.Parse(data[14], System.Globalization.CultureInfo.InvariantCulture);

                y = float.Parse(data[15], System.Globalization.CultureInfo.InvariantCulture);
                x = float.Parse(data[16], System.Globalization.CultureInfo.InvariantCulture);
                z = float.Parse(data[17], System.Globalization.CultureInfo.InvariantCulture);

                glControl1.Invalidate();

                yuk = double.Parse(data[5], System.Globalization.CultureInfo.InvariantCulture);
                //pil = double.Parse(data[10], System.Globalization.CultureInfo.InvariantCulture);
                pil = double.Parse("6.7", System.Globalization.CultureInfo.InvariantCulture);
                sck = double.Parse(data[8], System.Globalization.CultureInfo.InvariantCulture);
                bas = double.Parse(data[2], System.Globalization.CultureInfo.InvariantCulture);

                chart_Yukseklik.Series["Yukseklik"].Points.AddY(yuk);
                chart_PilGerilimi.Series["PilGerilimi"].Points.AddY(pil);
                chart_Sicaklik.Series["Sicaklik"].Points.AddY(sck);
                chart_Basinc.Series["Basinc"].Points.AddY(bas);

                //dgv_degerler.rows.add(txt_takimno.text, txt_paketnumarasi.text, txt_gonderimsaati.text, txt_basinc.text, txt_yukseklik.text, txt_inishizi.text, txt_sicaklik.text, txt_pilgerilimi.text, txt_gps_latitude.text, txt_gps_longitude.text, txt_gps_altitude.text, txt_uydustatusu.text, txt_pitch.text, txt_roll.text, txt_yaw.text, txt_donussayisi.text, txt_videoaktarimbilgisi.text);

            //    for (int i = 0; i < 10; i++)
            //    {
            //        dgv_degerler.rows[i].visible = false;
            //    }
               
            //    dgv_degerler.rows.add(data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13], data[14], data[15], data[16]);              
            //    dgv_Degerler.ClearSelection();
            }

        }

        private void GetPorts()
        {
            var ports = SerialPort.GetPortNames();
            cmb_Ports.Items.AddRange(ports);
        }

        private void btn_PortBaglan_Click(object sender, EventArgs e)
        {
            MessageBox.Show(realdata);
            SerialPortBaglan();
        }

        int sayac = 0;

        float alfa = 0;

        private void tmr_PortControl_Tick(object sender, EventArgs e)
        {
            sayac++;
            if (sayac == 3)
            {
                tmr_PortControl.Stop();
                SerialPortBaglan();
                sayac = 0;
                tmr_PortControl.Start();
            }
        }

        float x, y, z;
        

        private void glControl1_Paint(object sender, PaintEventArgs e)  // dger verilerinde grafihi,ig ,nten,yot
        {
            float step = 1.0f;
            float topla = step;
            float radius = 5.0f;
            float dikey1 = radius, dikey2 = -radius;
            GL.Clear(ClearBufferMask.ColorBufferBit);
            GL.Clear(ClearBufferMask.DepthBufferBit);

            Matrix4 perspective = Matrix4.CreatePerspectiveFieldOfView(1.04f, 4 / 3, 1, 10000);
            Matrix4 lookat = Matrix4.LookAt(25, 0, 0, 0, 0, 0, 0, 1, 0);
            GL.MatrixMode(MatrixMode.Projection);
            GL.LoadIdentity();
            GL.LoadMatrix(ref perspective);
            GL.MatrixMode(MatrixMode.Modelview);
            GL.LoadIdentity();
            GL.LoadMatrix(ref lookat);
            GL.Viewport(0, 0, glControl1.Width, glControl1.Height);
            GL.Enable(EnableCap.DepthTest);
            GL.DepthFunc(DepthFunction.Less);

            GL.Rotate(x, 1.0, 0.0, 0.0);//ÖNEMLİ
            GL.Rotate(z, 0.0, 1.0, 0.0);
            GL.Rotate(y, 0.0, 0.0, 1.0);

            silindir(step, topla, radius, 3, -5);
            silindir(0.01f, topla, 0.5f, 9, 9.7f);
            silindir(0.01f, topla, 0.1f, 5, dikey1 + 5);
            koni(0.01f, 0.01f, radius, 3.0f, 3, 5);
            koni(0.01f, 0.01f, radius, 2.0f, -5.0f, -10.0f);
            Pervane(9.0f, 11.0f, 0.2f, 0.5f);

            GL.Begin(BeginMode.Lines);

            GL.Color3(Color.FromArgb(250, 0, 0));
            GL.Vertex3(-30.0, 0.0, 0.0);
            GL.Vertex3(30.0, 0.0, 0.0);


            GL.Color3(Color.FromArgb(0, 0, 0));
            GL.Vertex3(0.0, 30.0, 0.0);
            GL.Vertex3(0.0, -30.0, 0.0);

            GL.Color3(Color.FromArgb(0, 0, 250));
            GL.Vertex3(0.0, 0.0, 30.0);
            GL.Vertex3(0.0, 0.0, -30.0);

            GL.End();
            //GraphicsContext.CurrentContext.VSync = true;
            glControl1.SwapBuffers();
        }

        private void glControl1_Load(object sender, EventArgs e)
        {
            GL.ClearColor(0.0f, 0.0f, 0.0f, 0.0f);
            GL.Enable(EnableCap.DepthTest);//sonradan yazdık
        }

        private void silindir(float step, float topla, float radius, float dikey1, float dikey2)
        {
            float eski_step = 0.1f;
            GL.Begin(BeginMode.Quads);//Y EKSEN CIZIM DAİRENİN
            while (step <= 360)
            {
                if (step < 45)
                    GL.Color3(Color.FromArgb(255, 0, 0));
                else if (step < 90)
                    GL.Color3(Color.FromArgb(255, 255, 255));
                else if (step < 135)
                    GL.Color3(Color.FromArgb(255, 0, 0));
                else if (step < 180)
                    GL.Color3(Color.FromArgb(255, 255, 255));
                else if (step < 225)
                    GL.Color3(Color.FromArgb(255, 0, 0));
                else if (step < 270)
                    GL.Color3(Color.FromArgb(255, 255, 255));
                else if (step < 315)
                    GL.Color3(Color.FromArgb(255, 0, 0));
                else if (step < 360)
                    GL.Color3(Color.FromArgb(255, 255, 255));


                float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey1, ciz1_y);

                float ciz2_x = (float)(radius * Math.Cos((step + 2) * Math.PI / 180F));
                float ciz2_y = (float)(radius * Math.Sin((step + 2) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey1, ciz2_y);

                GL.Vertex3(ciz1_x, dikey2, ciz1_y);
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }
            GL.End();
            GL.Begin(BeginMode.Lines);
            step = eski_step;
            topla = step;
            while (step <= 180)// UST KAPAK
            {
                if (step < 45)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 90)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 135)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 180)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 225)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 270)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 315)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 360)
                    GL.Color3(Color.FromArgb(250, 250, 200));


                float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey1, ciz1_y);

                float ciz2_x = (float)(radius * Math.Cos((step + 180) * Math.PI / 180F));
                float ciz2_y = (float)(radius * Math.Sin((step + 180) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey1, ciz2_y);

                GL.Vertex3(ciz1_x, dikey1, ciz1_y);
                GL.Vertex3(ciz2_x, dikey1, ciz2_y);
                step += topla;
            }
            step = eski_step;
            topla = step;
            while (step <= 180)//ALT KAPAK
            {
                if (step < 45)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 90)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 135)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 180)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 225)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 270)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 315)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 360)
                    GL.Color3(Color.FromArgb(250, 250, 200));

                float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey2, ciz1_y);

                float ciz2_x = (float)(radius * Math.Cos((step + 180) * Math.PI / 180F));
                float ciz2_y = (float)(radius * Math.Sin((step + 180) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);

                GL.Vertex3(ciz1_x, dikey2, ciz1_y);
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }
            GL.End();
        }

      


      
        private void chart_PilGerilimi_Click(object sender, EventArgs e)
        {

        }

        private void chart_Yukseklik_Click(object sender, EventArgs e)
        {

        }

        private void chart_Sicaklik_Click(object sender, EventArgs e)
        {

        }

        private void chart_Basinc_Click(object sender, EventArgs e)
        {

        }

        


        private void gMapProvider_Load(object sender, EventArgs e)
        {
            map.MapProvider = GMapProviders.BingSatelliteMap;
            map.DragButton = MouseButtons.Left;
           
            //double lat = Convert.ToDouble(txtLat.Text);
            //double longt = Convert.ToDouble(txtLong.Text);
            //map.Position = new PointLatLng(lat, longt);

           map.Position = new PointLatLng(41.2108, 32.6602);
            map.MinZoom = 5;
            map.Zoom = 25;
            map.MaxZoom = 200;
        }


        private void btn_AyrilKomutu_Click(object sender, EventArgs e)
        {
            SerialPort.Write("A,X*");
        }

        private void btnManuelTahrik_Click(object sender, EventArgs e)
        {
            SerialPort.Write("M,YX*");
        }

        private void btnKilitle_Click(object sender, EventArgs e)
        {
            SerialPort.Write("T,X*");
        }

        private void cmb_Ports_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void koni(float step, float topla, float radius1, float radius2, float dikey1, float dikey2)
        {
            float eski_step = 0.1f;
            GL.Begin(BeginMode.Lines);//Y EKSEN CIZIM DAİRENİN
            while (step <= 360)
            {
                if (step < 45)
                    GL.Color3(1.0, 1.0, 1.0);
                else if (step < 90)
                    GL.Color3(1.0, 0.0, 0.0);
                else if (step < 135)
                    GL.Color3(1.0, 1.0, 1.0);
                else if (step < 180)
                    GL.Color3(1.0, 0.0, 0.0);
                else if (step < 225)
                    GL.Color3(1.0, 1.0, 1.0);
                else if (step < 270)
                    GL.Color3(1.0, 0.0, 0.0);
                else if (step < 315)
                    GL.Color3(1.0, 1.0, 1.0);
                else if (step < 360)
                    GL.Color3(1.0, 0.0, 0.0);


                float ciz1_x = (float)(radius1 * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius1 * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey1, ciz1_y);

                float ciz2_x = (float)(radius2 * Math.Cos(step * Math.PI / 180F));
                float ciz2_y = (float)(radius2 * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }
            GL.End();

            GL.Begin(BeginMode.Lines);
            step = eski_step;
            topla = step;
            while (step <= 180)// UST KAPAK
            {
                if (step < 45)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 90)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 135)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 180)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 225)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 270)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 315)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 360)
                    GL.Color3(Color.FromArgb(250, 250, 200));


                float ciz1_x = (float)(radius2 * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius2 * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey2, ciz1_y);

                float ciz2_x = (float)(radius2 * Math.Cos((step + 180) * Math.PI / 180F));
                float ciz2_y = (float)(radius2 * Math.Sin((step + 180) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);

                GL.Vertex3(ciz1_x, dikey2, ciz1_y);
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }
            step = eski_step;
            topla = step;
            GL.End();
        }

       

        private void Pervane(float yukseklik, float uzunluk, float kalinlik, float egiklik)
        {
            float radius = 10, angle = 45.0f;
            GL.Begin(BeginMode.Quads);

            GL.Color3(Color.Red);
            GL.Vertex3(uzunluk, yukseklik, kalinlik);
            GL.Vertex3(uzunluk, yukseklik + egiklik, -kalinlik);
            GL.Vertex3(0.0, yukseklik + egiklik, -kalinlik);
            GL.Vertex3(0.0, yukseklik, kalinlik);

            GL.Color3(Color.Red);
            GL.Vertex3(-uzunluk, yukseklik + egiklik, kalinlik);
            GL.Vertex3(-uzunluk, yukseklik, -kalinlik);
            GL.Vertex3(0.0, yukseklik, -kalinlik);
            GL.Vertex3(0.0, yukseklik + egiklik, kalinlik);

            GL.Color3(Color.White);
            GL.Vertex3(kalinlik, yukseklik, -uzunluk);
            GL.Vertex3(-kalinlik, yukseklik + egiklik, -uzunluk);
            GL.Vertex3(-kalinlik, yukseklik + egiklik, 0.0);//+
            GL.Vertex3(kalinlik, yukseklik, 0.0);//-

            GL.Color3(Color.White);
            GL.Vertex3(kalinlik, yukseklik + egiklik, +uzunluk);
            GL.Vertex3(-kalinlik, yukseklik, +uzunluk);
            GL.Vertex3(-kalinlik, yukseklik, 0.0);
            GL.Vertex3(kalinlik, yukseklik + egiklik, 0.0);
            GL.End();

        }

       

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    alfa = (alfa + 5) % 360;
        //    //OpenGlControl.Refresh();
        //}

       
        public static int i;

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lbl_Altitude_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void picBoxVideo_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        public static int count;
        private void Frm_Home_Load(object sender, EventArgs e)
        {
            //dgv_Degerler.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dgv_Degerler.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;

            //dgv_Degerler.Rows.Add(10);
            //dgv_Degerler.ClearSelection();
            GetPorts();
            count = 0;
            i = 2;
            //dgv_Degerler.Rows.Add(txt_TakimNo.Text, txt_PaketNumarasi.Text, txt_GonderimSaati.Text, txt_Basinc.Text, txt_Yukseklik.Text, txt_inisHizi.Text, txt_Sicaklik.Text, txt_PilGerilimi.Text, txt_GPS_Latitude.Text, txt_GPS_Longitude.Text, txt_GPS_Altitude.Text, txt_UyduStatusu.Text, txt_Pitch.Text, txt_Roll.Text, txt_Yaw.Text, txt_DonusSayisi.Text, txt_VideoAktarimBilgisi.Text);


            // tmr_PortControl.Start();

            //timer1.Start();

        }

        private void degeralyaz()
        {
            data = realdata.Split(',');
/*
           // txt_TakimNo.Text = "467887";
            txt_PaketNumarasi.Text = data[0];
            txt_GonderimSaati.Text = data[1];
            txt_Basinc1.Text = data[2];
            txt_Basinc2.Text = data[3];
            txt_Yukseklik1.Text = data[4];
            txt_Yukseklik2.Text = data[5];
            txt_IrtifaFarki.Text = data[6];
            txt_inisHizi.Text = data[7];
            txt_Sicaklik.Text = data[8];
           // txt_PilGerilimi.Text = "6";
            txt_GPS1_Latitude.Text = data[9];
            txt_GPS1_Longitude.Text = data[10];
            txt_GPS1_Altitude.Text = data[11];
            txt_GPS2_Latitude.Text = data[12];
            txt_GPS2_Longitude.Text = data[13];
            txt_GPS2_Altitude.Text = data[14];
           // txt_UyduStatusu.Text = "Yükselişte";
            txt_Pitch.Text = data[15];
            txt_Roll.Text = data[16];
            txt_Yaw.Text = data[17];
            txt_DonusSayisi.Text = data[18];


            // txt_VideoAktarimBilgisi.Text = "Record";
*/

            //pitch = float.Parse(data[12], System.Globalization.CultureInfo.InvariantCulture);
            //roll = float.Parse(data[13], System.Globalization.CultureInfo.InvariantCulture);
            //yaw = float.Parse(data[14], System.Globalization.CultureInfo.InvariantCulture);

            //x = float.Parse(data[12], System.Globalization.CultureInfo.InvariantCulture);
            //y = float.Parse(data[13], System.Globalization.CultureInfo.InvariantCulture);
            //z = float.Parse(data[14], System.Globalization.CultureInfo.InvariantCulture);
            y = float.Parse(data[15], System.Globalization.CultureInfo.InvariantCulture);
            x = float.Parse(data[16], System.Globalization.CultureInfo.InvariantCulture);
            z = float.Parse(data[17], System.Globalization.CultureInfo.InvariantCulture);
            glControl1.Invalidate();

            Console.WriteLine(x + "," + y + "," + z);
            //dgv_Degerler.Rows.Add(data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13], data[14], data[15], data[16]);
            //dgv_Degerler.ClearSelection();
           
            
            yuk = double.Parse(data[5], System.Globalization.CultureInfo.InvariantCulture);
            //pil = double.Parse(data[10], System.Globalization.CultureInfo.InvariantCulture);
            pil = double.Parse("6.7", System.Globalization.CultureInfo.InvariantCulture);
            sck = double.Parse(data[8], System.Globalization.CultureInfo.InvariantCulture);
            bas = double.Parse(data[2], System.Globalization.CultureInfo.InvariantCulture);

            chart_Yukseklik.Series["Yukseklik"].Points.AddY(yuk);
            chart_PilGerilimi.Series["PilGerilimi"].Points.AddY(pil);
            chart_Sicaklik.Series["Sicaklik"].Points.AddY(sck);
            chart_Basinc.Series["Basinc"].Points.AddY(bas);

            //OpenGlControl.Refresh();

        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            degeralyaz();
        }

        private void Frm_Home_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (SerialPort.IsOpen)
            {
                SerialPort.Close();
            }
        }
        //----------------------------excel veri yazma ---------------------------------------------
      
        
        public void WriteToExcel()
        {
            if (count%50==0)
            {

            }
            Excel.Application excelapp = new Excel.Application();
            excelapp.Workbooks.Add();
            excelapp.Cells[count + 1, 1] = "TAKIM NO";
            excelapp.Cells[count + 1, 2] = "PAKET NUMARASI";
            excelapp.Cells[count + 1, 3] = "GÖNDERİM SAATİ";
            excelapp.Cells[count + 1, 4] = "BASINÇ1";
            excelapp.Cells[count + 1, 5] = "BASINÇ2";
            excelapp.Cells[count + 1, 6] = "YÜKSEKLİK1"; 
            excelapp.Cells[count + 1, 7] = "YÜKSEKLİK2";
            excelapp.Cells[count + 1, 8] = "İRTİFA FARKI";
            excelapp.Cells[count + 1, 9] = "İNİŞ HIZI";
            excelapp.Cells[count + 1, 10] = "SICAKLIK";
            excelapp.Cells[count + 1, 11] = "PİL GERİLİMİ";
            excelapp.Cells[count + 1, 12] = "GPS1 LATITUDE";
            excelapp.Cells[count + 1, 13] = "GPS1 LONGTITUDE";
            excelapp.Cells[count + 1, 14] = "GPS1 ALTITUDE";
            excelapp.Cells[count + 1, 15] = "GPS2 LATITUDE";
            excelapp.Cells[count + 1, 16] = "GPS2 LONGTITUDE";
            excelapp.Cells[count + 1, 17] = "GPS2 ALTITUDE";
            excelapp.Cells[count + 1, 18] = "UYDU STATÜSÜ";
            excelapp.Cells[count + 1, 19] = "PITCH";
            excelapp.Cells[count + 1, 20] = "ROLL";
            excelapp.Cells[count + 1, 21] = "YAW";
            excelapp.Cells[count + 1, 22] = "DÖNÜŞ SAYISI";
            excelapp.Cells[count + 1, 23] = "VİDEO AKTARIM BİLGİSİ";


            for (int i = 0; i < 23; i++)
            {
                values[i][data_count] = data[i];
                data_count++;

            }

            ////yazdırmak için tüm eklenmiş veriyi gezer
            for (int i = 0; i < 23; i++)
            {
                for (int j = 0; j < data_count; j++)
                {
                    excelapp.Cells[count + i, 1] = "467887";
                    excelapp.Cells[count + i, 2] = values[0][0];
                    excelapp.Cells[count + i, 3] = values[1][0];
                    excelapp.Cells[count + i, 4] = values[2][0];
                    excelapp.Cells[count + i, 5] = values[3][0];
                    excelapp.Cells[count + i, 6] = values[4][0];
                    excelapp.Cells[count + i, 7] = values[5][0];
                    excelapp.Cells[count + i, 8] = values[6][0];
                    excelapp.Cells[count + i, 9] = values[7][0];
                    excelapp.Cells[count + i, 10] = "6.7";
                    excelapp.Cells[count + i, 11] = values[8][0];
                    excelapp.Cells[count + i, 12] = values[9][0];
                    excelapp.Cells[count + i, 13] = values[10][0];
                    excelapp.Cells[count + i, 14] = values[11][0];
                    excelapp.Cells[count + i, 15] = values[12][0];
                    excelapp.Cells[count + i, 16] = values[13][0];
                    excelapp.Cells[count + i, 17] = values[14][0];
                    excelapp.Cells[count + i, 18] = "Yükselişte";
                    excelapp.Cells[count + i, 19] = values[15][0];
                    excelapp.Cells[count + i, 20] = values[16][0];
                    excelapp.Cells[count + i, 21] = values[17][0];
                    excelapp.Cells[count + i, 22] = values[18][0];
                    excelapp.Cells[count + i, 23] = "Record";
                }
            }
            excelapp.Visible = true;
            count += 23;


            ////in loop-----------------------------------------------------------------------------------------------------------
            //excelapp.Cells[count + i, 1] = txt_TakimNo.Text; //values[0][0] 2.satır values[0][1]
            //excelapp.Cells[count + i, 2] = txt_PaketNumarasi.Text;//values[1][0] 2.satır valus[1][1]
            //excelapp.Cells[count + i, 3] = txt_GonderimSaati.Text;
            //excelapp.Cells[count + i, 4] = txt_Basinc1.Text;
            //excelapp.Cells[count + i, 5] = txt_Basinc2.Text;
            //excelapp.Cells[count + i, 6] = txt_Yukseklik1.Text;
            //excelapp.Cells[count + i, 7] = txt_Yukseklik2.Text;
            //excelapp.Cells[count + i, 8] = txt_IrtifaFarki.Text;
            //excelapp.Cells[count + i, 9] = txt_inisHizi.Text;
            //excelapp.Cells[count + i, 10] = txt_Sicaklik.Text;
            //excelapp.Cells[count + i, 11] = txt_PilGerilimi.Text;
            //excelapp.Cells[count + i, 12] = txt_GPS1_Latitude.Text;
            //excelapp.Cells[count + i, 13] = txt_GPS1_Longitude.Text;
            //excelapp.Cells[count + i, 14] = txt_GPS1_Altitude.Text;
            //excelapp.Cells[count + i, 15] = txt_GPS2_Latitude.Text;
            //excelapp.Cells[count + i, 16] = txt_GPS2_Longitude.Text;
            //excelapp.Cells[count + i, 17] = txt_GPS2_Altitude.Text;
            //excelapp.Cells[count + i, 18] = txt_UyduStatusu.Text;
            //excelapp.Cells[count + i, 19] = txt_Pitch.Text;
            //excelapp.Cells[count + i, 20] = txt_Roll.Text;
            //excelapp.Cells[count + i, 21] = txt_Yaw.Text;
            //excelapp.Cells[count + i, 22] = txt_DonusSayisi.Text;
            //excelapp.Cells[count + i, 23] = txt_VideoAktarimBilgisi.Text;

            //excelapp.Visible = true;
            //count += 23;
            //i++;----------------------------------------------------------------------------------------------------------------------------
        }

        private void btn_Excel_Click(object sender, EventArgs e)
        {
            WriteToExcel();

        }
        //------------------------------------------------------------------------------------------------------------
    }
}
