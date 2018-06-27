using AForge.Imaging;
using MouseKeyboardActivityMonitor;
using MouseKeyboardActivityMonitor.WinApi;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FMacro
{
    public partial class Main : Form
    {
        public Boolean stop = false;

        public Point pt;

        private static Bitmap bmpScreenShot;
        private static Graphics gfxScreenShot;

        private static Color newcolor;
        private static Color oldcolor;

        public int customWidth = 140;
        public int customHeight = 140;

        private readonly MouseHookListener m_MouseHookManager;

        public int sPointX = 0;
        public int sPointY = 0;

        public int mPointX = 0;
        public int mPointY = 0;

        public int bmPointX = 0;
        public int bmPointY = 0;

        public int ePointX = 0;
        public int ePointY = 0;

        public Point prePosition;
        public Rectangle currentRect;

        public Main()
        {
            InitializeComponent();

            // Windows Form에서 Key 입력 이벤트가 발생할 수 있도록 설정.
            this.KeyPreview = true;
            // Key 누름 이벤트를 생성.
            this.KeyDown += new KeyEventHandler(Main_KeyDown);
            //this.KeyPress += new KeyPressEventHandler(Main_KeyPress);

            m_MouseHookManager = new MouseHookListener(new GlobalHooker());

            m_MouseHookManager.MouseDown += HookManager_MouseDown;
            m_MouseHookManager.MouseUp += HookManager_MouseUp;
        }

        private void HookManager_MouseDown(object sender, MouseEventArgs e)
        {
            Console.WriteLine(e.Button.ToString() + " Pressed");
            sPointX = e.X;
            sPointY = e.Y;
        }
        private void HookManager_MouseUp(object sender, MouseEventArgs e)
        {
            Console.WriteLine(e.Button.ToString() + " Released");
            ePointX = e.Y;
            ePointY = e.Y;

            BindSquareImage();

            m_MouseHookManager.Enabled = false;
        }

        public void Main_KeyDown(object sender, KeyEventArgs e)
        {
            //MessageBox.Show("KeyDown : " + e.KeyCode.ToString());

            if (e.KeyCode == Keys.F4)
            {
                stop = false;
                startMacro();
            }
            // F5를 누르면 Loop가 중지된다.
            else if (e.KeyCode == Keys.F5)
            {
                stop = true;
                //MessageBox.Show("STOP");
            }
            // F6을 누르면 현재 실행중인 쓰레드 정보를 보여준다.
            else if (e.KeyCode == Keys.F6)
            {
                checkCurrentThread();
            }
        }
        
        public void Main_KeyPress(object sender, KeyPressEventArgs e)
        {
            //MessageBox.Show("KeyPress : " + e.KeyChar.ToString());
            /*
            if (e.KeyChar >= 48 && e.KeyChar <= 57)
            {
                //MessageBox.Show("Form.KeyPress: '" + e.KeyChar.ToString() + "' pressed.");

                switch (e.KeyChar)
                {
                    case (char)49:
                    case (char)52:
                    case (char)55:
                        //MessageBox.Show("Form.KeyPress: '" + e.KeyChar.ToString() + "' consumed.");
                        e.Handled = true;
                        break;
                }
            }
            */
        }

        public void BindSquareImage()
        {
            int width = Math.Abs(ePointX - sPointX);
            int height = Math.Abs(ePointY - sPointY);

            bmpScreenShot = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            gfxScreenShot = Graphics.FromImage(bmpScreenShot);
            gfxScreenShot.CopyFromScreen(sPointX, sPointY, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);

            Graphics grp = pictureBox1.CreateGraphics();
            grp.DrawImage(bmpScreenShot, 0, 0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                startMacro();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void startMacro()
        {
            try
            {
                // 쓰레드에서 실행 될 함수를 가진 객체.
                ThreadStart ts = new ThreadStart(runMacro);

                // ThreadStart 객체를 가진 상태로 쓰레드 생성.
                Thread th = new Thread(ts);
                //MessageBox.Show("ManagedThreadID : " + th.ManagedThreadId + " | ThreadName : " + th.Name);

                // Thread 실행.
                th.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void runMacro()
        {
            // 시작 위치.
            int x = 0;
            // 시작 위치.
            int y = 0;
            // 해상도 넓이
            int screen_x = System.Windows.Forms.SystemInformation.VirtualScreen.Width;
            // 해상도 높이
            int screen_y = System.Windows.Forms.SystemInformation.VirtualScreen.Height;
            
            // 최초 위치.
            Cursor.Position = new Point(0, 0);

            // Delay Time : milliseconds ( 30 = 0.03 Sec )
            int milliseconds = 100; // 30;

            while (!stop)
            {
                // 우측으로 50 pixel 이동
                x = x + 50;

                // 해상도 넓이 보다 큰 경우
                if (x >= screen_x)
                {
                    // 해상도 높이 보다 큰 경우
                    if (y >= screen_y)
                    {
                        // 1 회 종료.
                        MessageBox.Show("Finish", "FMacro Alert", MessageBoxButtons.OK);
                        break;
                    }

                    // 시작 위치 초기화
                    x = 0;

                    // 아래로 50 pixel 이동
                    y = y + 50;

                    // 1행 검색 종료 후, 다음 행으로 이동.
                    Cursor.Position = new System.Drawing.Point(0, y);
                }

                // 마우스 위치 이동
                pt = new Point(x, y);

                Cursor.Position = pt;

                GetCursorRGB(pt.X, pt.Y);

                if (oldcolor != newcolor)
                {
                    Console.WriteLine("Old Color : " + oldcolor + " | New Color : " + newcolor);

                    BindCursorImage(pt.X, pt.Y, customWidth, customHeight);

                    //break;
                }

                oldcolor = newcolor;

                // 1회 이동 후 Delay
                Thread.Sleep(milliseconds);
            }
        }

        private void BindCursorImage(int x, int y, int cx, int cy)
        {
            bmpScreenShot = new Bitmap(cx, cy, PixelFormat.Format24bppRgb);
            gfxScreenShot = Graphics.FromImage(bmpScreenShot);
            gfxScreenShot.CopyFromScreen(x, y, 0, 0, new Size(cx, cy), CopyPixelOperation.SourceCopy);

            Graphics grp = pictureBox2.CreateGraphics();
            grp.DrawImage(bmpScreenShot, 0, 0);
            grp.Dispose();

            matchTemplate();
        }

        private void GetCursorRGB(int x, int y)
        {
            bmpScreenShot = new Bitmap(1, 1, PixelFormat.Format32bppArgb);
            gfxScreenShot = Graphics.FromImage(bmpScreenShot);
            gfxScreenShot.CopyFromScreen(x, y, 0, 0, new Size(1, 1), CopyPixelOperation.SourceCopy);
            Color c = bmpScreenShot.GetPixel(0, 0);
            newcolor = c;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            m_MouseHookManager.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //checkCurrentThread();
            matchTemplate();
        }

        public void matchTemplate()
        {
            try
            {
                pictureBox3.InitialImage = null;

                Bitmap _templateimage = new Bitmap(pictureBox1.ClientSize.Width, pictureBox1.ClientSize.Height, PixelFormat.Format24bppRgb);
                Bitmap _srcimage = new Bitmap(pictureBox2.ClientSize.Width, pictureBox2.ClientSize.Height, PixelFormat.Format24bppRgb);
                Bitmap _dsimage = (Bitmap)_srcimage.Clone();
                ExhaustiveTemplateMatching _tm = new ExhaustiveTemplateMatching(0.9f);
                TemplateMatch[] _mts = _tm.ProcessImage(_dsimage, _templateimage);
                /*if (_mts.Length > 0)
                {
                    MessageBox.Show(_mts.Length.ToString());
                }*/
                BitmapData _bmpdata = _dsimage.LockBits(new Rectangle(0, 0, _dsimage.Width, _dsimage.Height), ImageLockMode.ReadWrite, _dsimage.PixelFormat);
                foreach (var _match in _mts)
                {
                    Drawing.Rectangle(_bmpdata, _match.Rectangle, Color.Red);
                }

                _dsimage.UnlockBits(_bmpdata);

                pictureBox3.Image = _dsimage;

                stop = true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private static void checkCurrentThread()
        {
            Process proc = Process.GetCurrentProcess();
            ProcessThreadCollection ptc = proc.Threads;
            Console.WriteLine("현재 프로세스에서 실행중인 스레드 수 : {0}", ptc.Count);
            ThreadInfo(ptc);
        }

        private static void ThreadInfo(ProcessThreadCollection ptc)
        {
            int i = 1;
            foreach (ProcessThread pt in ptc)
            {
                Console.WriteLine("******* {0} 번째 스레드 정보 *******", i++);
                Console.WriteLine("ThreadId : {0}", pt.Id);            //스레드 ID
                Console.WriteLine("시작시간 : {0}", pt.StartTime);    //스레드 시작시간
                Console.WriteLine("우선순위 : {0}", pt.BasePriority);  //스레드 우선순위
                Console.WriteLine("상태 : {0}", pt.ThreadState);      //스레드 상태
                Console.WriteLine();
            }
        }
    }
}
