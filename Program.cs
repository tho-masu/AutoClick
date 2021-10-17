using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using static System.Console;

namespace AutoClick
{
    class Program
    {
        const int MOUSEEVENTF_LEFTDOWN = 0x2;
        const int MOUSEEVENTF_LEFTUP = 0x4;
        static System.Timers.Timer timer1;
        static System.Timers.Timer timer2;
        static int x = 820; //初期値
        static int y = 280; //初期値
        static int movex = 750;
        static int movey = 350;
        static string message = "no message"; //初期値
        static string switch_window_name = "Chrome"; //初期値

        [STAThread] //Clipboard.SetTextの利用のために必要
        static void Main(string[] args)
        {
            // スクリーンの左上隅を基準(0,0)とする座標を指定する。
            WriteLine("x軸を入力してください");
            string X = ReadLine();
            if (!String.IsNullOrWhiteSpace(X))
            {
                x = int.Parse(X);
            }
            WriteLine($"設定値はx={x}です");

            WriteLine("y軸を入力してください");
            string Y = ReadLine();
            if (!String.IsNullOrWhiteSpace(Y))
            {
                y = int.Parse(Y);
            }
            WriteLine($"設定値はy={y}です");

            WriteLine("文字列を入力してください");
            //WriteLine("ペーストしたい文字列をクリップボードにコピーしてください");
            string msg = ReadLine();
            if (!String.IsNullOrWhiteSpace(msg))
            {
                message = msg;
            }
            WriteLine($"設定値はmessage=\"{message}\"です");
            Clipboard.SetText(message);

            WriteLine("時間分秒をhhMMssの形で入力してください");
            string time = ReadLine();
            if (time.Length != 6)
            {
                WriteLine("書式が正しくありません");
                ReadKey();
                Environment.Exit(0);
            }
            int hour = int.Parse(time.Substring(0, 2));
            int minute = int.Parse(time.Substring(2, 2));
            int second = int.Parse(time.Substring(4, 2));

            WriteLine("キーを押すと処理を開始します");
            ReadKey(true);
            WriteLine("処理開始");
            SetTimer(hour, minute, second);
            WriteLine("キーを押すとプログラムを終了します");
            ReadKey(true);
        }

        static void SetTimer(int hour, int minute, int second)
        {
            timer1 = new System.Timers.Timer();
            timer2 = new System.Timers.Timer();
            TimeSpan time1 = DateTime.Today + new TimeSpan(hour, minute, second) - DateTime.Now;
            timer1.Interval = time1.TotalMilliseconds; // Intervalの設定単位はミリ秒
            timer2.Interval = 2000;
            timer1.Elapsed += PushButton; // タイマイベント処理(時間経過後の処理)を登録
            timer2.Elapsed += MoveCursor;
            timer1.Enabled = true; // <-- これを呼ばないとタイマは開始しません
            timer2.Enabled = true;


        }

        private static void PushButton(object sender, System.Timers.ElapsedEventArgs e)
        {
            timer1.Enabled = false;
            timer2.Enabled = false;
            NativeMethods.SetCursorPos(x, y);
            Console.WriteLine("呼ばれました");
            NativeMethods.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            NativeMethods.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            //ここからウィンドウ切り替え＆自動投稿処理
            SwitchWindow(switch_window_name); //switch_window_nameが含まれるウインドウに切り替え
            //切り替えに時間がかかる場合があるので少し待つ
            Thread.Sleep(7000);
            PasteAndPost();
        }

        private static void MoveCursor(object sender, System.Timers.ElapsedEventArgs e)
        {
            NativeMethods.SetCursorPos(movex, movey);
            movex = movex > 820 ? 750 : movex + 10;
        }

        private static void SwitchWindow(string name)
        {
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses())
            {
                //{name}がメインウィンドウのタイトルに含まれているか調べる
                if (0 <= p.MainWindowTitle.IndexOf(name))
                {
                    //ウィンドウをアクティブにする
                    NativeMethods.SetForegroundWindow(p.MainWindowHandle);
                    break;
                }
            }
        }

        private static void PasteAndPost()
        {
            SendKeys.SendWait("^V");
            Thread.Sleep(1000);
            SendKeys.SendWait("{ENTER}");
        }

    }

    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}