using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ProjectJarvis
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer JARVIS = new SpeechSynthesizer();
        public NotifyIcon mynotify = new NotifyIcon();
        int count = 1;
        int temp = 0;

        //Random rnd = new Random();

        //[DllImport("user32")]
        //public static extern bool ExitWindowsEx(uint uFlags, uint dwReason);

        [DllImport("user32")]
        public static extern void LockWorkStation();
        public Form1()
        {

            InitializeComponent();

            string[] commands = (File.ReadAllLines(@"C:\command.txt"));
            lstCommands.Items.Clear();
            lstCommands.SelectionMode = SelectionMode.None;
            lstCommands.Visible = false;
            foreach (string command in commands)
            {
                lstCommands.Items.Add(command);
            }



            string comport = "COM10";
            serialPort1.PortName = comport;
            serialPort1.BaudRate = 9600;
            //if (!serialPort1.IsOpen)
            //serialPort1.Open();
            try
            {
                if (!serialPort1.IsOpen)
                    serialPort1.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //for using system's default microphone . select audio input from devices ,files, or stream//
            _recognizer.SetInputToDefaultAudioDevice();

            //Input Command file
            _recognizer.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"C:\command.txt")))));
            _recognizer.LoadGrammar(new DictationGrammar());
            _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(_recognizer_SpeechRecognized);
            //this specifies recognition does not terminate after completion
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);
            //JARVIS.Speak("goodbye jarvis");
        }

        void _recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //int ranNum = rnd.Next(1, 10);
            string speech = e.Result.Text;
            if (temp == 1)
            {
                speech = speech.Replace(" ", "+");
                Process.Start("https://www.google.com.bd/?gws_rd=cr,ssl&ei=eoGUU4GvC9COuASioIGQAQ#q=" + speech);
                temp = 0;
                return;
            }
            switch (speech)
            {
                //GREETINGS
                case "hi":
                case "hey":
                case "hello":
                case "hello jarvis":
                    //if (ranNum < 6)
                    //{
                    JARVIS.Speak("Hello sir");
                    //}
                    //else if (ranNum > 5)
                    //{
                    //JARVIS.Speak("Hi");
                    //}
                    break;
                case "goodbye":
                    JARVIS.Speak("See you later Sir!");
                    Close();
                    break;

                //WEBSITES
                case "open google":
                    Process.Start("http://www.google.com");
                    break;
                case "open stackoverflow":
                    Process.Start("http://www.stackoverflow.com");
                    break;
                case "open MSDN":
                    Process.Start("https://msdn.microsoft.com/en-US/");
                    break;
                case "open yahoo":
                    Process.Start("https://www.yahoo.com");
                    break;
                case "open tech shop":
                    Process.Start("http://www.techshopbd.com/");
                    break;
                case "open speed test":
                    Process.Start("http://www.speedtest.net/");
                    break;
                case "open wiki":
                    Process.Start("http://www.wikipedia.org/");
                    break;
                case "open daily star":
                    Process.Start("http://www.thedailystar.net/");
                    break;

                //Search Google by speaking "find"
                case "find":
                    {
                        JARVIS.Speak("What do you want to search sir !!");
                        temp = 1;
                    }
                    break;

                //SHELL COMMANDS
                case "run unity":
                    Process.Start(@"C:\Program Files (x86)\Unity\Editor\Unity.exe");
                    JARVIS.Speak("Loading Unity 3d engine!!");
                    break;
                case "run calculator":
                    Process proc = new Process();
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = "calc";
                    JARVIS.Speak("Loading Calculator");
                    proc.Start();
                    break;
                case "run firefox":
                    Process.Start(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe");
                    JARVIS.Speak("Loading Mozilla Firefox");
                    break;
                case "run skype":
                    Process.Start(@"C:\Program Files (x86)\Skype\Phone\Skype.exe");
                    JARVIS.Speak("Loading Skype");
                    break;
                case "run IE":
                    Process.Start(@"C:\Program Files\Internet Explorer\iexplore.exe");
                    JARVIS.Speak("Loading Internet Explorer");
                    break;
                case "run command":
                    Process proc1 = new Process();
                    proc1.StartInfo.FileName = "cmd.exe";
                    proc1.Start();
                    break;

                //CLOSE PROGRAMS
                case "close unity":
                    foreach (Process Proc1 in Process.GetProcesses())
                    {
                        if (Proc1.ProcessName.Equals("Unity"))
                            Proc1.Kill();
                    }
                    break;
                case "close calculator":
                    foreach (Process Proc2 in Process.GetProcesses())
                    {
                        if (Proc2.ProcessName.Equals("calc"))
                            Proc2.Kill();
                    }
                    break;
                case "close firefox":
                    foreach (Process Proc3 in Process.GetProcesses())
                    {
                        if (Proc3.ProcessName.Equals("firefox"))
                            Proc3.Kill();
                    }
                    break;
                case "close IE":
                    foreach (Process Proc in Process.GetProcesses())
                    {
                        if (Proc.ProcessName.Equals("iexplore"))
                            Proc.Kill();
                    }
                    break;

                //CONDITION OF DAY
                case "what time is it":
                    DateTime now = DateTime.Now;
                    string time = now.GetDateTimeFormats('t')[0];
                    JARVIS.Speak(time);
                    break;
                case "what day is it":
                    JARVIS.Speak(DateTime.Today.ToString("dddd"));
                    break;
                case "whats the date":
                case "whats todays date":
                    JARVIS.Speak(DateTime.Today.ToString("dd-MM-yyyy"));
                    break;

                case "switch window":
                    SendKeys.Send("%{TAB " + count + "}");
                    count += 1;
                    break;

                case "out of the way":
                    if (WindowState == FormWindowState.Normal || WindowState == FormWindowState.Maximized)
                    {
                        WindowState = FormWindowState.Minimized;
                        JARVIS.Speak("My apologies");
                    }
                    break;
                case "come back":
                    if (WindowState == FormWindowState.Minimized)
                    {
                        JARVIS.Speak("Alright Sir");
                        WindowState = FormWindowState.Normal;
                    }
                    break;
                case "show commands":
                    string[] commands = (File.ReadAllLines(@"C:\Users\imran\Desktop\New folder\command.txt"));
                    JARVIS.Speak("OK Sir");
                    commandsLabel.Visible = true;
                    lstCommands.Items.Clear();
                    lstCommands.SelectionMode = SelectionMode.None;
                    lstCommands.Visible = true;
                    foreach (string command in commands)
                    {
                        lstCommands.Items.Add(command);
                    }
                    break;
                case "hide listbox":
                    lstCommands.Visible = false;
                    commandsLabel.Visible = false;
                    break;

                //SHUTDOWN RESTART LOG OFF
                case "shutdown":
                    Process.Start("shutdown", "/s /t 0");
                    break;
                //case "log off":
                //JARVIS.Speak("logging off");
                //ExitWindowsEx(0, 0);
                //break;
                case "restart":
                    Process.Start("shutdown", "/r /t 0");
                    break;
                case "lock PC":
                    JARVIS.Speak("locking");
                    LockWorkStation();
                    break;

                //Hardware Control
                case "fan on":
                    //arduino pin no.11 status : high
                    serialPort1.Write("a");
                    JARVIS.Speak("ok, sir");
                    break;
                case "fan off":
                    //arduino pin no.11 status : low
                    serialPort1.Write("A");
                    JARVIS.Speak("fan is off, sir");
                    break;

                case "light on":
                    //arduino pin no.12 status : high
                    serialPort1.Write("b");
                    JARVIS.Speak("ok, sir");
                    break;
                case "light off":
                    //arduino pin no.12 status : low
                    serialPort1.Write("B");
                    JARVIS.Speak("light is off, sir");
                    break;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            JARVIS.Speak("goodbye sir");

            this.Close();
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {

            if (this.WindowState == FormWindowState.Minimized)
            {
                mynotify.Visible = true;
                this.ShowInTaskbar = false;
                mynotify.Icon = SystemIcons.Application;
                mynotify.BalloonTipText = "voicEBot running in the background, double click to restore";
                mynotify.ShowBalloonTip(1000);

            }

            else if (this.WindowState == FormWindowState.Normal)
            {
                mynotify.Visible = false;
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            mynotify.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}


