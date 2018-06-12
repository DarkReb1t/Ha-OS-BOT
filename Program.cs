using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Speech.Recognition;
using System.Speech.Synthesis;
using System.Speech;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Net.NetworkInformation;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Net.Mail;

namespace HA_OS_BOT
{
    public static class CDRom
    {
        [DllImport("winmm.dll")]
        static extern Int32 mciSendString(String command, StringBuilder buffer, Int32 bufferSize, IntPtr hwndCallback);
        public static void OpenCDRom()
        {
            char dl = 'E';
            DriveInfo[] allDrives = DriveInfo.GetDrives();
            foreach (DriveInfo d in allDrives)
            {
                if (d.DriveType == DriveType.CDRom) dl = d.Name[0];
            }
            mciSendString("open " + dl + ": type CDAudio alias drive" + dl, null, 0, IntPtr.Zero);
            mciSendString("set drive" + dl + " door open", null, 0, IntPtr.Zero);
        }
        public static void CloseCDRom()
        {
            char dl = 'E';
            DriveInfo[] allDrives = DriveInfo.GetDrives();
            foreach (DriveInfo d in allDrives)
            {
                if (d.DriveType == DriveType.CDRom) dl = d.Name[0];
            }
            mciSendString("open " + dl + ": type CDAudio alias drive" + dl, null, 0, IntPtr.Zero);
            mciSendString("set drive" + dl + " door closed", null, 0, IntPtr.Zero);
        }
    }

    class INI
    {
        string Path; //Имя файла.

        [DllImport("kernel32")] // Подключаем kernel32.dll и описываем его функцию WritePrivateProfilesString
        static extern long WritePrivateProfileString(string Section, string Key, string Value, string FilePath);

        [DllImport("kernel32")] // Еще раз подключаем kernel32.dll, а теперь описываем функцию GetPrivateProfileString
        static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);

        // С помощью конструктора записываем пусть до файла и его имя.
        public INI(string IniPath)
        {
            Path = new FileInfo(IniPath).FullName.ToString();
        }

        //Читаем ini-файл и возвращаем значение указного ключа из заданной секции.
        public string ReadINI(string Section, string Key)
        {
            var RetVal = new StringBuilder(255);
            GetPrivateProfileString(Section, Key, "", RetVal, 255, Path);
            return RetVal.ToString();
        }
        //Записываем в ini-файл. Запись происходит в выбранную секцию в выбранный ключ.
        public void Write(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, Path);
        }

        //Удаляем ключ из выбранной секции.
        public void DeleteKey(string Key, string Section = null)
        {
            Write(Section, Key, null);
        }
        //Удаляем выбранную секцию
        public void DeleteSection(string Section = null)
        {
            Write(Section, null, null);
        }
        //Проверяем, есть ли такой ключ, в этой секции
        public bool KeyExists(string Key, string Section = null)
        {
            return ReadINI(Section, Key).Length > 0;
        }
    }

    class VarIS
    {
        public bool isBreak = false;
        public VarIS ()
        {
           
        }
        
    }
    class MySettings
    {
        public string myIP, myHost;
        //Конструктор класса с двумя текстовыми аргументами:
        public MySettings(string IP, string Host)
        {
            myIP = IP;
            myHost = Host;
        }

        //Открытый метод для отображения значений поля:
        public void show()
        {
            Console.WriteLine("Мой IP-адресс: {0}, мой Хост: {1}", myIP, myHost);
            Console.WriteLine();
        }
    }
    class PingCheck
    {
        string IP = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName()).AddressList[0].ToString();
        //Конструктор класса с одним аргументом:
        public PingCheck(string A)
        {
            Ping ping = new Ping();
            PingReply pingReply = null;
            string server = A;
            pingReply = ping.Send(server);
            if (pingReply.Status == IPStatus.Success)
            {
                if (IP == server) { Console.ForegroundColor = ConsoleColor.Blue; Console.WriteLine(server + " - Ваш ноут!"); Console.ForegroundColor = ConsoleColor.White; Console.WriteLine(); }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(server + " - Онлайн!"); //server          
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine();
                }
            }
        }
    }

    public class NetworkSettingsDemo
    {
        public static void ScanIP(int startIP, int endIP)
        {
            string IPAd="";
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface adapter in adapters)
            {
                IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
                GatewayIPAddressInformationCollection addresses = adapterProperties.GatewayAddresses;
                if (addresses.Count > 0)
                {
                    Console.WriteLine(adapter.Description);
                    foreach (GatewayIPAddressInformation address in addresses)
                    {
                        Console.WriteLine("  Адрес шлюза: {0}",
                          address.Address.ToString());

                        for (int i = 0, j = 0; i < address.Address.ToString().Length; i++)
                        {
                            if (j == 3 && address.Address.ToString()[i-1]=='.') break;
                            if (address.Address.ToString()[i] == '.') j++;
                            if (j < 3) IPAd += address.Address.ToString()[i];
                        }
                    }

                    Console.WriteLine();
                }
            }
            Console.WriteLine("Активные устройства в подсети:");
            for (int i = startIP; i <= endIP; i++)
            {
                string ipnum = IPAd + "." + i;
                PingCheck pch = new PingCheck(ipnum);
            }
            Console.WriteLine("Сканирование завершено");
            Console.WriteLine();

        }

        public static void ScanIP()
        {
            string IPAd = "0.0.0.0";
            int FirstAd = 0;
            string lol="";
            //Поля, содержащие IP адрес и Хост:
            VarIS varIS = new VarIS();
            //Вызываем метод show():
            //string adressA = "";
            Console.WriteLine("Шлюзы:");
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface adapter in adapters)
            {
                IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
                GatewayIPAddressInformationCollection addresses = adapterProperties.GatewayAddresses;
                if (addresses.Count > 0)
                {
                    Console.WriteLine(adapter.Description);
                    foreach (GatewayIPAddressInformation address in addresses)
                    {
                        Console.WriteLine("  Адрес шлюза: {0}",
                          address.Address.ToString());
                          IPAd = "";
                        for (int i=0,j=0;i<address.Address.ToString().Length;i++)
                        {                            
                            if (address.Address.ToString()[i] == '.') j++;
                            if (j == 3) i++;
                            if (j<3)IPAd += address.Address.ToString()[i];
                            else lol+=address.Address.ToString()[i];
                        }
                    }
                    Console.WriteLine(IPAd);
                    Console.WriteLine(lol);
                }
            }
            Console.WriteLine("Активные устройства в подсети:");

            for (int i = Convert.ToInt32(FirstAd); i <= Convert.ToInt32(FirstAd) + 15; i++)
            {
                if (varIS.isBreak) { varIS.isBreak = false; break; }
                string ipnum = IPAd + '.' + i;
                PingCheck pch = new PingCheck(ipnum);
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Сканирование завершено");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();

        }

        public static void TWeb()
        {

            SpeechSynthesizer ss = new SpeechSynthesizer();
            ss.Rate = 3;

            string IP = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName()).AddressList[0].ToString();

            string IPAd = "0.0.0.0";
            int FirstAd = 0;
            string lol = "";


            //Поля, содержащие IP адрес и Хост:
            VarIS varIS = new VarIS();
            //Вызываем метод show():
            //string adressA = "";
            Console.WriteLine("Шлюзы:");
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            Ping ping = new Ping();
            PingReply pingReply = null;
            foreach (NetworkInterface adapter in adapters)
            {
                IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
                GatewayIPAddressInformationCollection addresses = adapterProperties.GatewayAddresses;
                if (addresses.Count > 0)
                {
                    Console.WriteLine(adapter.Description);
                    foreach (GatewayIPAddressInformation address in addresses)
                    {
                        Console.WriteLine("  Адрес шлюза: {0}",
                          address.Address.ToString());
                            IPAd = "";
                        for (int i = 0, j = 0; i < address.Address.ToString().Length; i++)
                        {
                            if (address.Address.ToString()[i] == '.') j++;
                            if (j == 3) i++;
                            if (j < 3) IPAd += address.Address.ToString()[i];
                            else lol += address.Address.ToString()[i];
                        }
                    }
                    Console.WriteLine(IPAd);
                    Console.WriteLine(lol);
                }
            }
            int cIP=0;
            while (true)
            {
                int j = 0;
                Console.Clear();
                Console.WriteLine("Активные устройства в подсети:");

                for (int i = Convert.ToInt32(FirstAd); i <= Convert.ToInt32(FirstAd) + 15; i++)
                {
                    if (varIS.isBreak) { varIS.isBreak = false; break; }
                    string ipnum = IPAd + '.' + i;
                    pingReply = ping.Send(ipnum);
                    if (pingReply.Status == IPStatus.Success)
                    {
                        if (IP == ipnum) { Console.ForegroundColor = ConsoleColor.Blue; Console.WriteLine(ipnum + " Ваш ноут!"); Console.ForegroundColor = ConsoleColor.White; Console.WriteLine(); }
                        else
                        {
                            j++;
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine(ipnum + " Онлайн!"); //server          
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine();
                        }
                    }
                }
                if (cIP<j)
                {
                    if (j - cIP == 1)
                        ss.SpeakAsync("Найдено одно новоё устройство!");
                    else if (j - cIP < 5) ss.SpeakAsync("Найдено " + (j - cIP).ToString() + " новых устройства!");
                    else ss.SpeakAsync("Найдено " + (j-cIP).ToString() + " новых устройств!");
                }
                if (cIP>j)
                {
                    if (cIP-j == 1)
                        ss.SpeakAsync("Отключено одно устройство!");
                    else if (cIP - j < 5) ss.SpeakAsync("Отключено " + (j - cIP).ToString() + " устройства!");
                    else ss.SpeakAsync("Отключено " + (j - cIP).ToString() + " устройств!");
                }
                else
                {

                }
                cIP = j;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Сеть проверена)");
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine();
            }
        }

    }

    class Program
    {
        //[DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        //public static extern IntPtr GetForegroundWindow();
        //public static extern IntPtr GetWindowThreadProcessId(IntPtr a,out int b);
        static string IMGSS = "None";
        static void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            SpeechSynthesizer ss = new SpeechSynthesizer();
            ss.Rate = 3;
            if (e.Result.Confidence > 0.5)
            {
                Console.WriteLine("Выполняю>>" + e.Result.Text + ".");
                string str = e.Result.Text;

                if (str.IndexOf("Пожалуйста") > -1)
                {
                    ss.Speak("лол");
                }
                if (str.IndexOf("Пикни") > -1)
                {
                    Console.Beep();
                }
                if (str.IndexOf("Вот Это Шутка") > -1|| str.IndexOf("Ха ха ха") > -1)
                {
                    ss.Speak("Ха   Ха   Ха   Ха   Ха");
                }

                if (str == "Выключись")
                {
                    Process.Start("shutdown.exe", "-t 00 -s");
                }

                if (str.IndexOf("Открой")>-1 && str.IndexOf("Дисковод") > -1)
                {
                    CDRom.OpenCDRom();
                }

                if (str.IndexOf("Закрой") > -1 && str.IndexOf("Дисковод") > -1)
                {
                    CDRom.CloseCDRom();
                }

                if (str == "Пока")
                {
                    ss.Rate = 3;
                    ss.Speak("ПАКА!!!");
                    ss.Rate = 1;
                }
                if (str == "Широ")
                {
                   ss.Speak("Да, хазяин!");
                }
                if (str == "Как Тебя Зовут")
                {
                    Random rand = new Random();
                    int a = rand.Next(0, 10);
                    if (a>5) ss.Speak("Я, Широ");
                    else if (a < 5) ss.Speak("Меня зовут, Широ");
                }

                if (str.IndexOf("Привет") > -1 || str.IndexOf("Хай") > -1 || str.IndexOf("Хелоу") > -1 || str.IndexOf("Бонжур") > -1)
                {
                    Random rand = new Random();
                    int a = rand.Next(0, 10);
                    if (a > 5) ss.Speak("Привет");
                    else if (a < 5) ss.Speak("Хай");
                }
                if ((str.IndexOf("Скажи") > -1 || str.IndexOf("Сколько") > -1 || str.IndexOf("Который") > -1) && (str.IndexOf("Время") > -1 || str.IndexOf("Времени") > -1 || str.IndexOf("Час") > -1))
                {
                    ss.Speak(DateTime.Now.ToLongTimeString());
                }
                if ((str.IndexOf("Скажи") > -1 || str.IndexOf("Какое") > -1) && (str.IndexOf("Число") > -1) && (str.IndexOf("Завтра") > -1))
                {
                    ss.Speak((DateTime.Now.Day + 1).ToString() + "-ое");
                }
                else if ((str.IndexOf("Скажи") > -1 || str.IndexOf("Какое") > -1) && (str.IndexOf("Число") > -1) && (str.IndexOf("Послезавтра") > -1))
                {
                    ss.Speak((DateTime.Now.Day + 2).ToString() + "-ое");
                }
                else if ((str.IndexOf("Скажи") > -1 || str.IndexOf("Какое") > -1) && (str.IndexOf("Число") > -1))
                {
                    ss.Speak(DateTime.Now.Day.ToString() + "-ое");
                }
                if ((str.IndexOf("Скажи") > -1 || str.IndexOf("Какая") > -1) && (str.IndexOf("Дату") > -1 || str.IndexOf("Дата") > -1))
                {
                    ss.Speak(DateTime.Now.ToLongDateString());
                }
                if (str.IndexOf("Харэ") > -1 || str.IndexOf("Шухер") > -1 || str.IndexOf("Прекрати") > -1 || str.IndexOf("Стоп") > -1)
                {
                    VarIS var = new VarIS();
                    var.isBreak = true;
                }
                if ((str.IndexOf("Просканируй") > -1 || str.IndexOf("Сканируй") > -1 || str.IndexOf("Проскань") > -1) && (str.IndexOf("Сеть") > -1 || str.IndexOf("Айпи") > -1))
                {
                    ScanIp();
                }
                if ((str.IndexOf("Выведи") > -1 || str.IndexOf("Напечатай") > -1) && (str.IndexOf("Данные") > -1 || str.IndexOf("Инфу") > -1))
                {
                    //MySettings settings = new MySettings();

                }

                if ((str.IndexOf("Сделай") > -1 || str.IndexOf("Сфоткай") > -1 ) && (str.IndexOf("Скриншот") > -1 || str.IndexOf("Экран") > -1 || (str.IndexOf("Скрин") > -1)))
                {
                    ScreenShot();
                }
                if ((str.IndexOf("Покажи") > -1 || str.IndexOf("Открой") > -1) && (str.IndexOf("Скриншот") > -1 || (str.IndexOf("Скрин") > -1)))
                {
                    OpenScreenShot();
                }

                if (str.IndexOf("Транслируй") > -1 && str.IndexOf("Сеть") > -1)
                {
                    NetworkSettingsDemo.TWeb();
                }

                if (str.IndexOf("Закрой") > -1 && str.IndexOf("Это") > -1)
                {
                    
                }

                if ((str.IndexOf("Спасибо") > -1))
                {
                    ss.Speak("Пожалуйста");
                }

                //Ишем n-ное количество подстрок в строке
            }
        }

        static void Speech()
        {
            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ru-ru");
            SpeechRecognitionEngine sre = new SpeechRecognitionEngine(ci);
            sre.SetInputToDefaultAudioDevice();

            sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);



            GrammarBuilder gb = new GrammarBuilder();
            gb.Culture = ci;

            gb.Append(new Choices(new string[] { " ", "Привет", "Пожалуйста", "Спасибо" , "Всем", "Пока", "Пикни" }));
            gb.Append(new Choices(new string[] { " ", "Привет", "Пожалуйста", "Всем", "Пока" }));
            gb.Append(new Choices(new string[] { " ", "Вот Это Шутка", "Ха ха ха", "Широ" }));
            gb.Append(new Choices(new string[] { " ", "Харэ", "Шухер", "Прекрати", "Стоп" }));
            gb.Append(new Choices(new string[] { " ", "Открой", "Закрой", "Напиши", "Транслируй", "Найди", "Сделай", "Сфоткай", "Скажи", "Выведи", "Скинь", "Напечатай", "Покажи", "Выключись", "Сколько", "Какая", "Какое", "Просканируй", "Проскань", "Сканируй" }));
            gb.Append(new Choices(new string[] { " ", "Как Тебя Зовут" }));
            gb.Append(new Choices(new string[] { " ", "Инфу", "Информацию", "Данные", "Айпи", "Сеть", "Дисковод", "Сети", "Скриншот", "Скрин", "Экран", "Это" }));
            //gb.Append(new Choices(new string[] { " ", "1-ого", "2-ого", "3-ого", "5-ого", "6-ого", "20-ого" }));
            gb.Append(new Choices(new string[] { " ", "Время", "Времени", "Дата", "Дату", "Число", "Сейчас", "Сегодня", "Завтра", "Послезавтра", "Который", "Час", "Компьютер", "Сети", "Сеть", "Доступ" }));
            gb.Append(new Choices(new string[] { " ", "Время", "Времени", "Дата", "Дату", "Число", "Сейчас", "Сегодня", "Завтра", "Послезавтра", "Который", "Час", "Компьютер" }));
            gb.Append(new Choices(new string[] { " " }));


            Grammar g = new Grammar(gb);
            sre.LoadGrammar(g);

            sre.RecognizeAsync(RecognizeMode.Multiple);
        }

        static void ScanIp()
        {
            NetworkSettingsDemo.ScanIP();
        }
        static void ScanIp(int q, int w)
        {
            NetworkSettingsDemo.ScanIP(q, w);
        }

        static void FirstStart()
        {

        }

        /*static void CloseThis()
        {
            // получаем хендел активного окна
            IntPtr hWnd = GetForegroundWindow();
            int pid;
            //получаем pid потока активного окна
            GetWindowThreadProcessId(hWnd, out pid);
            // ввыводим в listbox PID процесса и имя процесса
            using (Process p = Process.GetProcessById(pid))
            {
                string test = String.Format("Программа {1}", p.Id, p.ProcessName);
                hWnd = IntPtr.Zero;
                Console.WriteLine(test);
            }
        }*/

        static void ScreenShot()
        {
            string path = "C:\\Screenshots\\";
            if (!Directory.Exists(path))
            {
                DirectoryInfo dir = new DirectoryInfo(path);
                dir.Create();
            }
            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

            Graphics graphics = Graphics.FromImage(printscreen as Image);

            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);

            string fp = path + DateTime.Now.ToShortDateString() + ".png";
            int i = 0;
            while (true)
            {
                if (!File.Exists(fp)) break;
                else
                {
                    i++;
                    fp = path + DateTime.Now.ToShortDateString()+" ("+ i + ").png";
                }
            }
            printscreen.Save(fp, System.Drawing.Imaging.ImageFormat.Jpeg);
            IMGSS = fp;
            
        }
        static void OpenScreenShot()
        {
            Process.Start(IMGSS);
        }

        static void Main(string[] args)
        {
            Speech();
            Console.WriteLine("Loading complite");
            Console.Beep();
            Console.Title = "HA-OS-BOT";
            while (true)
            {
                
            }
        }
    }
 }

