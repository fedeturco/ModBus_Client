

// -------------------------------------------------------------------------------------------

// Copyright (c) 2024 Federico Turco

// Permission is hereby granted, free of charge, to any person
// obtaining a copy of this software and associated documentation
// files (the "Software"), to deal in the Software without
// restriction, including without limitation the rights to use,
// copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following
// conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
// OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
// HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
// OTHER DEALINGS IN THE SOFTWARE.

// -------------------------------------------------------------------------------------------




using LanguageLib; // Libreria custom per caricare etichette in lingue differenti
using Microsoft.Win32;
using ModBusMaster_Chicco;
// Classe con funzioni di conversione DEC-HEX
using Raccolta_funzioni_parser;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
//Libreria JSON
//using Newtonsoft.Json;
using System.IO;
using System.IO.Compression;
//Porta seriale
using System.IO.Ports;
using System.Linq;
// Ping
using System.Net.NetworkInformation;
//Process

//Sockets
using System.Net.Sockets;
using System.Reflection;
//Comandi apri/chiudi console
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
//Threading per server ModBus TCP
using System.Threading;
// Json LIBs
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shell;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //-----------------------------------------------------
        //-----------------Variabili globali-------------------
        //-----------------------------------------------------

        public String version = "beta";    // Eventuale etichetta, major.minor lo recupera dall'assembly
        public String title = "ModBus Client";

        String defaultPathToConfiguration = "Default";
        public String pathToConfiguration;
        public String localPath = "";

        SolidColorBrush colorDefaultReadCell = Brushes.DarkBlue;
        SolidColorBrush colorDefaultWriteCell = Brushes.LightGreen;
        SolidColorBrush colorErrorCell = Brushes.Orange;

        public String colorDefaultReadCellStr;
        public String colorDefaultWriteCellStr;
        public String colorErrorCellStr;

        SolidColorBrush colorDefaultReadCell_Light = Brushes.DarkBlue;
        SolidColorBrush colorDefaultWriteCell_Light = Brushes.LightGreen;
        SolidColorBrush colorErrorCell_Light = Brushes.Orange;

        SolidColorBrush colorDefaultReadCell_Dark = Brushes.DarkBlue;
        SolidColorBrush colorDefaultWriteCell_Dark = Brushes.LightGreen;
        SolidColorBrush colorErrorCell_Dark = Brushes.Orange;

        public ObservableCollection<ModBus_Item> list_coilsTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_inputsTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_inputRegistersTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_holdingRegistersTable = new ObservableCollection<ModBus_Item>();

        System.Windows.Forms.ColorDialog colorDialogBox = new System.Windows.Forms.ColorDialog();

        Parser P = new Parser();

        // Elementi per visualizzare/nascondere la finestra della console
        bool statoConsole = false;

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        // Disable Console Exit Button
        [DllImport("user32.dll")]
        static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll")]
        static extern IntPtr DeleteMenu(IntPtr hMenu, uint uPosition, uint uFlags);

        const uint SC_CLOSE = 0xF060;
        const uint MF_BYCOMMAND = (uint)0x00000000L;

        //Coda per i comandi seriali da inviare
        //Queue BufferSerialeOut = new Queue();

        public ModBus_Chicco ModBus;
        public ModBus_Def ModBus_Def = new ModBus_Def();

        SerialPort serialPort = new SerialPort();

        SaveFileDialog saveFileDialogBox;
        //OpenFileDialog openFileDialogBox;

        // Le liste seguenti contengono il registro già convertito in DEC duramte il caricamento del file Template.json
        public ObservableCollection<ModBus_Item> list_template_coilsTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_template_inputsTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_template_holdingRegistersTable = new ObservableCollection<ModBus_Item>();
        public ObservableCollection<ModBus_Item> list_template_inputRegistersTable = new ObservableCollection<ModBus_Item>();

        public ModBus_Item lastEditModbusItem;

        // Stati loop interrogazioni
        public int pauseLoop = 1000;
        public bool loopCoils01 = false;
        public bool loopCoilsRange = false;
        public bool loopInput02 = false;
        public bool loopInputRange = false;
        public bool loopInputRegister04 = false;
        public bool loopInputRegisterRange = false;
        public bool loopHolding03 = false;
        public bool loopHoldingRange = false;
        public bool loopThreadRunning = false;

        Thread threadLoopQuery;

        public string language = "IT";

        public bool logWindowIsOpen = false;
        public bool dequeueExit = false;

        Thread threadDequeue;

        public int readTimeout;

        public Language lang;

        public int LogLimitRichTextBox = 2000;

        bool scrolled_log = false;
        int count_log = 0;

        public bool lastCoilsCommandGroup = false;
        public bool lastInputsCommandGroup = false;
        public bool lastHoldingRegistersCommandGroup = false;
        public bool lastInputRegistersCommandGroup = false;

        public bool importingProfile = false;

        // Variabili di appoggio per gestire le chiamate a thread

        public bool useOffsetInTable = false;
        public bool correctModbusAddressAuto = false;
        public bool colorMode = false;
        public bool darkMode = false;
        public bool disableDeleteKey = false;
        public bool disableComboProfile = true;
        public bool useOnlyReadSingleRegisterForGroups = false;
        public bool sendCellEditOnlyOnChange = false;

        public String textBoxModbusAddress_ = "";
        public String comboBoxCoilsRegistri_ = "";
        public String textBoxCoilsOffset_ = "";
        public String comboBoxCoilsOffset_ = "";
        public String comboBoxCoilsAddress01_ = "";
        public String textBoxCoilsAddress01_ = "";
        public String textBoxCoilNumber_ = "";
        public String comboBoxCoilsRange_A_ = "";
        public String textBoxCoilsRange_A_ = "";
        //public String comboBoxCoilsRange_B_ = "";
        public String textBoxCoilsRange_B_ = "";
        public String comboBoxCoilsAddress05_ = "";
        public String textBoxCoilsAddress05_ = "";
        public String textBoxCoilsValue05_ = "";
        public String comboBoxCoilsAddress05_b_ = "";
        public String textBoxCoilsAddress05_b_ = "";
        public String textBoxCoilsValue05_b_ = "";
        public String comboBoxCoilsAddress15_A_ = "";
        public String textBoxCoilsAddress15_A_ = "";
        public String comboBoxCoilsAddress15_B_ = "";
        public String textBoxCoilsAddress15_B_ = "";
        public String textBoxCoilsValue15_ = "";

        public String comboBoxInputRegistri_ = "";
        public String comboBoxInputOffset_ = "";
        public String textBoxInputOffset_ = "";
        public String comboBoxInputAddress02_ = "";
        public String textBoxInputAddress02_ = "";
        public String textBoxInputNumber_ = "";
        public String comboBoxInputRange_A_ = "";
        public String textBoxInputRange_A_ = "";
        //public String comboBoxInputRange_B_ = "";
        public String textBoxInputRange_B_ = "";

        public String comboBoxInputRegRegistri_ = "";
        public String comboBoxInputRegValori_ = "";
        public String comboBoxInputRegOffset_ = "";
        public String textBoxInputRegOffset_ = "";
        public String comboBoxInputRegisterAddress04_ = "";
        public String textBoxInputRegisterAddress04_ = "";
        public String textBoxInputRegisterNumber_ = "";
        public String comboBoxInputRegisterRange_A_ = "";
        public String textBoxInputRegisterRange_A_ = "";
        //public String comboBoxInputRegisterRange_B_ = "";
        public String textBoxInputRegisterRange_B_ = "";

        public String comboBoxHoldingRegistri_ = "";
        public String comboBoxHoldingValori_ = "";
        public String comboBoxHoldingOffset_ = "";
        public String textBoxHoldingOffset_ = "";
        public String comboBoxHoldingAddress03_ = "";
        public String textBoxHoldingAddress03_ = "";
        public String textBoxHoldingRegisterNumber_ = "";
        public String comboBoxHoldingRange_A_ = "";
        public String textBoxHoldingRange_A_ = "";
        //public String comboBoxHoldingRange_B_ = "";
        public String textBoxHoldingRange_B_ = "";
        public String comboBoxHoldingAddress06_ = "";
        public String textBoxHoldingAddress06_ = "";
        public String comboBoxHoldingValue06_ = "";
        public String textBoxHoldingValue06_ = "";
        public String comboBoxHoldingAddress06_b_ = "";
        public String textBoxHoldingAddress06_b_ = "";
        public String comboBoxHoldingValue06_b_ = "";
        public String textBoxHoldingValue06_b_ = "";
        public String comboBoxHoldingAddress16_A_ = "";
        public String textBoxHoldingAddress16_A_ = "";
        public String comboBoxHoldingAddress16_B_ = "";
        public String textBoxHoldingAddress16_B_ = "";
        public String comboBoxHoldingValue16_ = "";
        public String textBoxHoldingValue16_ = "";

        String comboBoxHoldingGroup_ = "";
        String comboBoxInputRegisterGroup_ = "";
        String comboBoxInputGroup_ = "";
        String comboBoxCoilsGroup_ = "";

        // Dark mode
        public SolidColorBrush ForeGroundDark = new SolidColorBrush(Color.FromArgb(255, (byte)249, (byte)249, (byte)249));
        public SolidColorBrush BackGroundDark = new SolidColorBrush(Color.FromArgb(255, (byte)60, (byte)60, (byte)60));
        public SolidColorBrush BackGroundDark2 = new SolidColorBrush(Color.FromArgb(255, (byte)90, (byte)90, (byte)90));
        public SolidColorBrush BackGroundDarkButton = new SolidColorBrush(Color.FromArgb(255, (byte)60, (byte)60, (byte)60));

        // Test dark color v2
        //public SolidColorBrush BackGroundDark = new SolidColorBrush(Color.FromArgb(255, (byte)90, (byte)90, (byte)90));
        //public SolidColorBrush BackGroundDark2 = new SolidColorBrush(Color.FromArgb(255, (byte)60, (byte)60, (byte)60)); 

        public string ForeGroundDarkStr;
        public string BackGroundDarkStr;

        // Light mode
        public SolidColorBrush ForeGroundLight = new SolidColorBrush(Color.FromArgb(255, (byte)10, (byte)10, (byte)10));
        public SolidColorBrush BackGroundLight = new SolidColorBrush(Color.FromArgb(255, (byte)229, (byte)229, (byte)229));
        public SolidColorBrush BackGroundLight2 = new SolidColorBrush(Color.FromArgb(255, (byte)255, (byte)255, (byte)255));
        public SolidColorBrush BackGroundLightButton = new SolidColorBrush(Color.FromArgb(255, (byte)0xDD, (byte)0xDD, (byte)0xDD));

        public string ForeGroundLightStr;
        public string BackGroundLightStr;
        public string BackGroundLight2Str;

        public int MaxJsonLength = 104857600; // 200 MB, sufficienti per 65536*4 etichette Template.json
                                              // see https://learn.microsoft.com/it-it/dotnet/api/system.web.script.serialization.javascriptserializer.maxjsonlength?view=netframework-4.8.1

        public ThreadPriority threadPriority;

        public MainWindow()
        {
            localPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            InitializeComponent();

            version = Assembly.GetEntryAssembly().GetName().Version.Major.ToString() + "." + Assembly.GetEntryAssembly().GetName().Version.Minor.ToString();

            lang = new Language(this);

            threadPriority = ThreadPriority.Highest;

            // Creo evento di chiusura del form
            this.Closing += Form1_FormClosing;

            //this.textBoxHoldingRegisterNumber.KeyUp = new new KeyEventHandler(buttonReadHolding03_Click);

            pathToConfiguration = defaultPathToConfiguration;

            // Aspetti grafici
            comboBoxDiagnosticFunction.Items.Add("00 Return Query Data");
            comboBoxDiagnosticFunction.Items.Add("01 Restart Comunications Option");
            comboBoxDiagnosticFunction.Items.Add("02 Return Diagnostic Register");
            comboBoxDiagnosticFunction.Items.Add("03 Change ASCII Input Delimeter");
            comboBoxDiagnosticFunction.Items.Add("04 Force Listen Only Mode");
            comboBoxDiagnosticFunction.Items.Add("10 Clear Counters and Diagnostic Register");
            comboBoxDiagnosticFunction.Items.Add("11 Return Bus Message Count");
            comboBoxDiagnosticFunction.Items.Add("12 Return Bus Comunication Error Count");
            comboBoxDiagnosticFunction.Items.Add("13 Return Bus Exception Error Count");
            comboBoxDiagnosticFunction.Items.Add("14 Return Slave Message Count");
            comboBoxDiagnosticFunction.Items.Add("15 Return Slave No Response Count");
            comboBoxDiagnosticFunction.Items.Add("16 Return Slave NAK Count");
            comboBoxDiagnosticFunction.Items.Add("17 Return Slave Busy Count");
            comboBoxDiagnosticFunction.Items.Add("20 Clear Overrun Counter and Flag");

            comboBoxTlsVersion.Items.Add("1.2");
            comboBoxTlsVersion.Items.Add("1.3");

            pictureBoxSerial.Background = Brushes.LightGray;
            pictureBoxTcp.Background = Brushes.LightGray;

            dataGridViewCoils.ItemsSource = list_coilsTable;
            dataGridViewInput.ItemsSource = list_inputsTable;
            dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;
            dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

            // Aspetti grafici di default
            comboBoxSerialSpeed.SelectedIndex = 7;
            comboBoxSerialParity.SelectedIndex = 0;
            comboBoxSerialStop.SelectedIndex = 0;

            textBoxTcpClientIpAddress.Text = "192.168.1.100";
            textBoxTcpClientPort.Text = "502";

            comboBoxCoilsRegistri.SelectedIndex = 0;
            comboBoxCoilsOffset.SelectedIndex = 0;
            textBoxCoilsOffset.Text = "0";

            comboBoxInputRegistri.SelectedIndex = 0;
            comboBoxInputOffset.SelectedIndex = 0;
            textBoxInputOffset.Text = "0";

            comboBoxInputRegOffset.SelectedIndex = 0;
            comboBoxInputRegValori.SelectedIndex = 0;
            comboBoxInputRegRegistri.SelectedIndex = 0;
            textBoxInputRegOffset.Text = "0";

            comboBoxHoldingRegistri.SelectedIndex = 0;
            comboBoxHoldingValori.SelectedIndex = 0;
            comboBoxHoldingOffset.SelectedIndex = 0;
            textBoxHoldingOffset.Text = "0";

            comboBoxCoilsAddress01.SelectedIndex = 0;
            comboBoxCoilsRange_A.SelectedIndex = 0;
            //comboBoxCoilsRange_B.SelectedIndex = 0;
            comboBoxCoilsAddress05.SelectedIndex = 0;
            comboBoxCoilsAddress05_b.SelectedIndex = 0;
            //comboBoxCoilsValue05.SelectedIndex = 0;
            //comboBoxCoilsValue05_b.SelectedIndex = 0;
            comboBoxCoilsAddress15_A.SelectedIndex = 0;
            comboBoxCoilsAddress15_B.SelectedIndex = 0;
            //comboBoxCoilsValue15.SelectedIndex = 0;

            comboBoxInputAddress02.SelectedIndex = 0;
            comboBoxInputRange_A.SelectedIndex = 0;
            //comboBoxInputRange_B.SelectedIndex = 0;

            comboBoxInputRegisterAddress04.SelectedIndex = 0;
            comboBoxInputRegisterRange_A.SelectedIndex = 0;
            //comboBoxInputRegisterRange_B.SelectedIndex = 0;

            comboBoxHoldingAddress03.SelectedIndex = 0;
            comboBoxHoldingRange_A.SelectedIndex = 0;
            //comboBoxHoldingRange_B.SelectedIndex = 0;
            comboBoxHoldingAddress06.SelectedIndex = 0;
            comboBoxHoldingValue06.SelectedIndex = 0;
            comboBoxHoldingAddress06_b.SelectedIndex = 0;
            comboBoxHoldingValue06_b.SelectedIndex = 0;
            comboBoxHoldingAddress16_A.SelectedIndex = 0;
            comboBoxHoldingAddress16_B.SelectedIndex = 0;
            comboBoxHoldingValue16.SelectedIndex = 0;

            comboBoxTcpConnectionMode.Items.Add("Open on connect, close on disconnect");
            comboBoxTcpConnectionMode.Items.Add("Open and close socket on each request");
            comboBoxTcpConnectionMode.SelectedIndex = 0;

            pictureBoxRunningAs.Background = Brushes.LightGray;
            pictureBoxRx.Background = Brushes.LightGray;
            pictureBoxTx.Background = Brushes.LightGray;

            richTextBoxStatus.AppendText("\n");

            radioButtonModeSerial.IsChecked = true;

            checkBoxAddLinesToEnd.Visibility = Visibility.Hidden;

            // Disabilita il pulsante di chiusura della console
            disableConsoleExitButton();

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);

            ForeGroundDarkStr = ForeGroundDark.ToString();
            BackGroundDarkStr = BackGroundDark.ToString();

            ForeGroundLightStr = ForeGroundLight.ToString();
            BackGroundLightStr = BackGroundLight.ToString();
            BackGroundLight2Str = BackGroundLight2.ToString();

            // Certificato pfx (cert+key con password)
            // labelClientPasswordName.Visibility = Visibility.Visible;
            // textBoxCertificatePassword.Visibility = Visibility.Visible;
            // labelClientKeyName.Visibility = Visibility.Hidden;
            // labelClientKeyPath.Visibility = Visibility.Hidden;
            // buttonLoadClientKey.Visibility = Visibility.Hidden;

            CheckBoxModbusSecure_Checked(null, null);

            TaskbarItemInfo = new TaskbarItemInfo();
        }

        private void Form1_FormClosing(object sender, EventArgs e)
        {
            dequeueExit = true;

            //SaveConfiguration_v1(false);
            SaveConfiguration_v2(false);

            try
            {
                if (ModBus != null)
                {
                    ModBus.close(); // Se non attivo niente ModBUs risulta null
                }
            }
            catch
            {

            }

            try
            {
                if (threadLoopQuery != null)
                {
                    if (threadLoopQuery.IsAlive)
                    {
                        threadLoopQuery.Abort();
                    }
                }
            }
            catch
            {

            }
        }

        private void Form1_Load(object sender, RoutedEventArgs e)
        {
            threadDequeue = new Thread(new ThreadStart(LogDequeue));
            threadDequeue.IsBackground = true;
            threadDequeue.Start();

            richTextBoxPackets.Document.Blocks.Clear();
            richTextBoxPackets.AppendText("\n");
            richTextBoxPackets.Document.PageWidth = 5000;

            // Console.WriteLine(this.Title + "\n");

            this.Title = title + " " + version;

            try
            {
                // Aggiornamento lista porte seriale
                string[] SerialPortList = System.IO.Ports.SerialPort.GetPortNames();
                // comboBoxSerialPort.Items.Add("Seleziona porta seriale ...");

                foreach (String port in SerialPortList)
                {
                    comboBoxSerialPort.Items.Add(port);
                }

                comboBoxSerialPort.SelectedIndex = 0;
            }
            catch
            {
                Console.WriteLine("Nessuna porta seriale trovata");
            }

            // Menu lingua
            languageToolStripMenu.Items.Clear();

            foreach (string lang in Directory.GetFiles(localPath + "//Lang"))
            {
                var tmp = new MenuItem();

                if (lang.IndexOf("Mappings") == -1)
                {
                    tmp.Header = System.IO.Path.GetFileNameWithoutExtension(lang);
                    tmp.IsCheckable = true;
                    tmp.Click += MenuItemLanguage_Click;

                    languageToolStripMenu.Items.Add(tmp);
                }
            }

            comboBoxProfileHome.Items.Clear();

            foreach (String sub in Directory.GetDirectories(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\"))
            {
                comboBoxProfileHome.Items.Add(sub.Split('\\')[sub.Split('\\').Length - 1]);
            }

            comboBoxProfileHome.SelectedItem = pathToConfiguration;

            // Se esiste una nuova versione del file di configurazione uso l'ultima, altrimenti carico il modello precedente
            if (File.Exists(localPath + "\\Json\\" + pathToConfiguration + "\\Config.json"))
            {
                LoadConfiguration_v2();
            }
            else
            {
                LoadConfiguration_v1();
            }

            //lang.loadLanguageTemplate(language);

            Thread updateTables = new Thread(new ThreadStart(genera_tabelle_registri));
            updateTables.IsBackground = true;
            updateTables.Start();

            // ToolTips
            // toolTipHelp.SetToolTip(this.comboBoxHoldingAddress03, "Formato indirizzo\nDEC: decimale\nHEX: esadecimale (inserire il valore senza 0x)");

            if ((bool)radioButtonModeSerial.IsChecked)
            {
                buttonSerialActive.Focus();
            }
            else
            {
                buttonTcpActive.Focus();
            }

            changeEnableButtonsConnect(false);

            // Command line parameters
            string[] argv = Environment.GetCommandLineArgs();
            for (int i = 0; i < argv.Length; i++)
            {
                // debug
                // Console.WriteLine("{0}", argv[i]);

                // -h
                // --help
                if (argv[i].IndexOf("-h") != -1 || argv[i].IndexOf("--help") != -1)
                {
                    Console.WriteLine("Command line parameters:");
                    Console.WriteLine("");
                    Console.WriteLine("Load profile:");
                    Console.WriteLine("--profile \"Profile name\"");
                    Console.WriteLine(":");
                    Console.WriteLine("TCP connection:");
                    Console.WriteLine("--tcp 192.168.1.40 502");
                    Console.WriteLine("");
                    Console.WriteLine("RTU connection:");
                    Console.WriteLine("--rtu COM3 19200 8N1");
                    Console.WriteLine("");
                    Console.WriteLine("--profile \"Custom\" --tcp 192.168.1.40 502");
                    Console.WriteLine("--profile \"Custom\" --rtu COM3 9600 8E2");
                    Console.WriteLine("");
                    apriConsole();
                }

                // --profile "test 2"
                if (argv[i].IndexOf("--profile") != -1)
                {
                    radioButtonModeTcp.IsChecked = true;

                    if ((i + 1) < argv.Length)
                    {
                        LoadProfile(argv[i + 1].Replace("\"", ""));
                    }
                    else
                    {
                        Console.WriteLine("Not enough parameter for --profile");
                    }
                }

                // --tcp 127.0.0.1 502
                if (argv[i].IndexOf("--tcp") != -1)
                {
                    radioButtonModeTcp.IsChecked = true;

                    if ((i + 2) < argv.Length)
                    {
                        textBoxTcpClientIpAddress.Text = argv[i + 1];
                        textBoxTcpClientPort.Text = argv[i + 2];

                        buttonTcpActive_Click(null, null);
                    }
                    else
                    {
                        Console.WriteLine("Not enough parameter for --tcp");
                    }
                }

                // --rtu COM3 19200 8N1
                if (argv[i].IndexOf("--rtu") != -1)
                {
                    radioButtonModeSerial.IsChecked = true;

                    if ((i + 1) < argv.Length)
                    {
                        for (int ii = 0; ii < comboBoxSerialPort.Items.Count; ii++)
                        {
                            if (comboBoxSerialPort.Items[ii].ToString().IndexOf(argv[i + 1]) != -1)
                                comboBoxSerialPort.SelectedIndex = ii;
                        }

                        for (int ii = 0; ii < comboBoxSerialSpeed.Items.Count; ii++)
                        {
                            if (comboBoxSerialSpeed.Items[ii].ToString().IndexOf(argv[i + 2]) != -1)
                                comboBoxSerialSpeed.SelectedIndex = ii;
                        }

                        for (int ii = 0; ii < comboBoxSerialStop.Items.Count; ii++)
                        {
                            if (comboBoxSerialStop.Items[ii].ToString().IndexOf(argv[i + 3].Substring(2)) != -1)
                                comboBoxSerialStop.SelectedIndex = ii;
                        }

                        if (argv[i + 3].Substring(1, 1) == "N")
                            comboBoxSerialParity.SelectedIndex = 0;
                        if (argv[i + 3].Substring(1, 1) == "E")
                            comboBoxSerialParity.SelectedIndex = 1;
                        if (argv[i + 3].Substring(1, 1) == "O")
                            comboBoxSerialParity.SelectedIndex = 2;

                        buttonTcpActive_Click(null, null);
                    }
                    else
                    {
                        Console.WriteLine("Not enough parameter for --rtu");
                    }
                }


                if (argv[i].IndexOf("--tab") != -1)
                {
                    if ((i + 1) < argv.Length)
                    {
                        tabControlMain.SelectedIndex = int.Parse(argv[i + 1]);
                    }
                    else
                    {
                        Console.WriteLine("Not enough parameter for --tab");
                    }
                }
            }

            // Aggiorno grafica colori
            CheckBoxDarkMode_Checked(null, null);

            disableComboProfile = false;
        }

        private void radioButtonModeSerial_CheckedChanged(object sender, RoutedEventArgs e)
        {

            // Serial ON

            // radioButtonModeASCII.IsEnabled = radioButtonModeSerial.IsChecked;
            // radioButtonModeRTU.IsEnabled = radioButtonModeSerial.IsChecked;

            comboBoxSerialParity.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialPort.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialSpeed.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            comboBoxSerialStop.IsEnabled = (bool)radioButtonModeSerial.IsChecked;

            buttonUpdateSerialList.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            //buttonSerialActive.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
            buttonSerialActive.Visibility = (bool)radioButtonModeSerial.IsChecked ? Visibility.Visible : Visibility.Hidden;


            // Tcp OFF
            // radioButtonTcpSlave.IsEnabled = !radioButtonModeSerial.IsChecked;

            //richTextBoxStatus.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            //buttonTcpActive.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            buttonTcpActive.Visibility = !(bool)radioButtonModeSerial.IsChecked ? Visibility.Visible : Visibility.Hidden;
            ImageLogoRTU.Visibility = (bool)radioButtonModeSerial.IsChecked ? Visibility.Visible : Visibility.Hidden;
            ImageLogoTCP.Visibility = !(bool)radioButtonModeSerial.IsChecked ? Visibility.Visible : Visibility.Hidden;

            textBoxTcpClientIpAddress.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            textBoxTcpClientPort.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            buttonPingIp.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            buttonLoadClientCertificate.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            buttonLoadClientKey.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            textBoxCertificatePassword.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            comboBoxTlsVersion.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
            checkBoxModbusSecure.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;
        }

        public void buttonSerialActive_Click(object sender, RoutedEventArgs e)
        {
            if (pictureBoxSerial.Background == Brushes.LightGray)
            {
                // Attivazione comunicazione seriale
                pictureBoxSerial.Background = Brushes.Lime;
                pictureBoxRunningAs.Background = Brushes.Lime;


                textBlockSerialActive.Text = lang.languageTemplate["strings"]["disconnect"];
                // holdingSuiteToolStripMenuItem.IsEnabled = true;

                menuItemToolBit.IsEnabled = true;
                menuItemToolWord.IsEnabled = true;
                menuItemToolByte.IsEnabled = true;
                salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = false;
                caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = false;
                gestisciDatabaseToolStripMenuItem.IsEnabled = false;


                try
                {
                    // ---------------------------------------------------------------------------------
                    // ----------------------Apertura comunicazione seriale-----------------------------
                    // ---------------------------------------------------------------------------------

                    // Create a new SerialPort object with default settings.
                    serialPort = new SerialPort();
                    serialPort.PortName = comboBoxSerialPort.SelectedItem.ToString();

                    // debug
                    //Console.WriteLine(comboBoxSerialSpeed.SelectedValue.ToString());
                    //Console.WriteLine(comboBoxSerialSpeed.SelectedItem.ToString());

                    serialPort.BaudRate = int.Parse(comboBoxSerialSpeed.SelectedValue.ToString().Split(' ')[1]);

                    // DEBUG
                    //Console.WriteLine("comboBoxSerialParity.SelectedIndex:" + comboBoxSerialParity.SelectedIndex.ToString());

                    switch (comboBoxSerialParity.SelectedIndex)
                    {
                        case 0:
                            serialPort.Parity = Parity.None;
                            break;
                        case 1:
                            serialPort.Parity = Parity.Even;
                            break;
                        case 2:
                            serialPort.Parity = Parity.Odd;
                            break;
                        default:
                            serialPort.Parity = Parity.None;
                            break;
                    }

                    serialPort.DataBits = 8;

                    // DEBUG
                    Console.WriteLine("comboBoxSerialStop.SelectedIndex:" + comboBoxSerialStop.SelectedIndex.ToString());

                    switch (comboBoxSerialStop.SelectedIndex)
                    {
                        case 0:
                            serialPort.StopBits = StopBits.One;
                            break;
                        case 1:
                            serialPort.StopBits = StopBits.OnePointFive;
                            break;
                        case 2:
                            serialPort.StopBits = StopBits.Two;
                            break;
                        default:
                            serialPort.StopBits = StopBits.One;
                            break;
                    }

                    serialPort.Handshake = Handshake.None;

                    // Timeout porta
                    serialPort.ReadTimeout = 50;
                    serialPort.WriteTimeout = 50;

                    ModBus = new ModBus_Chicco(
                        serialPort,
                        textBoxTcpClientIpAddress.Text,
                        textBoxTcpClientPort.Text,
                        ModBus_Def.TYPE_RTU,
                        pictureBoxTx,
                        pictureBoxRx,
                        "", "", "", -1);
                    ModBus.open();

                    serialPort.Open();
                    richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["connectedTo"] + " " + comboBoxSerialPort.SelectedItem.ToString());

                    radioButtonModeSerial.IsEnabled = false;
                    radioButtonModeTcp.IsEnabled = false;


                    comboBoxSerialPort.IsEnabled = false;
                    comboBoxSerialSpeed.IsEnabled = false;
                    comboBoxSerialParity.IsEnabled = false;
                    comboBoxSerialStop.IsEnabled = false;
                    languageToolStripMenu.IsEnabled = false;

                    if (pathToConfiguration != defaultPathToConfiguration)
                        this.Title = title + " " + version + " - File: " + pathToConfiguration + " - Port: " + comboBoxSerialPort.SelectedItem.ToString();
                    else
                        this.Title = title + " " + version + " - Port: " + comboBoxSerialPort.SelectedItem.ToString();
                }
                catch (Exception err)
                {
                    pictureBoxSerial.Background = Brushes.LightGray;
                    pictureBoxRunningAs.Background = Brushes.LightGray;


                    textBlockSerialActive.Text = lang.languageTemplate["strings"]["connect"];
                    //holdingSuiteToolStripMenuItem.IsEnabled = false;

                    menuItemToolBit.IsEnabled = false;
                    menuItemToolWord.IsEnabled = false;
                    menuItemToolByte.IsEnabled = false;
                    salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = true;
                    caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = true;
                    gestisciDatabaseToolStripMenuItem.IsEnabled = true;

                    comboBoxSerialPort.IsEnabled = true;
                    comboBoxSerialSpeed.IsEnabled = true;
                    comboBoxSerialParity.IsEnabled = true;
                    comboBoxSerialStop.IsEnabled = true;
                    languageToolStripMenu.IsEnabled = true;

                    Console.WriteLine("Errore apertura porta seriale");
                    Console.WriteLine(err);

                    richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["failedToConnect"]);

                    if (pathToConfiguration != defaultPathToConfiguration)
                        this.Title = title + " " + version + " - File: " + pathToConfiguration;
                    else
                        this.Title = title + " " + version;
                }
            }
            else
            {
                // Disattivazione comunicazione seriale
                pictureBoxSerial.Background = Brushes.LightGray;
                pictureBoxRunningAs.Background = Brushes.LightGray;


                textBlockSerialActive.Text = lang.languageTemplate["strings"]["connect"];

                radioButtonModeSerial.IsEnabled = true;
                radioButtonModeTcp.IsEnabled = true;

                menuItemToolBit.IsEnabled = false;
                menuItemToolWord.IsEnabled = false;
                menuItemToolByte.IsEnabled = false;
                salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = true;
                caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = true;
                gestisciDatabaseToolStripMenuItem.IsEnabled = true;

                comboBoxSerialPort.IsEnabled = true;
                comboBoxSerialSpeed.IsEnabled = true;
                comboBoxSerialParity.IsEnabled = true;
                comboBoxSerialStop.IsEnabled = true;
                languageToolStripMenu.IsEnabled = true;

                // Fermo eventuali loop
                disableAllLoops();

                // ---------------------------------------------------------------------------------
                // ----------------------Chiusura comunicazione seriale-----------------------------
                // ---------------------------------------------------------------------------------
                serialPort.Close();
                richTextBoxAppend(richTextBoxStatus, "Port closed");
            }

            changeEnableButtonsConnect(pictureBoxRunningAs.Background == Brushes.Lime);
        }

        private void buttonUpdateSerialList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string[] SerialPortList = System.IO.Ports.SerialPort.GetPortNames();

                //comboBoxSerialPort.Items.Add("Seleziona porta seriale ...");
                comboBoxSerialPort.Items.Clear();

                foreach (String port in SerialPortList)
                {
                    comboBoxSerialPort.Items.Add(port);
                }

                comboBoxSerialPort.SelectedIndex = 0;
            }
            catch
            {
                Console.WriteLine("Nessuna porta seriale disponibile");
            }
        }

        // Visualizza console programma da menu tendina
        private void apriConsoleToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            apriConsole();
        }

        // Nasconde console programma da menu tendina
        private void chiudiConsoleToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            chiudiConsole();
        }

        public void chiudiConsole()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            statoConsole = false;
        }

        public void apriConsole()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_SHOW);

            statoConsole = true;
        }

        // Disabilita il pulsante di chiusura della console
        public void disableConsoleExitButton()
        {
            IntPtr handle = GetConsoleWindow();
            IntPtr exitButton = GetSystemMenu(handle, false);
            if (exitButton != null) DeleteMenu(exitButton, SC_CLOSE, MF_BYCOMMAND);
        }

        // ----------------------------------------------------------------------------------
        // ---------------------------SALVATAGGIO CONFIGURAZIONE-----------------------------
        // ----------------------------------------------------------------------------------

        /*public void SaveConfiguration_v1(bool alert)   //Se alert true visualizza un messaggio di info salvataggio avvenuto
        {
            //DEBUG
            //MessageBox.Show("Salvataggio configurazione");

            try
            {
                // Caricamento variabili
                var config = new SAVE();

                config.modbusAddress = textBoxModbusAddress.Text;
                config.usingSerial = (bool)radioButtonModeSerial.IsChecked;

                //config.serialMaster = radioButtonSerialMaster.IsChecked;
                //config.serialRTU = radioButtonModeRTU.IsChecked;

                // Serial port
                config.serialPort = comboBoxSerialPort.SelectedIndex;
                config.serialSpeed = comboBoxSerialSpeed.SelectedIndex;
                config.serialParity = comboBoxSerialParity.SelectedIndex;
                config.serialStop = comboBoxSerialStop.SelectedIndex;

                // TCP
                config.tcpClientIpAddress = textBoxTcpClientIpAddress.Text;
                config.tcpClientPort = textBoxTcpClientPort.Text;
                //config.tcpServerIpAddress = textBoxTcpServerIpAddress.Text;
                //config.tcpServerPort = textBoxTcpServerPort.Text;

                // GRAFICA
                // TabPage1 (Coils)
                config.textBoxCoilsAddress01 = textBoxCoilsAddress01.Text;
                config.textBoxCoilNumber = textBoxCoilNumber.Text;
                config.textBoxCoilsRange_A = textBoxCoilsRange_A.Text;
                config.textBoxCoilsRange_B = textBoxCoilsRange_B.Text;
                config.textBoxCoilsAddress05 = textBoxCoilsAddress05.Text;
                config.textBoxCoilsValue05 = textBoxCoilsValue05.Text;
                config.textBoxCoilsAddress15_A = textBoxCoilsAddress15_A.Text;
                config.textBoxCoilsAddress15_B = textBoxCoilsAddress15_B.Text;
                config.textBoxCoilsValue15 = textBoxCoilsValue15.Text;
                config.textBoxGoToCoilAddress = textBoxGoToCoilAddress.Text;

                // TabPage2 (inputs)
                config.textBoxInputAddress02 = textBoxInputAddress02.Text;
                config.textBoxInputNumber = textBoxInputNumber.Text;
                config.textBoxInputRange_A = textBoxInputRange_A.Text;
                config.textBoxInputRange_B = textBoxInputRange_B.Text;
                config.textBoxGoToInputAddress = textBoxGoToInputAddress.Text;

                // TabPage3 (input registers)
                config.textBoxInputRegisterAddress04 = textBoxInputRegisterAddress04.Text;
                config.textBoxInputRegisterNumber = textBoxInputRegisterNumber.Text;
                config.textBoxInputRegisterRange_A = textBoxInputRegisterRange_A.Text;
                config.textBoxInputRegisterRange_B = textBoxInputRegisterRange_B.Text;
                config.textBoxGoToInputRegisterAddress = textBoxGoToInputRegisterAddress.Text;

                // TabPage4 (holding registers)
                config.textBoxHoldingAddress03 = textBoxHoldingAddress03.Text;
                config.textBoxHoldingRegisterNumber = textBoxHoldingRegisterNumber.Text;
                config.textBoxHoldingRange_A = textBoxHoldingRange_A.Text;
                config.textBoxHoldingRange_B = textBoxHoldingRange_B.Text;
                config.textBoxHoldingAddress06 = textBoxHoldingAddress06.Text;
                config.textBoxHoldingValue06 = textBoxHoldingValue06.Text;
                config.textBoxHoldingAddress16_A = textBoxHoldingAddress16_A.Text;
                config.textBoxHoldingAddress16_B = textBoxHoldingAddress16_B.Text;
                config.textBoxHoldingValue16 = textBoxHoldingValue16.Text;
                config.textBoxGoToHoldingAddress = textBoxGoToHoldingAddress.Text;

                config.statoConsole = statoConsole;

                // Funzioni aggiunte in seguito
                config.comboBoxCoilsAddress01_ = comboBoxCoilsAddress01.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsRange_A_ = comboBoxCoilsRange_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsRange_B_ = comboBoxCoilsRange_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsAddress05_ = comboBoxCoilsAddress05.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsValue05_ = "DEC"; // comboBoxCoilsValue05.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsAddress15_A_ = comboBoxCoilsAddress15_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsAddress15_B_ = comboBoxCoilsAddress15_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxCoilsValue15_ = "DEC"; // comboBoxCoilsValue15.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputAddress02_ = comboBoxInputAddress02.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRange_A_ = comboBoxInputRange_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRange_B_ = comboBoxInputRange_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegisterAddress04_ = comboBoxInputRegisterAddress04.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegisterRange_A_ = comboBoxInputRegisterRange_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegisterRange_B_ = comboBoxInputRegisterRange_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingAddress03_ = comboBoxHoldingAddress03.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingRange_A_ = comboBoxHoldingRange_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingRange_B_ = comboBoxHoldingRange_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingAddress06_ = comboBoxHoldingAddress06.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingValue06_ = comboBoxHoldingValue06.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingAddress16_A_ = comboBoxHoldingAddress16_A.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingAddress16_B_ = comboBoxHoldingAddress16_B.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingValue16_ = comboBoxHoldingValue16.SelectedValue.ToString().Split(' ')[1];

                //comboBox visualizzazione tabelle
                config.comboBoxCoilsRegistri_ = comboBoxCoilsRegistri.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegistri_ = comboBoxInputRegistri.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegRegistri_ = comboBoxInputRegRegistri.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegValori_ = comboBoxInputRegValori.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingRegistri_ = comboBoxHoldingRegistri.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingValori_ = comboBoxHoldingValori.SelectedValue.ToString().Split(' ')[1];

                config.comboBoxCoilsOffset_ = comboBoxCoilsOffset.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputOffset_ = comboBoxInputOffset.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxInputRegOffset_ = comboBoxInputRegOffset.SelectedValue.ToString().Split(' ')[1];
                config.comboBoxHoldingOffset_ = comboBoxHoldingOffset.SelectedValue.ToString().Split(' ')[1];

                // Funzioni doppie per force coil e preset holding aggiunte dopo
                config.comboBoxCoilsAddress05_b_ = comboBoxCoilsAddress05_b.SelectedValue.ToString().Split(' ')[1];
                config.textBoxCoilsAddress05_b_ = textBoxCoilsAddress05_b.Text;
                config.comboBoxCoilsValue05_b_ = "DEC"; //comboBoxCoilsValue05_b.SelectedValue.ToString().Split(' ')[1];
                config.textBoxCoilsValue05_b_ = textBoxCoilsValue05_b.Text;

                config.comboBoxHoldingAddress06_b_ = comboBoxHoldingAddress06_b.SelectedValue.ToString().Split(' ')[1];
                config.textBoxHoldingAddress06_b_ = textBoxHoldingAddress06_b.Text;
                config.comboBoxHoldingValue06_b_ = comboBoxHoldingValue06_b.SelectedValue.ToString().Split(' ')[1];
                config.textBoxHoldingValue06_b_ = textBoxHoldingValue06_b.Text;

                config.textBoxCoilsOffset_ = textBoxCoilsOffset.Text;
                config.textBoxInputOffset_ = textBoxInputOffset.Text;
                config.textBoxInputRegOffset_ = textBoxInputRegOffset.Text;
                config.textBoxHoldingOffset_ = textBoxHoldingOffset.Text;

                config.checkBoxUseOffsetInTables_ = (bool)checkBoxUseOffsetInTables.IsChecked;
                config.checkBoxUseOffsetInTextBox_ = (bool)checkBoxUseOffsetInTextBox.IsChecked;
                config.checkBoxFollowModbusProtocol_ = (bool)checkBoxFollowModbusProtocol.IsChecked;
                //config.checkBoxSavePackets_ = (bool)checkBoxSavePackets.IsChecked;
                config.checkBoxCloseConsolAfterBoot_ = (bool)checkBoxCloseConsolAfterBoot.IsChecked;
                config.checkBoxCellColorMode_ = (bool)checkBoxCellColorMode.IsChecked;
                config.checkBoxViewTableWithoutOffset_ = (bool)checkBoxViewTableWithoutOffset.IsChecked;

                //config.textBoxSaveLogPath_ = textBoxSaveLogPath.Text;


                config.comboBoxDiagnosticFunction_ = comboBoxDiagnosticFunction.SelectedValue.ToString().Split(' ')[1];

                config.textBoxDiagnosticFunctionManual_ = textBoxDiagnosticFunctionManual.Text;

                config.colorDefaultReadCell_ = colorDefaultReadCell.ToString();
                config.colorDefaultWriteCell_ = colorDefaultWriteCell.ToString();
                config.colorErrorCell_ = colorErrorCell.ToString();

                config.TextBoxPollingInterval_ = TextBoxPollingInterval.Text;
                config.TextBoxPollingInterval_ = TextBoxPollingInterval.Text;

                config.CheckBoxSendValuesOnEditCoillsTable_ = (bool)CheckBoxSendValuesOnEditCoillsTable.IsChecked;
                config.CheckBoxSendValuesOnEditHoldingTable_ = (bool)CheckBoxSendValuesOnEditHoldingTable.IsChecked;

                config.language = language;

                config.textBoxReadTimeout = textBoxReadTimeout.Text;

                var jss = new JavaScriptSerializer();
                jss.MaxJsonLength = this.MaxJsonLength;
                jss.RecursionLimit = 1000;
                string file_content = jss.Serialize(config);

                File.WriteAllText("Json/" + pathToConfiguration + "/CONFIGURAZIONE.json", file_content);

                if (alert)
                {
                    MessageBox.Show(lang.languageTemplate["strings"]["infoSaveConfig"], "Info");
                }

                Console.WriteLine("Salvata configurazione");
            }
            catch(Exception err)
            {
                Console.WriteLine("Errore salvataggio configurazione");
                Console.WriteLine(err);
            }
        }*/

        // ----------------------------------------------------------------------------------
        // ---------------------------CARICAMENTO CONFIGURAZIONE-----------------------------
        // ----------------------------------------------------------------------------------

        public void LoadConfiguration_v1()
        {
            try
            {
                string file_content = File.ReadAllText("Json/" + pathToConfiguration + "/CONFIGURAZIONE.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                SAVE config = jss.Deserialize<SAVE>(file_content);

                textBoxModbusAddress.Text = config.modbusAddress;

                // Scheda configurazione seriale
                radioButtonModeSerial.IsChecked = config.usingSerial;
                radioButtonModeTcp.IsChecked = !config.usingSerial;

                //radioButtonSerialMaster.IsChecked = config.serialMaster;
                //radioButtonSerialSlave.IsChecked = !config.serialMaster;
                //radioButtonModeRTU.IsChecked = config.serialRTU;
                //radioButtonModeASCII.IsChecked = !config.serialRTU;

                comboBoxSerialSpeed.SelectedIndex = config.serialSpeed;
                comboBoxSerialParity.SelectedIndex = config.serialParity;
                comboBoxSerialStop.SelectedIndex = config.serialStop;

                // Scheda configurazione TCP
                //radioButtonTcpMaster.IsChecked = config.serialMaster;
                //radioButtonTcpSlave.IsChecked = !config.serialMaster;

                textBoxTcpClientIpAddress.Text = config.tcpClientIpAddress;
                textBoxTcpClientPort.Text = config.tcpClientPort;
                //textBoxTcpServerIpAddress.Text = config.tcpServerIpAddress;
                //textBoxTcpServerPort.Text = config.tcpServerPort;

                comboBoxSerialPort.SelectedIndex = config.serialPort;

                // GRAFICA
                // TabPage1 (Coils)
                textBoxCoilsAddress01.Text = config.textBoxCoilsAddress01;
                textBoxCoilNumber.Text = config.textBoxCoilNumber;
                textBoxCoilsRange_A.Text = config.textBoxCoilsRange_A;
                textBoxCoilsRange_B.Text = config.textBoxCoilsRange_B;
                textBoxCoilsAddress05.Text = config.textBoxCoilsAddress05;
                textBoxCoilsValue05.Text = config.textBoxCoilsValue05;
                textBoxCoilsAddress15_A.Text = config.textBoxCoilsAddress15_A;
                textBoxCoilsAddress15_B.Text = config.textBoxCoilsAddress15_B;
                textBoxCoilsValue15.Text = config.textBoxCoilsValue15;
                textBoxGoToCoilAddress.Text = config.textBoxGoToCoilAddress;

                // TabPage2 (inputs)
                textBoxInputAddress02.Text = config.textBoxInputAddress02;
                textBoxInputNumber.Text = config.textBoxInputNumber;
                textBoxInputRange_A.Text = config.textBoxInputRange_A;
                textBoxInputRange_B.Text = config.textBoxInputRange_B;
                textBoxGoToInputAddress.Text = config.textBoxGoToInputAddress;

                // TabPage3 (input registers)
                textBoxInputRegisterAddress04.Text = config.textBoxInputRegisterAddress04;
                textBoxInputRegisterNumber.Text = config.textBoxInputRegisterNumber;
                textBoxInputRegisterRange_A.Text = config.textBoxInputRegisterRange_A;
                textBoxInputRegisterRange_B.Text = config.textBoxInputRegisterRange_B;
                textBoxGoToInputRegisterAddress.Text = config.textBoxGoToInputRegisterAddress;

                // TabPage4 (holding registers)
                textBoxHoldingAddress03.Text = config.textBoxHoldingAddress03;
                textBoxHoldingRegisterNumber.Text = config.textBoxHoldingRegisterNumber;
                textBoxHoldingRange_A.Text = config.textBoxHoldingRange_A;
                textBoxHoldingRange_B.Text = config.textBoxHoldingRange_B;
                textBoxHoldingAddress06.Text = config.textBoxHoldingAddress06;
                textBoxHoldingValue06.Text = config.textBoxHoldingValue06;
                textBoxHoldingAddress16_A.Text = config.textBoxHoldingAddress16_A;
                textBoxHoldingAddress16_B.Text = config.textBoxHoldingAddress16_B;
                textBoxHoldingValue16.Text = config.textBoxHoldingValue16;
                textBoxGoToHoldingAddress.Text = config.textBoxGoToHoldingAddress;

                statoConsole = config.statoConsole;

                // Funzioni aggiunte in seguito
                comboBoxCoilsAddress01.SelectedIndex = config.comboBoxCoilsAddress01_ == "HEX" ? 1 : 0;
                comboBoxCoilsRange_A.SelectedIndex = config.comboBoxCoilsRange_A_ == "HEX" ? 1 : 0;
                //comboBoxCoilsRange_B.SelectedIndex = config.comboBoxCoilsRange_B_ == "HEX" ? 1 : 0;
                comboBoxCoilsAddress05.SelectedIndex = config.comboBoxCoilsAddress05_ == "HEX" ? 1 : 0;
                //comboBoxCoilsValue05.SelectedIndex = config.comboBoxCoilsValue05_ == "HEX" ? 1 : 0;
                comboBoxCoilsAddress15_A.SelectedIndex = config.comboBoxCoilsAddress15_A_ == "HEX" ? 1 : 0;
                comboBoxCoilsAddress15_B.SelectedIndex = config.comboBoxCoilsAddress15_B_ == "HEX" ? 1 : 0;
                //comboBoxCoilsValue15.SelectedIndex = config.comboBoxCoilsValue15_ == "HEX" ? 1 : 0;
                comboBoxInputAddress02.SelectedIndex = config.comboBoxInputAddress02_ == "HEX" ? 1 : 0;
                comboBoxInputRange_A.SelectedIndex = config.comboBoxInputRange_A_ == "HEX" ? 1 : 0;
                //comboBoxInputRange_B.SelectedIndex = config.comboBoxInputRange_B_ == "HEX" ? 1 : 0;
                comboBoxInputRegisterAddress04.SelectedIndex = config.comboBoxInputRegisterAddress04_ == "HEX" ? 1 : 0;
                comboBoxInputRegisterRange_A.SelectedIndex = config.comboBoxInputRegisterRange_A_ == "HEX" ? 1 : 0;
                //comboBoxInputRegisterRange_B.SelectedIndex = config.comboBoxInputRegisterRange_B_ == "HEX" ? 1 : 0;
                comboBoxHoldingAddress03.SelectedIndex = config.comboBoxHoldingAddress03_ == "HEX" ? 1 : 0;
                comboBoxHoldingRange_A.SelectedIndex = config.comboBoxHoldingRange_A_ == "HEX" ? 1 : 0;
                //comboBoxHoldingRange_B.SelectedIndex = config.comboBoxHoldingRange_B_ == "HEX" ? 1 : 0;
                comboBoxHoldingAddress06.SelectedIndex = config.comboBoxHoldingAddress06_ == "HEX" ? 1 : 0;
                comboBoxHoldingValue06.SelectedIndex = config.comboBoxHoldingValue06_ == "HEX" ? 1 : 0;
                comboBoxHoldingAddress16_A.SelectedIndex = config.comboBoxHoldingAddress16_A_ == "HEX" ? 1 : 0;
                comboBoxHoldingAddress16_B.SelectedIndex = config.comboBoxHoldingAddress16_B_ == "HEX" ? 1 : 0;
                comboBoxHoldingValue16.SelectedIndex = config.comboBoxHoldingValue16_ == "HEX" ? 1 : 0;

                comboBoxCoilsOffset.SelectedIndex = config.comboBoxCoilsOffset_ == "HEX" ? 1 : 0;
                comboBoxInputOffset.SelectedIndex = config.comboBoxInputOffset_ == "HEX" ? 1 : 0;
                comboBoxInputRegOffset.SelectedIndex = config.comboBoxInputRegOffset_ == "HEX" ? 1 : 0;
                comboBoxHoldingOffset.SelectedIndex = config.comboBoxHoldingOffset_ == "HEX" ? 1 : 0;

                //comboBox visualizzazione tabelle
                comboBoxCoilsRegistri.SelectedIndex = config.comboBoxCoilsRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegistri.SelectedIndex = config.comboBoxInputRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegRegistri.SelectedIndex = config.comboBoxInputRegRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegValori.SelectedIndex = config.comboBoxInputRegValori_ == "HEX" ? 1 : 0;
                comboBoxHoldingRegistri.SelectedIndex = config.comboBoxHoldingRegistri_ == "HEX" ? 1 : 0;
                comboBoxHoldingValori.SelectedIndex = config.comboBoxHoldingValori_ == "HEX" ? 1 : 0;

                // Da eliminare
                //textBoxCoilsOffset.Text = config.textBoxCoilsOffset_;
                //textBoxInputOffset.Text = config.textBoxInputOffset_;
                //textBoxInputRegOffset.Text = config.textBoxInputRegOffset_;
                //textBoxHoldingOffset.Text = config.textBoxHoldingOffset_;

                // Funzioni doppie per force coil e preset holding aggiunte dopo
                comboBoxCoilsAddress05_b.SelectedIndex = config.comboBoxCoilsAddress05_b_ == "HEX" ? 1 : 0;
                textBoxCoilsAddress05_b.Text = config.textBoxCoilsAddress05_b_;
                //comboBoxCoilsValue05_b.SelectedIndex = config.comboBoxCoilsValue05_b_ == "HEX" ? 1 : 0;
                textBoxCoilsValue05_b.Text = config.textBoxCoilsValue05_b_;

                comboBoxHoldingAddress06_b.SelectedIndex = config.comboBoxHoldingAddress06_b_ == "HEX" ? 1 : 0;
                textBoxHoldingAddress06_b.Text = config.textBoxHoldingAddress06_b_;
                comboBoxHoldingValue06_b.SelectedIndex = config.comboBoxHoldingValue06_b_ == "HEX" ? 1 : 0;
                textBoxHoldingValue06_b.Text = config.textBoxHoldingValue06_b_;

                textBoxCoilsOffset.Text = config.textBoxCoilsOffset_;
                textBoxInputOffset.Text = config.textBoxInputOffset_;
                textBoxInputRegOffset.Text = config.textBoxInputRegOffset_;
                textBoxHoldingOffset.Text = config.textBoxHoldingOffset_;

                checkBoxUseOffsetInTextBox.IsChecked = config.checkBoxUseOffsetInTextBox_;
                checkBoxCloseConsolAfterBoot.IsChecked = config.checkBoxCloseConsolAfterBoot_;
                checkBoxCellColorMode.IsChecked = config.checkBoxCellColorMode_;
                checkBoxViewTableWithoutOffset.IsChecked = config.checkBoxViewTableWithoutOffset_;

                comboBoxDiagnosticFunction.SelectedIndex = 0;

                textBoxDiagnosticFunctionManual.Text = config.textBoxDiagnosticFunctionManual_;

                BrushConverter bc = new BrushConverter();

                colorDefaultReadCell = (SolidColorBrush)bc.ConvertFromString(config.colorDefaultReadCell_);
                colorDefaultWriteCell = (SolidColorBrush)bc.ConvertFromString(config.colorDefaultWriteCell_);
                colorErrorCell = (SolidColorBrush)bc.ConvertFromString(config.colorErrorCell_);

                labelColorCellRead.Background = colorDefaultReadCell;
                labelColorCellWrote.Background = colorDefaultWriteCell;
                labelColorCellError.Background = colorErrorCell;

                colorDefaultReadCellStr = colorDefaultReadCell.ToString();
                colorDefaultWriteCellStr = colorDefaultWriteCell.ToString();
                colorErrorCellStr = colorErrorCell.ToString();

                if (!(config.TextBoxPollingInterval_ is null))
                {
                    TextBoxPollingInterval.Text = config.TextBoxPollingInterval_;
                }

                if (config.CheckBoxSendValuesOnEditCoillsTable_ != null)
                {
                    CheckBoxSendValuesOnEditCoillsTable.IsChecked = config.CheckBoxSendValuesOnEditCoillsTable_;
                }

                if (config.CheckBoxSendValuesOnEditHoldingTable_ != null)
                {
                    CheckBoxSendValuesOnEditHoldingTable.IsChecked = config.CheckBoxSendValuesOnEditHoldingTable_;
                }


                if (statoConsole)
                {
                    apriConsole();
                }

                // Scelgo quale richTextBox di status abilitare
                //richTextBoxStatus.IsEnabled = (bool)radioButtonModeSerial.IsChecked;
                //richTextBoxStatus.IsEnabled = !(bool)radioButtonModeSerial.IsChecked;

                if (config.language != null)
                {
                    language = config.language;

                    foreach (MenuItem tmp in languageToolStripMenu.Items)
                    {
                        if (tmp.Header.ToString().IndexOf(language) != -1)
                        {
                            tmp.IsChecked = true;
                        }
                    }
                }

                if (config.textBoxReadTimeout != null)
                {
                    textBoxReadTimeout.Text = config.textBoxReadTimeout;
                }

                textBoxCurrentLanguage.Text = language;
                //lang.loadLanguageTemplate(language);

                Console.WriteLine("Caricata configurazione precedente\n");
            }
            catch
            {
                Console.WriteLine("Error loading configuration\n");
            }

            // Al termine del caricamento della configurazione carico il template
            try
            {
                string file_content = File.ReadAllText(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + pathToConfiguration + "\\Template.json");
                JavaScriptSerializer jss = new JavaScriptSerializer();
                jss.MaxJsonLength = this.MaxJsonLength;
                TEMPLATE template = jss.Deserialize<TEMPLATE>(file_content);

                list_template_coilsTable.Clear();
                list_template_inputsTable.Clear();
                list_template_inputRegistersTable.Clear();
                list_template_holdingRegistersTable.Clear();

                UInt16 tmp = 0;

                // Coils
                int template_coilsOffset = int.Parse(template.textBoxCoilsOffset_, template.comboBoxCoilsOffset_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                // Inputs
                int template_inputsOffset = int.Parse(template.textBoxInputOffset_, template.comboBoxInputOffset_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                // Input registers
                int template_inputRegistersOffset = int.Parse(template.textBoxInputRegOffset_, template.comboBoxInputRegOffset_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                // Holding registers
                int template_HoldingOffset = int.Parse(template.textBoxHoldingOffset_, template.comboBoxHoldingOffset_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                // Tabella coils
                for (int i = 0; i < template.dataGridViewCoils.Count(); i++)
                {
                    if (UInt16.TryParse(template.dataGridViewCoils[i].Register, template.comboBoxCoilsRegistri_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out tmp))
                    {
                        template.dataGridViewCoils[i].RegisterUInt = (UInt16)(tmp + template_coilsOffset);
                        template.dataGridViewCoils[i].Register = tmp.ToString();
                        list_template_coilsTable.Add(template.dataGridViewCoils[i]);
                    }
                }

                // Tabella inputs
                for (int i = 0; i < template.dataGridViewInput.Count(); i++)
                {
                    if (UInt16.TryParse(template.dataGridViewInput[i].Register, template.comboBoxInputRegistri_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out tmp))
                    {
                        template.dataGridViewInput[i].RegisterUInt = (UInt16)(tmp + template_inputsOffset);
                        template.dataGridViewInput[i].Register = tmp.ToString();
                        list_template_inputsTable.Add(template.dataGridViewInput[i]);
                    }
                }

                // Tabella input registers
                for (int i = 0; i < template.dataGridViewInputRegister.Count(); i++)
                {
                    if (UInt16.TryParse(template.dataGridViewInputRegister[i].Register, template.comboBoxInputRegRegistri_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out tmp))
                    {
                        template.dataGridViewInputRegister[i].RegisterUInt = (UInt16)(tmp + template_inputRegistersOffset);
                        template.dataGridViewInputRegister[i].Register = template.dataGridViewInputRegister[i].RegisterUInt.ToString();
                        list_template_inputRegistersTable.Add(template.dataGridViewInputRegister[i]);
                    }
                }

                // Tabella holdings
                for (int i = 0; i < template.dataGridViewHolding.Count(); i++)
                {
                    if (UInt16.TryParse(template.dataGridViewHolding[i].Register, template.comboBoxHoldingRegistri_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out tmp))
                    {
                        template.dataGridViewHolding[i].RegisterUInt = (UInt16)(tmp + template_HoldingOffset);
                        template.dataGridViewHolding[i].Register = template.dataGridViewHolding[i].RegisterUInt.ToString();
                        list_template_holdingRegistersTable.Add(template.dataGridViewHolding[i]);
                    }
                }

                // Tabella groups
                comboBoxHoldingGroup.Items.Clear();
                comboBoxInputRegisterGroup.Items.Clear();
                comboBoxInputGroup.Items.Clear();
                comboBoxCoilsGroup.Items.Clear();

                if (template.Groups != null)
                {
                    comboBoxHoldingGroup.IsEnabled = template.Groups.Count() > 0;
                    comboBoxInputRegisterGroup.IsEnabled = template.Groups.Count() > 0;
                    comboBoxInputGroup.IsEnabled = template.Groups.Count() > 0;
                    comboBoxCoilsGroup.IsEnabled = template.Groups.Count() > 0;

                    if (template.Groups.Count() > 0)
                    {
                        try
                        {
                            foreach (Group_Item gr in template.Groups.OrderBy(x => int.Parse(x.Group)))
                            {
                                KeyValuePair<Group_Item, String> kp = new KeyValuePair<Group_Item, String>(gr, gr.Group + " - " + gr.Label);

                                if (template.dataGridViewHolding.First<ModBus_Item>(x => x.Group.IndexOf(kp.Key.Group) != -1) != null)
                                    comboBoxHoldingGroup.Items.Add(kp);

                                if (template.dataGridViewInputRegister.First<ModBus_Item>(x => x.Group.IndexOf(kp.Key.Group) != -1) != null)
                                    comboBoxInputRegisterGroup.Items.Add(kp);

                                if (template.dataGridViewInput.First<ModBus_Item>(x => x.Group.IndexOf(kp.Key.Group) != -1) != null)
                                    comboBoxInputGroup.Items.Add(kp);

                                if (template.dataGridViewCoils.First<ModBus_Item>(x => x.Group.IndexOf(kp.Key.Group) != -1) != null)
                                    comboBoxCoilsGroup.Items.Add(kp);
                            }
                        }
                        catch (Exception err)
                        {
                            Console.WriteLine("Error loading group items\n");
                            Console.WriteLine(err);
                        }
                    }
                }
                else
                {
                    comboBoxHoldingGroup.IsEnabled = false;
                    comboBoxInputRegisterGroup.IsEnabled = false;
                    comboBoxInputGroup.IsEnabled = false;
                    comboBoxCoilsGroup.IsEnabled = false;
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("Error loading configuration\n");
                Console.WriteLine(err);
            }

        }

        private void salvaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            SaveConfiguration_v2(true);
        }

        private void genera_tabelle_registri()
        {
            buttonReadCoils01.Dispatcher.Invoke((Action)delegate
            {
                // Disattivazione pulsanti fino al termine della generaizone delle tabelle
                /*buttonReadCoils01.IsEnabled = false;
                buttonReadCoilsRange.IsEnabled = false;
                buttonWriteCoils05.IsEnabled = false;
                buttonWriteCoils15.IsEnabled = false;*/
                buttonGoToCoilAddress.IsEnabled = false;

                /*buttonReadInput02.IsEnabled = false;
                buttonReadInputRange.IsEnabled = false;*/
                buttonGoToInputAddress.IsEnabled = false;

                /*buttonReadInputRegister04.IsEnabled = false;
                buttonReadInputRegisterRange.IsEnabled = false;*/
                buttonGoToInputRegisterAddress.IsEnabled = false;

                /*buttonReadHolding03.IsEnabled = false;
                buttonReadHoldingRange.IsEnabled = false;
                buttonWriteHolding06.IsEnabled = false;
                buttonWriteHolding16.IsEnabled = false;*/
                buttonGoToHoldingAddress.IsEnabled = false;

                list_coilsTable.Clear();
                list_inputsTable.Clear();
                list_inputRegistersTable.Clear();
                list_holdingRegistersTable.Clear();
            });

            ModBus_Item row = new ModBus_Item();

            buttonReadCoils01.Dispatcher.Invoke((Action)delegate
            {
                // Attivazione pulsanti al termine della generaizone delle tabelle
                /*buttonReadCoils01.IsEnabled = true;
                buttonReadCoilsRange.IsEnabled = true;
                buttonWriteCoils05.IsEnabled = true;
                buttonWriteCoils15.IsEnabled = true;*/
                buttonGoToCoilAddress.IsEnabled = true;

                /*buttonReadInput02.IsEnabled = true;
                buttonReadInputRange.IsEnabled = true;*/
                buttonGoToInputAddress.IsEnabled = true;

                /*buttonReadInputRegister04.IsEnabled = true;
                buttonReadInputRegisterRange.IsEnabled = true;*/
                buttonGoToInputRegisterAddress.IsEnabled = true;

                /*buttonReadHolding03.IsEnabled = true;
                buttonReadHoldingRange.IsEnabled = true;
                buttonWriteHolding06.IsEnabled = true;
                buttonWriteHolding16.IsEnabled = true;*/
                buttonGoToHoldingAddress.IsEnabled = true;


                // Dopo 2.5s secondi chiudo la console

                if ((bool)checkBoxCloseConsolAfterBoot.IsChecked)
                {
                    var handle = GetConsoleWindow();
                    ShowWindow(handle, SW_HIDE);
                }
            });
            Thread.Sleep(500);
            Thread.CurrentThread.Abort();
        }

        public void buttonTcpActive_Click(object sender, RoutedEventArgs e)
        {
            buttonTcpActive.IsEnabled = false;

            Thread t = new Thread(new ThreadStart(openCloseTcpConnection));
            t.Start();
        }

        private void openCloseTcpConnection()
        {
            String ip_address = "";
            String port = "";
            bool check = false;
            int TCPMode = 0;

            this.Dispatcher.Invoke((Action)delegate
            {
                ip_address = textBoxTcpClientIpAddress.Text;
                port = textBoxTcpClientPort.Text;
                check = pictureBoxTcp.Background == Brushes.LightGray;
                if (check)
                    richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["connectingTo"] + " " + ip_address + ":" + port);

                if((bool)checkBoxModbusSecure.IsChecked)
                    TCPMode = ModBus_Def.TYPE_TCP_SECURE;
                else
                    TCPMode = comboBoxTcpConnectionMode.SelectedIndex == 0 ? ModBus_Def.TYPE_TCP_SOCK : ModBus_Def.TYPE_TCP_REOPEN;
            });

            if (check)
            {
                try
                {
                    // Open, test the connection and close
                    // TcpClient client = new TcpClient(ip_address, int.Parse(port));
                    // client.Close();

                    // If TCPMode check if certificate file exists
                    if(TCPMode == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (!File.Exists(textBoxClientCertificatePath.Text))
                            {
                                if (textBoxClientCertificatePath.Text.Length > 0)
                                {
                                    MessageBox.Show(lang.languageTemplate["strings"]["certificateFileNotFound"] + textBoxClientCertificatePath.Text, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                                    richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["certificateFileNotFound"] + textBoxClientCertificatePath.Text);
                                }
                                else
                                {
                                    MessageBox.Show(lang.languageTemplate["strings"]["certificateFileNotFound"].Replace(":", ""), "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                                    richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["certificateFileNotFound"].Replace(":", ""));
                                }
                                throw new Exception();
                            }
                            if (textBoxClientCertificatePath.Text.IndexOf(".pfx") == -1)
                            {
                                if (!File.Exists(textBoxClientKeyPath.Text))
                                {
                                    if (textBoxClientKeyPath.Text.Length > 0)
                                    {
                                        MessageBox.Show(lang.languageTemplate["strings"]["certificateFileNotFound"] + textBoxClientKeyPath.Text, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                                        richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["certificateFileNotFound"] + textBoxClientKeyPath.Text);
                                    }
                                    else
                                    {
                                        MessageBox.Show(lang.languageTemplate["strings"]["certificateFileNotFound"].Replace(":", ""), "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                                        richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["certificateFileNotFound"].Replace(":", ""));
                                    }

                                    throw new Exception();
                                }
                            }
                        });
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        // Initialise Modbus Object
                        ModBus = new ModBus_Chicco(
                            serialPort, 
                            textBoxTcpClientIpAddress.Text, 
                            textBoxTcpClientPort.Text, 
                            TCPMode, 
                            pictureBoxTx, 
                            pictureBoxRx,
                            textBoxClientCertificatePath.Text,
                            textBoxCertificatePassword.Text,
                            textBoxClientKeyPath.Text,
                            comboBoxTlsVersion.SelectedIndex);
                        ModBus.open();

                        pictureBoxTcp.Background = Brushes.Lime;
                        pictureBoxRunningAs.Background = Brushes.Lime;
                        radioButtonModeSerial.IsEnabled = false;
                        radioButtonModeTcp.IsEnabled = false;

                        richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["connectedTo"] + " " + ip_address + ":" + port);

                        textBlockTcpActive.Text = lang.languageTemplate["strings"]["disconnect"];
                        menuItemToolBit.IsEnabled = true;
                        menuItemToolWord.IsEnabled = true;
                        menuItemToolByte.IsEnabled = true;
                        salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = false;
                        caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = false;
                        gestisciDatabaseToolStripMenuItem.IsEnabled = false;

                        textBoxTcpClientIpAddress.IsEnabled = false;
                        textBoxTcpClientPort.IsEnabled = false;
                        languageToolStripMenu.IsEnabled = false;

                        buttonLoadClientCertificate.IsEnabled = false;
                        buttonLoadClientKey.IsEnabled = false;
                        textBoxCertificatePassword.IsEnabled = false;
                        comboBoxTlsVersion.IsEnabled = false;
                        checkBoxModbusSecure.IsEnabled = false;

                        checkBoxModbusSecure.IsEnabled = false;
                        buttonLoadClientCertificate.IsEnabled = false;
                        buttonLoadClientKey.IsEnabled = false;
                        textBoxCertificatePassword.IsEnabled = false;
                        comboBoxTlsVersion.IsEnabled = false;

                        if (pathToConfiguration != defaultPathToConfiguration)
                            this.Title = title + " " + version + " - File: " + pathToConfiguration + " - Ip: " + textBoxTcpClientIpAddress.Text + ":" + textBoxTcpClientPort.Text;
                        else
                            this.Title = title + " " + version + " - Ip: " + textBoxTcpClientIpAddress.Text + ":" + textBoxTcpClientPort.Text;
                    });
                }
                catch(Exception err)
                {
                    Console.WriteLine(err);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        if (err is CryptographicException)
                            if (err.ToString().IndexOf("password is not correct") != -1)
                                MessageBox.Show(lang.languageTemplate["strings"]["certificatePasswordNotValid"] + textBoxClientCertificatePath.Text, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);

                        richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["failedToConnect"] + " " + ip_address + ":" + port);
                        changeEnableButtonsConnect(false);
                        buttonTcpActive.IsEnabled = true;

                        if (pathToConfiguration != defaultPathToConfiguration)
                            this.Title = title + " " + version + " - File: " + pathToConfiguration;
                        else
                            this.Title = title + " " + version;
                    });

                    return;
                }
            }
            else
            {
                // Close connection
                this.Dispatcher.Invoke((Action)delegate
                {
                    pictureBoxTcp.Background = Brushes.LightGray;
                    pictureBoxRunningAs.Background = Brushes.LightGray;

                    richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["disconnected"]);

                    textBlockTcpActive.Text = lang.languageTemplate["strings"]["connect"];
                    menuItemToolBit.IsEnabled = false;
                    menuItemToolWord.IsEnabled = false;
                    menuItemToolByte.IsEnabled = false;
                    salvaConfigurazioneNelDatabaseToolStripMenuItem.IsEnabled = true;
                    caricaConfigurazioneDalDatabaseToolStripMenuItem.IsEnabled = true;
                    gestisciDatabaseToolStripMenuItem.IsEnabled = true;

                    radioButtonModeSerial.IsEnabled = true;
                    radioButtonModeTcp.IsEnabled = true;

                    textBoxTcpClientIpAddress.IsEnabled = true;
                    textBoxTcpClientPort.IsEnabled = true;
                    languageToolStripMenu.IsEnabled = true;

                    checkBoxModbusSecure.IsEnabled = true;
                    buttonLoadClientCertificate.IsEnabled = true;
                    buttonLoadClientKey.IsEnabled = true;
                    textBoxCertificatePassword.IsEnabled = true;
                    comboBoxTlsVersion.IsEnabled = true;

                    // Close the connection
                    ModBus.close();

                    // Fermo eventuali loop
                    disableAllLoops();
                });
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                changeEnableButtonsConnect(pictureBoxRunningAs.Background == Brushes.Lime);
                buttonTcpActive.IsEnabled = true;
            });
        }

        //----------------------------------------------------------------------------------
        //------------------------------- COILS --------------------------------------------
        //----------------------------------------------------------------------------------

        private void buttonReadCoils01_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                buttonReadCoils01.IsEnabled = false;
            }

            lastCoilsCommandGroup = false;

            Thread t = new Thread(new ThreadStart(readCoils));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readCoils()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_) + P.uint_parser(textBoxCoilsAddress01_, comboBoxCoilsAddress01_);

                if (ModBus.type == ModBus_Def.TYPE_RTU && uint.Parse(textBoxCoilNumber_) > 125)
                    MessageBox.Show(lang.languageTemplate["strings"]["maxRegNumber"], "Info");
                else if (uint.Parse(textBoxCoilNumber_) > 123)
                    MessageBox.Show(lang.languageTemplate["strings"]["maxRegNumber"], "Info");
                else
                {
                    UInt16[] response = ModBus.readCoilStatus_01(byte.Parse(textBoxModbusAddress_), address_start, uint.Parse(textBoxCoilNumber_), readTimeout);

                    if (response != null)
                    {
                        if (response.Length > 0)
                        {
                            // Cancello la tabella e inserisco le nuove righe
                            insertRowsTable(list_coilsTable, 
                                list_template_coilsTable, 
                                P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_), 
                                address_start, 
                                response, 
                                colorDefaultReadCellStr, 
                                comboBoxCoilsOffset_,
                                comboBoxCoilsRegistri_,
                                "DEC", 
                                true, 
                                false);
                        }
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonReadCoils01.IsEnabled = true;

                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_coilsTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if(ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonReadCoils01.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_coilsTable, false);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_coilsTable, (ModbusException)ex, false);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_coilsTable, false);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadCoils01.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_coilsTable, false);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadCoils01.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });

                Console.WriteLine(ex);
            }
        }

        private void buttonReadCoilsRange_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                buttonReadCoilsRange.IsEnabled = false;
            }

            lastCoilsCommandGroup = false;
            list_coilsTable.Clear();

            Thread t = new Thread(new ThreadStart(readColisRange));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readColisRange()
        {
            try
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    TaskbarItemInfo.ProgressValue = 0;
                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
                });

                uint address_start = P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_) + P.uint_parser(textBoxCoilsRange_A_, comboBoxCoilsRange_A_);
                uint coil_len = P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_) + P.uint_parser(textBoxCoilsRange_B_, comboBoxCoilsRange_A_) - address_start + 1;
                uint read_len = uint.Parse(textBoxCoilNumber_);
                uint repeatQuery = coil_len / read_len;

                if (coil_len % read_len != 0)
                {
                    repeatQuery += 1;
                }

                UInt16[] response = new UInt16[coil_len];

                for (int i = 0; i < repeatQuery; i++)
                {
                    if (i == (repeatQuery - 1) && coil_len % read_len != 0)
                    {
                        UInt16[] read = ModBus.readCoilStatus_01(byte.Parse(textBoxModbusAddress_), address_start + (uint)(read_len * i), coil_len % read_len, readTimeout);

                        // Timeout
                        if (read is null)
                        {
                            SetTableTimeoutError(list_coilsTable, false);
                        }

                        // CRC error
                        if (read.Length == 0)
                        {
                            SetTableCrcError(list_coilsTable, false);
                        }

                        Array.Copy(read, 0, response, read_len * i, coil_len % read_len);

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(i) / (double)(repeatQuery);
                        });
                    }
                    else
                    {
                        UInt16[] read = ModBus.readCoilStatus_01(byte.Parse(textBoxModbusAddress_), address_start + (uint)(read_len * i), read_len, readTimeout);

                        Array.Copy(read, 0, response, read_len * i, read_len);

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(i) / (double)(repeatQuery);
                        });
                    }
                }

                // Cancello la tabella e inserisco le nuove righe
                insertRowsTable(
                    list_coilsTable, 
                    list_template_coilsTable, 
                    P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_), 
                    address_start, 
                    response, 
                    colorDefaultReadCellStr, 
                    comboBoxCoilsOffset_, 
                    comboBoxCoilsRegistri_, 
                    "DEC", 
                    true, 
                    false);

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonReadCoilsRange.IsEnabled = true;

                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_coilsTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonReadCoilsRange.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_coilsTable, false);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_coilsTable, (ModbusException)ex, false);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_coilsTable, false);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadCoilsRange.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_coilsTable, false);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadCoilsRange.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                });

                Console.WriteLine(ex);
            }
        }

        public void SetTableInternalError(ObservableCollection<ModBus_Item> list_, bool clear)
        {
            ModBus_Item tmp = new ModBus_Item();

            tmp.Register = "Internal";
            tmp.Value = "Error";
            tmp.Foreground = ForeGroundLightStr;
            tmp.Background = Brushes.Red.ToString();

            this.Dispatcher.Invoke((Action)delegate
            {
                if (clear)
                    list_.Clear();

                list_.Add(tmp);
            });
        }

        public void SetTableCrcError(ObservableCollection<ModBus_Item> list_, bool clear)
        {
            ModBus_Item tmp = new ModBus_Item();

            tmp.Register = "CRC";
            tmp.Value = "Error";
            tmp.Foreground = ForeGroundLightStr;
            tmp.Background = Brushes.Tomato.ToString();

            this.Dispatcher.Invoke((Action)delegate
            {
                if (clear)
                    list_.Clear();

                list_.Add(tmp);
            });
        }

        public void SetTableTimeoutError(ObservableCollection<ModBus_Item> list_, bool clear)
        {
            ModBus_Item tmp = new ModBus_Item();

            tmp.Register = "Timeout";
            tmp.Value = "";
            tmp.Foreground = ForeGroundLightStr;
            tmp.Background = Brushes.Violet.ToString();

            this.Dispatcher.Invoke((Action)delegate
            {
                if (clear)
                    list_.Clear();

                list_.Add(tmp);
            });
        }

        public void SetTableModBusError(ObservableCollection<ModBus_Item> list_, ModbusException err, bool clear)
        {
            ModBus_Item tmp = new ModBus_Item();

            Console.WriteLine("err.ToString(): " + err.ToString());

            tmp.Register = "ErrCode:";
            tmp.Value = err.ToString().Split('-')[0].Split(':')[2];
            tmp.ValueBin = err.ToString().Split('-')[1].Split('\n')[0].Replace("\r", "");
            tmp.Notes = tmp.ValueBin;
            tmp.Foreground = ForeGroundLightStr;
            tmp.Background = Brushes.OrangeRed.ToString();

            this.Dispatcher.Invoke((Action)delegate
            {
                if (clear)
                    list_.Clear();

                list_.Add(tmp);
            });
        }
        public void SetTableStringError(ObservableCollection<ModBus_Item> list_, ModbusException err, bool clear)
        {
            ModBus_Item tmp = new ModBus_Item();

            Console.WriteLine("err.ToString(): " + err.ToString());

            tmp.Register = "Protocol";
            tmp.Value = "Error";
            tmp.ValueBin = err.Message;
            tmp.Notes = err.Message;
            tmp.Foreground = ForeGroundLightStr;
            tmp.Background = Brushes.PaleVioletRed.ToString();

            this.Dispatcher.Invoke((Action)delegate
            {
                if (clear)
                    list_.Clear();

                list_.Add(tmp);
            });
        }

        public void SetTableDisconnectError(ObservableCollection<ModBus_Item> list_, bool clear)
        {
            ModBus_Item tmp = new ModBus_Item();

            tmp.Register = "Sock err";
            tmp.Value = "";
            tmp.ValueBin = "Socket closed";
            tmp.Notes = tmp.ValueBin;
            tmp.Foreground = ForeGroundLightStr;
            tmp.Background = Brushes.LightBlue.ToString();

            this.Dispatcher.Invoke((Action)delegate
            {
                pictureBoxTx.Background = Brushes.LightGray;
                pictureBoxRx.Background = Brushes.LightGray;

                if (clear)
                    list_.Clear();

                list_.Add(tmp);
            });
        }

        private void buttonWriteCoils05_Click(object sender, RoutedEventArgs e)
        {
            buttonWriteCoils05.IsEnabled = false;
            lastCoilsCommandGroup = false;

            Thread t = new Thread(new ThreadStart(writeCoil_01));
            t.Priority = threadPriority;
            t.Start();
        }

        public void writeCoil_01()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_) + P.uint_parser(textBoxCoilsAddress05_, comboBoxCoilsAddress05_);

                bool? result = ModBus.forceSingleCoil_05(byte.Parse(textBoxModbusAddress_), address_start, uint.Parse(textBoxCoilsValue05_), readTimeout);

                if (result != null)
                {
                    if (result == true)
                    {
                        UInt16[] value = { UInt16.Parse(textBoxCoilsValue05_) };

                        // Cancello la tabella e inserisco le nuove righe
                        insertRowsTable(
                            list_coilsTable, 
                            list_template_coilsTable, 
                            P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_), 
                            address_start, 
                            value, 
                            colorDefaultWriteCellStr, 
                            comboBoxCoilsOffset_, 
                            comboBoxCoilsRegistri_, 
                            "DEC", 
                            true, 
                            true);
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteCoils05.IsEnabled = true;

                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });
            }
            catch (Exception ex)
            {
                if(ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_coilsTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonWriteCoils05.IsEnabled = true;
                        });
                    }
                }
                else if(ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_coilsTable, false);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_coilsTable, (ModbusException)ex, false);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_coilsTable, false);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteCoils05.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_coilsTable, true);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteCoils05.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });

                Console.WriteLine(ex);
            }
        }

        private void buttonWriteCoils05_B_Click(object sender, RoutedEventArgs e)
        {
            buttonWriteCoils05_B.IsEnabled = false;
            lastCoilsCommandGroup = false;

            Thread t = new Thread(new ThreadStart(writeCoil_02));
            t.Priority = threadPriority;
            t.Start();
        }

        public void writeCoil_02()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_) + P.uint_parser(textBoxCoilsAddress05_b_, comboBoxCoilsAddress05_b_);

                bool? result = ModBus.forceSingleCoil_05(byte.Parse(textBoxModbusAddress_), address_start, uint.Parse(textBoxCoilsValue05_b_), readTimeout);

                if (result != null)
                {
                    if (result == true)
                    {
                        UInt16[] value = { UInt16.Parse(textBoxCoilsValue05_b_) };

                        // Cancello la tabella e inserisco le nuove righe
                        insertRowsTable(
                            list_coilsTable, 
                            list_template_coilsTable, 
                            P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_), 
                            address_start, 
                            value, 
                            colorDefaultWriteCellStr, 
                            comboBoxCoilsOffset_, 
                            comboBoxCoilsRegistri_, 
                            "DEC", 
                            true, 
                            true);
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteCoils05_B.IsEnabled = true;

                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_coilsTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonWriteCoils05_B.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_coilsTable, false);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_coilsTable, (ModbusException)ex, false);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_coilsTable, false);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteCoils05_B.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_coilsTable, true);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteCoils05_B.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });

                Console.WriteLine(ex);
            }
        }

        private void buttonWriteCoils15_Click(object sender, RoutedEventArgs e)
        {
            buttonWriteCoils15.IsEnabled = false;
            lastCoilsCommandGroup = false;

            Thread t = new Thread(new ThreadStart(writeMultipleCoils));
            t.Priority = threadPriority;
            t.Start();
        }

        public void writeMultipleCoils()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxCoilsOffset_, comboBoxHoldingOffset_) + P.uint_parser(textBoxCoilsAddress15_A_, comboBoxCoilsAddress15_A_);

                uint stop = uint.Parse(textBoxCoilsAddress15_B_);

                bool[] buffer = new bool[stop];

                for (int i = 0; i < stop; i++)
                {
                    buffer[i] = uint.Parse(textBoxCoilsValue15_.Substring(i, 1)) > 0;
                }

                bool? result = ModBus.forceMultipleCoils_15(byte.Parse(textBoxModbusAddress_), address_start, buffer, readTimeout);

                if (result != null)
                {
                    if (result == true)
                    {
                        UInt16[] value = new UInt16[stop];

                        for (int i = 0; i < stop; i++)
                        {
                            value[i] = uint.Parse(textBoxCoilsValue15_.Substring(i, 1)) > 0 ? (UInt16)(1) : (UInt16)(0);
                        }

                        // Cancello la tabella e inserisco le nuove righe
                        insertRowsTable(
                            list_coilsTable, 
                            list_template_coilsTable, 
                            P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_), 
                            address_start, 
                            value, 
                            colorDefaultWriteCellStr, 
                            comboBoxCoilsOffset_, 
                            comboBoxCoilsRegistri_, 
                            null, 
                            true, 
                            true);
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteCoils15.IsEnabled = true;

                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_coilsTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonWriteCoils15.IsEnabled = true;
                        });
                    }
                }
                else if(ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_coilsTable, true);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_coilsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_coilsTable, true);
                    }
                }
                else
                {
                    SetTableInternalError(list_coilsTable, true);
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });

                Console.WriteLine(ex);
            }
        }

        private void buttonGoToCoilAddress_Click(object sender, RoutedEventArgs e)
        {
            int index = int.Parse(textBoxGoToCoilAddress.Text);

            //dataGridViewCoils.Scroo = dataGridViewCoils.Rows[index].Cells[0];
        }

        //----------------------------------------------------------------------------------
        //--------------------------------DIGITAL INPUTS------------------------------------
        //----------------------------------------------------------------------------------

        // Read digital input
        private void buttonReadInput02_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                buttonReadInput02.IsEnabled = false;
            }

            lastInputsCommandGroup = false;

            Thread t = new Thread(new ThreadStart(readInputs));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readInputs()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_) + P.uint_parser(textBoxInputAddress02_, comboBoxInputAddress02_);

                if (address_start > 9999 && correctModbusAddressAuto)    // Se indirizzo espresso in 10001+ imposto offset a 0
                {
                    address_start = address_start - 10001;
                }

                if (ModBus.type == ModBus_Def.TYPE_RTU && uint.Parse(textBoxInputNumber_) > 125)
                {
                    MessageBox.Show(lang.languageTemplate["strings"]["maxRegNumber"], "Info");
                }
                else if (uint.Parse(textBoxInputNumber_) > 123)
                {
                    MessageBox.Show(lang.languageTemplate["strings"]["maxRegNumber"], "Info");
                }
                else
                {
                    UInt16[] response = ModBus.readInputStatus_02(byte.Parse(textBoxModbusAddress_), address_start, uint.Parse(textBoxInputNumber_), readTimeout);

                    if (response != null)
                    {
                        if (response.Length > 0)
                        {
                            // Cancello la tabella e inserisco le nuove righe
                            insertRowsTable(
                                list_inputsTable, 
                                list_template_inputsTable, 
                                P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_), 
                                address_start, 
                                response, 
                                colorDefaultReadCellStr, 
                                comboBoxInputOffset_, 
                                comboBoxInputRegistri_, 
                                "DEC", 
                                true, 
                                false);
                        }
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonReadInput02.IsEnabled = true;

                    dataGridViewInput.ItemsSource = null;
                    dataGridViewInput.ItemsSource = list_inputsTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_inputsTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonReadInput02.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_inputsTable, true);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_inputsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_inputsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_inputsTable, true);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInput02.IsEnabled = true;
                    });
                }
                else
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInput02.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewInput.ItemsSource = null;
                    dataGridViewInput.ItemsSource = list_inputsTable;
                });

                Console.WriteLine(ex);
            }
        }

        // Read input range
        private void buttonReadInputRange_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                buttonReadInputRange.IsEnabled = false;
            }

            lastInputsCommandGroup = false;
            list_inputsTable.Clear();

            Thread t = new Thread(new ThreadStart(readInputsRange));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readInputsRange()
        {
            try
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    TaskbarItemInfo.ProgressValue = 0;
                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
                });

                uint address_start = P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_) + P.uint_parser(textBoxInputRange_A_, comboBoxInputRange_A_);
                uint input_len = P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_) + P.uint_parser(textBoxInputRange_B_, comboBoxInputRange_A_) - address_start + 1;

                if (address_start > 9999 && correctModbusAddressAuto)    // Se indirizzo espresso in 10001+ imposto offset a 0
                {
                    address_start = address_start - 10001;
                }

                uint read_len = uint.Parse(textBoxInputNumber_);
                uint repeatQuery = input_len / read_len;

                if (input_len % read_len != 0)
                {
                    repeatQuery += 1;
                }

                UInt16[] response = new UInt16[input_len];

                for (int i = 0; i < repeatQuery; i++)
                {
                    if (i == (repeatQuery - 1) && input_len % read_len != 0)
                    {
                        UInt16[] read = ModBus.readInputStatus_02(byte.Parse(textBoxModbusAddress_), address_start + (uint)(read_len * i), input_len % read_len, readTimeout);

                        // Timeout
                        if (read is null)
                        {
                            SetTableTimeoutError(list_inputsTable, false);
                            return;
                        }

                        // CRC error
                        if (read.Length == 0)
                        {
                            SetTableCrcError(list_inputsTable, false);
                            return;
                        }

                        Array.Copy(read, 0, response, read_len * i, input_len % read_len);

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(i) / (double)(repeatQuery);
                        });
                    }
                    else
                    {
                        UInt16[] read = ModBus.readInputStatus_02(byte.Parse(textBoxModbusAddress_), address_start + (uint)(read_len * i), read_len, readTimeout);

                        Array.Copy(read, 0, response, read_len * i, read_len);

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(i) / (double)(repeatQuery);
                        });
                    }
                }

                // Cancello la tabella e inserisco le nuove righe
                insertRowsTable(
                    list_inputsTable, 
                    list_template_inputsTable, 
                    P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_), 
                    address_start, 
                    response, 
                    colorDefaultReadCellStr, 
                    comboBoxInputOffset_, 
                    comboBoxInputRegistri_, 
                    "DEC", 
                    true, 
                    false);

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonReadInputRange.IsEnabled = true;

                    dataGridViewInput.ItemsSource = null;
                    dataGridViewInput.ItemsSource = list_inputsTable;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_inputsTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonReadInputRange.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_inputsTable, false);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_inputsTable, (ModbusException)ex, false);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_inputsTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_inputsTable, false);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRange.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_inputsTable, false);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRange.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewInput.ItemsSource = null;
                    dataGridViewInput.ItemsSource = list_inputsTable;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                });

                Console.WriteLine(ex);
            }
        }

        // Go to digital input
        private void buttonGoToInputAddress_Click(object sender, RoutedEventArgs e)
        {
            int index = int.Parse(textBoxGoToInputAddress.Text);

            if (index > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)
                index = index - 10001;

            //dataGridViewInput.FirstDisplayedCell = dataGridViewInput.Rows[index].Cells[0];
        }

        // ----------------------------------------------------------------------------------
        // ---------------------------INPUT REGISTERS----------------------------------------
        // ----------------------------------------------------------------------------------

        // Read input register FC04
        private void buttonReadInputRegister04_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                buttonReadInputRegister04.IsEnabled = false;
            }

            lastInputRegistersCommandGroup = false;

            Thread t = new Thread(new ThreadStart(readInputRegisters));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readInputRegisters()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_) + P.uint_parser(textBoxInputRegisterAddress04_, comboBoxInputRegisterAddress04_);

                if (ModBus.type == ModBus_Def.TYPE_RTU && uint.Parse(textBoxInputRegisterNumber_) > 125)
                    MessageBox.Show(lang.languageTemplate["strings"]["maxRegNumber"], "Info");
                else if (uint.Parse(textBoxInputRegisterNumber_) > 123)
                    MessageBox.Show(lang.languageTemplate["strings"]["maxRegNumber"], "Info");
                else
                {
                    if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                    {
                        address_start = address_start - 30001;
                    }

                    UInt16[] response = ModBus.readInputRegister_04(byte.Parse(textBoxModbusAddress_), address_start, uint.Parse(textBoxInputRegisterNumber_), readTimeout);

                    if (response != null)
                    {
                        if (response.Length > 0)
                        {
                            // Cancello la tabella e inserisco le nuove righe
                            insertRowsTable(
                                list_inputRegistersTable, 
                                list_template_inputRegistersTable, 
                                P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_), 
                                address_start, 
                                response, 
                                colorDefaultReadCellStr, 
                                comboBoxInputRegOffset_, 
                                comboBoxInputRegRegistri_, 
                                comboBoxInputRegValori_, 
                                true, 
                                false);
                        }
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonReadInputRegister04.IsEnabled = true;

                    dataGridViewInputRegister.ItemsSource = null;
                    dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_inputRegistersTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonReadInputRegister04.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_inputRegistersTable, true);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_inputRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_inputRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_inputRegistersTable, true);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRegister04.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_inputRegistersTable, true);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRegister04.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewInputRegister.ItemsSource = null;
                    dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;
                });

                Console.WriteLine(ex);
            }
        }

        // Read input register range
        private void buttonReadInputRegisterRange_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                buttonReadInputRegisterRange.IsEnabled = false;
            }

            lastInputRegistersCommandGroup = false;
            list_inputRegistersTable.Clear();

            Thread t = new Thread(new ThreadStart(readInputRegistersRange));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readInputRegistersRange()
        {
            try
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    TaskbarItemInfo.ProgressValue = 0;
                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
                });

                uint address_start = P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_) + P.uint_parser(textBoxInputRegisterRange_A_, comboBoxInputRegisterRange_A_);
                uint register_len = P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_) + P.uint_parser(textBoxInputRegisterRange_B_, comboBoxInputRegisterRange_A_) - address_start + 1;

                if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 30001;
                }

                uint read_len = uint.Parse(textBoxInputRegisterNumber_);
                uint repeatQuery = register_len / read_len;

                if (register_len % read_len != 0)
                {
                    repeatQuery += 1;
                }

                UInt16[] response = new UInt16[register_len];

                for (int i = 0; i < repeatQuery; i++)
                {
                    if (i == (repeatQuery - 1) && register_len % read_len != 0)
                    {
                        UInt16[] read = ModBus.readInputRegister_04(byte.Parse(textBoxModbusAddress_), address_start + (uint)(read_len * i), register_len % read_len, readTimeout);

                        // Timeout
                        if (read is null)
                        {
                            SetTableTimeoutError(list_inputRegistersTable, false);
                            return;
                        }

                        // CRC error
                        if (read.Length == 0)
                        {
                            SetTableCrcError(list_inputRegistersTable, false);
                            return;
                        }

                        Array.Copy(read, 0, response, read_len * i, register_len % read_len);

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(i) / (double)(repeatQuery);
                        });
                    }
                    else
                    {
                        UInt16[] read = ModBus.readInputRegister_04(byte.Parse(textBoxModbusAddress_), address_start + (uint)(read_len * i), read_len, readTimeout);

                        Array.Copy(read, 0, response, read_len * i, read_len);

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(i) / (double)(repeatQuery);
                        });
                    }
                }

                if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 30001;
                }

                // Cancello la tabella e inserisco le nuove righe
                insertRowsTable(
                    list_inputRegistersTable, 
                    list_template_inputRegistersTable, 
                    P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_), 
                    address_start, 
                    response, 
                    colorDefaultReadCellStr, 
                    comboBoxInputRegOffset_, 
                    comboBoxInputRegRegistri_, 
                    comboBoxInputRegValori_, 
                    true, 
                    false);

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonReadInputRegisterRange.IsEnabled = true;

                    dataGridViewInputRegister.ItemsSource = null;
                    dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_inputRegistersTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonReadInputRegisterRange.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_inputRegistersTable, false);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_inputRegistersTable, (ModbusException)ex, false);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_inputRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_inputRegistersTable, false);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRegisterRange.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_inputRegistersTable, false);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRegisterRange.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewInputRegister.ItemsSource = null;
                    dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                });

                Console.WriteLine(ex);
            }
        }

        // Go to input register
        private void buttonGoToInputRegisterAddress_Click(object sender, RoutedEventArgs e)
        {
            int index = int.Parse(textBoxGoToInputRegisterAddress.Text);

            if (index > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)
                index = index - 30001;

            //dataGridViewInputRegister.FirstDisplayedCell = dataGridViewInputRegister.Rows[index].Cells[0];
        }

        // ----------------------------------------------------------------------------------
        // ---------------------------HOLDING REGISTER---------------------------------------
        // ----------------------------------------------------------------------------------

        // Read holding register
        private void buttonReadHolding03_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                buttonReadHolding03.IsEnabled = false;
            }

            lastHoldingRegistersCommandGroup = false;

            Thread t = new Thread(new ThreadStart(readHoldingRegisters));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readHoldingRegisters()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(textBoxHoldingAddress03_, comboBoxHoldingAddress03_);

                if (ModBus.type == ModBus_Def.TYPE_RTU && uint.Parse(textBoxHoldingRegisterNumber_) > 125)
                    MessageBox.Show(lang.languageTemplate["strings"]["maxRegNumber"], "Info");
                else if (uint.Parse(textBoxHoldingRegisterNumber_) > 123)
                    MessageBox.Show(lang.languageTemplate["strings"]["maxRegNumber"], "Info");
                else
                {
                    if (address_start > 9999 && correctModbusAddressAuto)    // Se indirizzo espresso in 40001+ imposto offset a 0
                    {
                        address_start = address_start - 40001;
                    }

                    UInt16[] response = ModBus.readHoldingRegister_03(byte.Parse(textBoxModbusAddress_), P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(textBoxHoldingAddress03_, comboBoxHoldingAddress03_), uint.Parse(textBoxHoldingRegisterNumber_), readTimeout);

                    if (response != null)
                    {
                        if (response.Length > 0)
                        {
                            // Cancello la tabella e inserisco le nuove righe
                            insertRowsTable(
                                list_holdingRegistersTable,
                                list_template_holdingRegistersTable,
                                P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_),
                                address_start,
                                response,
                                colorDefaultReadCellStr,
                                comboBoxHoldingOffset_,
                                comboBoxHoldingRegistri_,
                                comboBoxHoldingValori_,
                                true,
                                false);
                        }
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonReadHolding03.IsEnabled = true;

                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_holdingRegistersTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonReadHolding03.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_holdingRegistersTable, true);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_holdingRegistersTable, true);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadHolding03.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_holdingRegistersTable, true);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadHolding03.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });

                Console.WriteLine(ex);
            }
        }

        // Preset single register
        private void buttonWriteHolding06_Click(object sender, RoutedEventArgs e)
        {
            buttonWriteHolding06.IsEnabled = false;
            lastHoldingRegistersCommandGroup = false;

            Thread T = new Thread(new ThreadStart(witeHoldingRegister_01));
            T.Start();
        }

        public void witeHoldingRegister_01()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(textBoxHoldingAddress06_, comboBoxHoldingAddress06_);

                if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                bool? result = ModBus.presetSingleRegister_06(byte.Parse(textBoxModbusAddress_), address_start, P.uint_parser(textBoxHoldingValue06_, comboBoxHoldingValue06_), readTimeout);

                if (result != null)
                {
                    if (result == true)
                    {
                        UInt16[] value = { (UInt16)P.uint_parser(textBoxHoldingValue06_, comboBoxHoldingValue06_) };

                        // Cancello la tabella e inserisco le nuove righe
                        insertRowsTable(list_holdingRegistersTable,
                            list_template_holdingRegistersTable,
                            P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_),
                            address_start,
                            value,
                            colorDefaultWriteCellStr,
                            comboBoxHoldingOffset_,
                            comboBoxHoldingRegistri_,
                            comboBoxHoldingValori_,
                            true,
                            true);
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteHolding06.IsEnabled = true;

                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_holdingRegistersTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonWriteHolding06.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_holdingRegistersTable, true);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_holdingRegistersTable, true);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteHolding06.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_holdingRegistersTable, true);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteHolding06.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });

                Console.WriteLine(ex);
            }
        }

        // Preset multiple register
        private void buttonWriteHolding16_Click(object sender, RoutedEventArgs e)
        {
            buttonWriteHolding16.IsEnabled = false;
            lastHoldingRegistersCommandGroup = false;

            Thread t = new Thread(new ThreadStart(writeMultipleRegisters));
            t.Priority = threadPriority;
            t.Start();
        }

        public void writeMultipleRegisters()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(textBoxHoldingAddress16_A_, comboBoxHoldingAddress16_A_);

                if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                uint word_count = P.uint_parser(textBoxHoldingAddress16_B_, comboBoxHoldingAddress16_B_);

                UInt16[] buffer = new UInt16[word_count];

                if (comboBoxHoldingValue16_ == "HEX")
                {
                    if (textBoxHoldingValue16_.Split(' ').Length != (word_count * 2))
                    {
                        MessageBox.Show("Formato stringa inserito non corretto" + word_count.ToString(), "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        for (int i = 0; i < word_count; i += 1)
                        {
                            uint HB = P.uint_parser(textBoxHoldingValue16_.Split(' ')[i * 2], comboBoxHoldingValue16_);
                            uint LB = P.uint_parser(textBoxHoldingValue16_.Split(' ')[i * 2 + 1], comboBoxHoldingValue16_);

                            buffer[i] = (UInt16)((HB << 8) + LB);
                        }

                        UInt16[] writtenRegs = ModBus.presetMultipleRegisters_16(byte.Parse(textBoxModbusAddress_), address_start, buffer, readTimeout);

                        if (writtenRegs.Length == word_count)
                        {
                            // Cancello la tabella e inserisco le nuove righe
                            insertRowsTable(list_holdingRegistersTable, 
                                list_template_holdingRegistersTable, 
                                P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_), 
                                address_start, 
                                writtenRegs, 
                                colorDefaultWriteCellStr, 
                                comboBoxHoldingOffset_,
                                comboBoxHoldingRegistri_, 
                                comboBoxHoldingValori_, 
                                true, 
                                true);
                        }
                        else
                        {
                            MessageBox.Show(lang.languageTemplate["strings"]["errSetReg"], "Alert");
                            list_holdingRegistersTable[(int)(address_start)].Foreground = ForeGroundLight.ToString();
                            list_holdingRegistersTable[(int)(address_start)].Background = Brushes.Red.ToString();
                        }
                    }
                }
                else
                {
                    if (textBoxHoldingValue16_.Split(' ').Length != word_count)
                    {
                        MessageBox.Show("Formato stringa inserito non corretto", "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        for (int i = 0; i < word_count; i += 1)
                        {
                            UInt16 WD = (UInt16)P.uint_parser(textBoxHoldingValue16_.Split(' ')[i], comboBoxHoldingValue16_);

                            buffer[i] = WD;
                        }

                        UInt16[] writtenRegs = ModBus.presetMultipleRegisters_16(byte.Parse(textBoxModbusAddress_), address_start, buffer, readTimeout);

                        if (writtenRegs.Length == word_count)
                        {
                            // Cancello la tabella e inserisco le nuove righe
                            insertRowsTable(list_holdingRegistersTable, 
                                list_template_holdingRegistersTable, 
                                P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_), 
                                address_start, 
                                writtenRegs, 
                                colorDefaultWriteCellStr, 
                                comboBoxHoldingOffset_,
                                comboBoxHoldingRegistri_, 
                                comboBoxHoldingValori_, 
                                true, 
                                true);
                        }
                        else
                        {
                            MessageBox.Show(lang.languageTemplate["strings"]["errSetReg"], "Alert");
                            list_holdingRegistersTable[(int)(address_start)].Foreground = ForeGroundLight.ToString();
                            list_holdingRegistersTable[(int)(address_start)].Background = Brushes.Red.ToString();
                        }
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteHolding16.IsEnabled = true;

                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_holdingRegistersTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonWriteHolding16.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_holdingRegistersTable, true);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_holdingRegistersTable, true);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteHolding16.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_holdingRegistersTable, true);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteHolding16.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });

                Console.WriteLine(ex);
            }
        }

        // Read holding register range
        private void buttonReadHoldingRange_Click(object sender, RoutedEventArgs e)
        {
            if (sender != null)
            {
                buttonReadHoldingRange.IsEnabled = false;
            }

            lastHoldingRegistersCommandGroup = false;
            list_holdingRegistersTable.Clear();

            Thread t = new Thread(new ThreadStart(readHoldingRegistersRange));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readHoldingRegistersRange()
        {
            try
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    TaskbarItemInfo.ProgressValue = 0;
                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
                });

                uint address_start = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(textBoxHoldingRange_A_, comboBoxHoldingRange_A_);
                uint register_len = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(textBoxHoldingRange_B_, comboBoxHoldingRange_A_) - address_start + 1;

                if (address_start > 9999 && correctModbusAddressAuto)    // Se indirizzo espresso in 40001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                uint read_len = uint.Parse(textBoxHoldingRegisterNumber_);
                uint repeatQuery = register_len / read_len;

                if (register_len % read_len != 0)
                {
                    repeatQuery += 1;
                }

                UInt16[] response = new UInt16[register_len];

                for (int i = 0; i < repeatQuery; i++)
                {
                    if (i == (repeatQuery - 1) && register_len % read_len != 0)
                    {
                        UInt16[] read = ModBus.readHoldingRegister_03(byte.Parse(textBoxModbusAddress_), address_start + (uint)(read_len * i), register_len % read_len, readTimeout);

                        // Timeout
                        if (read is null)
                        {
                            SetTableTimeoutError(list_holdingRegistersTable, false);
                            return;
                        }

                        // CRC error
                        if (read.Length == 0)
                        {
                            SetTableCrcError(list_holdingRegistersTable, false);
                            return;
                        }

                        Array.Copy(read, 0, response, read_len * i, register_len % read_len);

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(i) / (double)(repeatQuery);
                        });
                    }
                    else
                    {
                        UInt16[] read = ModBus.readHoldingRegister_03(byte.Parse(textBoxModbusAddress_), address_start + (uint)(read_len * i), read_len, readTimeout);

                        // Timeout
                        if (read is null)
                        {
                            SetTableTimeoutError(list_holdingRegistersTable, false);
                            return;
                        }

                        // CRC error
                        if (read.Length == 0)
                        {
                            SetTableCrcError(list_holdingRegistersTable, false);
                            return;
                        }

                        Array.Copy(read, 0, response, read_len * i, read_len);

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(i) / (double)(repeatQuery);
                        });
                    }
                }

                if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                // Cancello la tabella e inserisco le nuove righe
                insertRowsTable(list_holdingRegistersTable, 
                    list_template_holdingRegistersTable, 
                    P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_), 
                    address_start, 
                    response, 
                    colorDefaultReadCellStr,
                    comboBoxHoldingOffset_,
                    comboBoxHoldingRegistri_, 
                    comboBoxHoldingValori_, 
                    true, 
                    false);

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonReadHoldingRange.IsEnabled = true;

                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_holdingRegistersTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonReadHoldingRange.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_holdingRegistersTable, false);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, false);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_holdingRegistersTable, false);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadHoldingRange.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_holdingRegistersTable, false);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadHoldingRange.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                });

                Console.WriteLine(ex);
            }
        }

        // Go to holding register
        private void buttonGoToHoldingAddress_Click(object sender, RoutedEventArgs e)
        {
            int index = int.Parse(textBoxGoToHoldingAddress.Text);

            if (index > 9999 && (bool)checkBoxUseOffsetInTextBox.IsChecked)
                index = index - 40001;

            //dataGridViewHolding.FirstDisplayedCell = dataGridViewHolding.Rows[index].Cells[0];
        }

        // ----------------------------------------------------------------------------------
        // --------------------------- OTHERS MODBUS ----------------------------------------
        // ----------------------------------------------------------------------------------

        private void buttonSendDiagnosticQuery_Click(object sender, RoutedEventArgs e)
        {
            byte[] diagnostic_codes = { 0x00, 0x01, 0x02, 0x03, 0x04, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E, 0x0F, 0x10, 0x11, 0x12, 0x14 };

            //DEBUG
            //MessageBox.Show(comboBoxDiagnosticFunction.SelectedIndex.ToString());
            //SelectedIndex considera -1 la cella vuota e 0 il primo elemento del menu a tendina

            if (comboBoxDiagnosticFunction.SelectedIndex >= 0)
            {
                try
                {
                    textBoxDiagnosticResponse.Text = ModBus.diagnostics_08(byte.Parse(textBoxModbusAddress.Text), diagnostic_codes[comboBoxDiagnosticFunction.SelectedIndex], UInt16.Parse(textBoxDiagnosticData.Text), readTimeout);
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                    {
                        textBoxDiagnosticData.Text = "Socket error, disconnected";

                        if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonSerialActive_Click(null, null);
                            });
                        }
                        else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonTcpActive_Click(null, null);
                            });
                        }
                    }
                    else if (ex is ModbusException)
                    {
                        if (ex.Message.IndexOf("Timed out") != -1)
                        {
                            textBoxDiagnosticData.Text = "Timeout";
                        }
                        if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                        {
                            textBoxDiagnosticResponse.Text = "ErrCode: " + ex.ToString().Split('-')[0].Split(':')[2] + " - " + ex.ToString().Split('-')[1].Split('\n')[0].Replace("\r", "");
                        }
                        if (ex.Message.IndexOf("CRC Error") != -1)
                        {
                            textBoxDiagnosticData.Text = "CRC Error";
                        }
                    }
                    else
                    {
                        textBoxDiagnosticResponse.Text = "Error executing command";
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        dataGridViewHolding.ItemsSource = null;
                        dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                    });

                    Console.WriteLine(ex);
                }
            }
            else
            {
                MessageBox.Show("Valore scelto non valido", "Alert");
            }
        }

        private void buttonSendManualDiagnosticQuery_Click(object sender, RoutedEventArgs e)
        {
            // Elimino eventuali spazi in fondo
            while (textBoxDiagnosticFunctionManual.Text[textBoxDiagnosticFunctionManual.Text.Length - 1] == ' ')
            {
                textBoxDiagnosticFunctionManual.Text = textBoxDiagnosticFunctionManual.Text.Substring(0, textBoxDiagnosticFunctionManual.Text.Length - 1);
            }

            if ((bool)radioButtonModeSerial.IsChecked)
            {
                try
                {
                    string[] bytes = textBoxDiagnosticFunctionManual.Text.Split(' ');
                    byte[] buffer = new byte[textBoxDiagnosticFunctionManual.Text.Split(' ').Count()];

                    for (int i = 0; i < buffer.Length; i++)
                    {
                        buffer[i] = byte.Parse(bytes[i], System.Globalization.NumberStyles.HexNumber);
                    }

                    serialPort.ReadTimeout = readTimeout;
                    serialPort.Write(buffer, 0, buffer.Length);

                    try
                    {
                        byte[] response = new byte[100];
                        int Length = serialPort.Read(response, 0, response.Length);

                        textBoxManualDiagnosticResponse.Text = "";

                        for (int i = 0; i < Length; i++)
                        {
                            textBoxManualDiagnosticResponse.Text += response[i].ToString("X").PadLeft(2, '0') + " ";
                        }
                    }
                    catch
                    {
                        textBoxManualDiagnosticResponse.Text = "Read timed out";
                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine(err);
                }
            }
            else
            {
                try
                {
                    TcpClient client = new TcpClient(textBoxTcpClientIpAddress.Text, int.Parse(textBoxTcpClientPort.Text));

                    string[] bytes = textBoxDiagnosticFunctionManual.Text.Split(' ');
                    byte[] buffer = new byte[textBoxDiagnosticFunctionManual.Text.Split(' ').Count()];

                    for (int i = 0; i < buffer.Length; i++)
                    {
                        buffer[i] = byte.Parse(bytes[i], System.Globalization.NumberStyles.HexNumber);
                    }

                    NetworkStream stream = client.GetStream();
                    stream.ReadTimeout = readTimeout;
                    stream.Write(buffer, 0, buffer.Length);

                    Thread.Sleep(50);

                    try
                    {
                        byte[] response = new byte[100];
                        int Length = stream.Read(response, 0, response.Length);

                        textBoxManualDiagnosticResponse.Text = "";

                        for (int i = 0; i < Length; i++)
                        {
                            textBoxManualDiagnosticResponse.Text += response[i].ToString("X").PadLeft(2, '0') + " ";
                        }
                    }
                    catch
                    {
                        textBoxManualDiagnosticResponse.Text = "Read timed out";
                    }

                    client.Close();
                }
                catch (Exception err)
                {
                    Console.WriteLine(err);
                }

            }
        }

        private void buttonCalcCrcRtu_Click(object sender, RoutedEventArgs e)
        {
            // Elimino eventuali spazi in fondo
            while (textBoxDiagnosticFunctionManual.Text[textBoxDiagnosticFunctionManual.Text.Length - 1] == ' ')
            {
                textBoxDiagnosticFunctionManual.Text = textBoxDiagnosticFunctionManual.Text.Substring(0, textBoxDiagnosticFunctionManual.Text.Length - 1);
            }

            try
            {
                string[] bytes = textBoxDiagnosticFunctionManual.Text.Split(' ');
                byte[] buffer = new byte[textBoxDiagnosticFunctionManual.Text.Split(' ').Count()];

                for (int i = 0; i < buffer.Length; i++)
                {
                    buffer[i] = byte.Parse(bytes[i], System.Globalization.NumberStyles.HexNumber);
                }

                byte[] crc = ModBus.Calcolo_CRC(buffer, buffer.Length);

                textBoxDiagnosticFunctionManual.Text += " " + crc[0].ToString("X").PadLeft(2, '0');
                textBoxDiagnosticFunctionManual.Text += " " + crc[1].ToString("X").PadLeft(2, '0');
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }
        }

        private void ButtonExampleTCP_Click(object sender, RoutedEventArgs e)
        {
            textBoxDiagnosticFunctionManual.Text = "00 01 00 00 00 06 01 03 00 02 00 04";
        }

        private void ButtonExampleRTU_Click(object sender, RoutedEventArgs e)
        {
            textBoxDiagnosticFunctionManual.Text = "01 03 00 02 00 04";
        }

        // Funzione che dal valore costruisce il mapping dei bit associato se è stato fornito un mapping
        // Mapping: b0:Status ON/OFF,b1: .....,
        public string GetMappingValue2(UInt16[] value_list, int list_index, string mappings, out string convertedValue)
        {
            convertedValue = "";

            // If mappings null sets mappings to empty string
            if (mappings == null)
                mappings = "";

            try
            {
                string[] labels = new string[16];
                string result = "";

                // 8 bytes tmp per conversioni UInt32/UInt64
                UInt16[] values_ = { 0, 0, 0, 0 };

                int type = 0; // 0 -> Not found, 1 -> bitmap, 2 -> byte map
                int a = -3;

                if (mappings.IndexOf('+') != -1)
                {
                    a = 0; // Sposto le due word da prendere in la di 1
                }

                int index_start = list_index + a;

                for (int i = 0; i < 4; i++)
                {
                    if ((index_start + i) >= 0 && (index_start + i) < value_list.Length)
                        values_[i] = value_list[index_start + i];
                }

                if (mappings.IndexOf(':') == -1 && mappings.Length > 0)
                    mappings += ":";

                foreach (string match in mappings.Split(';'))
                {
                    string test = match.Split(':')[0];

                    if (match.Split(':').Length > 1)
                    {
                        // byte (low byte or high byte)
                        if (test.ToLower().IndexOf("byte") == 0)
                        {
                            // Soluzione bug sul fatto che ragiono a blocchi di 8 byte ma prendo gli utlimi 4
                            if (test.ToLower().IndexOf("+") != -1)
                                values_[3] = values_[0];

                            if (test.ToLower().IndexOf("-") != -1)
                                labels[0] = "value (byte): " + ((byte)((values_[3] >> 8) & 0xFF)).ToString(); // High Byte
                            else
                                labels[0] = "value (byte): " + ((byte)(values_[3] & 0xFF)).ToString();      // Low Byte

                            convertedValue = labels[0].Replace("value ", "");
                            type = 4;
                        }

                        // bitmap (type 1)
                        else if (test.IndexOf("b") == 0)
                        {
                            int index = int.Parse(test.Substring(1));

                            labels[index] = match.Split(':')[1];
                            type = 1;

                            convertedValue = "(bitmask)";
                        }

                        // bytemap (type 2)
                        else if (test.IndexOf("B") == 0)
                        {
                            int index = int.Parse(test.Substring(1));

                            labels[index] = match.Split(':')[1];
                            type = 2;

                            convertedValue = String.Format("(mask): H: {0} L: {1}", ((values_[3]) >> 8), (values_[3] & 0xFF));
                        }

                        // enum (type 10)
                        else if (test.IndexOf("E") == 0)
                        {
                            if (UInt16.Parse(test.Substring(1)) == values_[3])
                            {
                                convertedValue = String.Format("(enum): {1}", match.Split(':')[0], match.Split(':')[1]);
                                type = 10;
                            }

                            result += labels[0] + String.Format("{0}: {1}", match.Split(':')[0], match.Split(':')[1]);
                            labels[0] = "\n";
                        }

                        // float (type 3)
                        else if (test.IndexOf("F") == 0 || test.ToLower().IndexOf("float") == 0)
                        {
                            byte[] tmp = new byte[4];

                            // Soluzione bug sul fatto che ragiono a blocchi di 8 byte ma prendo gli utlimi 4
                            if (test.ToLower().IndexOf("+") != -1)
                            {
                                values_[2] = values_[0];
                                values_[3] = values_[1];
                            }

                            if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                            {
                                tmp[0] = (byte)(values_[2] & 0xFF);
                                tmp[1] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[3] & 0xFF);
                                tmp[3] = (byte)((values_[3] >> 8) & 0xFF);
                            }
                            else
                            {
                                tmp[0] = (byte)(values_[3] & 0xFF);
                                tmp[1] = (byte)((values_[3] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[2] & 0xFF);
                                tmp[3] = (byte)((values_[2] >> 8) & 0xFF);
                            }

                            labels[0] = "value (float32): " + BitConverter.ToSingle(tmp, 0).ToString(System.Globalization.CultureInfo.InvariantCulture); // + " " + match.Split(':')[1];
                            convertedValue = labels[0].Replace("value ", "");
                            type = 3;
                        }

                        // double (type 7)
                        else if (test.IndexOf("d") == 0 || test.ToLower().IndexOf("double") == 0)
                        {
                            byte[] tmp = new byte[8];

                            if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                            {
                                tmp[0] = (byte)(values_[0] & 0xFF);
                                tmp[1] = (byte)((values_[0] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[1] & 0xFF);
                                tmp[3] = (byte)((values_[1] >> 8) & 0xFF);
                                tmp[4] = (byte)(values_[2] & 0xFF);
                                tmp[5] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[6] = (byte)(values_[3] & 0xFF);
                                tmp[7] = (byte)((values_[3] >> 8) & 0xFF);
                            }
                            else
                            {
                                tmp[0] = (byte)(values_[3] & 0xFF);
                                tmp[1] = (byte)((values_[3] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[2] & 0xFF);
                                tmp[3] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[4] = (byte)(values_[1] & 0xFF);
                                tmp[5] = (byte)((values_[1] >> 8) & 0xFF);
                                tmp[6] = (byte)(values_[0] & 0xFF);
                                tmp[7] = (byte)((values_[0] >> 8) & 0xFF);
                            }

                            labels[0] = "value (double64): " + BitConverter.ToDouble(tmp, 0).ToString(System.Globalization.CultureInfo.InvariantCulture); // + " " + match.Split(':')[1];
                            convertedValue = labels[0].Replace("value ", "");
                            type = 7;
                        }

                        // uint64 (type 6)
                        else if (test.ToLower().IndexOf("uint64") == 0)
                        {
                            byte[] tmp = new byte[8];

                            if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                            {
                                tmp[0] = (byte)(values_[0] & 0xFF);
                                tmp[1] = (byte)((values_[0] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[1] & 0xFF);
                                tmp[3] = (byte)((values_[1] >> 8) & 0xFF);
                                tmp[4] = (byte)(values_[2] & 0xFF);
                                tmp[5] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[6] = (byte)(values_[3] & 0xFF);
                                tmp[7] = (byte)((values_[3] >> 8) & 0xFF);
                            }
                            else
                            {
                                tmp[0] = (byte)(values_[3] & 0xFF);
                                tmp[1] = (byte)((values_[3] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[2] & 0xFF);
                                tmp[3] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[4] = (byte)(values_[1] & 0xFF);
                                tmp[5] = (byte)((values_[1] >> 8) & 0xFF);
                                tmp[6] = (byte)(values_[0] & 0xFF);
                                tmp[7] = (byte)((values_[0] >> 8) & 0xFF);
                            }

                            labels[0] = "value (uint64): " + BitConverter.ToUInt64(tmp, 0).ToString(); // + "" + match.Split(':')[1];
                            convertedValue = labels[0].Replace("value ", "");
                            type = 6;
                        }

                        // int64 (type 6)
                        else if (test.ToLower().IndexOf("int64") == 0)
                        {
                            byte[] tmp = new byte[8];

                            if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                            {
                                tmp[0] = (byte)(values_[0] & 0xFF);
                                tmp[1] = (byte)((values_[0] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[1] & 0xFF);
                                tmp[3] = (byte)((values_[1] >> 8) & 0xFF);
                                tmp[4] = (byte)(values_[2] & 0xFF);
                                tmp[5] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[6] = (byte)(values_[3] & 0xFF);
                                tmp[7] = (byte)((values_[3] >> 8) & 0xFF);
                            }
                            else
                            {
                                tmp[0] = (byte)(values_[3] & 0xFF);
                                tmp[1] = (byte)((values_[3] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[2] & 0xFF);
                                tmp[3] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[4] = (byte)(values_[1] & 0xFF);
                                tmp[5] = (byte)((values_[1] >> 8) & 0xFF);
                                tmp[6] = (byte)(values_[0] & 0xFF);
                                tmp[7] = (byte)((values_[0] >> 8) & 0xFF);
                            }

                            labels[0] = "value (int64): " + BitConverter.ToInt64(tmp, 0).ToString(); // + "" + match.Split(':')[1];
                            convertedValue = labels[0].Replace("value ", "");
                            type = 6;
                        }

                        // uint32 (type 5)
                        else if (test.ToLower().IndexOf("uint32") == 0)
                        {
                            byte[] tmp = new byte[4];

                            // Soluzione bug sul fatto che ragiono a blocchi di 8 byte ma prendo gli utlimi 4
                            if (test.ToLower().IndexOf("+") != -1)
                            {
                                values_[2] = values_[0];
                                values_[3] = values_[1];
                            }

                            if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                            {
                                tmp[0] = (byte)(values_[2] & 0xFF);
                                tmp[1] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[3] & 0xFF);
                                tmp[3] = (byte)((values_[3] >> 8) & 0xFF);
                            }
                            else
                            {
                                tmp[0] = (byte)(values_[3] & 0xFF);
                                tmp[1] = (byte)((values_[3] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[2] & 0xFF);
                                tmp[3] = (byte)((values_[2] >> 8) & 0xFF);
                            }

                            labels[0] = "value (uint32): " + BitConverter.ToUInt32(tmp, 0).ToString(); // + "" + match.Split(':')[1];
                            convertedValue = labels[0].Replace("value ", "");
                            type = 5;
                        }

                        // int32 (type 5)
                        else if (test.ToLower().IndexOf("int32") == 0)
                        {
                            byte[] tmp = new byte[4];

                            // Soluzione bug sul fatto che ragiono a blocchi di 8 byte ma prendo gli utlimi 4
                            if (test.ToLower().IndexOf("+") != -1)
                            {
                                values_[2] = values_[0];
                                values_[3] = values_[1];
                            }

                            if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                            {
                                tmp[0] = (byte)(values_[2] & 0xFF);
                                tmp[1] = (byte)((values_[2] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[3] & 0xFF);
                                tmp[3] = (byte)((values_[3] >> 8) & 0xFF);
                            }
                            else
                            {
                                tmp[0] = (byte)(values_[3] & 0xFF);
                                tmp[1] = (byte)((values_[3] >> 8) & 0xFF);
                                tmp[2] = (byte)(values_[2] & 0xFF);
                                tmp[3] = (byte)((values_[2] >> 8) & 0xFF);
                            }

                            labels[0] = "value (int32): " + BitConverter.ToInt32(tmp, 0).ToString(); // + "" + match.Split(':')[1];
                            convertedValue = labels[0].Replace("value ", "");
                            type = 5;
                        }

                        // uint16 (type 4)
                        else if (test.ToLower().IndexOf("uint") == 0 || test.ToLower().IndexOf("uint16") == 0)
                        {
                            byte[] tmp = new byte[2];

                            // Soluzione bug sul fatto che ragiono a blocchi di 8 byte ma prendo gli utlimi 4
                            if (test.ToLower().IndexOf("+") != -1)
                            {
                                values_[3] = values_[0];
                            }

                            tmp[0] = (byte)(values_[3] & 0xFF);
                            tmp[1] = (byte)((values_[3] >> 8) & 0xFF);

                            labels[0] = "value (uint16): " + BitConverter.ToUInt16(tmp, 0).ToString(); // + "" + match.Split(':')[1];
                            convertedValue = labels[0].Replace("value ", "");
                            type = 4;
                        }

                        // int16 (type 4)
                        else if (test.ToLower().IndexOf("int") == 0 || test.ToLower().IndexOf("int16") == 0)
                        {
                            byte[] tmp = new byte[2];

                            // Soluzione bug sul fatto che ragiono a blocchi di 8 byte ma prendo gli utlimi 4
                            if (test.ToLower().IndexOf("+") != -1)
                            {
                                values_[3] = values_[0];
                            }

                            tmp[0] = (byte)(values_[3] & 0xFF);
                            tmp[1] = (byte)((values_[3] >> 8) & 0xFF);

                            labels[0] = "value (int16): " + BitConverter.ToInt16(tmp, 0).ToString(); // + "" + match.Split(':')[1];
                            convertedValue = labels[0].Replace("value ", "");
                            type = 4;
                        }

                        // String (type 255)
                        else if (test.ToLower().IndexOf("string") == 0)
                        {
                            int length = int.Parse(test.Split(')')[0].Split('(')[1].Split('.')[0]);
                            int offset = 0;

                            if (test.Split(')')[0].Split('(')[1].IndexOf('.') != -1)
                            {
                                int.TryParse(test.Split(')')[0].Split('(')[1].Split('.')[1], out offset);
                            }

                            byte[] tmp = new byte[length];
                            String output = "";

                            int start = 0;
                            int stop = 0;

                            start = list_index + ((offset) / 2);
                            stop = list_index + (length / 2) + ((offset) / 2);

                            for (int i = start; i < stop; i += 1)
                            {
                                UInt16 currrValue = value_list[i];

                                try
                                {
                                    ASCIIEncoding ascii = new ASCIIEncoding();

                                    if (test.ToLower().IndexOf("_swap") != -1)
                                    {
                                        output += ascii.GetString(new byte[] { (byte)(currrValue & 0xFF), (byte)((currrValue >> 8) & 0xFF) });
                                    }
                                    else
                                    {
                                        output += ascii.GetString(new byte[] { (byte)((currrValue >> 8) & 0xFF), (byte)(currrValue & 0xFF) });
                                    }
                                }
                                catch (Exception err)
                                {
                                    Console.WriteLine(err);
                                }
                            }

                            labels[0] = "value (string): " + output;
                            convertedValue = labels[0].Replace("value ", "");
                            type = 7;
                        }
                    }

                    // etichetta generica senza mapping
                    if (mappings.Length < 2)
                    {
                        labels[0] = "value (dec): " + values_[3].ToString() + "\nvalue (hex): 0x" + values_[3].ToString("X").PadLeft(4, '0') + "\nvalue (bin): " + Convert.ToString(values_[3] >> 8, 2).PadLeft(8, '0') + " " + Convert.ToString((UInt16)(values_[3] & 0x00FF), 2).PadLeft(8, '0');
                        //convertedValue = labels[0];
                        type = 255;
                    }
                }

                // bitmap
                if (type == 1)
                {
                    /*if (value_list[index_start + Math.Abs(a)].Notes != null)
                    {
                        if (value_list[index_start + Math.Abs(a)].Notes.Length > 0)
                        {
                            result = value_list[index_start + Math.Abs(a)].Notes + "\n\n";
                        }
                    }*/

                    for (int i = 15; i >= 0; i--)
                    {
                        if ((values_[3] & (1 << i)) > 0)
                        {
                            if (i < 10)
                            {
                                result += "bit   " + i.ToString() + ":  1 - " + labels[i];
                            }
                            else
                            {
                                result += "bit " + i.ToString() + ":  1 - " + labels[i];
                            }
                        }
                        else
                        {
                            if (i < 10)
                            {
                                result += "bit   " + i.ToString() + ":  0 - " + labels[i];
                            }
                            else
                            {
                                result += "bit " + i.ToString() + ":  0 - " + labels[i];
                            }
                        }

                        if (i > 0)
                        {
                            result += "\n";
                        }
                    }
                }

                // bytemap
                else if (type == 2)
                {
                    /*if (value_list[index_start + Math.Abs(a)].Notes != null)
                    {
                        if (value_list[index_start + Math.Abs(a)].Notes.Length > 0)
                        {
                            result = value_list[index_start + Math.Abs(a)].Notes + "\n\n";
                        }
                    }*/

                    for (int i = 1; i >= 0; i--)
                    {
                        result += "byte " + i.ToString() + ": " + ((values_[3] >> (i * 8)) & 0xFF).ToString() + " - " + labels[i];

                        if (i == 1)
                        {
                            result += "\n";
                        }
                    }
                }

                // conversioni interi
                else if (type == 3 || type == 4 || type == 5 || type == 6 || type == 7)
                {
                    /*if (value_list[index_start - a].Notes != null)
                    {
                        if (value_list[index_start - a].Notes.Length > 0)
                        {
                            result = value_list[index_start - a].Notes + "\n\n";
                        }
                    }*/

                    result += labels[0];
                }

                // etichetta generica
                else if (type == 255)
                {
                    /*if (value_list[index_start - a].Notes != null)
                    {
                        if (value_list[index_start - a].Notes.Length > 0)
                        {
                            result = value_list[index_start - a].Notes + "\n\n";
                        }
                    }*/

                    result += labels[0];
                }

                return result;
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
                return "";
            }
        }

        // Funzione inserimento righe nelle collections
        public void insertRowsTable(ObservableCollection<ModBus_Item> tab_1, IEnumerable<ModBus_Item> template, uint user_offset, uint address_start, UInt16[] response, String cellBackGround, String formatOffset, String formatRegister, String formatVal, bool clearTable, bool colorAlways)
        {
            if (clearTable)
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    tab_1.Clear();
                });
            }

            if (response != null)
            {
                for (int i = 0; i < response.Length; i++)
                {
                    ModBus_Item row = new ModBus_Item();

                    // Cella con numero del registro
                    if (formatOffset == "DEC" || formatOffset == null)
                    {
                        row.Offset = ((useOffsetInTable ? user_offset : 0)).ToString();
                    }
                    else
                    {
                        row.Offset = "0x" + ((useOffsetInTable ? user_offset : 0)).ToString("X").PadLeft(4, '0');
                    }

                    // Valore reale registro modbus
                    row.RegisterUInt = (UInt16)(address_start + i);

                    // Cella con numero del registro
                    if (formatRegister == "DEC" || formatRegister == null)
                    {
                        row.Register = (address_start + i - (useOffsetInTable ? user_offset : 0)).ToString();
                    }
                    else
                    {
                        row.Register = "0x" + (address_start + i - (useOffsetInTable ? user_offset : 0)).ToString("X").PadLeft(4, '0');
                    }

                    // Cella con valore letto
                    if (formatVal == "DEC" || formatVal == null)
                    {
                        row.Value = (response[i]).ToString();
                    }
                    else
                    {
                        row.Value = "0x" + (response[i]).ToString("X").PadLeft(4, '0');
                    }

                    // Colorazione celle
                    if (colorMode)
                    {
                        if (response[i] > 0 || colorAlways)
                        {
                            row.Foreground = darkMode ? ForeGroundDarkStr : ForeGroundLightStr;
                            row.Background = cellBackGround;
                        }
                        else
                        {
                            row.Foreground = darkMode ? ForeGroundDarkStr : ForeGroundLightStr;
                            row.Background = darkMode ? BackGroundDarkStr : Brushes.White.ToString();
                        }
                    }
                    else
                    {
                        if (i % 2 == 0 || colorAlways)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                row.Foreground = darkMode ? BackGroundDarkStr : ForeGroundLightStr;
                                row.Background = cellBackGround;
                            });
                        }
                    }

                    ModBus_Item found = template.FirstOrDefault(x => x.RegisterUInt == (address_start + i));
                    if (found != null)
                    {
                        String convertedValue;
                        row.Mappings = GetMappingValue2(response, i, found.Mappings, out convertedValue);
                        row.Notes = found.Notes;
                        row.ValueConverted = convertedValue;
                    }
                    else
                    {
                        String convertedValue;
                        row.Mappings = GetMappingValue2(response, i, "", out convertedValue);
                        row.ValueConverted = convertedValue;
                    }

                    // Il valore in binario lo metto sempre tanto poi nelle coils ed inputs è nascosto
                    row.ValueBin = Convert.ToString(response[i] >> 8, 2).PadLeft(8, '0') + " " + Convert.ToString((UInt16)(response[i] << 8) >> 8, 2).PadLeft(8, '0');

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        tab_1.Add(row);
                    });
                }
            }
        }

        // ----------------------------------------------------------------------------------
        // --------------------------- OTHERS -----------------------------------------------
        // ----------------------------------------------------------------------------------

        // Altri pulsanti della grafica
        private void guidaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(localPath + "\\Manuals\\ModBus_Client_" + textBoxCurrentLanguage.Text + ".pdf");
            }
            catch
            {
                MessageBox.Show("File \"" + localPath + "\\Manuals\\Guida_ModBus_Client_" + textBoxCurrentLanguage.Text + ".pdf\" not found", "Alert", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void infoToolStripMenuItem1_Click(object sender, RoutedEventArgs e)
        {
            Info info = new Info(title, version);
            info.Show();
        }

        private void richTextBoxAppend(RichTextBox richTextBox, String append)
        {
            richTextBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + " " + append + "\n");
            richTextBox.ScrollToEnd();
        }

        private void buttonClearSerialStatus_Click(object sender, RoutedEventArgs e)
        {
            richTextBoxStatus.Document.Blocks.Clear();
            richTextBoxStatus.AppendText("\n");
        }

        private void esciToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            //SaveConfiguration_v1(false);
            SaveConfiguration_v2(false);
            this.Close();
        }

        private void impostazioniToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            tabControlMain.SelectedIndex = 7;
        }

        private void buttonColorCellRead_Click(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (darkMode)
                    colorDefaultReadCell_Dark = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                else
                    colorDefaultReadCell_Light = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));

                colorDefaultReadCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellRead.Background = colorDefaultReadCell;
                colorDefaultReadCellStr = colorDefaultReadCell.ToString();
            }
        }

        private void buttonCellWrote_Click(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (darkMode)
                    colorDefaultWriteCell_Dark = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                else
                    colorDefaultWriteCell_Light = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));

                colorDefaultWriteCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellWrote.Background = colorDefaultWriteCell;
            }
        }

        private void buttonColorCellError_Click(object sender, RoutedEventArgs e)
        {
            if (colorDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (darkMode)
                    colorErrorCell_Dark = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                else
                    colorErrorCell_Light = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));

                colorErrorCell = new SolidColorBrush(Color.FromArgb(colorDialogBox.Color.A, colorDialogBox.Color.R, colorDialogBox.Color.G, colorDialogBox.Color.B));
                labelColorCellError.Background = colorErrorCell;
            }
        }

        // Salvataggio pacchetti inviati
        private void buttonExportSentPackets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                saveFileDialogBox = new SaveFileDialog();

                saveFileDialogBox.DefaultExt = ".txt";
                saveFileDialogBox.AddExtension = false;
                saveFileDialogBox.FileName = "Pacchetti_inviati_" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "." + DateTime.Now.Month.ToString().PadLeft(2, '0') + "." + DateTime.Now.Year.ToString().PadLeft(2, '0');
                saveFileDialogBox.Filter = "Text|*.txt|Log|*.log";
                saveFileDialogBox.Title = "Salva Log pacchetti inviati";
                saveFileDialogBox.ShowDialog();

                if (saveFileDialogBox.FileName != "")
                {
                    StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                    TextRange textRange = new TextRange(
                        // TextPointer to the start of content in the RichTextBox.
                        richTextBoxPackets.Document.ContentStart,
                        // TextPointer to the end of content in the RichTextBox.
                        richTextBoxPackets.Document.ContentEnd
                    );


                    writer.Write(textRange.Text);
                    writer.Dispose();
                    writer.Close();
                }
            }
            catch { }
        }

        // Salvataggio pacchetti ricevuti
        private void buttonExportReceivedPackets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                saveFileDialogBox = new SaveFileDialog();

                saveFileDialogBox.DefaultExt = ".txt";
                saveFileDialogBox.AddExtension = false;
                saveFileDialogBox.FileName = "Pacchetti_ricevuti_" + DateTime.Now.Day.ToString().PadLeft(2, '0') + "." + DateTime.Now.Month.ToString().PadLeft(2, '0') + "." + DateTime.Now.Year.ToString().PadLeft(2, '0');
                saveFileDialogBox.Filter = "Text|*.txt|Log|*.log";
                saveFileDialogBox.Title = "Salva Log pacchetti inviati";
                saveFileDialogBox.ShowDialog();

                if (saveFileDialogBox.FileName != "")
                {
                    StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                    TextRange textRange = new TextRange(
                        // TextPointer to the start of content in the RichTextBox.
                        richTextBoxPackets.Document.ContentStart,
                        // TextPointer to the end of content in the RichTextBox.
                        richTextBoxPackets.Document.ContentEnd
                    );


                    writer.Write(textRange.Text);
                    writer.Dispose();
                    writer.Close();
                }
            }
            catch { }
        }

        private void buttonWriteHolding06_b_Click(object sender, RoutedEventArgs e)
        {
            buttonWriteHolding06_b.IsEnabled = false;
            lastHoldingRegistersCommandGroup = false;

            Thread t = new Thread(new ThreadStart(writeHoldingRegister_02));
            t.Priority = threadPriority;
            t.Start();
        }

        public void writeHoldingRegister_02()
        {
            try
            {
                uint address_start = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(textBoxHoldingAddress06_b_, comboBoxHoldingAddress06_b_);

                if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                {
                    address_start = address_start - 40001;
                }

                bool? result = ModBus.presetSingleRegister_06(byte.Parse(textBoxModbusAddress_), address_start, P.uint_parser(textBoxHoldingValue06_b_, comboBoxHoldingValue06_b_), readTimeout);

                if (result != null)
                {
                    if (result == true)
                    {
                        UInt16[] value = { (UInt16)P.uint_parser(textBoxHoldingValue06_b_, comboBoxHoldingValue06_b_) };

                        // Cancello la tabella e inserisco le nuove righe
                        insertRowsTable(list_holdingRegistersTable, 
                            list_template_holdingRegistersTable, 
                            P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_), 
                            address_start, 
                            value, 
                            colorDefaultWriteCellStr, 
                            comboBoxHoldingOffset_,
                            comboBoxHoldingRegistri_, 
                            comboBoxHoldingValori_, 
                            true, 
                            true);
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteHolding06_b.IsEnabled = true;

                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });
            }
            catch (Exception ex)
            {
                if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                {
                    SetTableDisconnectError(list_holdingRegistersTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonSerialActive_Click(null, null);
                        });
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (ModBus.ClientActive)
                                buttonTcpActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonWriteHolding06_b.IsEnabled = true;
                        });
                    }
                }
                else if (ex is ModbusException)
                {
                    if (ex.Message.IndexOf("Timed out") != -1)
                    {
                        SetTableTimeoutError(list_holdingRegistersTable, true);
                    }
                    if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                    {
                        SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                    {
                        SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                    }
                    if (ex.Message.IndexOf("CRC Error") != -1)
                    {
                        SetTableCrcError(list_holdingRegistersTable, true);
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteHolding06_b.IsEnabled = true;
                    });
                }
                else
                {
                    SetTableInternalError(list_holdingRegistersTable, true);

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonWriteHolding06_b.IsEnabled = true;
                    });
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });

                Console.WriteLine(ex);
            }
        }

        private void checkBoxViewTableWithoutOffset_CheckedChanged(object sender, RoutedEventArgs e)
        {
            useOffsetInTable = (bool)checkBoxViewTableWithoutOffset.IsChecked;
        }

        //-----------------------------------------------------------------------------------------
        //------------------FUNZIONE PER AGGIORNARE ELEMENTI GRAFICA DA ALTRI FORM-----------------
        //-----------------------------------------------------------------------------------------

        public void toolSTripMenuEnable(bool enable)
        {
            //simFORMToolStripMenuItem.IsEnabled = enable;
        }

        private void buttonClearAll_Click(object sender, RoutedEventArgs e)
        {
            count_log = 0;

            richTextBoxPackets.Document.Blocks.Clear();
            richTextBoxPackets.AppendText("\n");
        }

        private void gestisciDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            DatabaseManager window = new DatabaseManager(this);
            window.ShowDialog();

            if ((bool)window.DialogResult)
            {
                comboBoxProfileHome.Items.Clear();

                foreach (String sub in Directory.GetDirectories(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\"))
                {
                    comboBoxProfileHome.Items.Add(sub.Split('\\')[sub.Split('\\').Length - 1]);
                }

                disableComboProfile = true;
                comboBoxProfileHome.SelectedValue = window.SelectedProfile;
                LoadProfile(window.SelectedProfile);
                disableComboProfile = false;
            }
        }

        private void comboBoxHoldingValue06_b_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingValue06_b_ = comboBoxHoldingValue06_b.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void menuItemToolBit_Click(object sender, RoutedEventArgs e)
        {
            ComandiBit sim = new ComandiBit(ModBus, this);
            sim.Show();
        }

        private void menuItemToolByte_Click(object sender, RoutedEventArgs e)
        {
            ComandiByte sim_byte = new ComandiByte(ModBus, this);
            sim_byte.Show();
        }

        private void menuItemToolWord_Click(object sender, RoutedEventArgs e)
        {
            ComandiWord sim_word = new ComandiWord(ModBus, this);
            sim_word.Show();
        }

        private void salvaConfigurazioneAttualeNelDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            String prevoiusPath = pathToConfiguration;

            Salva_impianto form_save = new Salva_impianto(this);
            form_save.ShowDialog();

            //Controllo il risultato del form
            if ((bool)form_save.DialogResult)
            {
                //SaveConfiguration_v1(false);
                SaveConfiguration_v2(false);

                pathToConfiguration = form_save.path;

                if (pathToConfiguration != defaultPathToConfiguration)
                    this.Title = title + " " + version + " - File: " + pathToConfiguration;
                else
                    this.Title = title + " " + version;

                Directory.CreateDirectory(localPath + "\\Json\\" + pathToConfiguration);

                String[] fileNames = Directory.GetFiles(localPath + "\\Json\\" + prevoiusPath + "\\");

                for (int i = 0; i < fileNames.Length; i++)
                {
                    String newFile = localPath + "\\Json\\" + pathToConfiguration + fileNames[i].Substring(fileNames[i].LastIndexOf('\\'));

                    Console.WriteLine("Copying file: " + fileNames[i] + " to " + newFile);
                    File.Copy(fileNames[i], newFile);
                }

                disableComboProfile = true;
                comboBoxProfileHome.Items.Clear();

                foreach (String sub in Directory.GetDirectories(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\"))
                {
                    comboBoxProfileHome.Items.Add(sub.Split('\\')[sub.Split('\\').Length - 1]);
                }

                comboBoxProfileHome.SelectedValue = form_save.path;
                LoadProfile(form_save.path);
                disableComboProfile = false;
            }
        }

        private void caricaConfigurazioneDaDatabaseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Carica_impianto form_load = new Carica_impianto(defaultPathToConfiguration, this);
            form_load.ShowDialog();

            // Check form result
            if ((bool)form_load.DialogResult)
            {
                // Old way
                // LoadProfile(form_load.path);

                disableComboProfile = true;
                comboBoxProfileHome.SelectedValue = form_load.path;
                LoadProfile(form_load.path);
                disableComboProfile = false;
            }
        }

        public void LoadProfile(string profile)
        {
            if (Directory.Exists(localPath + "\\Json\\" + pathToConfiguration))
            {
                SaveConfiguration_v2(false);
            }

            pathToConfiguration = profile;

            if (pathToConfiguration != defaultPathToConfiguration)
                this.Title = title + " " + version + " - File: " + pathToConfiguration;
            else
                this.Title = title + " " + version;

            richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["selectedProfile"] + " " + profile);

            // Se esiste una nuova versione del file di configurazione uso l'ultima, altrimenti carico il modello precedente
            if (File.Exists(localPath + "\\Json\\" + pathToConfiguration + "\\Config.json"))
            {
                LoadConfiguration_v2();
            }
            else
            {
                LoadConfiguration_v1();
            }

            if ((bool)radioButtonModeSerial.IsChecked)
            {
                buttonSerialActive.Focus();
            }
            else
            {
                buttonTcpActive.Focus();
            }
        }

        private void apriTemplateEditor_Click(object sender, RoutedEventArgs e)
        {
            if (importingProfile)
            {
                MessageBox.Show(lang.languageTemplate["strings"]["importingProfile"], "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                TemplateEditor window = new TemplateEditor(this);
                window.Show();
            }
        }

        private void caricaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            //SaveConfiguration_v1(false);
            SaveConfiguration_v2(false);

            // Se esiste una nuova versione del file di configurazione uso l'ultima, altrimenti carico il modello precedente
            if (File.Exists(localPath + "\\Json\\" + pathToConfiguration + "\\Config.json"))
            {
                LoadConfiguration_v2();
            }
            else
            {
                LoadConfiguration_v1();
            }
        }

        private void buttonLoopCoils01_Click(object sender, RoutedEventArgs e)
        {
            loopCoils01 = !loopCoils01;
            buttonLoopCoils01.Content = loopCoils01 ? "Stop" : "Loop";
            buttonLoopCoils01.Background = loopCoils01 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopCoilsRange_Click(object sender, RoutedEventArgs e)
        {
            loopCoilsRange = !loopCoilsRange;
            buttonLoopCoilsRange.Content = loopCoilsRange ? "Stop" : "Loop";
            buttonLoopCoilsRange.Background = loopCoilsRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopInput02_Click(object sender, RoutedEventArgs e)
        {
            loopInput02 = !loopInput02;
            buttonLoopInput02.Content = loopInput02 ? "Stop" : "Loop";
            buttonLoopInput02.Background = loopInput02 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopInputRange_Click(object sender, RoutedEventArgs e)
        {
            loopInputRange = !loopInputRange;
            buttonLoopInputRange.Content = loopInputRange ? "Stop" : "Loop";
            buttonLoopInputRange.Background = loopInputRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopInputRegister04_Click(object sender, RoutedEventArgs e)
        {
            loopInputRegister04 = !loopInputRegister04;
            buttonLoopInputRegister04.Content = loopInputRegister04 ? "Stop" : "Loop";
            buttonLoopInputRegister04.Background = loopInputRegister04 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopInputRegisterRange_Click(object sender, RoutedEventArgs e)
        {
            loopInputRegisterRange = !loopInputRegisterRange;
            buttonLoopInputRegisterRange.Content = loopInputRegisterRange ? "Stop" : "Loop";
            buttonLoopInputRegisterRange.Background = loopInputRegisterRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopHolding03_Click(object sender, RoutedEventArgs e)
        {
            loopHolding03 = !loopHolding03;
            buttonLoopHolding03.Content = loopHolding03 ? "Stop" : "Loop";
            buttonLoopHolding03.Background = loopHolding03 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        private void buttonLoopHoldingRange_Click(object sender, RoutedEventArgs e)
        {
            loopHoldingRange = !loopHoldingRange;
            buttonLoopHoldingRange.Content = loopHoldingRange ? "Stop" : "Loop";
            buttonLoopHoldingRange.Background = loopHoldingRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));
            checkLoop();
        }

        public void disableAllLoops()
        {
            // Fermo eventuali loop
            loopCoils01 = false;
            loopCoilsRange = false;
            loopInput02 = false;
            loopInputRange = false;
            loopInputRegister04 = false;
            loopInputRegisterRange = false;
            loopHolding03 = false;
            loopHoldingRange = false;

            buttonLoopCoils01.Content = loopCoils01 ? "Stop" : "Loop";
            buttonLoopCoils01.Background = loopCoils01 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));

            buttonLoopCoilsRange.Content = loopCoilsRange ? "Stop" : "Loop";
            buttonLoopCoilsRange.Background = loopCoilsRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));

            buttonLoopInput02.Content = loopInput02 ? "Stop" : "Loop";
            buttonLoopInput02.Background = loopInput02 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));

            buttonLoopInputRange.Content = loopInputRange ? "Stop" : "Loop";
            buttonLoopInputRange.Background = loopInputRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));

            buttonLoopInputRegister04.Content = loopInputRegister04 ? "Stop" : "Loop";
            buttonLoopInputRegister04.Background = loopInputRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));

            buttonLoopInputRegisterRange.Content = loopInputRegisterRange ? "Stop" : "Loop";
            buttonLoopInputRegisterRange.Background = loopInputRegisterRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));

            buttonLoopHolding03.Content = loopHolding03 ? "Stop" : "Loop";
            buttonLoopHolding03.Background = loopHolding03 ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));

            buttonLoopHoldingRange.Content = loopHoldingRange ? "Stop" : "Loop";
            buttonLoopHoldingRange.Background = loopHoldingRange ? Brushes.LightGreen : new SolidColorBrush(Color.FromArgb(0xFF, 0xDD, 0xDD, 0xDD));

            checkLoop();
        }

        public void loopPollingRegisters()
        {
            while (loopThreadRunning)
            {
                // Coils
                if (loopCoils01)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadCoils01_Click(null, null);
                    });
                }

                if (loopCoilsRange)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadCoilsRange_Click(null, null);
                    });
                }

                // Inputs
                if (loopInput02)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInput02_Click(null, null);
                    });
                }

                if (loopInputRange)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRange_Click(null, null);
                    });
                }

                // Input Registers
                if (loopInputRegister04)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRegister04_Click(null, null);
                    });
                }

                if (loopInputRegisterRange)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadInputRegisterRange_Click(null, null);
                    });
                }

                // Holding Registers
                if (loopHolding03)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadHolding03_Click(null, null);
                    });
                }

                if (loopHoldingRange)
                {
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        buttonReadHoldingRange_Click(null, null);
                    });
                }

                // Pausa tra una query e la successiva
                Thread.Sleep(pauseLoop);
            }
        }

        public bool checkLoopStartStop()
        {
            if (loopCoils01 || loopCoilsRange || loopInput02 || loopInputRange || loopInputRegister04 || loopInputRegisterRange || loopHolding03 || loopHoldingRange)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void checkLoop()
        {
            if (checkLoopStartStop() != loopThreadRunning)
            {
                // Thread già attivo, lo fermo
                if (loopThreadRunning)
                {
                    Console.WriteLine("Fermo il thread di polling");
                    loopThreadRunning = false;

                    Thread.Sleep(100);

                    try
                    {
                        if (threadLoopQuery.IsAlive)
                        {
                            Console.WriteLine("threadLoopQuery Abort");
                            threadLoopQuery.Abort();
                        }
                    }
                    catch (Exception err)
                    {
                        Console.WriteLine(err);
                    }


                }

                // Thread non attivo, lo avvio
                else
                {
                    Console.WriteLine("Avvio il thread di polling");
                    loopThreadRunning = true;
                    threadLoopQuery = new Thread(new ThreadStart(loopPollingRegisters));
                    threadLoopQuery.IsBackground = true;
                    threadLoopQuery.Priority = ThreadPriority.Normal;
                    threadLoopQuery.Start();
                }
            }
        }

        private void TextBoxPollingINterval_TextChanged(object sender, TextChangedEventArgs e)
        {
            int interval = 0;

            if (int.TryParse(TextBoxPollingInterval.Text, out interval))
            {
                if (interval >= 500)
                {
                    pauseLoop = interval;
                }
            }
        }

        private void buttonPingIp_Click(object sender, RoutedEventArgs e)
        {
            Ping p1 = new Ping();
            try
            {
                PingReply PR = p1.Send(textBoxTcpClientIpAddress.Text, 500);

                // check when the ping is not success
                if (!PR.Status.ToString().Equals("Success"))
                {
                    buttonPingIp.Background = Brushes.PaleVioletRed;

                    // Rimosso box per comodita
                    // DoEvents();
                    // MessageBox.Show("Ping failed", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                    richTextBoxAppend(richTextBoxStatus, "Ping failed");
                }
                else
                {
                    buttonPingIp.Background = Brushes.LightGreen;

                    // Rimosso box per comodita
                    // DoEvents();
                    // MessageBox.Show("Ping ok.\nResponse time: " + PR.RoundtripTime + "ms", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                    richTextBoxAppend(richTextBoxStatus, "Ping Ok - Response time: " + PR.RoundtripTime + "ms");
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err);

                richTextBoxAppend(richTextBoxStatus, "Network unreachable");
            }
        }

        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(delegate { }));
        }

        private void buttonClearHoldingReg_Click(object sender, RoutedEventArgs e)
        {
            list_holdingRegistersTable.Clear();
            //dataGridViewHolding.Items.Refresh();
        }

        private void buttonClearInputReg_Click(object sender, RoutedEventArgs e)
        {
            list_inputRegistersTable.Clear();
            //dataGridViewInputRegister.Items.Clear();
        }

        private void buttonClearInput_Click(object sender, RoutedEventArgs e)
        {
            list_inputsTable.Clear();
            //dataGridViewInput.Items.Clear();
        }

        private void buttonClearCoils_Click(object sender, RoutedEventArgs e)
        {
            list_coilsTable.Clear();
            //dataGridViewCoils.Items.Clear();
        }

        private void licenseToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            License window = new License();
            window.Show();
        }

        private void CheckBoxPinWIndow_Checked(object sender, RoutedEventArgs e)
        {
            this.Topmost = (bool)CheckBoxPinWIndow.IsChecked;
        }

        private void dataGridViewHolding_KeyUp(object sender, DataGridCellEditEndingEventArgs e)
        {
            if ((bool)CheckBoxSendValuesOnEditHoldingTable.IsChecked)
            {
                try
                {
                    var tmp = e.EditingElement as TextBox;
                    lastEditModbusItem = (ModBus_Item)e.Row.Item;

                    String[] params_ = new String[2];

                    params_[0] = tmp.Text;
                    params_[1] = e.Column.Header.ToString();

                    Thread t = new Thread(new ParameterizedThreadStart(writeRegisterDatagrid));
                    t.Priority = threadPriority;
                    t.Start(params_);
                }
                catch (Exception err)
                {
                    Console.WriteLine(err);
                }
            }

            disableDeleteKey = false;
        }
        public void writeRegisterDatagrid(object params_)
        {
            try
            {
                if (lastEditModbusItem != null)
                {
                    // Debug
                    Console.WriteLine("Register: " + lastEditModbusItem.Register);
                    //Console.WriteLine("Value: " + lastEditModbusItem.Value);
                    //Console.WriteLine("Notes: " + lastEditModbusItem.Notes);

                    String[] params__ = params_ as String[];

                    String obj = params__[0];
                    String columnName = params__[1];

                    Console.WriteLine("Value: " + obj);

                    // Converted Value (colonna valori convertiti, applico la conversione e invio le word corrispondenti)
                    if (columnName.ToLower().IndexOf("converted") != -1)
                    {
                        if (!lastEditModbusItem.ValueConverted.Equals(obj) || !sendCellEditOnlyOnChange)
                        {
                            // Estraggo il datatype
                            ModBus_Item tmp = list_template_holdingRegistersTable.FirstOrDefault(x => x.RegisterUInt == lastEditModbusItem.RegisterUInt);

                            String test = "";
                            int offset = 0;     // Offset starting word 32bit / 64bit variables

                            if (tmp != null)
                            {
                                if (tmp.Mappings != null)
                                {
                                    test = tmp.Mappings.Replace(" ", "");
                                }
                            }

                            if (obj.IndexOf(":") != -1)
                                obj = obj.Split(':')[1];

                            UInt16[] toSend = null;

                            // byte (low byte or high byte)
                            if (test.ToLower().IndexOf("byte") == 0)
                            {
                                toSend = new UInt16[1];
                                toSend[0] = (UInt16)(UInt16.Parse(obj) & 0x00FF);
                            }

                            // bitmap (type 1)
                            else if (test.IndexOf("b") == 0)
                            {
                                MessageBox.Show("Datatype not suppported", "Info", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }

                            // bytemap (type 2)
                            else if (test.IndexOf("B") == 0)
                            {
                                MessageBox.Show("Datatype not suppported", "Info", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }

                            // float (type 3)
                            else if (test.IndexOf("F") == 0 || test.ToLower().IndexOf("float") == 0)
                            {
                                byte[] conv = BitConverter.GetBytes(float.Parse(obj, System.Globalization.CultureInfo.InvariantCulture));
                                toSend = new UInt16[2];

                                if (test.ToLower().IndexOf("+") == -1)
                                    offset = -1;

                                if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                                {
                                    toSend[1] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[0] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                                else
                                {
                                    toSend[0] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[1] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                            }

                            // double (type 7)
                            else if (test.IndexOf("d") == 0 || test.ToLower().IndexOf("double") == 0)
                            {
                                byte[] conv = BitConverter.GetBytes(double.Parse(obj, System.Globalization.CultureInfo.InvariantCulture));
                                toSend = new UInt16[2];

                                if (test.ToLower().IndexOf("+") == -1)
                                    offset = -2;

                                if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                                {
                                    toSend[3] = (UInt16)((conv[7] << 8) + (conv[6]));
                                    toSend[2] = (UInt16)((conv[5] << 8) + (conv[4]));
                                    toSend[1] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[0] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                                else
                                {
                                    toSend[0] = (UInt16)((conv[7] << 8) + (conv[6]));
                                    toSend[1] = (UInt16)((conv[5] << 8) + (conv[4]));
                                    toSend[2] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[3] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                            }

                            // uint64 (type 6)
                            else if (test.ToLower().IndexOf("uint64") == 0)
                            {
                                byte[] conv = BitConverter.GetBytes(UInt64.Parse(obj));
                                toSend = new UInt16[4];

                                if (test.ToLower().IndexOf("+") == -1)
                                    offset = -3;

                                if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                                {
                                    toSend[3] = (UInt16)((conv[7] << 8) + (conv[6]));
                                    toSend[2] = (UInt16)((conv[5] << 8) + (conv[4]));
                                    toSend[1] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[0] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                                else
                                {
                                    toSend[0] = (UInt16)((conv[7] << 8) + (conv[6]));
                                    toSend[1] = (UInt16)((conv[5] << 8) + (conv[4]));
                                    toSend[2] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[3] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                            }

                            // int64 (type 6)
                            else if (test.ToLower().IndexOf("int64") == 0)
                            {
                                byte[] conv = BitConverter.GetBytes(Int64.Parse(obj));
                                toSend = new UInt16[4];

                                if (test.ToLower().IndexOf("+") == -1)
                                    offset = -3;

                                if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                                {
                                    toSend[3] = (UInt16)((conv[7] << 8) + (conv[6]));
                                    toSend[2] = (UInt16)((conv[5] << 8) + (conv[4]));
                                    toSend[1] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[0] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                                else
                                {
                                    toSend[0] = (UInt16)((conv[7] << 8) + (conv[6]));
                                    toSend[1] = (UInt16)((conv[5] << 8) + (conv[4]));
                                    toSend[2] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[3] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                            }

                            // uint32 (type 5)
                            else if (test.ToLower().IndexOf("uint32") == 0)
                            {
                                byte[] conv = BitConverter.GetBytes(UInt32.Parse(obj));
                                toSend = new UInt16[2];

                                if (test.ToLower().IndexOf("+") == -1)
                                    offset = -1;

                                if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                                {
                                    toSend[1] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[0] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                                else
                                {
                                    toSend[0] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[1] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                            }

                            // int32 (type 5)
                            else if (test.ToLower().IndexOf("int32") == 0)
                            {
                                byte[] conv = BitConverter.GetBytes(Int32.Parse(obj));
                                toSend = new UInt16[2];

                                if (test.ToLower().IndexOf("+") == -1)
                                    offset = -1;

                                if (test.ToLower().IndexOf("-") != -1 || test.ToLower().IndexOf("_swap") != -1)
                                {
                                    toSend[1] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[0] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                                else
                                {
                                    toSend[0] = (UInt16)((conv[3] << 8) + (conv[2]));
                                    toSend[1] = (UInt16)((conv[1] << 8) + (conv[0]));
                                }
                            }

                            // uint16 (type 4)
                            else if (test.ToLower().IndexOf("uint") == 0 || test.ToLower().IndexOf("uint16") == 0)
                            {
                                toSend = new UInt16[1];
                                toSend[0] = UInt16.Parse(obj);
                            }

                            // int16 (type 4)
                            else if (test.ToLower().IndexOf("int") == 0 || test.ToLower().IndexOf("int16") == 0)
                            {
                                byte[] conv = BitConverter.GetBytes(Int16.Parse(obj));
                                toSend = new UInt16[1];

                                toSend[0] = (UInt16)((conv[0] << 0) + (conv[1] << 8));
                            }

                            // String (type 255)
                            else if (test.ToLower().IndexOf("string") == 0)
                            {
                                MessageBox.Show("Datatype not suppported", "Info", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                            else
                            {
                                UInt16 dummy;

                                if (UInt16.TryParse(obj, out dummy))
                                {
                                    toSend = new UInt16[1];
                                    toSend[0] = dummy;
                                }
                            }

                            if (toSend != null)
                            {
                                long address_start = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(lastEditModbusItem.Register, comboBoxHoldingRegistri_) + offset;

                                if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                                {
                                    address_start = address_start - 40001;
                                }

                                uint value_ = P.uint_parser(lastEditModbusItem.Value, comboBoxHoldingValori_);

                                if (address_start >= 0)
                                {
                                    UInt16[] result = null;

                                    if (toSend.Length == 1)
                                    {
                                        if ((bool)ModBus.presetSingleRegister_06(byte.Parse(textBoxModbusAddress_), (UInt16)(address_start), toSend[0], readTimeout))
                                        {
                                            result = new UInt16[] { toSend[0] };
                                        }
                                    }
                                    else
                                    {
                                        result = ModBus.presetMultipleRegisters_16(byte.Parse(textBoxModbusAddress_), (UInt16)(address_start), toSend, readTimeout);
                                    }

                                    if (result != null)
                                    {
                                        for (int i = 0; i < result.Length; i++)
                                        {
                                            ModBus_Item select = list_holdingRegistersTable.FirstOrDefault(x => x.RegisterUInt == (UInt16)(address_start) + i);
                                            if (select != null)
                                            {
                                                this.Dispatcher.Invoke((Action)delegate
                                                {
                                                    select.Value = comboBoxHoldingValori_ == "HEX" ? "0x" + result[i].ToString("X").PadLeft(4, '0') : result[i].ToString();
                                                    select.ValueBin = Convert.ToString(result[i] >> 8, 2).PadLeft(8, '0') + " " + Convert.ToString((UInt16)(result[i] << 8) >> 8, 2).PadLeft(8, '0');
                                                    select.Foreground = ForeGroundLight.ToString();
                                                    select.Background = colorDefaultWriteCell.ToString();

                                                    ModBus_Item found = list_template_holdingRegistersTable.FirstOrDefault(x => x.RegisterUInt == (address_start + i));
                                                    if (found != null)
                                                    {
                                                        String convertedValue;
                                                        select.Mappings = GetMappingValue2(result, i, found.Mappings, out convertedValue);
                                                        select.ValueConverted = convertedValue;
                                                    }
                                                });
                                            }
                                        }
                                    }
                                    else
                                    {
                                        SetTableTimeoutError(list_holdingRegistersTable, true);
                                    }
                                }
                                else
                                {
                                    SetTableInternalError(list_holdingRegistersTable, true);
                                }
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            dataGridViewHolding.CommitEdit();   // Stackover consiglia due volte one evitare exception
                            dataGridViewHolding.CommitEdit();
                            dataGridViewHolding.Items.Refresh();
                            dataGridViewHolding.SelectedItem = lastEditModbusItem;
                        });
                    }

                    // Binary value
                    else if (columnName.ToLower().IndexOf("binary") != -1)
                    {
                        if (!lastEditModbusItem.ValueBin.Equals(obj) || !sendCellEditOnlyOnChange)
                        {
                            UInt16 dummy_ = Convert.ToUInt16(obj.Replace(" ", ""), 2); ;
                            uint address_start = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(lastEditModbusItem.Register, comboBoxHoldingRegistri_);

                            if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                            {
                                address_start = address_start - 40001;
                            }

                            bool? result = ModBus.presetSingleRegister_06(byte.Parse(textBoxModbusAddress_), address_start, dummy_, readTimeout);

                            if (result != null)
                            {
                                if (result == true)
                                {
                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        lastEditModbusItem.Value = comboBoxHoldingValori_.IndexOf("HEX") == -1 ? dummy_.ToString() : dummy_.ToString("x");
                                        lastEditModbusItem.ValueBin = Convert.ToString(dummy_ >> 8, 2).PadLeft(8, '0') + " " + Convert.ToString((UInt16)(dummy_ << 8) >> 8, 2).PadLeft(8, '0');
                                        lastEditModbusItem.Foreground = ForeGroundLight.ToString();
                                        lastEditModbusItem.Background = colorDefaultWriteCell.ToString();

                                        // Rimosso perche non e' possibile convertire valori > 1 word
                                        // ModBus_Item found = list_template_holdingRegistersTable.FirstOrDefault(x => x.RegisterUInt == (address_start));
                                        // if (found != null)
                                        // {
                                        //     lastEditModbusItem.Mappings = GetMappingValue2(new UInt16[] { dummy_ }, 0, found.Mappings, out string convertedValue);
                                        //     lastEditModbusItem.ValueConverted = convertedValue;
                                        // }

                                        lastEditModbusItem.Mappings = "";
                                        lastEditModbusItem.ValueConverted = "";
                                    });
                                }
                                else
                                {
                                    SetTableCrcError(list_holdingRegistersTable, true);
                                }
                            }
                            else
                            {
                                SetTableTimeoutError(list_holdingRegistersTable, true);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            dataGridViewHolding.CommitEdit();   // Stackover consiglia due volte one evitare exception
                            dataGridViewHolding.CommitEdit();
                            dataGridViewHolding.Items.Refresh();
                            dataGridViewHolding.SelectedItem = lastEditModbusItem;
                        });
                    }

                    // Standard value
                    else if (columnName.ToLower().IndexOf("value") != -1)
                    {
                        // Lavoro con int32, per lavorare in negativo
                        Int32 dummy_ = 0;

                        if (Int32.TryParse(obj.Replace("0x", "").Replace("x", "").Replace("h", ""), (obj.ToLower().IndexOf("x") != -1 || obj.ToLower().IndexOf("h") != -1 || comboBoxHoldingValori_ == "HEX") ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out dummy_))
                        {
                            if (dummy_ < -32768 || dummy_ > 0xFFFF)
                            {
                                MessageBox.Show(lang.languageTemplate["strings"]["valueNotValid"] + " " + obj, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                            }

                            if (!lastEditModbusItem.Value.Equals(obj) || !sendCellEditOnlyOnChange)
                            {
                                uint address_start = P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_) + P.uint_parser(lastEditModbusItem.Register, comboBoxHoldingRegistri_);

                                if (address_start > 9999 && correctModbusAddressAuto)    //Se indirizzo espresso in 30001+ imposto offset a 0
                                {
                                    address_start = address_start - 40001;
                                }

                                UInt16 value_ = (UInt16)dummy_;

                                bool? result = ModBus.presetSingleRegister_06(byte.Parse(textBoxModbusAddress_), address_start, value_, readTimeout);

                                if (result != null)
                                {
                                    if (result == true)
                                    {
                                        this.Dispatcher.Invoke((Action)delegate
                                        {
                                            lastEditModbusItem.Value = (obj.ToLower().IndexOf("x") != -1 || obj.ToLower().IndexOf("h") != -1 || comboBoxHoldingValori_ == "HEX") ? "0x" + value_.ToString("X").PadLeft(4, '0') : obj;
                                            lastEditModbusItem.ValueBin = Convert.ToString(value_ >> 8, 2).PadLeft(8, '0') + " " + Convert.ToString((UInt16)(value_ << 8) >> 8, 2).PadLeft(8, '0');
                                            lastEditModbusItem.Foreground = ForeGroundLight.ToString();
                                            lastEditModbusItem.Background = colorDefaultWriteCell.ToString();
                                        });
                                    }
                                    else
                                    {
                                        SetTableCrcError(list_holdingRegistersTable, true);
                                    }
                                }
                                else
                                {
                                    SetTableTimeoutError(list_holdingRegistersTable, true);
                                }
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            dataGridViewHolding.CommitEdit();   // Stackover consiglia due volte one evitare exception
                            dataGridViewHolding.CommitEdit();
                            dataGridViewHolding.Items.Refresh();
                            dataGridViewHolding.SelectedItem = lastEditModbusItem;
                        });
                    }
                }
            }
            catch (InvalidOperationException err)
            {
                if (err.Message.IndexOf("non-connected socket") != -1)
                {
                    SetTableDisconnectError(list_holdingRegistersTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonSerialActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonTcpActive_Click(null, null);
                        });
                    }
                }

                Console.WriteLine(err);
            }
            catch (ModbusException err)
            {
                if (err.Message.IndexOf("Timed out") != -1)
                {
                    SetTableTimeoutError(list_holdingRegistersTable, true);
                }
                if (err.Message.IndexOf("ModBus ErrCode") != -1)
                {
                    SetTableModBusError(list_holdingRegistersTable, err, true);
                }
                if (err.Message.IndexOf("ModbusProtocolError") != -1)
                {
                    SetTableStringError(list_holdingRegistersTable, (ModbusException)err, true);
                }
                if (err.Message.IndexOf("CRC Error") != -1)
                {
                    SetTableCrcError(list_holdingRegistersTable, true);
                }

                Console.WriteLine(err);

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteHolding06_b.IsEnabled = true;

                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });
            }
            catch (Exception err)
            {
                SetTableInternalError(list_holdingRegistersTable, true);
                Console.WriteLine(err);

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteHolding06_b.IsEnabled = true;

                    dataGridViewHolding.ItemsSource = null;
                    dataGridViewHolding.ItemsSource = list_holdingRegistersTable;
                });
            }
        }

        public void MenuItemLanguage_Click(object sender, EventArgs e)
        {
            var currMenuItem = (MenuItem)sender;

            language = currMenuItem.Header.ToString();

            // Passo fuori le lingue disponibili nel menu
            foreach (MenuItem tmp in languageToolStripMenu.Items)
            {
                tmp.IsChecked = tmp == currMenuItem;
            }

            // Carico template selezionato
            textBoxCurrentLanguage.Text = currMenuItem.Header.ToString();
            //lang.loadLanguageTemplate(currMenuItem.Header.ToString());
        }

        private void dataGridViewCoils_KeyUp(object sender, DataGridCellEditEndingEventArgs e)
        {
            if ((bool)CheckBoxSendValuesOnEditCoillsTable.IsChecked)
            {
                try
                {
                    var tmp = e.EditingElement as TextBox;
                    lastEditModbusItem = (ModBus_Item)e.Row.Item;

                    String[] params_ = new String[2];

                    params_[0] = tmp.Text;
                    params_[1] = e.Column.Header.ToString();

                    UInt16 out_;
                    if (UInt16.TryParse(tmp.Text, out out_))
                    {
                        Thread t = new Thread(new ParameterizedThreadStart(writeCoilDataGrid));
                        t.Priority = threadPriority;
                        t.Start(params_);
                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine(err);
                }
            }
        }

        public void writeCoilDataGrid(object params_)
        {
            try
            {
                if (lastEditModbusItem != null)
                {
                    String[] params__ = params_ as String[];

                    String obj = params__[0];
                    String columnName = params__[1];

                    // Converted Value (colonna valori convertiti, applico la conversione e invio le word corrispondenti)
                    if (columnName.ToLower().IndexOf("value") != -1)
                    {

                        // Debug
                        Console.WriteLine("Register: " + lastEditModbusItem.Register);
                        Console.WriteLine("Value: " + obj);
                        Console.WriteLine("Notes: " + lastEditModbusItem.Notes);

                        if (!lastEditModbusItem.Value.Equals(obj) || !sendCellEditOnlyOnChange)
                        {
                            UInt16 newValue = 0;

                            if (UInt16.TryParse(obj, out newValue))
                            {
                                uint address_start = P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_) + P.uint_parser(lastEditModbusItem.Register, comboBoxCoilsAddress05_);

                                bool? result = ModBus.forceSingleCoil_05(byte.Parse(textBoxModbusAddress_), address_start, newValue, readTimeout);

                                if (result != null)
                                {
                                    if (result == true)
                                    {
                                        this.Dispatcher.Invoke((Action)delegate
                                        {
                                            lastEditModbusItem.Foreground = ForeGroundLight.ToString();
                                            lastEditModbusItem.Background = colorDefaultWriteCell.ToString();

                                            /*if(index + 1 < dataGridViewCoils.Items.Count)
                                                dataGridViewCoils.SelectedItem = list_coilsTable[index + 1];*/

                                            dataGridViewCoils.Items.Refresh();
                                        });
                                    }
                                    else
                                    {
                                        SetTableCrcError(list_coilsTable, true);
                                    }
                                }
                                else
                                {
                                    SetTableTimeoutError(list_coilsTable, true);
                                }
                            }
                        }
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        dataGridViewCoils.CommitEdit();   // Stackover consiglia due volte one evitare exception
                        dataGridViewCoils.CommitEdit();
                        dataGridViewCoils.Items.Refresh();
                        dataGridViewCoils.SelectedItem = lastEditModbusItem;
                    });
                }
            }
            catch (InvalidOperationException err)
            {
                if (err.Message.IndexOf("non-connected socket") != -1)
                {
                    SetTableDisconnectError(list_coilsTable, true);

                    if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonSerialActive_Click(null, null);
                        });
                    }
                    else
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            buttonTcpActive_Click(null, null);
                        });
                    }
                }

                Console.WriteLine(err);
            }
            catch (ModbusException err)
            {
                if (err.Message.IndexOf("Timed out") != -1)
                {
                    SetTableTimeoutError(list_coilsTable, true);
                }
                if (err.Message.IndexOf("ModBus ErrCode") != -1)
                {
                    SetTableModBusError(list_coilsTable, err, true);
                }
                if (err.Message.IndexOf("ModbusProtocolError") != -1)
                {
                    SetTableStringError(list_coilsTable, (ModbusException)err, true);
                }
                if (err.Message.IndexOf("CRC Error") != -1)
                {
                    SetTableCrcError(list_coilsTable, true);
                }

                Console.WriteLine(err);

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteCoils05_B.IsEnabled = true;

                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });
            }
            catch (Exception err)
            {
                SetTableInternalError(list_coilsTable, true);

                Console.WriteLine(err);

                this.Dispatcher.Invoke((Action)delegate
                {
                    buttonWriteCoils05_B.IsEnabled = true;

                    dataGridViewCoils.ItemsSource = null;
                    dataGridViewCoils.ItemsSource = list_coilsTable;
                });
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {

            // Vincolato al ctrl
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                switch (e.Key)
                {
                    case Key.D1:
                        tabControlMain.SelectedIndex = 0;
                        break;

                    case Key.D2:
                        tabControlMain.SelectedIndex = 1;
                        break;

                    case Key.D3:
                        tabControlMain.SelectedIndex = 2;
                        break;

                    case Key.D4:
                        tabControlMain.SelectedIndex = 3;
                        break;

                    case Key.D5:
                        tabControlMain.SelectedIndex = 4;
                        break;

                    case Key.D6:
                        tabControlMain.SelectedIndex = 5;
                        break;

                    case Key.D7:
                        tabControlMain.SelectedIndex = 6;
                        break;

                    case Key.D8:
                        tabControlMain.SelectedIndex = 7;
                        break;

                    // Comandi read
                    case Key.R:

                        // Coils
                        if (tabControlMain.SelectedIndex == 1)
                        {
                            if (buttonReadCoils01.IsEnabled)
                                buttonReadCoils01_Click(sender, e);
                        }

                        // Inputs
                        if (tabControlMain.SelectedIndex == 2)
                        {
                            if (buttonReadInput02.IsEnabled)
                                buttonReadInput02_Click(sender, e);
                        }

                        // Holding registers
                        if (tabControlMain.SelectedIndex == 3)
                        {
                            if (buttonReadHolding03.IsEnabled)
                                buttonReadHolding03_Click(sender, e);
                        }

                        // Input registers
                        if (tabControlMain.SelectedIndex == 4)
                        {
                            if (buttonReadInputRegister04.IsEnabled)
                                buttonReadInputRegister04_Click(sender, e);
                        }

                        break;

                    // Comandi read all
                    case Key.E:

                        // Coils
                        if (tabControlMain.SelectedIndex == 1)
                        {
                            if (buttonReadCoilsRange.IsEnabled)
                                buttonReadCoilsRange_Click(sender, e);
                        }

                        // Inputs
                        if (tabControlMain.SelectedIndex == 2)
                        {
                            if (buttonReadInputRange.IsEnabled)
                                buttonReadInputRange_Click(sender, e);
                        }

                        // Holding registers
                        if (tabControlMain.SelectedIndex == 3)
                        {
                            if (buttonReadHoldingRange.IsEnabled)
                                buttonReadHoldingRange_Click(sender, e);
                        }

                        // Input registers
                        if (tabControlMain.SelectedIndex == 4)
                        {
                            if (buttonReadInputRegisterRange.IsEnabled)
                                buttonReadInputRegisterRange_Click(sender, e);
                        }

                        break;

                    // Comandi polling
                    case Key.P:

                        // Polling
                        if (tabControlMain.SelectedIndex == 0)
                        {
                            buttonPingIp_Click(sender, e);
                        }

                        // Coils
                        if (tabControlMain.SelectedIndex == 1)
                        {
                            if (buttonLoopCoils01.IsEnabled)
                                buttonLoopCoils01_Click(sender, e);
                        }

                        // Inputs
                        if (tabControlMain.SelectedIndex == 2)
                        {
                            if (buttonLoopInput02.IsEnabled)
                                buttonLoopInput02_Click(sender, e);
                        }

                        // Holding registers
                        if (tabControlMain.SelectedIndex == 3)
                        {
                            if (buttonLoopHolding03.IsEnabled)
                                buttonLoopHolding03_Click(sender, e);
                        }

                        // Input registers
                        if (tabControlMain.SelectedIndex == 4)
                        {
                            if (buttonLoopInputRegister04.IsEnabled)
                                buttonLoopInputRegister04_Click(sender, e);
                        }

                        break;

                    // Finestra statistiche
                    case Key.J:
                        statisticsToolStripMenuItem_Click(null, null);
                        break;

                    // Comandi polling
                    case Key.K:

                        // Coils
                        if (tabControlMain.SelectedIndex == 1)
                        {
                            if (buttonLoopCoilsRange.IsEnabled)
                                buttonLoopCoilsRange_Click(sender, e);
                        }

                        // Inputs
                        if (tabControlMain.SelectedIndex == 2)
                        {
                            if (buttonLoopInputRange.IsEnabled)
                                buttonLoopInputRange_Click(sender, e);
                        }

                        // Holding registers
                        if (tabControlMain.SelectedIndex == 3)
                        {
                            if (buttonLoopHoldingRange.IsEnabled)
                                buttonLoopHoldingRange_Click(sender, e);
                        }

                        // Input registers
                        if (tabControlMain.SelectedIndex == 4)
                        {
                            if (buttonLoopInputRegisterRange.IsEnabled)
                                buttonLoopInputRegisterRange_Click(sender, e);
                        }

                        break;

                    // Comandi read template group
                    case Key.G:

                        // Coils
                        if (tabControlMain.SelectedIndex == 1)
                        {
                            if (buttonReadCoilsTemplateGroup.IsEnabled)
                                buttonReadCoilsTemplateGroup_Click(sender, e);
                        }

                        // Inputs
                        if (tabControlMain.SelectedIndex == 2)
                        {
                            if (buttonReadInputTemplateGroup.IsEnabled)
                                buttonReadInputTemplateGroup_Click(sender, e);
                        }

                        // Holding registers
                        if (tabControlMain.SelectedIndex == 3)
                        {
                            if (buttonReadHoldingTemplateGroup.IsEnabled)
                                buttonReadHoldingTemplateGroup_Click(sender, e);
                        }

                        // Input registers
                        if (tabControlMain.SelectedIndex == 4)
                        {
                            if (buttonReadInputRegisterTemplateGroup.IsEnabled)
                                buttonReadInputRegisterTemplateGroup_Click(sender, e);
                        }

                        break;

                    // Comandi read template group
                    case Key.H:

                        // Coils
                        if (tabControlMain.SelectedIndex == 1)
                        {
                            if (buttonReadCoilsTemplate.IsEnabled)
                                buttonReadCoilsTemplate_Click(sender, e);
                        }

                        // Inputs
                        if (tabControlMain.SelectedIndex == 2)
                        {
                            if (buttonReadInputTemplate.IsEnabled)
                                buttonReadInputTemplate_Click(sender, e);
                        }

                        // Holding registers
                        if (tabControlMain.SelectedIndex == 3)
                        {
                            if (buttonReadHoldingTemplate.IsEnabled)
                                buttonReadHoldingTemplate_Click(sender, e);
                        }

                        // Input registers
                        if (tabControlMain.SelectedIndex == 4)
                        {
                            if (buttonReadInputRegisterTemplate.IsEnabled)
                                buttonReadInputRegisterTemplate_Click(sender, e);
                        }

                        break;

                    // Carica profilo
                    case Key.O:
                        caricaConfigurazioneDaDatabaseToolStripMenuItem_Click(sender, e);
                        break;

                    // Apri log
                    case Key.L:
                        logToolStripMenu_Click(sender, e);
                        break;

                    // DB Manager
                    case Key.D:
                        if (gestisciDatabaseToolStripMenuItem.IsEnabled)
                            gestisciDatabaseToolStripMenuItem_Click(sender, e);
                        break;

                    // Info
                    case Key.I:
                        infoToolStripMenuItem1_Click(sender, e);
                        break;

                    // Salva
                    case Key.S:

                        // Salva su database
                        if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
                        {
                            salvaConfigurazioneAttualeNelDatabaseToolStripMenuItem_Click(sender, e);
                        }
                        // Salva config. corrente
                        else
                        {
                            salvaToolStripMenuItem_Click(sender, e);
                        }

                        break;

                    // Mode
                    case Key.M:

                        if ((bool)radioButtonModeSerial.IsEnabled && (bool)radioButtonModeTcp.IsEnabled)
                        {
                            if ((bool)radioButtonModeSerial.IsChecked)
                            {
                                radioButtonModeTcp.IsChecked = true;
                            }
                            else
                            {
                                radioButtonModeSerial.IsChecked = true;
                            }
                        }
                        break;

                    // Connect
                    case Key.B:
                        if ((bool)radioButtonModeSerial.IsChecked)
                        {
                            buttonSerialActive_Click(sender, e);
                        }
                        else
                        {
                            buttonTcpActive_Click(sender, e);
                        }
                        break;

                    // Chiudi finestra
                    case Key.W:
                    case Key.Q:
                        this.Close();
                        break;

                    // Template
                    case Key.T:
                        apriTemplateEditor_Click(sender, e);
                        break;

                    // Left arrow
                    /*case Key.Left:
                        if(tabControlMain.SelectedIndex > 0)
                        {
                            tabControlMain.SelectedIndex = tabControlMain.SelectedIndex - 1;
                        }
                        break;

                    // Right arrow
                    case Key.Right:
                        if (tabControlMain.SelectedIndex < 7)
                        {
                            tabControlMain.SelectedIndex = tabControlMain.SelectedIndex + 1;
                        }
                        break;*/

                    case Key.F:
                        ViewFullSizeTables.IsChecked = !ViewFullSizeTables.IsChecked;
                        ViewFullSizeTables_Checked(null, null);
                        break;

                    // Comandi export table
                    case Key.U:

                        // Coils
                        if (tabControlMain.SelectedIndex == 1)
                        {
                            if (buttonExportCoils.IsEnabled)
                                buttonExportCoils_Click(sender, e);
                        }

                        // Inputs
                        if (tabControlMain.SelectedIndex == 2)
                        {
                            if (buttonExportInput.IsEnabled)
                                buttonExportInput_Click(sender, e);
                        }

                        // Holding registers
                        if (tabControlMain.SelectedIndex == 3)
                        {
                            if (buttonExportHoldingReg.IsEnabled)
                                buttonExportHoldingReg_Click(sender, e);
                        }

                        // Input registers
                        if (tabControlMain.SelectedIndex == 4)
                        {
                            if (buttonExportInputReg.IsEnabled)
                                buttonExportInputReg_Click(sender, e);
                        }

                        break;

                    // Comandi import table
                    case Key.Y:

                        // Coils
                        if (tabControlMain.SelectedIndex == 1)
                        {
                            if (buttonImportCoils.IsEnabled)
                                buttonImportCoils_Click(sender, e);
                        }

                        // Holding registers
                        if (tabControlMain.SelectedIndex == 3)
                        {
                            if (buttonImportHoldingReg.IsEnabled)
                                buttonImportHoldingReg_Click(sender, e);
                        }

                        break;

                }

                if (e.Key == Key.C && (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)))
                {
                    if (!statoConsole)
                    {
                        apriConsole();

                        this.Focus();
                    }
                    else
                    {
                        chiudiConsole();
                    }
                }

            }

            // Non vincolato al ctrl
            switch (e.Key)
            {
                // Cancella tabella
                case Key.Delete:

                    // Home
                    if (tabControlMain.SelectedIndex == 0)
                    {
                        buttonClearSerialStatus_Click(sender, e);
                    }

                    // Coils
                    if (tabControlMain.SelectedIndex == 1)
                    {
                        if (!disableDeleteKey)
                            buttonClearCoils_Click(sender, e);
                    }

                    // Inputs
                    if (tabControlMain.SelectedIndex == 2)
                    {
                        if (!disableDeleteKey)
                            buttonClearInput_Click(sender, e);
                    }

                    // Holding registers
                    if (tabControlMain.SelectedIndex == 3)
                    {
                        if (!disableDeleteKey)
                            buttonClearHoldingReg_Click(sender, e);
                    }

                    // Input registers
                    if (tabControlMain.SelectedIndex == 4)
                    {
                        if (!disableDeleteKey)
                            buttonClearInputReg_Click(sender, e);
                    }

                    // Log
                    if (tabControlMain.SelectedIndex == 6)
                    {
                        buttonClearAll_Click(sender, e);
                    }

                    break;
            }
        }

        private void logToolStripMenu_Click(object sender, RoutedEventArgs e)
        {
            if (!logWindowIsOpen)
            {
                logWindowIsOpen = true;
                LogView window = new LogView(this);
                window.Show();
            }
            else
            {
                MessageBox.Show(lang.languageTemplate["strings"]["logIsAlreadyOpen"], "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void textBoxHoldingAddress16_B_TextChanged(object sender, TextChangedEventArgs e)
        {
            string tmp = "";
            int stop = 0;

            System.Globalization.NumberStyles style = comboBoxHoldingAddress16_B_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;

            if (int.TryParse(textBoxHoldingAddress16_B.Text, style, null, out stop))
            {
                if (comboBoxHoldingValue16_ == "HEX")
                {
                    stop = stop * 2;
                }

                for (int i = 0; i < stop; i++)
                {
                    if (comboBoxHoldingValue16_ == "HEX")
                    {
                        tmp += "00";
                    }
                    else
                    {
                        tmp += "0";
                    }

                    if (i < (stop - 1))
                    {
                        tmp += " ";
                    }
                }

                textBoxHoldingValue16.Text = tmp;
            }

            textBoxHoldingAddress16_B_ = textBoxHoldingAddress16_B.Text;
        }

        public void LogDequeue()
        {
            while (!dequeueExit)
            {
                String content;

                if (this.ModBus != null)
                {
                    if (this.ModBus.log2.TryDequeue(out content))
                    {
                        richTextBoxPackets.Dispatcher.Invoke((Action)delegate
                        {
                            if (count_log > LogLimitRichTextBox)
                            {
                                // Arrivato al limite tolgo una riga ogni volta che aggiungo una riga
                                richTextBoxPackets.Document.Blocks.Remove(richTextBoxPackets.Document.Blocks.FirstBlock);
                            }
                            else
                            {
                                count_log += 1;
                            }

                            richTextBoxPackets.AppendText(content);
                        });

                        scrolled_log = false;
                    }
                    else
                    {
                        if (!scrolled_log)
                        {
                            richTextBoxPackets.Dispatcher.Invoke((Action)delegate
                            {
                                richTextBoxPackets.ScrollToEnd();
                            });

                            scrolled_log = true;
                        }


                        Thread.Sleep(100);
                    }
                }
                else
                {
                    Thread.Sleep(100);
                }
            }
        }

        private void TextBoxReadTimeout_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!int.TryParse(textBoxReadTimeout.Text, out readTimeout))
            {
                textBoxReadTimeout.Text = "1000";
            }
        }

        private void buttonExportHoldingReg_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                saveFileDialogBox = new SaveFileDialog();

                saveFileDialogBox.DefaultExt = "csv";
                saveFileDialogBox.AddExtension = false;
                if (lastHoldingRegistersCommandGroup && comboBoxHoldingGroup.SelectedItem != null)
                {
                    KeyValuePair<Group_Item, String> kp = (KeyValuePair<Group_Item, String>)comboBoxHoldingGroup.SelectedItem;
                    saveFileDialogBox.FileName = "" + pathToConfiguration + "_HoldingRegisters" + kp.Value.Split('-')[1].Replace(" ", "_");
                }
                else
                    saveFileDialogBox.FileName = "" + pathToConfiguration + "_HoldingRegisters";
                saveFileDialogBox.Filter = "CSV|*.csv|JSON|*.json";
                saveFileDialogBox.Title = title;

                if ((bool)saveFileDialogBox.ShowDialog())
                {
                    if (saveFileDialogBox.FileName.IndexOf("csv") != -1)
                    {
                        String file_content = "";

                        for (int i = 0; i < (list_holdingRegistersTable.Count); i++)
                        {
                            file_content += list_holdingRegistersTable[i].Offset + ",";
                            file_content += list_holdingRegistersTable[i].Register + ",";
                            file_content += list_holdingRegistersTable[i].Value + ",";
                            file_content += list_holdingRegistersTable[i].Notes + "\n";
                        }

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                    else
                    {
                        // File JSON
                        dataGridJson save = new dataGridJson();

                        save.items = new ModBusItem_Save[list_holdingRegistersTable.Count];

                        for (int i = 0; i < (list_holdingRegistersTable.Count); i++)
                        {
                            try
                            {
                                ModBusItem_Save item = new ModBusItem_Save();

                                item.Offset = list_holdingRegistersTable[i].Offset;
                                item.Register = list_holdingRegistersTable[i].Register;
                                item.Value = list_holdingRegistersTable[i].Value;
                                item.Notes = list_holdingRegistersTable[i].Notes;

                                save.items[i] = item;
                            }
                            catch { }
                        }


                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        jss.MaxJsonLength = this.MaxJsonLength;
                        string file_content = jss.Serialize(save);

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void buttonExportInputReg_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                saveFileDialogBox = new SaveFileDialog();

                saveFileDialogBox.DefaultExt = "csv";
                saveFileDialogBox.AddExtension = false;
                if (lastInputRegistersCommandGroup && comboBoxInputRegisterGroup.SelectedItem != null)
                {
                    KeyValuePair<Group_Item, String> kp = (KeyValuePair<Group_Item, String>)comboBoxInputRegisterGroup.SelectedItem;
                    saveFileDialogBox.FileName = "" + pathToConfiguration + "_InputRegisters" + kp.Value.Split('-')[1].Replace(" ", "_");
                }
                else
                    saveFileDialogBox.FileName = "" + pathToConfiguration + "_InputRegisters";
                saveFileDialogBox.Filter = "CSV|*.csv|JSON|*.json";
                saveFileDialogBox.Title = title;

                if ((bool)saveFileDialogBox.ShowDialog())
                {
                    if (saveFileDialogBox.FileName.IndexOf("csv") != -1)
                    {
                        String file_content = "";

                        for (int i = 0; i < (list_inputRegistersTable.Count); i++)
                        {
                            file_content += list_inputRegistersTable[i].Offset + ",";
                            file_content += list_inputRegistersTable[i].Register + ",";
                            file_content += list_inputRegistersTable[i].Value + ",";
                            file_content += list_inputRegistersTable[i].Notes + "\n";
                        }

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                    else
                    {
                        // File JSON
                        dataGridJson save = new dataGridJson();

                        save.items = new ModBusItem_Save[list_inputRegistersTable.Count];

                        for (int i = 0; i < (list_inputRegistersTable.Count); i++)
                        {
                            try
                            {
                                ModBusItem_Save item = new ModBusItem_Save();

                                item.Offset = list_inputRegistersTable[i].Offset;
                                item.Register = list_inputRegistersTable[i].Register;
                                item.Value = list_inputRegistersTable[i].Value;
                                item.Notes = list_inputRegistersTable[i].Notes;

                                save.items[i] = item;
                            }
                            catch { }
                        }


                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        jss.MaxJsonLength = this.MaxJsonLength;
                        string file_content = jss.Serialize(save);

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void buttonExportInput_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                saveFileDialogBox = new SaveFileDialog();

                saveFileDialogBox.DefaultExt = "csv";
                saveFileDialogBox.AddExtension = false;
                if (lastInputsCommandGroup && comboBoxInputGroup.SelectedItem != null)
                {
                    KeyValuePair<Group_Item, String> kp = (KeyValuePair<Group_Item, String>)comboBoxInputGroup.SelectedItem;
                    saveFileDialogBox.FileName = "" + pathToConfiguration + "_Inputs" + kp.Value.Split('-')[1].Replace(" ", "_");
                }
                else
                    saveFileDialogBox.FileName = "" + pathToConfiguration + "_Inputs";
                saveFileDialogBox.Filter = "CSV|*.csv|JSON|*.json";
                saveFileDialogBox.Title = title;

                if ((bool)saveFileDialogBox.ShowDialog())
                {
                    if (saveFileDialogBox.FileName.IndexOf("csv") != -1)
                    {
                        String file_content = "";

                        for (int i = 0; i < (list_inputsTable.Count); i++)
                        {
                            file_content += list_inputsTable[i].Offset + ",";
                            file_content += list_inputsTable[i].Register + ",";
                            file_content += list_inputsTable[i].Value + ",";
                            file_content += list_inputsTable[i].Notes + "\n";
                        }

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                    else
                    {
                        // File JSON
                        dataGridJson save = new dataGridJson();

                        save.items = new ModBusItem_Save[list_inputsTable.Count];

                        for (int i = 0; i < (list_inputsTable.Count); i++)
                        {
                            try
                            {
                                ModBusItem_Save item = new ModBusItem_Save();

                                item.Offset = list_inputsTable[i].Offset;
                                item.Register = list_inputsTable[i].Register;
                                item.Value = list_inputsTable[i].Value;
                                item.Notes = list_inputsTable[i].Notes;

                                save.items[i] = item;
                            }
                            catch { }
                        }


                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        jss.MaxJsonLength = this.MaxJsonLength;
                        string file_content = jss.Serialize(save);

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void buttonExportCoils_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                saveFileDialogBox = new SaveFileDialog();

                saveFileDialogBox.DefaultExt = "csv";
                saveFileDialogBox.AddExtension = false;
                if (lastCoilsCommandGroup && comboBoxCoilsGroup.SelectedItem != null)
                {
                    KeyValuePair<Group_Item, String> kp = (KeyValuePair<Group_Item, String>)comboBoxCoilsGroup.SelectedItem;
                    saveFileDialogBox.FileName = "" + pathToConfiguration + "_Coils" + kp.Value.Split('-')[1].Replace(" ", "_");
                }
                else
                    saveFileDialogBox.FileName = "" + pathToConfiguration + "_Coils";
                saveFileDialogBox.Filter = "CSV|*.csv|JSON|*.json";
                saveFileDialogBox.Title = title;

                if ((bool)saveFileDialogBox.ShowDialog())
                {
                    if (saveFileDialogBox.FileName.IndexOf("csv") != -1)
                    {
                        String file_content = "";

                        for (int i = 0; i < (list_coilsTable.Count); i++)
                        {
                            file_content += list_coilsTable[i].Offset + ",";
                            file_content += list_coilsTable[i].Register + ",";
                            file_content += list_coilsTable[i].Value + ",";
                            file_content += list_coilsTable[i].Notes + "\n";
                        }

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                    else
                    {
                        // File JSON
                        dataGridJson save = new dataGridJson();

                        save.items = new ModBusItem_Save[list_coilsTable.Count];

                        for (int i = 0; i < (list_coilsTable.Count); i++)
                        {
                            try
                            {
                                ModBusItem_Save item = new ModBusItem_Save();

                                item.Offset = list_coilsTable[i].Offset;
                                item.Register = list_coilsTable[i].Register;
                                item.Value = list_coilsTable[i].Value;
                                item.Notes = list_coilsTable[i].Notes;

                                save.items[i] = item;
                            }
                            catch { }
                        }


                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        jss.MaxJsonLength = this.MaxJsonLength;
                        string file_content = jss.Serialize(save);

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        // Test salvataggio configurazione con descrittore JSON
        public void SaveConfiguration_v2(bool alert)
        {
            JavaScriptSerializer jss = new JavaScriptSerializer();

            dynamic toSave = jss.DeserializeObject(File.ReadAllText(localPath + "\\Config\\SettingsToSave.json"));

            Dictionary<string, Dictionary<string, object>> file_ = new Dictionary<string, Dictionary<string, object>>();

            foreach (KeyValuePair<string, object> row in toSave["toSave"])
            {
                // row.key = "textBoxes"
                // row.value = {  }

                switch (row.Key)
                {
                    case "textBoxes":

                        Dictionary<string, object> textBoxes = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            // sub.key = "textBoxModbusAddess_1"
                            // sub.value = { "..." }

                            // debug
                            //Console.WriteLine("sub.key: " + sub.Key);
                            //Console.WriteLine("sub.value: " + sub.Value);

                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                // prop.key = "key"
                                // prop.value = "nomeVariabile"

                                // debug
                                //Console.WriteLine("prop.key: " + prop.Key);
                                //Console.WriteLine("prop.value: " + prop.Value as String);

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        textBoxes.Add(prop.Value as String, (this.FindName(sub.Key) as TextBox).Text);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    textBoxes.Add(sub.Key, (this.FindName(sub.Key) as TextBox).Text);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("textBoxes", textBoxes);
                        break;

                    case "checkBoxes":

                        Dictionary<string, object> checkBoxes = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        checkBoxes.Add(prop.Value as String, (bool)(this.FindName(sub.Key) as CheckBox).IsChecked);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    checkBoxes.Add(sub.Key, (bool)(this.FindName(sub.Key) as CheckBox).IsChecked);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("checkBoxes", checkBoxes);
                        break;

                    case "menuItems":

                        Dictionary<string, object> menuItems = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        menuItems.Add(prop.Value as String, (bool)(this.FindName(sub.Key) as MenuItem).IsChecked);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    menuItems.Add(sub.Key, (bool)(this.FindName(sub.Key) as MenuItem).IsChecked);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("menuItems", menuItems);
                        break;

                    case "radioButtons":

                        Dictionary<string, object> radioButtons = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        radioButtons.Add(prop.Value as String, (this.FindName(sub.Key) as RadioButton).IsChecked);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    radioButtons.Add(sub.Key, (this.FindName(sub.Key) as RadioButton).IsChecked);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("radioButtons", radioButtons);
                        break;

                    case "comboBoxes":

                        Dictionary<string, object> comboBoxes = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        comboBoxes.Add(prop.Value as String, (this.FindName(sub.Key) as ComboBox).SelectedIndex);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    comboBoxes.Add(sub.Key, (this.FindName(sub.Key) as ComboBox).SelectedIndex);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("comboBoxes", comboBoxes);
                        break;
                }
            }

            // Altre variabili custom
            Dictionary<string, object> others = new Dictionary<string, object>();

            others.Add("colorDefaultReadCell", colorDefaultReadCell.ToString());
            others.Add("colorDefaultWriteCell", colorDefaultWriteCell.ToString());
            others.Add("colorErrorCell", colorErrorCell.ToString());

            file_.Add("others", others);

            File.WriteAllText(localPath + "/Json/" + pathToConfiguration + "/Config.json", jss.Serialize(file_));

            if (alert)
            {
                MessageBox.Show(lang.languageTemplate["strings"]["infoSaveConfig"], "Info", MessageBoxButton.OK, MessageBoxImage.None);
            }
        }

        public void LoadConfiguration_v2()
        {
            JavaScriptSerializer jss = new JavaScriptSerializer();

            importingProfile = true;

            dynamic toSave = jss.DeserializeObject(File.ReadAllText(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config\\SettingsToSave.json"));
            dynamic loaded = jss.DeserializeObject(File.ReadAllText(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + pathToConfiguration + "\\Config.json"));

            foreach (KeyValuePair<string, object> row in toSave["toSave"])
            {
                // row.key = "textBoxes"
                // row.value = {  }

                switch (row.Key)
                {
                    case "textBoxes":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        // debug
                                        //Console.WriteLine(" -- ");
                                        //Console.WriteLine("sub.Key: " + sub.Key);
                                        //Console.WriteLine("prop.Value: " + prop.Value);
                                        //Console.WriteLine(loaded[row.Key][prop.Value.ToString()]);

                                        try
                                        {
                                            (this.FindName(sub.Key) as TextBox).Text = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as TextBox).Text = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "checkBoxes":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        try
                                        {
                                            (this.FindName(sub.Key) as CheckBox).IsChecked = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as CheckBox).IsChecked = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "menuItems":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        try
                                        {
                                            (this.FindName(sub.Key) as MenuItem).IsChecked = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as MenuItem).IsChecked = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "radioButtons":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        try
                                        {
                                            (this.FindName(sub.Key) as RadioButton).IsChecked = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as RadioButton).IsChecked = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "comboBoxes":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        try
                                        {
                                            (this.FindName(sub.Key) as ComboBox).SelectedIndex = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as ComboBox).SelectedIndex = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;
                }

                // Altre variabili custom
                BrushConverter bc = new BrushConverter();

                // Light Mode
                if (loaded["others"].ContainsKey("colorDefaultReadCell_Light"))
                    colorDefaultReadCell_Light = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorDefaultReadCell_Light"]);
                else
                    colorDefaultReadCell_Light = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorDefaultReadCell"]);

                if (loaded["others"].ContainsKey("colorDefaultWriteCell_Light"))
                    colorDefaultWriteCell_Light = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorDefaultWriteCell_Light"]);
                else
                    colorDefaultWriteCell_Light = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorDefaultWriteCell"]);

                if (loaded["others"].ContainsKey("colorErrorCell_Light"))
                    colorErrorCell_Light = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorErrorCell_Light"]);
                else
                    colorErrorCell_Light = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorErrorCell"]);

                // Dark Mode
                if (loaded["others"].ContainsKey("colorDefaultReadCell_Dark"))
                    colorDefaultReadCell_Dark = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorDefaultReadCell_Dark"]);
                else
                    colorDefaultReadCell_Dark = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorDefaultReadCell"]);

                if (loaded["others"].ContainsKey("colorDefaultWriteCell_Dark"))
                    colorDefaultWriteCell_Dark = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorDefaultWriteCell_Dark"]);
                else
                    colorDefaultWriteCell_Dark = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorDefaultWriteCell"]);

                if (loaded["others"].ContainsKey("colorErrorCell_Dark"))
                    colorErrorCell_Dark = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorErrorCell_Dark"]);
                else
                    colorErrorCell_Dark = (SolidColorBrush)bc.ConvertFromString(loaded["others"]["colorErrorCell"]);

                if (darkMode)
                {
                    labelColorCellRead.Background = colorDefaultReadCell_Dark;
                    labelColorCellWrote.Background = colorDefaultWriteCell_Dark;
                    labelColorCellError.Background = colorErrorCell_Dark;

                    colorDefaultReadCell = colorDefaultReadCell_Dark;
                    colorDefaultWriteCell = colorDefaultWriteCell_Dark;
                    colorErrorCell = colorErrorCell_Dark;

                    colorDefaultReadCellStr = colorDefaultReadCell.ToString();
                    colorDefaultWriteCellStr = colorDefaultWriteCell.ToString();
                    colorErrorCellStr = colorErrorCell.ToString();
                }
                else
                {
                    labelColorCellRead.Background = colorDefaultReadCell_Light;
                    labelColorCellWrote.Background = colorDefaultWriteCell_Light;
                    labelColorCellError.Background = colorErrorCell_Light;

                    colorDefaultReadCell = colorDefaultReadCell_Light;
                    colorDefaultWriteCell = colorDefaultWriteCell_Light;
                    colorErrorCell = colorErrorCell_Light;

                    colorDefaultReadCellStr = colorDefaultReadCell.ToString();
                    colorDefaultWriteCellStr = colorDefaultWriteCell.ToString();
                    colorErrorCellStr = colorErrorCell.ToString();
                }
            }

            // Al termine del caricamento della configurazione carico il template
            Thread t = new Thread(new ThreadStart(LoadTemplaeGroups));
            t.Start();
        }

        public void LoadTemplaeGroups()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                TaskbarItemInfo.ProgressValue = 0.01;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            string file_content = File.ReadAllText(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + pathToConfiguration + "\\Template.json");
            JavaScriptSerializer jss = new JavaScriptSerializer();
            jss.MaxJsonLength = this.MaxJsonLength;
            TEMPLATE template = jss.Deserialize<TEMPLATE>(file_content);

            list_template_coilsTable.Clear();
            list_template_inputsTable.Clear();
            list_template_inputRegistersTable.Clear();
            list_template_holdingRegistersTable.Clear();

            UInt16 tmp = 0;

            // Coils
            int template_coilsOffset = int.Parse(template.textBoxCoilsOffset_, template.comboBoxCoilsOffset_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

            // Inputs
            int template_inputsOffset = int.Parse(template.textBoxInputOffset_, template.comboBoxInputOffset_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

            // Input registers
            int template_inputRegistersOffset = int.Parse(template.textBoxInputRegOffset_, template.comboBoxInputRegOffset_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

            // Holding registers
            int template_HoldingOffset = int.Parse(template.textBoxHoldingOffset_, template.comboBoxHoldingOffset_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

            // Tabella coils
            for (int i = 0; i < template.dataGridViewCoils.Count(); i++)
            {
                if (UInt16.TryParse(template.dataGridViewCoils[i].Register, template.comboBoxCoilsRegistri_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out tmp))
                {
                    template.dataGridViewCoils[i].RegisterUInt = (UInt16)(tmp + template_coilsOffset);
                    template.dataGridViewCoils[i].Register = tmp.ToString();
                    list_template_coilsTable.Add(template.dataGridViewCoils[i]);
                }
            }

            // Tabella inputs
            for (int i = 0; i < template.dataGridViewInput.Count(); i++)
            {
                if (UInt16.TryParse(template.dataGridViewInput[i].Register, template.comboBoxInputRegistri_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out tmp))
                {
                    template.dataGridViewInput[i].RegisterUInt = (UInt16)(tmp + template_inputsOffset);
                    template.dataGridViewInput[i].Register = tmp.ToString();
                    list_template_inputsTable.Add(template.dataGridViewInput[i]);
                }
            }

            // Tabella input registers
            for (int i = 0; i < template.dataGridViewInputRegister.Count(); i++)
            {
                if (UInt16.TryParse(template.dataGridViewInputRegister[i].Register, template.comboBoxInputRegRegistri_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out tmp))
                {
                    template.dataGridViewInputRegister[i].RegisterUInt = (UInt16)(tmp + template_inputRegistersOffset);
                    template.dataGridViewInputRegister[i].Register = template.dataGridViewInputRegister[i].RegisterUInt.ToString();
                    list_template_inputRegistersTable.Add(template.dataGridViewInputRegister[i]);
                }
            }

            // Tabella holdings
            for (int i = 0; i < template.dataGridViewHolding.Count(); i++)
            {
                if (UInt16.TryParse(template.dataGridViewHolding[i].Register, template.comboBoxHoldingRegistri_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out tmp))
                {
                    template.dataGridViewHolding[i].RegisterUInt = (UInt16)(tmp + template_HoldingOffset);
                    template.dataGridViewHolding[i].Register = template.dataGridViewHolding[i].RegisterUInt.ToString();
                    list_template_holdingRegistersTable.Add(template.dataGridViewHolding[i]);
                }
            }

            // Tabella groups
            this.Dispatcher.Invoke((Action)delegate
            {
                comboBoxHoldingGroup.Items.Clear();
                comboBoxInputRegisterGroup.Items.Clear();
                comboBoxInputGroup.Items.Clear();
                comboBoxCoilsGroup.Items.Clear();
            });

            if (template.Groups != null)
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    comboBoxHoldingGroup.IsEnabled = template.Groups.Count() > 0;
                    comboBoxInputRegisterGroup.IsEnabled = template.Groups.Count() > 0;
                    comboBoxInputGroup.IsEnabled = template.Groups.Count() > 0;
                    comboBoxCoilsGroup.IsEnabled = template.Groups.Count() > 0;
                });

                if (template.Groups.Count() > 0)
                {
                    int counter = 0;

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["loadingProfile"]);
                    });

                    foreach (Group_Item gr in template.Groups.OrderBy(x => int.Parse(x.Group)))
                    {
                        try
                        {
                            KeyValuePair<Group_Item, String> kp = new KeyValuePair<Group_Item, String>(gr, gr.Group + " - " + gr.Label);

                            // Apparently fast but actually really slow when loading big templates
                            if (template.dataGridViewCoils.FirstOrDefault<ModBus_Item>(x => (x.Group != null ? x.Group : "").Split(';').FirstOrDefault<string>(y => String.Compare(y, kp.Key.Group) == 0) != null) != null)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    comboBoxCoilsGroup.Items.Add(kp);
                                });
                            }

                            if (template.dataGridViewInput.FirstOrDefault<ModBus_Item>(x => (x.Group != null ? x.Group : "").Split(';').FirstOrDefault<string>(y => String.Compare(y, kp.Key.Group) == 0) != null) != null)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    comboBoxInputGroup.Items.Add(kp);
                                });
                            }

                            if (template.dataGridViewInputRegister.FirstOrDefault<ModBus_Item>(x => (x.Group != null ? x.Group : "").Split(';').FirstOrDefault<string>(y => String.Compare(y, kp.Key.Group) == 0) != null) != null)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    comboBoxInputRegisterGroup.Items.Add(kp);
                                });
                            }

                            if (template.dataGridViewHolding.FirstOrDefault<ModBus_Item>(x => (x.Group != null ? x.Group : "").Split(';').FirstOrDefault<string>(y => String.Compare(y, kp.Key.Group) == 0) != null) != null)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    comboBoxHoldingGroup.Items.Add(kp);
                                });
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (((double)counter / (double)template.Groups.Count) > 0.01)
                                {
                                    TaskbarItemInfo.ProgressValue = (double)counter / (double)template.Groups.Count;
                                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
                                }
                            });

                            counter++;
                        }
                        catch
                        {
                            break;
                        }
                    }

                    try
                    {
                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = 0;
                            TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;

                            richTextBoxAppend(richTextBoxStatus, lang.languageTemplate["strings"]["loadedProfile"]);
                        });
                    }
                    catch
                    {
                    }
                }
            }
            else
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    comboBoxHoldingGroup.IsEnabled = false;
                    comboBoxInputRegisterGroup.IsEnabled = false;
                    comboBoxInputGroup.IsEnabled = false;
                    comboBoxCoilsGroup.IsEnabled = false;
                });
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                try
                {
                    TaskbarItemInfo.ProgressValue = 0;
                    TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                }
                catch
                {
                }
            });

            importingProfile = false;
        }

        public void changeColumnVisibility(object sender, RoutedEventArgs e)
        {
            DataGridTextColumn toEdit = (DataGridTextColumn)this.FindName((sender as MenuItem).Name.Replace("view", "dataGrid"));

            if (toEdit != null)
            {
                if ((sender as MenuItem).IsChecked)
                {
                    toEdit.Visibility = Visibility.Visible;
                }
                else
                {
                    toEdit.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void textBoxCoilsAddress15_B_TextChanged(object sender, TextChangedEventArgs e)
        {
            string tmp = "";
            int stop = 0;

            System.Globalization.NumberStyles style = comboBoxCoilsAddress15_B_ == "HEX" ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer;

            if (int.TryParse(textBoxCoilsAddress15_B.Text, style, null, out stop))
            {
                for (int i = 0; i < stop; i++)
                {
                    tmp += "0";
                }

                textBoxCoilsValue15.Text = tmp;
            }

            textBoxCoilsAddress15_B_ = textBoxCoilsAddress15_B.Text;
        }

        private void TextBoxCurrentLanguage_TextChanged(object sender, TextChangedEventArgs e)
        {
            language = textBoxCurrentLanguage.Text;

            lang.loadLanguageTemplate(language);
        }

        public void changeEnableButtonsConnect(bool enabled)
        {
            buttonReadCoils01.IsEnabled = enabled;
            buttonLoopCoils01.IsEnabled = enabled;
            buttonReadCoilsRange.IsEnabled = enabled;
            buttonLoopCoilsRange.IsEnabled = enabled;
            buttonWriteCoils05.IsEnabled = enabled;
            buttonWriteCoils05_B.IsEnabled = enabled;
            buttonWriteCoils15.IsEnabled = enabled;
            buttonImportCoils.IsEnabled = enabled;

            buttonReadInput02.IsEnabled = enabled;
            buttonLoopInput02.IsEnabled = enabled;
            buttonReadInputRange.IsEnabled = enabled;
            buttonLoopInputRange.IsEnabled = enabled;

            buttonReadInputRegister04.IsEnabled = enabled;
            buttonLoopInputRegister04.IsEnabled = enabled;
            buttonReadInputRegisterRange.IsEnabled = enabled;
            buttonLoopInputRegisterRange.IsEnabled = enabled;

            buttonReadHolding03.IsEnabled = enabled;
            buttonLoopHolding03.IsEnabled = enabled;
            buttonReadHoldingRange.IsEnabled = enabled;
            buttonLoopHoldingRange.IsEnabled = enabled;
            buttonWriteHolding06.IsEnabled = enabled;
            buttonWriteHolding06_b.IsEnabled = enabled;
            buttonWriteHolding16.IsEnabled = enabled;
            buttonImportHoldingReg.IsEnabled = enabled;

            buttonReadHoldingTemplateGroup.IsEnabled = enabled;
            buttonReadHoldingTemplate.IsEnabled = enabled;
            buttonReadInputRegisterTemplateGroup.IsEnabled = enabled;
            buttonReadInputRegisterTemplate.IsEnabled = enabled;
            buttonReadInputTemplateGroup.IsEnabled = enabled;
            buttonReadInputTemplate.IsEnabled = enabled;
            buttonReadCoilsTemplateGroup.IsEnabled = enabled;
            buttonReadCoilsTemplate.IsEnabled = enabled;

            buttonSendDiagnosticQuery.IsEnabled = enabled;
            buttonSendManualDiagnosticQuery.IsEnabled = enabled;

            comboBoxTcpConnectionMode.IsEnabled = !enabled;

            menuItemImportCoils.IsEnabled = enabled;
            menuItemImportHoldingRegisters.IsEnabled = enabled;
            menuItemImportCoilsClipboard.IsEnabled = enabled;
            menuItemImportHoldingRegistersClipboard.IsEnabled = enabled;

            comboBoxProfileHome.IsEnabled = !enabled;
        }

        private void comboBoxHoldingAddress03_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingAddress03_ = comboBoxHoldingAddress03.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxHoldingAddress03_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingAddress03_ = textBoxHoldingAddress03.Text;
        }

        private void textBoxHoldingRegisterNumber_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingRegisterNumber_ = textBoxHoldingRegisterNumber.Text;
        }

        private void comboBoxHoldingRange_A_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingRange_A_ = comboBoxHoldingRange_A.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxHoldingRange_A_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingRange_A_ = textBoxHoldingRange_A.Text;
        }

        /*private void comboBoxHoldingRange_B_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingRange_B_ = comboBoxHoldingRange_B.SelectedIndex == 0 ? "DEC" : "HEX";
        }*/

        private void textBoxHoldingRange_B_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingRange_B_ = textBoxHoldingRange_B.Text;
        }

        private void comboBoxHoldingAddress06_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingAddress06_ = comboBoxHoldingAddress06.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxHoldingAddress06_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingAddress06_ = textBoxHoldingAddress06.Text;
        }

        private void comboBoxHoldingValue06_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingValue06_ = comboBoxHoldingValue06.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxHoldingValue06_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingValue06_ = textBoxHoldingValue06.Text;
        }

        private void comboBoxHoldingAddress06_b_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingAddress06_b_ = comboBoxHoldingAddress06_b.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxHoldingAddress06_b_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingAddress06_b_ = textBoxHoldingAddress06_b.Text;
        }

        private void textBoxHoldingValue06_b_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingValue06_b_ = textBoxHoldingValue06_b.Text;
        }

        private void comboBoxHoldingAddress16_A_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingAddress16_A_ = comboBoxHoldingAddress16_A.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxHoldingAddress16_A_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingAddress16_A_ = textBoxHoldingAddress16_A.Text;
        }

        private void comboBoxHoldingAddress16_B_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingAddress16_B_ = comboBoxHoldingAddress16_B.SelectedIndex == 0 ? "DEC" : "HEX";

            textBoxHoldingAddress16_B_TextChanged(null, null);
        }

        private void comboBoxHoldingValue16_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingValue16_ = comboBoxHoldingValue16.SelectedIndex == 0 ? "DEC" : "HEX";

            textBoxHoldingAddress16_B_TextChanged(null, null);
        }

        private void textBoxHoldingValue16_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingValue16_ = textBoxHoldingValue16.Text;
        }

        private void comboBoxHoldingRegistri_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingRegistri_ = comboBoxHoldingRegistri.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void comboBoxHoldingValori_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingValori_ = comboBoxHoldingValori.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void comboBoxHoldingOffset_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxHoldingOffset_ = comboBoxHoldingOffset.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxHoldingOffset_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxHoldingOffset_ = textBoxHoldingOffset.Text;
        }

        private void textBoxModbusAddress_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxModbusAddress_ = textBoxModbusAddress.Text;
        }

        private void checkBoxCellColorMode_Checked(object sender, RoutedEventArgs e)
        {
            colorMode = (bool)checkBoxCellColorMode.IsChecked;
        }

        private void comboBoxCoilsRegistri_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsRegistri_ = comboBoxCoilsRegistri.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void comboBoxCoilsOffset_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsOffset_ = comboBoxCoilsOffset.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxCoilsOffset_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsOffset_ = textBoxCoilsOffset.Text;
        }

        private void comboBoxCoilsAddress01_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsAddress01_ = comboBoxCoilsAddress01.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxCoilsAddress01_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsAddress01_ = textBoxCoilsAddress01.Text;
        }

        private void textBoxCoilNumber_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilNumber_ = textBoxCoilNumber.Text;
        }

        private void comboBoxCoilsRange_A_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsRange_A_ = comboBoxCoilsRange_A.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxCoilsRange_A_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsRange_A_ = textBoxCoilsRange_A.Text;
        }

        /*private void comboBoxCoilsRange_B_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsRange_B_ = comboBoxCoilsRange_B.SelectedIndex == 0 ? "DEC" : "HEX";
        }*/

        private void textBoxCoilsRange_B_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsRange_B_ = textBoxCoilsRange_B.Text;
        }

        private void comboBoxCoilsAddress05_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsAddress05_ = comboBoxCoilsAddress05.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxCoilsAddress05_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsAddress05_ = textBoxCoilsAddress05.Text;
        }

        private void textBoxCoilsValue05_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsValue05_ = textBoxCoilsValue05.Text;
        }

        private void comboBoxCoilsAddress05_b_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsAddress05_b_ = comboBoxCoilsAddress05_b.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxCoilsAddress05_b_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsAddress05_b_ = textBoxCoilsAddress05_b.Text;
        }

        private void textBoxCoilsValue05_b_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsValue05_b_ = textBoxCoilsValue05_b.Text;
        }

        private void comboBoxCoilsAddress15_A_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsAddress15_A_ = comboBoxCoilsAddress15_A.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxCoilsAddress15_A_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsAddress15_A_ = textBoxCoilsAddress15_A.Text;
        }

        private void comboBoxCoilsAddress15_B_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxCoilsAddress15_B_ = comboBoxCoilsAddress15_B.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxCoilsValue15_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxCoilsValue15_ = textBoxCoilsValue15.Text;
        }

        private void comboBoxInputRegistri_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRegistri_ = comboBoxInputRegistri.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void comboBoxInputOffset_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputOffset_ = comboBoxInputOffset.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxInputOffset_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputOffset_ = textBoxInputOffset.Text;
        }

        private void comboBoxInputAddress02_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputAddress02_ = comboBoxInputAddress02.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxInputAddress02_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputAddress02_ = textBoxInputAddress02.Text;
        }

        private void textBoxInputNumber_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputNumber_ = textBoxInputNumber.Text;
        }

        private void comboBoxInputRange_A_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRange_A_ = comboBoxInputRange_A.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxInputRange_A_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputRange_A_ = textBoxInputRange_A.Text;
        }

        /*private void comboBoxInputRange_B_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRange_B_ = comboBoxInputRange_B.SelectedIndex == 0 ? "DEC" : "HEX";
        }*/

        private void textBoxInputRange_B_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputRange_B_ = textBoxInputRange_B.Text;
        }

        private void comboBoxInputRegRegistri_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRegRegistri_ = comboBoxInputRegRegistri.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void comboBoxInputRegValori_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRegValori_ = comboBoxInputRegValori.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void comboBoxInputRegOffset_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRegOffset_ = comboBoxInputRegOffset.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxInputRegOffset_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputRegOffset_ = textBoxInputRegOffset.Text;
        }

        private void comboBoxInputRegisterAddress04_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRegisterAddress04_ = comboBoxInputRegisterAddress04.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxInputRegisterAddress04_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputRegisterAddress04_ = textBoxInputRegisterAddress04.Text;
        }

        private void textBoxInputRegisterNumber_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputRegisterNumber_ = textBoxInputRegisterNumber.Text;
        }

        private void comboBoxInputRegisterRange_A_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRegisterRange_A_ = comboBoxInputRegisterRange_A.SelectedIndex == 0 ? "DEC" : "HEX";
        }

        private void textBoxInputRegisterRange_A_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputRegisterRange_A_ = textBoxInputRegisterRange_A.Text;
        }

        /*private void comboBoxInputRegisterRange_B_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            comboBoxInputRegisterRange_B_ = comboBoxInputRegisterRange_B.SelectedIndex == 0 ? "DEC" : "HEX";
        }*/

        private void textBoxInputRegisterRange_B_TextChanged(object sender, TextChangedEventArgs e)
        {
            textBoxInputRegisterRange_B_ = textBoxInputRegisterRange_B.Text;
        }

        private void checkBoxUseOffsetInTextBox_Checked(object sender, RoutedEventArgs e)
        {
            correctModbusAddressAuto = (bool)checkBoxUseOffsetInTextBox.IsChecked;
        }

        private void CheckBoxDarkMode_Checked(object sender, RoutedEventArgs e)
        {
            darkMode = (bool)CheckBoxDarkMode.IsChecked;

            if (darkMode)
            {
                colorDefaultReadCell = colorDefaultReadCell_Dark;
                colorDefaultWriteCell = colorDefaultWriteCell_Dark;
                colorErrorCell = colorErrorCell_Dark;

                colorDefaultReadCellStr = colorDefaultReadCell.ToString();
                colorDefaultWriteCellStr = colorDefaultWriteCell.ToString();
                colorErrorCellStr = colorErrorCell.ToString();
            }
            else
            {
                colorDefaultReadCell = colorDefaultReadCell_Light;
                colorDefaultWriteCell = colorDefaultWriteCell_Light;
                colorErrorCell = colorErrorCell_Light;

                colorDefaultReadCellStr = colorDefaultReadCell.ToString();
                colorDefaultWriteCellStr = colorDefaultWriteCell.ToString();
                colorErrorCellStr = colorErrorCell.ToString();
            }

            // BorderBrush
            Setter SetBorderBrush = new Setter();
            SetBorderBrush.Property = BorderBrushProperty;
            SetBorderBrush.Value = darkMode ? BackGroundDark : BackGroundLight2;

            // Background
            Setter SetBackgroundProperty = new Setter();
            SetBackgroundProperty.Property = BackgroundProperty;
            SetBackgroundProperty.Value = darkMode ? BackGroundDark : BackGroundLight2;

            // Stile custom per cella standard
            Style NewStyle = new Style();
            NewStyle.Setters.Add(SetBorderBrush);
            NewStyle.Setters.Add(SetBackgroundProperty);

            // Main
            this.Background = darkMode ? BackGroundDark2 : BackGroundLight2;
            ToolBarTrayMain.Background = darkMode ? BackGroundDark2 : Brushes.White;

            textBoxModbusAddress.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxModbusAddress.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            tabControlMain.Background = darkMode ? BackGroundDark2 : BackGroundLight2;
            GridConnection.Background = darkMode ? BackGroundDark : BackGroundLight;

            labelSerialRtu.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelPortRtu.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelBaudRate.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelParity.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelStopBits.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            labelTcp.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelIp.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelPortTcp.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            ToolBarMain.Background = darkMode ? BackGroundDark : new SolidColorBrush(Color.FromArgb(255, (byte)238, (byte)245, (byte)253));

            labelConnection.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            radioButtonModeSerial.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            radioButtonModeTcp.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelModBusAddress.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelRunning.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            CheckBoxPinWIndow.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelTx.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelRx.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelLoadProfile.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            checkBoxModbusSecure.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelClientCertificateName.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelClientPasswordName.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelClientTLSVersion.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelClientCertificatePath.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelClientKeyPath.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelClientKeyName.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            buttonPingIp.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonPingIp.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoadClientCertificate.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoadClientCertificate.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoadClientKey.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoadClientKey.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonSerialActive.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonSerialActive.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonTcpActive.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonTcpActive.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonClearSerialStatus.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonClearSerialStatus.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            buttonNewProfile.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonNewProfile.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonImportZip.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonImportZip.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonExportZip.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonExportZip.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            ButtonFullSize.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            ButtonFullSize.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            // Coils
            dataGridTabCoilsRegisters.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabCoilsValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabCoilsNotes.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabCoilsConvertedValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            dataGridTabCoilsRegisters.CellStyle = NewStyle;
            dataGridTabCoilsValues.CellStyle = NewStyle;
            dataGridTabCoilsNotes.CellStyle = NewStyle;
            dataGridTabCoilsConvertedValues.CellStyle = NewStyle;

            GridCoils.Background = darkMode ? BackGroundDark : BackGroundLight;

            labelCoils_0.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_1.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_2.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_5.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_8.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_13.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            CheckBoxSendValuesOnEditCoillsTable.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            StackPanelCoils0.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelCoils1.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelCoils2.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelCoils3.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelCoils4.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelCoils5.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));

            labelCoils_3.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_4.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_6.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_7.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_9.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_10.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_11.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_12.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_14.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_15.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelCoils_16.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            dataGridViewCoils.Background = darkMode ? BackGroundDark : BackGroundLight;

            buttonReadCoilsTemplateGroup.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadCoilsTemplateGroup.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadCoilsTemplate.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadCoilsTemplate.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadCoils01.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadCoils01.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoopCoils01.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoopCoils01.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadCoilsRange.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadCoilsRange.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoopCoilsRange.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoopCoilsRange.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonWriteCoils05.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonWriteCoils05.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonWriteCoils05_B.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonWriteCoils05_B.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonWriteCoils15.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonWriteCoils15.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonClearCoils.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonClearCoils.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonExportCoils.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonExportCoils.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonImportCoils.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonImportCoils.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            // Inputs
            dataGridTabInputsRegisters.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabInputsValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabInputsNotes.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabInputsConvertedValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            dataGridTabInputsRegisters.CellStyle = NewStyle;
            dataGridTabInputsValues.CellStyle = NewStyle;
            dataGridTabInputsNotes.CellStyle = NewStyle;
            dataGridTabInputsConvertedValues.CellStyle = NewStyle;

            GridInputs.Background = darkMode ? BackGroundDark : BackGroundLight;

            labelInputs_0.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputs_1.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputs_2.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputs_5.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            StackPanelInputs0.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelInputs1.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));

            labelInputs_3.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputs_4.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputs_6.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputs_7.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            dataGridViewInput.Background = darkMode ? BackGroundDark : BackGroundLight;

            buttonReadInputTemplateGroup.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadInputTemplateGroup.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadInputTemplate.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadInputTemplate.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadInput02.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadInput02.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoopInput02.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoopInput02.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadInputRange.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadInputRange.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoopInputRange.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoopInputRange.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonClearInput.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonClearInput.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonExportInput.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonExportInput.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            // Input register
            dataGridTabInputRegistersRegisters.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabInputRegistersValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabInputRegistersBinaryValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabInputRegistersConvertedValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabInputRegistersNotes.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabInputRegistersMappings.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            dataGridTabInputRegistersRegisters.CellStyle = NewStyle;
            dataGridTabInputRegistersValues.CellStyle = NewStyle;
            dataGridTabInputRegistersBinaryValues.CellStyle = NewStyle;
            dataGridTabInputRegistersConvertedValues.CellStyle = NewStyle;
            dataGridTabInputRegistersNotes.CellStyle = NewStyle;
            dataGridTabInputRegistersMappings.CellStyle = NewStyle;

            GridInputRegisters.Background = darkMode ? BackGroundDark : BackGroundLight;

            labelInputRegisters_0.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputRegisters_1.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputRegisters_2.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputRegisters_3.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputRegisters_6.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            StackPanelInputRegisters0.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelInputRegisters1.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));

            labelInputRegisters_4.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputRegisters_5.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputRegisters_7.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelInputRegisters_8.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            dataGridViewInputRegister.Background = darkMode ? BackGroundDark : BackGroundLight;

            buttonReadInputRegisterTemplateGroup.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadInputRegisterTemplateGroup.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadInputRegisterTemplate.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadInputRegisterTemplate.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadInputRegister04.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadInputRegister04.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoopInputRegister04.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoopInputRegister04.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadInputRegisterRange.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadInputRegisterRange.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoopInputRegisterRange.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoopInputRegisterRange.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonClearInputReg.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonClearInputReg.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonExportInputReg.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonExportInputReg.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            // Holding register
            dataGridTabHoldingRegistersRegisters.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabHoldingRegistersValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabHoldingRegistersBinaryValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabHoldingRegistersConvertedValues.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabHoldingRegistersNotes.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            dataGridTabHoldingRegistersMappings.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            dataGridTabHoldingRegistersRegisters.CellStyle = NewStyle;
            dataGridTabHoldingRegistersValues.CellStyle = NewStyle;
            dataGridTabHoldingRegistersBinaryValues.CellStyle = NewStyle;
            dataGridTabHoldingRegistersConvertedValues.CellStyle = NewStyle;
            dataGridTabHoldingRegistersNotes.CellStyle = NewStyle;
            dataGridTabHoldingRegistersMappings.CellStyle = NewStyle;

            GridHoldingRegisters.Background = darkMode ? BackGroundDark : BackGroundLight;

            labelHoldingRegisters_0.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_1.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_2.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_3.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_6.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_9.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_14.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            CheckBoxSendValuesOnEditHoldingTable.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            StackPanelHoldingRegisters0.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelHoldingRegisters1.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelHoldingRegisters2.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelHoldingRegisters3.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelHoldingRegisters4.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));
            StackPanelHoldingRegisters5.Background = darkMode ? BackGroundDark2 : new SolidColorBrush(Color.FromArgb(255, (byte)240, (byte)248, (byte)255));

            labelHoldingRegisters_4.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_5.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_7.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_8.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_10.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_11.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_12.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_13.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_15.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_16.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelHoldingRegisters_17.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            dataGridViewHolding.Background = darkMode ? BackGroundDark : BackGroundLight;

            buttonReadHoldingTemplateGroup.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadHoldingTemplateGroup.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadHoldingTemplate.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadHoldingTemplate.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadHolding03.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadHolding03.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoopHolding03.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoopHolding03.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonReadHoldingRange.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonReadHoldingRange.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonLoopHoldingRange.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonLoopHoldingRange.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonWriteHolding06.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonWriteHolding06.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonWriteHolding06_b.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonWriteHolding06_b.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonClearHoldingReg.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonClearHoldingReg.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonExportHoldingReg.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonExportHoldingReg.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonImportHoldingReg.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonImportHoldingReg.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            // Diagnostic
            GridDiagnostic.Background = darkMode ? BackGroundDark : BackGroundLight;

            labelDiagnostic_0.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_1.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_2.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_3.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_4.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_5.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_6.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_7.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_8.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_9.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelDiagnostic_10.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            textBoxDiagnosticResponse.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxDiagnosticFunctionManual.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxManualDiagnosticResponse.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            textBoxDiagnosticResponse.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxDiagnosticFunctionManual.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxManualDiagnosticResponse.Background = darkMode ? BackGroundDark : BackGroundLight2;

            buttonSendDiagnosticQuery.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonSendDiagnosticQuery.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonCalcCrcRtu.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonCalcCrcRtu.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonSendManualDiagnosticQuery.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonSendManualDiagnosticQuery.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            ButtonExampleTCP.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            ButtonExampleTCP.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            ButtonExampleRTU.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            ButtonExampleRTU.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            // Log
            GridLog.Background = darkMode ? BackGroundDark : BackGroundLight;

            checkBoxAddLinesToEnd.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            buttonExportSentPackets.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonExportSentPackets.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonClearSent.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonClearSent.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonClearAll.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonClearAll.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            // Settings
            GridSettings.Background = darkMode ? BackGroundDark : BackGroundLight;
            labelSettings_0.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelSettings_1.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelSettings_2.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            checkBoxUseOffsetInTextBox.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            checkBoxCellColorMode.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            checkBoxViewTableWithoutOffset.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            checkBoxCloseConsolAfterBoot.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            CheckBoxDarkMode.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            CheckBoxUseOnlyReadSingleRegistersForGroups.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelColorCellRead.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelColorCellWrote.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelColorCellError.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelSettings_3.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelSettings_4.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelSettings_5.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            labelSettings_6.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            buttonColorCellRead.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonColorCellRead.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonCellWrote.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonCellWrote.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            buttonColorCellError.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            buttonColorCellError.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            ButtonResetLightMode.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            ButtonResetLightMode.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;
            ButtonResetDarkMode.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            ButtonResetDarkMode.Background = darkMode ? BackGroundDarkButton : BackGroundLightButton;

            textBoxReadTimeout.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxReadTimeout.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            TextBoxPollingInterval.Background = darkMode ? BackGroundDark : BackGroundLight2;
            TextBoxPollingInterval.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;

            // Menu
            menuStrip.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            menuStrip.Background = darkMode ? BackGroundDark : BackGroundLight;

            // Controlli
            textBoxCoilsOffset.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsOffset.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsAddress01.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsAddress01.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilNumber.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilNumber.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsRange_A.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsRange_A.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsRange_B.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsRange_B.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsAddress05.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsAddress05.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsValue05.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsValue05.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsAddress05_b.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsAddress05_b.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsValue05_b.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsValue05_b.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsAddress15_A.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsAddress15_A.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsAddress15_B.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsAddress15_B.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxCoilsValue15.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxCoilsValue15.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxGoToCoilAddress.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxGoToCoilAddress.Background = darkMode ? BackGroundDark : BackGroundLight2;

            textBoxInputOffset.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputOffset.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxInputAddress02.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputAddress02.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxInputNumber.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputNumber.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxInputRange_A.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputRange_A.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxInputRange_B.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputRange_B.Background = darkMode ? BackGroundDark : BackGroundLight2;

            textBoxInputRegOffset.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputRegOffset.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxInputRegisterAddress04.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputRegisterAddress04.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxInputRegisterNumber.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputRegisterNumber.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxInputRegisterRange_A.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputRegisterRange_A.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxInputRegisterRange_B.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxInputRegisterRange_B.Background = darkMode ? BackGroundDark : BackGroundLight2;

            textBoxHoldingOffset.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingOffset.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingAddress03.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingAddress03.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingRegisterNumber.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingRegisterNumber.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingRange_A.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingRange_A.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingRange_B.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingRange_B.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingAddress06.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingAddress06.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingValue06.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingValue06.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingAddress06_b.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingAddress06_b.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingValue06_b.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingValue06_b.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingAddress16_A.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingAddress16_A.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingAddress16_B.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingAddress16_B.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxHoldingValue16.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxHoldingValue16.Background = darkMode ? BackGroundDark : BackGroundLight2;
            textBoxGoToHoldingAddress.Foreground = darkMode ? ForeGroundDark : ForeGroundLight;
            textBoxGoToHoldingAddress.Background = darkMode ? BackGroundDark : BackGroundLight2;

            // Cancello le raccolte delle tabelle
            list_holdingRegistersTable.Clear();
            list_inputRegistersTable.Clear();
            list_inputsTable.Clear();
            list_coilsTable.Clear();
        }

        private void ButtonResetLightMode_Click(object sender, RoutedEventArgs e)
        {
            colorDefaultReadCell_Light = new SolidColorBrush(Color.FromArgb(0xFF, 0x87, 0xCE, 0xFA));
            colorDefaultWriteCell_Light = Brushes.LightGreen;
            colorErrorCell_Light = Brushes.Orange;

            labelColorCellRead.Background = colorDefaultReadCell_Light;
            labelColorCellWrote.Background = colorDefaultWriteCell_Light;
            labelColorCellError.Background = colorErrorCell_Light;

            CheckBoxDarkMode.IsChecked = false;
        }

        private void ButtonResetDarkMode_Click(object sender, RoutedEventArgs e)
        {
            colorDefaultReadCell_Dark = new SolidColorBrush(Color.FromArgb(0xFF, 85, 85, 85));
            colorDefaultWriteCell_Dark = new SolidColorBrush(Color.FromArgb(0xFF, 0, 128, 0));
            colorErrorCell_Dark = Brushes.Orange;

            labelColorCellRead.Background = colorDefaultReadCell_Dark;
            labelColorCellWrote.Background = colorDefaultWriteCell_Dark;
            labelColorCellError.Background = colorErrorCell_Dark;

            CheckBoxDarkMode.IsChecked = true;
        }

        private void buttonImportCoils_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".txt";
            openFileDialog.AddExtension = false;
            openFileDialog.Filter = "CSV|*.csv|JSON|*.json";
            openFileDialog.Title = lang.languageTemplate["strings"]["openFileCoils"];
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();

            if (openFileDialog.FileNames.Length > 0)
            {
                PreviewImport previewImport = new PreviewImport(this, openFileDialog.FileNames, 5);
                previewImport.Show();
            }
        }

        private void buttonImportHoldingReg_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".txt";
            openFileDialog.AddExtension = false;
            openFileDialog.Filter = "CSV|*.csv|JSON|*.json";
            openFileDialog.Title = lang.languageTemplate["strings"]["openFileHoldingRegisters"];
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();

            if (openFileDialog.FileNames.Length > 0)
            {
                PreviewImport previewImport = new PreviewImport(this, openFileDialog.FileNames, 6);
                previewImport.Show();
            }
        }

        private void buttonReadHoldingTemplate_Click(object sender, RoutedEventArgs e)
        {
            buttonReadHoldingTemplate.IsEnabled = false;
            lastHoldingRegistersCommandGroup = false;

            Thread t = new Thread(new ThreadStart(readHoldingTemplateAll));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readHoldingTemplateAll()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                list_holdingRegistersTable.Clear();
                TaskbarItemInfo.ProgressValue = 0;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            List<ModBus_Item> toRead = new List<ModBus_Item>();
            List<uint> toRemove = new List<uint>();
            UInt32 counter = 0;

            if (useOnlyReadSingleRegisterForGroups)
            {
                foreach (ModBus_Item item in list_template_holdingRegistersTable.OrderBy(x => x.RegisterUInt))
                {
                    try
                    {
                        counter++;

                        if (item == null)
                            continue;

                        if (toRemove.Contains(item.RegisterUInt))
                            continue;

                        uint address_start = item.RegisterUInt;
                        uint num_regs = 1;

                        if (item.Mappings != null)
                        {
                            if (item.Mappings.ToLower().IndexOf("int32") != -1 || item.Mappings.ToLower().IndexOf("float") != -1)
                                num_regs = 2;

                            if (item.Mappings.ToLower().IndexOf("int64") != -1 || item.Mappings.ToLower().IndexOf("double") != -1)
                                num_regs = 4;

                            if (item.Mappings.ToLower().IndexOf("string") != -1)
                            {
                                int offset = 0;

                                if (item.Mappings.IndexOf('.') != -1)
                                    int.TryParse(item.Mappings.Split('.')[1].ToLower().Split(')')[0], out offset);
                                
                                address_start = (uint)(address_start + (offset / 2 + offset % 2));

                                num_regs = uint.Parse(item.Mappings.Split('.')[0].ToLower().Replace("string(", "").Replace(")", ""));
                                num_regs = num_regs / 2 + num_regs % 2;
                            }
                            else
                            {
                                if (item.Mappings.IndexOf("+") == -1)
                                    address_start = address_start - num_regs + 1;
                            }

                            if (num_regs > 1)
                            {
                                for (int i = 1; i < num_regs; i++)
                                {
                                    toRemove.Add((uint)(item.RegisterUInt + i));
                                }
                            }
                        }

                        UInt16[] response = ModBus.readHoldingRegister_03(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_holdingRegistersTable,
                                    list_template_holdingRegistersTable,
                                    P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxHoldingOffset_,
                                    comboBoxHoldingRegistri_,
                                    comboBoxHoldingValori_,
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_holdingRegistersTable.Count);
                        });
                    }
                    catch (Exception ex)
                    {
                        if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                        {
                            SetTableDisconnectError(list_holdingRegistersTable, true);

                            if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonSerialActive_Click(null, null);
                                });
                            }
                            else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonTcpActive_Click(null, null);
                                });
                            }
                        }
                        else if (ex is ModbusException)
                        {
                            if (ex.Message.IndexOf("Timed out") != -1)
                            {
                                SetTableTimeoutError(list_holdingRegistersTable, false);
                            }
                            if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, false);
                            }
                            if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                            {
                                SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                            }
                            if (ex.Message.IndexOf("CRC Error") != -1)
                            {
                                SetTableCrcError(list_holdingRegistersTable, false);
                            }
                        }
                        else
                        {
                            SetTableInternalError(list_holdingRegistersTable, false);
                        }

                        Console.WriteLine(ex);
                    }
                }
            }
            else
            {
                uint address_start = 0xFFFF;
                uint num_regs = 0;

                uint curr_start = 0xFFFF;
                uint curr_len = 0;

                try
                {
                    foreach (ModBus_Item item in list_template_holdingRegistersTable.OrderBy(x => x.RegisterUInt))
                    {
                        counter++;

                        if (item == null)
                            continue;

                        if (toRemove.Contains(item.RegisterUInt))
                            continue;

                        curr_start = item.RegisterUInt;
                        curr_len = 1;

                        if (item.Mappings != null)
                        {
                            if (item.Mappings.ToLower().IndexOf("int32") != -1 || item.Mappings.ToLower().IndexOf("float") != -1)
                                curr_len = 2;

                            if (item.Mappings.ToLower().IndexOf("int64") != -1 || item.Mappings.ToLower().IndexOf("double") != -1)
                                curr_len = 4;

                            if (item.Mappings.ToLower().IndexOf("string") != -1)
                            {
                                int offset = 0;

                                if (item.Mappings.IndexOf('.') != -1)
                                    int.TryParse(item.Mappings.Split('.')[1].ToLower().Split(')')[0], out offset);
                                
                                address_start = (uint)(address_start + (offset / 2 + offset % 2));

                                curr_len = uint.Parse(item.Mappings.Split('.')[0].ToLower().Replace("string(", "").Replace(")", ""));
                                curr_len = curr_len / 2 + curr_len % 2;
                            }
                            else
                            {
                                if (item.Mappings.IndexOf("+") == -1)
                                    curr_start = curr_start - curr_len + 1;
                            }

                            if (curr_len > 1)
                            {
                                for (int i = 1; i < curr_len; i++)
                                {
                                    toRemove.Add((uint)(item.RegisterUInt + i));
                                }
                            }
                        }

                        if (address_start != 0xFFFF && ((curr_start > address_start + num_regs) || (num_regs >= UInt16.Parse(textBoxHoldingRegisterNumber_))))
                        {
                            UInt16[] response = ModBus.readHoldingRegister_03(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_holdingRegistersTable,
                                        list_template_holdingRegistersTable,
                                        P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxHoldingOffset_,
                                        comboBoxHoldingRegistri_,
                                        comboBoxHoldingValori_,
                                        false,
                                        false);
                                }
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_holdingRegistersTable.Count);
                            });

                            address_start = 0xFFFF;
                            num_regs = 0;
                        }

                        num_regs += curr_len;

                        if (address_start == 0xFFFF)
                            address_start = curr_start;
                    }

                    if (num_regs == 0)
                        num_regs = curr_len;

                    if (address_start != 0xFFFF)
                    {
                        UInt16[] response = ModBus.readHoldingRegister_03(
                        byte.Parse(textBoxModbusAddress_),
                        address_start,
                        num_regs,
                        readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_holdingRegistersTable,
                                    list_template_holdingRegistersTable,
                                    P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxHoldingOffset_,
                                    comboBoxHoldingRegistri_,
                                    comboBoxHoldingValori_,
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_holdingRegistersTable.Count);
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                    {
                        SetTableDisconnectError(list_holdingRegistersTable, true);

                        if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonSerialActive_Click(null, null);
                            });
                        }
                        else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonTcpActive_Click(null, null);
                            });
                        }
                    }
                    else if (ex is ModbusException)
                    {
                        if (ex.Message.IndexOf("Timed out") != -1)
                        {
                            SetTableTimeoutError(list_holdingRegistersTable, false);
                        }
                        if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                        {
                            SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, false);
                        }
                        if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                        {
                            SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                        }
                        if (ex.Message.IndexOf("CRC Error") != -1)
                        {
                            SetTableCrcError(list_holdingRegistersTable, false);
                        }
                    }
                    else
                    {
                        SetTableInternalError(list_holdingRegistersTable, false);
                    }

                    Console.WriteLine(ex);
                }
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                buttonReadHoldingTemplate.IsEnabled = true;

                dataGridViewHolding.ItemsSource = null;
                dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
            });
        }

        private void buttonReadHoldingTemplateGroup_Click(object sender, RoutedEventArgs e)
        {
            if (comboBoxHoldingGroup.IsEnabled)
            {
                buttonReadHoldingTemplateGroup.IsEnabled = false;
                lastHoldingRegistersCommandGroup = true;

                Thread t = new Thread(new ThreadStart(readHoldingTemplateGroup));
                t.Priority = threadPriority;
                t.Start();
            }
        }

        public void readHoldingTemplateGroup()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                list_holdingRegistersTable.Clear();
                TaskbarItemInfo.ProgressValue = 0;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            List<ModBus_Item> toRead = new List<ModBus_Item>();
            List<uint> toRemove = new List<uint>();
            UInt32 counter = 0;

            foreach (ModBus_Item item in list_template_holdingRegistersTable.OrderBy(x => x.RegisterUInt))
            {
                try
                {
                    if (item == null)
                        continue;

                    if (item.Group == null && (comboBoxHoldingGroup_.Length > 1))
                        continue;

                    if (item.Group == null && (comboBoxHoldingGroup_.Length > 0))
                    {
                        if (int.Parse(comboBoxHoldingGroup_) != 0)
                            continue;
                    }

                    if (toRemove.Contains(item.RegisterUInt))
                        continue;

                    if (item.Group != null)
                    {
                        bool found = false;

                        foreach (string str in item.Group.Split(';'))
                        {
                            int dummy = -1;
                            if (int.TryParse(str, out dummy))
                            {
                                if (dummy == int.Parse(comboBoxHoldingGroup_))
                                {
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            continue;
                    }

                    toRead.Add(item);
                }
                catch (Exception err)
                {
                    Console.WriteLine(err);
                }
            }

            if (useOnlyReadSingleRegisterForGroups)
            {
                foreach (ModBus_Item item in toRead)
                {
                    try
                    {
                        uint address_start = item.RegisterUInt;
                        uint num_regs = 1;
                        counter++;

                        if (item.Mappings != null)
                        {
                            if (item.Mappings.ToLower().IndexOf("int32") != -1 || item.Mappings.ToLower().IndexOf("float") != -1)
                                num_regs = 2;

                            if (item.Mappings.ToLower().IndexOf("int64") != -1 || item.Mappings.ToLower().IndexOf("double") != -1)
                                num_regs = 4;

                            if (item.Mappings.ToLower().IndexOf("string") != -1)
                            {
                                int offset = 0;

                                if (item.Mappings.IndexOf('.') != -1)
                                    int.TryParse(item.Mappings.Split('.')[1].ToLower().Split(')')[0], out offset);
                                
                                address_start = (uint)(address_start + (offset / 2 + offset % 2));

                                num_regs = uint.Parse(item.Mappings.Split('.')[0].ToLower().Replace("string(", "").Replace(")", ""));
                                num_regs = num_regs / 2 + num_regs % 2;
                            }
                            else
                            {
                                if (item.Mappings.IndexOf("+") == -1)
                                    address_start = address_start - num_regs + 1;
                            }

                            if (num_regs > 1)
                            {
                                for (int i = 1; i < num_regs; i++)
                                {
                                    toRemove.Add((uint)(item.RegisterUInt + i));
                                }
                            }
                        }

                        UInt16[] response = ModBus.readHoldingRegister_03(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_holdingRegistersTable,
                                    list_template_holdingRegistersTable,
                                    P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxHoldingOffset_,
                                    comboBoxHoldingRegistri_,
                                    comboBoxHoldingValori_,
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                        });
                    }
                    catch (Exception ex)
                    {
                        if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                        {
                            SetTableDisconnectError(list_holdingRegistersTable, true);

                            if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonSerialActive_Click(null, null);
                                });
                            }
                            else
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonTcpActive_Click(null, null);
                                });
                            }
                        }
                        else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                        {
                            if (ex.Message.IndexOf("Timed out") != -1)
                            {
                                SetTableTimeoutError(list_holdingRegistersTable, false);
                            }
                            if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, false);
                            }
                            if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                            {
                                SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                            }
                            if (ex.Message.IndexOf("CRC Error") != -1)
                            {
                                SetTableCrcError(list_holdingRegistersTable, false);
                            }
                        }
                        else
                        {
                            SetTableInternalError(list_holdingRegistersTable, false);
                        }

                        Console.WriteLine(ex);
                    }
                }
            }
            else
            {
                uint address_start = 0xFFFF;
                uint num_regs = 0;

                uint curr_start = 0xFFFF;
                uint curr_len = 1;

                try
                {
                    foreach (ModBus_Item item in toRead)
                    {
                        if (item == null)
                            continue;

                        if (toRemove.Contains(item.RegisterUInt))
                            continue;

                        curr_start = item.RegisterUInt;
                        curr_len = 1;
                        counter++;

                        if (item.Mappings != null)
                        {
                            if (item.Mappings.ToLower().IndexOf("int32") != -1 || item.Mappings.ToLower().IndexOf("float") != -1)
                                curr_len = 2;

                            if (item.Mappings.ToLower().IndexOf("int64") != -1 || item.Mappings.ToLower().IndexOf("double") != -1)
                                curr_len = 4;

                            if (item.Mappings.ToLower().IndexOf("string") != -1)
                            {
                                int offset = 0;

                                if (item.Mappings.IndexOf('.') != -1)
                                    int.TryParse(item.Mappings.Split('.')[1].ToLower().Split(')')[0], out offset);

                                curr_start = (uint)(curr_start + (offset / 2 + offset % 2));

                                curr_len = uint.Parse(item.Mappings.Split('.')[0].ToLower().Replace("string(", "").Replace(")", ""));
                                curr_len = curr_len / 2 + curr_len % 2;
                            }
                            else
                            {
                                if (item.Mappings.IndexOf("+") == -1)
                                    curr_start = curr_start - curr_len + 1;
                            }

                            if (curr_len > 1)
                            {
                                for (int i = 1; i < curr_len; i++)
                                {
                                    toRemove.Add((uint)(curr_start + i));
                                }
                            }
                        }

                        if (address_start != 0xFFFF && ((curr_start > address_start + num_regs) || (num_regs >= UInt16.Parse(textBoxHoldingRegisterNumber_))))
                        {
                            UInt16[] response = ModBus.readHoldingRegister_03(
                                byte.Parse(textBoxModbusAddress_),
                                address_start,
                                num_regs,
                                readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_holdingRegistersTable,
                                        list_template_holdingRegistersTable,
                                        P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxHoldingOffset_,
                                        comboBoxHoldingRegistri_,
                                        comboBoxHoldingValori_,
                                        false,
                                        false);
                                }
                            }

                            address_start = 0xFFFF;
                            num_regs = 0;

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                            });
                        }

                        num_regs += curr_len;

                        if (address_start == 0xFFFF)
                            address_start = curr_start;
                    }

                    if (num_regs == 0)
                        num_regs = curr_len;

                    if (address_start != 0xFFFF)
                    {
                        UInt16[] response = ModBus.readHoldingRegister_03(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_holdingRegistersTable,
                                    list_template_holdingRegistersTable,
                                    P.uint_parser(textBoxHoldingOffset_, comboBoxHoldingOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxHoldingOffset_,
                                    comboBoxHoldingRegistri_,
                                    comboBoxHoldingValori_,
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                    {
                        SetTableDisconnectError(list_holdingRegistersTable, true);

                        if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonSerialActive_Click(null, null);
                            });
                        }
                        else
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonTcpActive_Click(null, null);
                            });
                        }
                    }
                    else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                    {
                        if (ex.Message.IndexOf("Timed out") != -1)
                        {
                            SetTableTimeoutError(list_holdingRegistersTable, false);
                        }
                        if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                        {
                            SetTableModBusError(list_holdingRegistersTable, (ModbusException)ex, false);
                        }
                        if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                        {
                            SetTableStringError(list_holdingRegistersTable, (ModbusException)ex, true);
                        }
                        if (ex.Message.IndexOf("CRC Error") != -1)
                        {
                            SetTableCrcError(list_holdingRegistersTable, false);
                        }
                    }
                    else
                    {
                        SetTableInternalError(list_holdingRegistersTable, false);
                    }

                    Console.WriteLine(ex);
                }
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                buttonReadHoldingTemplateGroup.IsEnabled = true;

                dataGridViewHolding.ItemsSource = null;
                dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
            });
        }

        private void comboBoxHoldingGroup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboBoxHoldingGroup.SelectedItem != null && buttonReadHoldingTemplateGroup.IsEnabled)
            {
                KeyValuePair<Group_Item, String> kp = (KeyValuePair<Group_Item, String>)comboBoxHoldingGroup.SelectedItem;
                comboBoxHoldingGroup_ = kp.Key.Group;

                buttonReadHoldingTemplateGroup_Click(null, null);
            }
        }

        private void buttonReadInputRegisterTemplate_Click(object sender, RoutedEventArgs e)
        {
            buttonReadInputRegisterTemplate.IsEnabled = false;
            lastInputRegistersCommandGroup = false;

            Thread t = new Thread(new ThreadStart(readInputRegisterTemplateAll));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readInputRegisterTemplateAll()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                list_inputRegistersTable.Clear();

                TaskbarItemInfo.ProgressValue = 0;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            UInt32 counter = 0;

            if (useOnlyReadSingleRegisterForGroups)
            {
                List<uint> toRemove = new List<uint>();
                foreach (ModBus_Item item in list_template_inputRegistersTable.OrderBy(x => x.RegisterUInt))
                {
                    try
                    {
                        counter++;

                        if (item == null)
                            continue;

                        if (toRemove.Contains(item.RegisterUInt))
                            continue;

                        uint address_start = item.RegisterUInt;
                        uint num_regs = 1;

                        if (item.Mappings != null)
                        {
                            if (item.Mappings.ToLower().IndexOf("int32") != -1 || item.Mappings.ToLower().IndexOf("float") != -1)
                                num_regs = 2;

                            if (item.Mappings.ToLower().IndexOf("int64") != -1 || item.Mappings.ToLower().IndexOf("double") != -1)
                                num_regs = 4;

                            if (item.Mappings.ToLower().IndexOf("string") != -1)
                            {
                                int offset = 0;

                                if (item.Mappings.IndexOf('.') != -1)
                                    int.TryParse(item.Mappings.Split('.')[1].ToLower().Split(')')[0], out offset);
                                
                                address_start = (uint)(address_start + (offset / 2 + offset % 2));

                                num_regs = uint.Parse(item.Mappings.Split('.')[0].ToLower().Replace("string(", "").Replace(")", ""));
                                num_regs = num_regs / 2 + num_regs % 2;
                            }
                            else
                            {
                                if (item.Mappings.IndexOf("+") == -1)
                                    address_start = address_start - num_regs + 1;
                            }

                            if (num_regs > 1)
                            {
                                for (int i = 1; i < num_regs; i++)
                                {
                                    toRemove.Add((uint)(item.RegisterUInt + i));
                                }
                            }
                        }

                        UInt16[] response = ModBus.readInputRegister_04(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_inputRegistersTable,
                                    list_template_inputRegistersTable,
                                    P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxInputRegOffset_,
                                    comboBoxInputRegRegistri_,
                                    comboBoxInputRegValori_,
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_inputRegistersTable.Count);
                        });
                    }
                    catch (Exception ex)
                    {
                        if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                        {
                            SetTableDisconnectError(list_inputRegistersTable, true);

                            if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    buttonSerialActive_Click(null, null);
                                });
                            }
                            else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    buttonTcpActive_Click(null, null);
                                });
                            }
                            else
                            {

                            }
                        }
                        else if (ex is ModbusException)
                        {
                            if (ex.Message.IndexOf("Timed out") != -1)
                            {
                                SetTableTimeoutError(list_inputRegistersTable, false);
                            }
                            if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                SetTableModBusError(list_inputRegistersTable, (ModbusException)ex, false);
                            }
                            if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                            {
                                SetTableStringError(list_inputRegistersTable, (ModbusException)ex, true);
                            }
                            if (ex.Message.IndexOf("CRC Error") != -1)
                            {
                                SetTableCrcError(list_inputRegistersTable, false);
                            }
                        }
                        else
                        {
                            SetTableInternalError(list_inputRegistersTable, false);
                        }

                        Console.WriteLine(ex);
                    }
                }
            }
            else
            {
                uint address_start = 0xFFFF;
                uint num_regs = 0;

                uint curr_start = 0xFFFF;
                uint curr_len = 1;

                try
                {
                    List<uint> toRemove = new List<uint>();
                    foreach (ModBus_Item item in list_template_inputRegistersTable.OrderBy(x => x.RegisterUInt))
                    {
                        counter++;

                        if (item == null)
                            continue;

                        if (toRemove.Contains(item.RegisterUInt))
                            continue;

                        curr_start = item.RegisterUInt;
                        curr_len = 1;

                        if (item.Mappings != null)
                        {
                            if (item.Mappings.ToLower().IndexOf("int32") != -1 || item.Mappings.ToLower().IndexOf("float") != -1)
                                curr_len = 2;

                            if (item.Mappings.ToLower().IndexOf("int64") != -1 || item.Mappings.ToLower().IndexOf("double") != -1)
                                curr_len = 4;

                            if (item.Mappings.ToLower().IndexOf("string") != -1)
                            {
                                int offset = 0;

                                if (item.Mappings.IndexOf('.') != -1)
                                    int.TryParse(item.Mappings.Split('.')[1].ToLower().Split(')')[0], out offset);

                                curr_start = (uint)(curr_start + (offset / 2 + offset % 2));

                                curr_len = uint.Parse(item.Mappings.Split('.')[0].ToLower().Replace("string(", "").Replace(")", ""));
                                curr_len = curr_len / 2 + curr_len % 2;
                            }
                            else
                            {
                                if (item.Mappings.IndexOf("+") == -1)
                                    curr_start = curr_start - curr_len + 1;
                            }

                            if (curr_len > 1)
                            {
                                for (int i = 1; i < curr_len; i++)
                                {
                                    toRemove.Add((uint)(item.RegisterUInt + i));
                                }
                            }
                        }

                        if (address_start != 0xFFFF && ((curr_start > address_start + num_regs) || (num_regs >= UInt16.Parse(textBoxInputRegisterNumber_))))
                        {
                            UInt16[] response = ModBus.readInputRegister_04(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_inputRegistersTable,
                                        list_template_inputRegistersTable,
                                        P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxInputRegOffset_,
                                        comboBoxInputRegRegistri_,
                                        comboBoxInputRegValori_,
                                        false,
                                        false);
                                }
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_inputRegistersTable.Count);
                            });

                            address_start = 0xFFFF;
                            num_regs = 0;
                        }

                        num_regs += curr_len;

                        if (address_start == 0xFFFF)
                            address_start = curr_start;
                    }

                    if (num_regs == 0)
                        num_regs = curr_len;

                    if (address_start != 0xFFFF)
                    {
                        UInt16[] response = ModBus.readInputRegister_04(
                        byte.Parse(textBoxModbusAddress_),
                        address_start,
                        num_regs,
                        readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_inputRegistersTable,
                                    list_template_inputRegistersTable,
                                    P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxInputRegOffset_,
                                    comboBoxInputRegRegistri_,
                                    comboBoxInputRegValori_,
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_inputRegistersTable.Count);
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                    {
                        SetTableDisconnectError(list_inputRegistersTable, true);

                        if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                buttonSerialActive_Click(null, null);
                            });
                        }
                        else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                buttonTcpActive_Click(null, null);
                            });
                        }
                        else
                        {

                        }
                    }
                    else if (ex is ModbusException)
                    {
                        if (ex.Message.IndexOf("Timed out") != -1)
                        {
                            SetTableTimeoutError(list_inputRegistersTable, false);
                        }
                        if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                        {
                            SetTableModBusError(list_inputRegistersTable, (ModbusException)ex, false);
                        }
                        if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                        {
                            SetTableStringError(list_inputRegistersTable, (ModbusException)ex, true);
                        }
                        if (ex.Message.IndexOf("CRC Error") != -1)
                        {
                            SetTableCrcError(list_inputRegistersTable, false);
                        }
                    }
                    else
                    {
                        SetTableInternalError(list_inputRegistersTable, false);
                    }

                    Console.WriteLine(ex);
                }
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                buttonReadInputRegisterTemplate.IsEnabled = true;

                dataGridViewInputRegister.ItemsSource = null;
                dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;

                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
            });
        }

        private void buttonReadInputRegisterTemplateGroup_Click(object sender, RoutedEventArgs e)
        {
            buttonReadInputRegisterTemplateGroup.IsEnabled = false;
            lastInputRegistersCommandGroup = true;

            Thread t = new Thread(new ThreadStart(readInputRegisterTemplateGroup));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readInputRegisterTemplateGroup()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                list_inputRegistersTable.Clear();

                TaskbarItemInfo.ProgressValue = 0;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            List<ModBus_Item> toRead = new List<ModBus_Item>();
            List<uint> toRemove = new List<uint>();
            UInt32 counter = 0;

            foreach (ModBus_Item item in list_template_inputRegistersTable.OrderBy(x => x.RegisterUInt).ToList<ModBus_Item>())
            {
                try
                {
                    if (item == null)
                        continue;

                    if (item.Group == null && (comboBoxInputRegisterGroup_.Length > 1))
                        continue;

                    if (item.Group == null && (comboBoxInputRegisterGroup_.Length > 0))
                    {
                        if (int.Parse(comboBoxInputRegisterGroup_) != 0)
                            continue;
                    }

                    if (toRemove.Contains(item.RegisterUInt))
                        continue;

                    if (item.Group != null)
                    {
                        bool found = false;

                        foreach (string str in item.Group.Split(';'))
                        {
                            int dummy = -1;
                            if (int.TryParse(str, out dummy))
                            {
                                if (dummy == int.Parse(comboBoxInputRegisterGroup_))
                                {
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            continue;
                    }

                    toRead.Add(item);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }

            if (useOnlyReadSingleRegisterForGroups)
            {
                foreach (ModBus_Item item in toRead)
                {
                    try
                    {
                        uint address_start = item.RegisterUInt;
                        uint num_regs = 1;
                        counter++;

                        if (item.Mappings != null)
                        {
                            if (item.Mappings.ToLower().IndexOf("int32") != -1 || item.Mappings.ToLower().IndexOf("float") != -1)
                                num_regs = 2;

                            if (item.Mappings.ToLower().IndexOf("int64") != -1 || item.Mappings.ToLower().IndexOf("double") != -1)
                                num_regs = 4;

                            if (item.Mappings.ToLower().IndexOf("string") != -1)
                            {
                                int offset = 0;

                                if (item.Mappings.IndexOf('.') != -1)
                                    int.TryParse(item.Mappings.Split('.')[1].ToLower().Split(')')[0], out offset);
                                
                                address_start = (uint)(address_start + (offset / 2 + offset % 2));

                                num_regs = uint.Parse(item.Mappings.Split('.')[0].ToLower().Replace("string(", "").Replace(")", ""));
                                num_regs = num_regs / 2 + num_regs % 2;
                            }
                            else
                            {
                                if (item.Mappings.IndexOf("+") == -1)
                                    address_start = address_start - num_regs + 1;
                            }

                            if (num_regs > 1)
                            {
                                for (int i = 1; i < num_regs; i++)
                                {
                                    toRemove.Add((uint)(item.RegisterUInt + i));
                                }
                            }
                        }

                        UInt16[] response = ModBus.readInputRegister_04(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_inputRegistersTable,
                                    list_template_inputRegistersTable,
                                    P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxInputRegOffset_,
                                    comboBoxInputRegRegistri_,
                                    comboBoxInputRegValori_,
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                        });
                    }
                    catch (Exception ex)
                    {
                        if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                        {
                            SetTableDisconnectError(list_inputRegistersTable, true);

                            if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    buttonSerialActive_Click(null, null);
                                });
                            }
                            else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    buttonTcpActive_Click(null, null);
                                });
                            }
                            else
                            {

                            }
                        }
                        else if (ex is ModbusException)
                        {
                            if (ex.Message.IndexOf("Timed out") != -1)
                            {
                                SetTableTimeoutError(list_inputRegistersTable, false);
                            }
                            if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                SetTableModBusError(list_inputRegistersTable, (ModbusException)ex, false);
                            }
                            if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                            {
                                SetTableStringError(list_inputRegistersTable, (ModbusException)ex, true);
                            }
                            if (ex.Message.IndexOf("CRC Error") != -1)
                            {
                                SetTableCrcError(list_inputRegistersTable, false);
                            }
                        }
                        else
                        {
                            SetTableInternalError(list_inputRegistersTable, false);
                        }

                        Console.WriteLine(ex);
                    }
                }
            }
            else
            {
                uint address_start = 0xFFFF;
                uint num_regs = 0;

                uint curr_start = 0xFFFF;
                uint curr_len = 1;

                try
                {
                    foreach (ModBus_Item item in toRead)
                    {
                        Console.WriteLine("curr_start {0}, item.RegisterUInt: {1}", curr_start, item.RegisterUInt);

                        if (item == null)
                            continue;

                        if (toRemove.Contains(item.RegisterUInt))
                            continue;
                        
                        curr_start = item.RegisterUInt;
                        curr_len = 1;
                        counter++;

                        if (item.Mappings != null)
                        {
                            if (item.Mappings.ToLower().IndexOf("int32") != -1 || item.Mappings.ToLower().IndexOf("float") != -1)
                                curr_len = 2;

                            if (item.Mappings.ToLower().IndexOf("int64") != -1 || item.Mappings.ToLower().IndexOf("double") != -1)
                                curr_len = 4;

                            if (item.Mappings.ToLower().IndexOf("string") != -1)
                            {
                                int offset = 0;

                                if (item.Mappings.IndexOf('.') != -1)
                                    int.TryParse(item.Mappings.Split('.')[1].ToLower().Split(')')[0], out offset);

                                curr_start = (uint)(curr_start + (offset / 2 + offset % 2));

                                curr_len = uint.Parse(item.Mappings.Split('.')[0].ToLower().Replace("string(", "").Replace(")", ""));
                                curr_len = curr_len / 2 + curr_len % 2;
                            }
                            else
                            {
                                if (item.Mappings.IndexOf("+") == -1)
                                    curr_start = curr_start - curr_len + 1;
                            }

                            Console.WriteLine("curr_start {0}, curr_len: {1}", curr_start, curr_len);

                            if (curr_len > 1)
                            {
                                for (int i = 1; i < curr_len; i++)
                                {
                                    toRemove.Add((uint)(item.RegisterUInt + i));
                                }
                            }
                        }

                        if (address_start != 0xFFFF && ((curr_start > address_start + num_regs) || (num_regs >= UInt16.Parse(textBoxInputRegisterNumber_))))
                        {
                            UInt16[] response = ModBus.readInputRegister_04(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_inputRegistersTable,
                                        list_template_inputRegistersTable,
                                        P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxInputRegOffset_,
                                        comboBoxInputRegRegistri_,
                                        comboBoxInputRegValori_,
                                        false,
                                        false);
                                }
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                            });

                            address_start = 0xFFFF;
                            num_regs = 0;
                        }

                        num_regs += curr_len;

                        if (address_start == 0xFFFF)
                            address_start = curr_start;
                    }

                    if (num_regs == 0)
                        num_regs = curr_len;

                    if (address_start != 0xFFFF)
                    {
                        UInt16[] response = ModBus.readInputRegister_04(
                        byte.Parse(textBoxModbusAddress_),
                        address_start,
                        num_regs,
                        readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_inputRegistersTable,
                                    list_template_inputRegistersTable,
                                    P.uint_parser(textBoxInputRegOffset_, comboBoxInputRegOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxInputRegOffset_,
                                    comboBoxInputRegRegistri_,
                                    comboBoxInputRegValori_,
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                    {
                        SetTableDisconnectError(list_inputRegistersTable, true);

                        if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                buttonSerialActive_Click(null, null);
                            });
                        }
                        else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                buttonTcpActive_Click(null, null);
                            });
                        }
                        else
                        {

                        }
                    }
                    else if (ex is ModbusException)
                    {
                        if (ex.Message.IndexOf("Timed out") != -1)
                        {
                            SetTableTimeoutError(list_inputRegistersTable, false);
                        }
                        if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                        {
                            SetTableModBusError(list_inputRegistersTable, (ModbusException)ex, false);
                        }
                        if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                        {
                            SetTableStringError(list_inputRegistersTable, (ModbusException)ex, true);
                        }
                        if (ex.Message.IndexOf("CRC Error") != -1)
                        {
                            SetTableCrcError(list_inputRegistersTable, false);
                        }
                    }
                    else
                    {
                        SetTableInternalError(list_inputRegistersTable, false);
                    }

                    Console.WriteLine(ex);
                }
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                buttonReadInputRegisterTemplateGroup.IsEnabled = true;

                dataGridViewInputRegister.ItemsSource = null;
                dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;

                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
            });
        }

        private void comboBoxInputRegisterGroup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboBoxInputRegisterGroup.SelectedItem != null)
            {
                KeyValuePair<Group_Item, String> kp = (KeyValuePair<Group_Item, String>)comboBoxInputRegisterGroup.SelectedItem;
                comboBoxInputRegisterGroup_ = kp.Key.Group;

                buttonReadInputRegisterTemplateGroup_Click(null, null);
            }
        }

        private void buttonReadInputTemplate_Click(object sender, RoutedEventArgs e)
        {
            buttonReadInputTemplate.IsEnabled = false;
            lastInputsCommandGroup = false;

            Thread t = new Thread(new ThreadStart(readInputTemplateAll));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readInputTemplateAll()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                list_inputsTable.Clear();

                TaskbarItemInfo.ProgressValue = 0;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            UInt32 counter = 0;

            if (useOnlyReadSingleRegisterForGroups)
            {
                foreach (ModBus_Item item in list_template_inputsTable.OrderBy(x => x.RegisterUInt))
                {
                    try
                    {
                        counter++;

                        if (item == null)
                            continue;

                        uint address_start = item.RegisterUInt;
                        uint num_regs = 1;

                        UInt16[] response = ModBus.readInputStatus_02(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_inputsTable,
                                    list_template_inputsTable,
                                    P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxInputOffset_,
                                    comboBoxInputRegistri_,
                                    "DEC",
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_inputsTable.Count);
                        });
                    }
                    catch (Exception ex)
                    {
                        if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                        {
                            SetTableDisconnectError(list_inputsTable, true);

                            if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonSerialActive_Click(null, null);
                                });
                            }
                            else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonTcpActive_Click(null, null);
                                });
                            }
                        }
                        else if (ex is ModbusException)
                        {
                            if (ex.Message.IndexOf("Timed out") != -1)
                            {
                                SetTableTimeoutError(list_inputsTable, false);
                            }
                            if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                SetTableModBusError(list_inputsTable, (ModbusException)ex, false);
                            }
                            if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                            {
                                SetTableStringError(list_inputsTable, (ModbusException)ex, true);
                            }
                            if (ex.Message.IndexOf("CRC Error") != -1)
                            {
                                SetTableCrcError(list_inputsTable, false);
                            }
                        }
                        else
                        {
                            SetTableInternalError(list_inputsTable, false);
                        }

                        Console.WriteLine(ex);
                    }
                }
            }
            else
            {
                uint address_start = 0xFFFF;
                uint num_regs = 0;

                uint curr_start = 0xFFFF;

                try
                {
                    foreach (ModBus_Item item in list_template_inputsTable.OrderBy(x => x.RegisterUInt))
                    {
                        counter++;

                        if (item == null)
                            continue;

                        curr_start = item.RegisterUInt;

                        if (address_start != 0xFFFF && ((curr_start > address_start + num_regs) || (num_regs >= UInt16.Parse(textBoxInputNumber_))))
                        {
                            UInt16[] response = ModBus.readInputStatus_02(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_inputsTable,
                                        list_template_inputsTable,
                                        P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxInputOffset_,
                                        comboBoxInputRegistri_,
                                        "DEC",
                                        false,
                                        false);
                                }
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_inputsTable.Count);
                            });

                            address_start = 0xFFFF;
                            num_regs = 0;
                        }

                        num_regs += 1;

                        if (address_start == 0xFFFF)
                            address_start = curr_start;
                    }

                    if (address_start != 0xFFFF)
                    {
                        UInt16[] response = ModBus.readInputStatus_02(
                        byte.Parse(textBoxModbusAddress_),
                        address_start,
                        num_regs,
                        readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_inputsTable,
                                    list_template_inputsTable,
                                    P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxInputOffset_,
                                    comboBoxInputRegistri_,
                                    "DEC",
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_inputsTable.Count);
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                    {
                        SetTableDisconnectError(list_inputsTable, true);

                        if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonSerialActive_Click(null, null);
                            });
                        }
                        else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonTcpActive_Click(null, null);
                            });
                        }
                    }
                    else if (ex is ModbusException)
                    {
                        if (ex.Message.IndexOf("Timed out") != -1)
                        {
                            SetTableTimeoutError(list_inputsTable, false);
                        }
                        if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                        {
                            SetTableModBusError(list_inputsTable, (ModbusException)ex, false);
                        }
                        if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                        {
                            SetTableStringError(list_inputsTable, (ModbusException)ex, true);
                        }
                        if (ex.Message.IndexOf("CRC Error") != -1)
                        {
                            SetTableCrcError(list_inputsTable, false);
                        }
                    }
                    else
                    {
                        SetTableInternalError(list_inputsTable, false);
                    }

                    Console.WriteLine(ex);
                }
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                buttonReadInputTemplate.IsEnabled = true;

                dataGridViewInput.ItemsSource = null;
                dataGridViewInput.ItemsSource = list_inputsTable;

                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
            });
        }

        private void buttonReadInputTemplateGroup_Click(object sender, RoutedEventArgs e)
        {
            buttonReadInputTemplateGroup.IsEnabled = false;
            lastInputsCommandGroup = true;

            Thread t = new Thread(new ThreadStart(readInputTemplateGroup));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readInputTemplateGroup()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                list_inputsTable.Clear();

                TaskbarItemInfo.ProgressValue = 0;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            List<ModBus_Item> toRead = new List<ModBus_Item>();
            List<uint> toRemove = new List<uint>();
            UInt32 counter = 0;

            foreach (ModBus_Item item in list_template_inputsTable.OrderBy(x => x.RegisterUInt))
            {
                try
                {
                    if (item == null)
                        continue;

                    if (item.Group == null && (comboBoxInputGroup_.Length > 1))
                        continue;

                    if (item.Group == null && (comboBoxInputGroup_.Length > 0))
                    {
                        if (int.Parse(comboBoxInputGroup_) != 0)
                            continue;
                    }

                    if (item.Group != null)
                    {
                        bool found = false;

                        foreach (string str in item.Group.Split(';'))
                        {
                            int dummy = -1;
                            if (int.TryParse(str, out dummy))
                            {
                                if (dummy == int.Parse(comboBoxInputGroup_))
                                {
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            continue;
                    }

                    toRead.Add(item);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }

            if (useOnlyReadSingleRegisterForGroups)
            {
                foreach (ModBus_Item item in toRead)
                {
                    try
                    {
                        uint address_start = item.RegisterUInt;
                        uint num_regs = 1;
                        counter++;

                        UInt16[] response = ModBus.readInputStatus_02(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_inputsTable,
                                    list_template_inputsTable,
                                    P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxInputOffset_,
                                    comboBoxInputRegistri_,
                                    "DEC",
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                        });
                    }
                    catch (Exception ex)
                    {
                        if (ex is InvalidOperationException)
                        {
                            SetTableDisconnectError(list_inputsTable, true);

                            if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonSerialActive_Click(null, null);
                                });
                            }
                            else
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonTcpActive_Click(null, null);
                                });
                            }
                        }
                        else if (ex is ModbusException)
                        {
                            if (ex.Message.IndexOf("Timed out") != -1)
                            {
                                SetTableTimeoutError(list_inputsTable, false);
                            }
                            if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                SetTableModBusError(list_inputsTable, (ModbusException)ex, false);
                            }
                            if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                            {
                                SetTableStringError(list_inputsTable, (ModbusException)ex, true);
                            }
                            if (ex.Message.IndexOf("CRC Error") != -1)
                            {
                                SetTableCrcError(list_inputsTable, false);
                            }
                        }
                        else
                        {
                            SetTableInternalError(list_inputsTable, false);
                        }

                        Console.WriteLine(ex);
                    }
                }
            }
            else
            {
                uint address_start = 0xFFFF;
                uint num_regs = 0;

                uint curr_start = 0xFFFF;

                try
                {
                    foreach (ModBus_Item item in toRead)
                    {
                        curr_start = item.RegisterUInt;
                        counter++;

                        if (address_start != 0xFFFF && ((curr_start > address_start + num_regs) || (num_regs >= UInt16.Parse(textBoxInputNumber_))))
                        {
                            UInt16[] response = ModBus.readInputStatus_02(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_inputsTable,
                                        list_template_inputsTable,
                                        P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxInputOffset_,
                                        comboBoxInputRegistri_,
                                        "DEC",
                                        false,
                                        false);
                                }
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                            });

                            address_start = 0xFFFF;
                            num_regs = 0;
                        }

                        num_regs += 1;

                        if (address_start == 0xFFFF)
                            address_start = curr_start;
                    }

                    if (address_start != 0xFFFF)
                    {
                        UInt16[] response = ModBus.readInputStatus_02(
                        byte.Parse(textBoxModbusAddress_),
                        address_start,
                        num_regs,
                        readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_inputsTable,
                                    list_template_inputsTable,
                                    P.uint_parser(textBoxInputOffset_, comboBoxInputOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxInputOffset_,
                                    comboBoxInputRegistri_,
                                    "DEC",
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException)
                    {
                        SetTableDisconnectError(list_inputsTable, true);

                        if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonSerialActive_Click(null, null);
                            });
                        }
                        else
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonTcpActive_Click(null, null);
                            });
                        }
                    }
                    else if (ex is ModbusException)
                    {
                        if (ex.Message.IndexOf("Timed out") != -1)
                        {
                            SetTableTimeoutError(list_inputsTable, false);
                        }
                        if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                        {
                            SetTableModBusError(list_inputsTable, (ModbusException)ex, false);
                        }
                        if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                        {
                            SetTableStringError(list_inputsTable, (ModbusException)ex, true);
                        }
                        if (ex.Message.IndexOf("CRC Error") != -1)
                        {
                            SetTableCrcError(list_inputsTable, false);
                        }
                    }
                    else
                    {
                        SetTableInternalError(list_inputsTable, false);
                    }

                    Console.WriteLine(ex);
                }
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                buttonReadInputTemplateGroup.IsEnabled = true;

                dataGridViewInput.ItemsSource = null;
                dataGridViewInput.ItemsSource = list_inputsTable;

                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
            });
        }

        private void comboBoxInputGroup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboBoxInputGroup.SelectedItem != null)
            {
                KeyValuePair<Group_Item, String> kp = (KeyValuePair<Group_Item, String>)comboBoxInputGroup.SelectedItem;
                comboBoxInputGroup_ = kp.Key.Group;

                buttonReadInputTemplateGroup_Click(null, null);
            }
        }

        private void buttonReadCoilsTemplate_Click(object sender, RoutedEventArgs e)
        {
            buttonReadCoilsTemplate.IsEnabled = false;
            lastCoilsCommandGroup = false;

            Thread t = new Thread(new ThreadStart(readCoilsTemplateAll));
            t.Priority = threadPriority;
            t.Start();
        }

        public void readCoilsTemplateAll()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                list_coilsTable.Clear();

                TaskbarItemInfo.ProgressValue = 0;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            UInt32 counter = 0;

            try
            {
                if (useOnlyReadSingleRegisterForGroups)
                {
                    foreach (ModBus_Item item in list_template_coilsTable.OrderBy(x => x.RegisterUInt))
                    {
                        try
                        {
                            counter++;

                            if (item == null)
                                continue;

                            uint address_start = item.RegisterUInt;
                            uint num_regs = 1;

                            UInt16[] response = ModBus.readCoilStatus_01(
                                byte.Parse(textBoxModbusAddress_),
                                address_start,
                                num_regs,
                                readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_coilsTable,
                                        list_template_coilsTable,
                                        P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxCoilsOffset_,
                                        comboBoxCoilsRegistri_,
                                        "DEC",
                                        false,
                                        false);
                                }
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_coilsTable.Count);
                            });
                        }
                        catch (Exception ex)
                        {
                            if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                            {
                                SetTableDisconnectError(list_coilsTable, true);

                                if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                                {
                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        if (ModBus.ClientActive)
                                            buttonSerialActive_Click(null, null);
                                    });
                                }
                                else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                                {
                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        if (ModBus.ClientActive)
                                            buttonTcpActive_Click(null, null);
                                    });
                                }
                                else
                                {
                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        buttonReadCoilsTemplate.IsEnabled = true;
                                    });
                                }
                            }
                            else if (ex is ModbusException)
                            {
                                if (ex.Message.IndexOf("Timed out") != -1)
                                {
                                    SetTableTimeoutError(list_coilsTable, false);
                                }
                                if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                                {
                                    SetTableModBusError(list_coilsTable, (ModbusException)ex, false);
                                }
                                if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                                {
                                    SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                                }
                                if (ex.Message.IndexOf("CRC Error") != -1)
                                {
                                    SetTableCrcError(list_coilsTable, false);
                                }
                            }
                            else
                            {
                                SetTableInternalError(list_coilsTable, false);
                            }

                            Console.WriteLine(ex);
                        }
                    }
                }
                else
                {
                    uint address_start = 0xFFFF;
                    uint num_regs = 0;

                    uint curr_start = 0xFFFF;

                    try
                    {
                        foreach (ModBus_Item item in list_template_coilsTable.OrderBy(x => x.RegisterUInt))
                        {
                            counter++;

                            if (item == null)
                                continue;

                            curr_start = item.RegisterUInt;

                            if (address_start != 0xFFFF && ((curr_start > address_start + num_regs) || (num_regs >= UInt16.Parse(textBoxCoilNumber_))))
                            {
                                UInt16[] response = ModBus.readCoilStatus_01(
                                byte.Parse(textBoxModbusAddress_),
                                address_start,
                                num_regs,
                                readTimeout);

                                if (response != null)
                                {
                                    if (response.Length > 0)
                                    {
                                        // Cancello la tabella e inserisco le nuove righe
                                        insertRowsTable(
                                            list_coilsTable,
                                            list_template_coilsTable,
                                            P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_),
                                            address_start,
                                            response,
                                            colorDefaultReadCellStr,
                                            comboBoxCoilsOffset_,
                                            comboBoxCoilsRegistri_,
                                            "DEC",
                                            false,
                                            false);
                                    }
                                }

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_coilsTable.Count);
                                });

                                address_start = 0xFFFF;
                                num_regs = 0;
                            }

                            num_regs += 1;

                            if (address_start == 0xFFFF)
                                address_start = curr_start;
                        }

                        if (address_start != 0xFFFF)
                        {
                            UInt16[] response = ModBus.readCoilStatus_01(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_coilsTable,
                                        list_template_coilsTable,
                                        P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxCoilsOffset_,
                                        comboBoxCoilsRegistri_,
                                        "DEC",
                                        false,
                                        false);
                                }
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(list_template_coilsTable.Count);
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                        {
                            SetTableDisconnectError(list_coilsTable, true);

                            if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonSerialActive_Click(null, null);
                                });
                            }
                            else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonTcpActive_Click(null, null);
                                });
                            }
                            else
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    buttonReadCoilsTemplate.IsEnabled = true;
                                });
                            }
                        }
                        else if (ex is ModbusException)
                        {
                            if (ex.Message.IndexOf("Timed out") != -1)
                            {
                                SetTableTimeoutError(list_coilsTable, false);
                            }
                            if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                SetTableModBusError(list_coilsTable, (ModbusException)ex, false);
                            }
                            if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                            {
                                SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                            }
                            if (ex.Message.IndexOf("CRC Error") != -1)
                            {
                                SetTableCrcError(list_coilsTable, false);
                            }
                        }
                        else
                        {
                            SetTableInternalError(list_coilsTable, false);
                        }

                        Console.WriteLine(ex);
                    }
                }
            }
            catch (Exception err)
            {
                SetTableInternalError(list_coilsTable, false);
                Console.WriteLine(err);
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                buttonReadCoilsTemplate.IsEnabled = true;

                dataGridViewCoils.ItemsSource = null;
                dataGridViewCoils.ItemsSource = list_coilsTable;

                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
            });
        }

        private void buttonReadCoilsTemplateGroup_Click(object sender, RoutedEventArgs e)
        {
            if (comboBoxCoilsGroup.IsEnabled)
            {
                buttonReadCoilsTemplateGroup.IsEnabled = false;
                lastCoilsCommandGroup = true;

                Thread t = new Thread(new ThreadStart(readCoilsTemplateGroup));
                t.Priority = threadPriority;
                t.Start();
            }
        }

        public void readCoilsTemplateGroup()
        {
            this.Dispatcher.Invoke((Action)delegate
            {
                list_coilsTable.Clear();

                TaskbarItemInfo.ProgressValue = 0;
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
            });

            List<ModBus_Item> toRead = new List<ModBus_Item>();
            List<uint> toRemove = new List<uint>();
            UInt32 counter = 0;

            foreach (ModBus_Item item in list_template_coilsTable.OrderBy(x => x.RegisterUInt))
            {
                try
                {
                    if (item == null)
                        continue;

                    if (item.Group == null && (comboBoxCoilsGroup_.Length > 1))
                        continue;

                    if (item.Group == null && (comboBoxCoilsGroup_.Length > 0))
                    {
                        if (int.Parse(comboBoxCoilsGroup_) != 0)
                            continue;
                    }

                    if (item.Group != null)
                    {
                        bool found = false;

                        foreach (string str in item.Group.Split(';'))
                        {
                            int dummy = -1;
                            if (int.TryParse(str, out dummy))
                            {
                                if (dummy == int.Parse(comboBoxCoilsGroup_))
                                {
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            continue;
                    }

                    toRead.Add(item);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }

            if (useOnlyReadSingleRegisterForGroups)
            {
                foreach (ModBus_Item item in toRead)
                {
                    try
                    {
                        uint address_start = item.RegisterUInt;
                        uint num_regs = 1;
                        counter++;

                        UInt16[] response = ModBus.readCoilStatus_01(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_coilsTable,
                                    list_template_coilsTable,
                                    P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxCoilsOffset_,
                                    comboBoxCoilsRegistri_,
                                    "DEC",
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                        });
                    }
                    catch (Exception ex)
                    {
                        if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                        {
                            SetTableDisconnectError(list_coilsTable, true);

                            if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonSerialActive_Click(null, null);
                                });
                            }
                            else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                            {
                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    if (ModBus.ClientActive)
                                        buttonTcpActive_Click(null, null);
                                });
                            }
                            else
                            {

                            }
                        }
                        else if (ex is ModbusException)
                        {
                            if (ex.Message.IndexOf("Timed out") != -1)
                            {
                                SetTableTimeoutError(list_coilsTable, false);
                            }
                            if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                SetTableModBusError(list_coilsTable, (ModbusException)ex, false);
                            }
                            if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                            {
                                SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                            }
                            if (ex.Message.IndexOf("CRC Error") != -1)
                            {
                                SetTableCrcError(list_coilsTable, false);
                            }
                        }
                        else
                        {
                            SetTableInternalError(list_coilsTable, false);
                        }

                        Console.WriteLine(ex);
                    }
                }
            }
            else
            {
                uint address_start = 0xFFFF;
                uint num_regs = 0;

                uint curr_start = 0xFFFF;

                try
                {
                    foreach (ModBus_Item item in toRead)
                    {
                        curr_start = item.RegisterUInt;
                        counter++;

                        if (address_start != 0xFFFF && ((curr_start > address_start + num_regs) || (num_regs >= UInt16.Parse(textBoxCoilNumber_))))
                        {
                            UInt16[] response = ModBus.readCoilStatus_01(
                            byte.Parse(textBoxModbusAddress_),
                            address_start,
                            num_regs,
                            readTimeout);

                            if (response != null)
                            {
                                if (response.Length > 0)
                                {
                                    // Cancello la tabella e inserisco le nuove righe
                                    insertRowsTable(
                                        list_coilsTable,
                                        list_template_coilsTable,
                                        P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_),
                                        address_start,
                                        response,
                                        colorDefaultReadCellStr,
                                        comboBoxCoilsOffset_,
                                        comboBoxCoilsRegistri_,
                                        "DEC",
                                        false,
                                        false);
                                }
                            }

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                            });

                            address_start = 0xFFFF;
                            num_regs = 0;
                        }

                        num_regs += 1;

                        if (address_start == 0xFFFF)
                            address_start = curr_start;
                    }

                    if (address_start != 0xFFFF)
                    {
                        UInt16[] response = ModBus.readCoilStatus_01(
                        byte.Parse(textBoxModbusAddress_),
                        address_start,
                        num_regs,
                        readTimeout);

                        if (response != null)
                        {
                            if (response.Length > 0)
                            {
                                // Cancello la tabella e inserisco le nuove righe
                                insertRowsTable(
                                    list_coilsTable,
                                    list_template_coilsTable,
                                    P.uint_parser(textBoxCoilsOffset_, comboBoxCoilsOffset_),
                                    address_start,
                                    response,
                                    colorDefaultReadCellStr,
                                    comboBoxCoilsOffset_,
                                    comboBoxCoilsRegistri_,
                                    "DEC",
                                    false,
                                    false);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            TaskbarItemInfo.ProgressValue = (double)(counter) / (double)(toRead.Count);
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException || ex is IOException || ex is SocketException)
                    {
                        SetTableDisconnectError(list_coilsTable, true);

                        if (ModBus.type == ModBus_Def.TYPE_RTU || ModBus.type == ModBus_Def.TYPE_ASCII)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonSerialActive_Click(null, null);
                            });
                        }
                        else if (ModBus.type == ModBus_Def.TYPE_TCP_SOCK || ModBus.type == ModBus_Def.TYPE_TCP_SECURE)
                        {
                            this.Dispatcher.Invoke((Action)delegate
                            {
                                if (ModBus.ClientActive)
                                    buttonTcpActive_Click(null, null);
                            });
                        }
                        else
                        {

                        }
                    }
                    else if (ex is ModbusException)
                    {
                        if (ex.Message.IndexOf("Timed out") != -1)
                        {
                            SetTableTimeoutError(list_coilsTable, false);
                        }
                        if (ex.Message.IndexOf("ModBus ErrCode") != -1)
                        {
                            SetTableModBusError(list_coilsTable, (ModbusException)ex, false);
                        }
                        if (ex.Message.IndexOf("ModbusProtocolError") != -1)
                        {
                            SetTableStringError(list_coilsTable, (ModbusException)ex, true);
                        }
                        if (ex.Message.IndexOf("CRC Error") != -1)
                        {
                            SetTableCrcError(list_coilsTable, false);
                        }
                    }
                    else
                    {
                        SetTableInternalError(list_coilsTable, false);
                    }

                    Console.WriteLine(ex);
                }
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                buttonReadCoilsTemplateGroup.IsEnabled = true;

                dataGridViewCoils.ItemsSource = null;
                dataGridViewCoils.ItemsSource = list_coilsTable;

                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
            });
        }

        private void comboBoxCoilsGroup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboBoxCoilsGroup.SelectedItem != null && buttonReadCoilsTemplateGroup.IsEnabled)
            {
                KeyValuePair<Group_Item, String> kp = (KeyValuePair<Group_Item, String>)comboBoxCoilsGroup.SelectedItem;
                comboBoxCoilsGroup_ = kp.Key.Group;

                buttonReadCoilsTemplateGroup_Click(null, null);
            }
        }

        private void ViewFullSizeTables_Checked(object sender, RoutedEventArgs e)
        {
            if (ViewFullSizeTables.IsChecked)
            {
                Thickness marg = new Thickness(10, 50, 10, 10);
                dataGridViewCoils.Margin = marg;
                dataGridViewInput.Margin = marg;
                dataGridViewInputRegister.Margin = marg;
                dataGridViewHolding.Margin = marg;
            }
            else
            {
                Thickness marg = new Thickness(10, 50, 415, 10);
                dataGridViewCoils.Margin = marg;
                dataGridViewInput.Margin = marg;
                dataGridViewInputRegister.Margin = marg;
                dataGridViewHolding.Margin = marg;
            }
        }

        private void ButtonFullSize_Click(object sender, RoutedEventArgs e)
        {
            ViewFullSizeTables.IsChecked = !ViewFullSizeTables.IsChecked;
            ViewFullSizeTables_Checked(null, null);
        }

        private void dataGridViewHolding_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            disableDeleteKey = true;
        }

        private void comboBoxProfileHome_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!disableComboProfile)
            {
                if (comboBoxProfileHome.SelectedItem != null)
                {
                    LoadProfile(comboBoxProfileHome.SelectedItem.ToString());
                }
            }
        }

        private void buttonLoadClientCertificate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".pfx";
            openFileDialog.AddExtension = false;
            openFileDialog.Filter = "PFX|*.pfx";
            openFileDialog.Title = "Client certificate";
            openFileDialog.ShowDialog();

            if (openFileDialog.FileName != "")
            {
                if (File.Exists(openFileDialog.FileName))
                {
                    textBoxClientCertificatePath.Text = openFileDialog.FileName;
                }
            }
        }

        private void buttonLoadClientKey_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".key";
            openFileDialog.AddExtension = false;
            openFileDialog.Filter = "KEY|*.key|PEM|*.pem";
            openFileDialog.Title = "Client certificate";
            openFileDialog.ShowDialog();

            if (openFileDialog.FileName != "")
            {
                if (File.Exists(openFileDialog.FileName))
                {
                    textBoxClientKeyPath.Text = openFileDialog.FileName;
                }
            }
        }

        private void textBoxClientCertificatePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            labelClientCertificatePath.Content = System.IO.Path.GetFileName(textBoxClientCertificatePath.Text);

            /*Predisposto per certificati .crt
            if(textBoxClientCertificatePath.Text.IndexOf(".pfx") != -1)
            {
                // Certificato pfx (cert+key con password)
                if(labelClientPasswordName != null)
                    labelClientPasswordName.Visibility = Visibility.Visible;
                if(textBoxCertificatePassword != null)
                    textBoxCertificatePassword.Visibility = Visibility.Visible;
                if(labelClientKeyName != null)
                    labelClientKeyName.Visibility = Visibility.Hidden;
                if(labelClientKeyPath != null)
                    labelClientKeyPath.Visibility = Visibility.Hidden;
                if(buttonLoadClientKey != null)
                    buttonLoadClientKey.Visibility = Visibility.Hidden;
            }
            else
            {
                // Certificato .cert + Chiave .key
                if (labelClientPasswordName != null)
                    labelClientPasswordName.Visibility = Visibility.Hidden;
                if (textBoxCertificatePassword != null)
                    textBoxCertificatePassword.Visibility = Visibility.Hidden;
                if (labelClientKeyName != null)
                    labelClientKeyName.Visibility = Visibility.Visible;
                if (labelClientKeyPath != null)
                    labelClientKeyPath.Visibility = Visibility.Visible;
                if (buttonLoadClientKey != null)
                    buttonLoadClientKey.Visibility = Visibility.Visible;
            }*/
        }

        private void textBoxClientKeyPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            labelClientKeyPath.Content = System.IO.Path.GetFileName(textBoxClientKeyPath.Text);
        }

        private void CheckBoxModbusSecure_Checked(object sender, RoutedEventArgs e)
        {
            labelClientCertificateName.Visibility = (bool)checkBoxModbusSecure.IsChecked ? Visibility.Visible : Visibility.Hidden;
            labelClientCertificatePath.Visibility = (bool)checkBoxModbusSecure.IsChecked ? Visibility.Visible : Visibility.Hidden;
            buttonLoadClientCertificate.Visibility = (bool)checkBoxModbusSecure.IsChecked ? Visibility.Visible : Visibility.Hidden;
            labelClientTLSVersion.Visibility = (bool)checkBoxModbusSecure.IsChecked ? Visibility.Visible : Visibility.Hidden;
            comboBoxTlsVersion.Visibility = (bool)checkBoxModbusSecure.IsChecked ? Visibility.Visible : Visibility.Hidden;

            if ((bool)checkBoxModbusSecure.IsChecked)
            {
                // Predisposto per uso con file .crt
                /*if (textBoxClientCertificatePath.Text.IndexOf(".pfx") != -1)
                {
                    // Certificato pfx (cert+key con password)
                    if (labelClientPasswordName != null)
                        labelClientPasswordName.Visibility = Visibility.Visible;
                    if (textBoxCertificatePassword != null)
                        textBoxCertificatePassword.Visibility = Visibility.Visible;
                    if (labelClientKeyName != null)
                        labelClientKeyName.Visibility = Visibility.Hidden;
                    if (labelClientKeyPath != null)
                        labelClientKeyPath.Visibility = Visibility.Hidden;
                    if (buttonLoadClientKey != null)
                        buttonLoadClientKey.Visibility = Visibility.Hidden;
                }
                else
                {
                    // Certificato .cert + Chiave .key
                    if (labelClientPasswordName != null)
                        labelClientPasswordName.Visibility = Visibility.Hidden;
                    if (textBoxCertificatePassword != null)
                        textBoxCertificatePassword.Visibility = Visibility.Hidden;
                    if (labelClientKeyName != null)
                        labelClientKeyName.Visibility = Visibility.Visible;
                    if (labelClientKeyPath != null)
                        labelClientKeyPath.Visibility = Visibility.Visible;
                    if (buttonLoadClientKey != null)
                        buttonLoadClientKey.Visibility = Visibility.Visible;
                }*/

                // Certificato pfx (cert+key con password)
                labelClientPasswordName.Visibility = Visibility.Visible;
                textBoxCertificatePassword.Visibility = Visibility.Visible;
                labelClientKeyName.Visibility = Visibility.Hidden;
                labelClientKeyPath.Visibility = Visibility.Hidden;
                buttonLoadClientKey.Visibility = Visibility.Hidden;
            }
            else
            {
                if (labelClientPasswordName != null)
                    labelClientPasswordName.Visibility = Visibility.Hidden;
                if (textBoxCertificatePassword != null)
                    textBoxCertificatePassword.Visibility = Visibility.Hidden;
                if (labelClientKeyName != null)
                    labelClientKeyName.Visibility = Visibility.Hidden;
                if (labelClientKeyPath != null)
                    labelClientKeyPath.Visibility = Visibility.Hidden;
                if (buttonLoadClientKey != null)
                    buttonLoadClientKey.Visibility = Visibility.Hidden;
            }
        }

        private void CheckBoxUseOnlyReadSingleRegistersForGroups_Checked(object sender, RoutedEventArgs e)
        {
            useOnlyReadSingleRegisterForGroups = (bool)CheckBoxUseOnlyReadSingleRegistersForGroups.IsChecked;
        }

        private void buttonImportZip_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog window = new OpenFileDialog();

            window.Filter = "ModBus Profile | *.mbp|Zip File | *.zip";
            window.DefaultExt = ".mbp";
            window.Multiselect = false;

            if ((bool)window.ShowDialog())
            {
                if (Directory.Exists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + System.IO.Path.GetFileNameWithoutExtension(window.FileName)))
                {
                    if (MessageBox.Show(lang.languageTemplate["strings"]["overwriteProfile"], "Info", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                    {
                        Directory.Delete(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + System.IO.Path.GetFileNameWithoutExtension(window.FileName), true);
                    }
                    else
                    {
                        return;
                    }
                }

                ZipFile.ExtractToDirectory(window.FileName, System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + System.IO.Path.GetFileNameWithoutExtension(window.FileName));
            }

            disableComboProfile = true;
            comboBoxProfileHome.Items.Clear();

            foreach (String sub in Directory.GetDirectories(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\"))
            {
                comboBoxProfileHome.Items.Add(sub.Split('\\')[sub.Split('\\').Length - 1]);
            }

            LoadProfile(System.IO.Path.GetFileNameWithoutExtension(window.FileName));
            comboBoxProfileHome.SelectedItem = System.IO.Path.GetFileNameWithoutExtension(window.FileName);
            disableComboProfile = false;
        }

        private void buttonExportZip_Click(object sender, RoutedEventArgs e)
        {
            String currItem = comboBoxProfileHome.SelectedItem.ToString();

            SaveFileDialog window = new SaveFileDialog();

            window.Filter = "ModBus Profile | *.mbp|Zip File | *.zip";
            window.DefaultExt = ".mbp";
            window.FileName = currItem + ".mbp";

            if ((bool)window.ShowDialog())
            {
                // Se il file lo esiste gia' lo elimino
                if (File.Exists(window.FileName))
                    File.Delete(window.FileName);

                ZipFile.CreateFromDirectory("Json\\" + currItem, window.FileName);
            }
        }

        private void buttonImportCoilsClipboard_Click(object sender, RoutedEventArgs e)
        {
            PreviewImport previewImport = new PreviewImport(this, null, 5);
            previewImport.Show();
        }

        private void buttonImportHoldingRegClipboard_Click(object sender, RoutedEventArgs e)
        {
            PreviewImport previewImport = new PreviewImport(this, null, 6);
            previewImport.Show();
        }

        private void coilsMenuItemDelete_Click(object sender, RoutedEventArgs e)
        {
            list_coilsTable.Clear();
        }

        private void coilsMenuItemCut_Click(object sender, RoutedEventArgs e)
        {
            CopyListToClipboard(list_coilsTable);
            list_coilsTable.Clear();
        }

        private void coilsMenuItemCopy_Click(object sender, RoutedEventArgs e)
        {
            CopyListToClipboard(list_coilsTable);
        }

        private void coilsMenuItemPaste_Click(object sender, RoutedEventArgs e)
        {
            PreviewImport previewImport = new PreviewImport(this, null, 5);
            previewImport.Show();
        }

        private void coilsMenuItemImport_Click(object sender, RoutedEventArgs e)
        {
            buttonImportCoils_Click(null, null);
        }

        private void coilsMenuItemExport_Click(object sender, RoutedEventArgs e)
        {
            buttonExportCoils_Click(null, null);
        }

        private void inputsMenuItemDelete_Click(object sender, RoutedEventArgs e)
        {
            list_inputsTable.Clear();
        }

        private void inputsMenuItemCut_Click(object sender, RoutedEventArgs e)
        {
            CopyListToClipboard(list_inputsTable);
            list_inputsTable.Clear();
        }

        private void inputsMenuItemCopy_Click(object sender, RoutedEventArgs e)
        {
            CopyListToClipboard(list_inputsTable);
        }

        private void inputsMenuItemExport_Click(object sender, RoutedEventArgs e)
        {
            buttonExportInput_Click(null, null);
        }

        private void holdingRegistersMenuItemDelete_Click(object sender, RoutedEventArgs e)
        {
            list_holdingRegistersTable.Clear();
        }

        private void holdingRegistersMenuItemCut_Click(object sender, RoutedEventArgs e)
        {
            CopyListToClipboard(list_holdingRegistersTable);
            list_holdingRegistersTable.Clear();
        }

        private void holdingRegistersMenuItemCopy_Click(object sender, RoutedEventArgs e)
        {
            CopyListToClipboard(list_holdingRegistersTable);
        }

        private void holdingRegistersMenuItemPaste_Click(object sender, RoutedEventArgs e)
        {
            PreviewImport previewImport = new PreviewImport(this, null, 6);
            previewImport.Show();
        }

        private void holdingRegistersMenuItemImport_Click(object sender, RoutedEventArgs e)
        {
            buttonImportHoldingReg_Click(null, null);
        }

        private void holdingRegistersMenuItemExport_Click(object sender, RoutedEventArgs e)
        {
            buttonExportHoldingReg_Click(null, null);
        }

        private void inputRegistersMenuItemDelete_Click(object sender, RoutedEventArgs e)
        {
            list_inputRegistersTable.Clear();
        }

        private void inputRegistersMenuItemCut_Click(object sender, RoutedEventArgs e)
        {
            CopyListToClipboard(list_inputRegistersTable);
            list_inputRegistersTable.Clear();
        }

        private void inputRegistersMenuItemCopy_Click(object sender, RoutedEventArgs e)
        {
            CopyListToClipboard(list_inputRegistersTable);
        }

        private void inputRegistersMenuItemExport_Click(object sender, RoutedEventArgs e)
        {
            buttonExportInputReg_Click(null, null);
        }

        public void CopyListToClipboard(ObservableCollection<ModBus_Item> list)
        {
            try
            {
                string clipContent = "";
                foreach (ModBus_Item item in list)
                {
                    clipContent += item.Offset;
                    clipContent += "\t";
                    clipContent += item.Register;
                    clipContent += "\t";
                    clipContent += item.Value;
                    clipContent += "\t";
                    clipContent += item.Notes;
                    clipContent += "\r\n";
                }

                Clipboard.SetText(clipContent);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private void CheckBoxSendCellEditOnlyOnChange_Checked(object sender, RoutedEventArgs e)
        {
            sendCellEditOnlyOnChange = (bool)CheckBoxSendCellEditOnlyOnChange.IsChecked;
        }

        private void statisticsToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (ModBus != null)
            {
                Statistics statistics = new Statistics(this.title, ModBus);
                statistics.Show();
            }
        }
    }

    // Classe per caricare dati dal file di configurazione json, vecchia versione salvataggio
    public class SAVE
    {
        // Variabili interfaccia 
        public bool usingSerial { get; set; } //True -> Serial, False -> TCP

        public string modbusAddress { get; set; }

        //Vatiabili configurazione seriale
        public int serialPort { get; set; }
        public int serialSpeed { get; set; }
        public int serialParity { get; set; }
        public int serialStop { get; set; }

        public bool serialMaster { get; set; } //True -> Master, False -> False
        public bool serialRTU { get; set; }

        //Variabili configurazione tcp
        public string tcpClientIpAddress { get; set; }
        public string tcpClientPort { get; set; }
        public string tcpServerIpAddress { get; set; }
        public string tcpServerPort { get; set; }

        //VARIABILI TABPAGE

        //TabPage1 (Coils)
        public string textBoxCoilsAddress01 { get; set; }
        public string textBoxCoilNumber { get; set; }
        public string textBoxCoilsRange_A { get; set; }
        public string textBoxCoilsRange_B { get; set; }
        public string textBoxCoilsAddress05 { get; set; }
        public string textBoxCoilsValue05 { get; set; }
        public string textBoxCoilsAddress15_A { get; set; }
        public string textBoxCoilsAddress15_B { get; set; }
        public string textBoxCoilsValue15 { get; set; }
        public string textBoxGoToCoilAddress { get; set; }


        //TabPage2 (inputs)
        public string textBoxInputAddress02 { get; set; }
        public string textBoxInputNumber { get; set; }
        public string textBoxInputRange_A { get; set; }
        public string textBoxInputRange_B { get; set; }
        public string textBoxGoToInputAddress { get; set; }

        //TabPage3 (input registers)
        public string textBoxInputRegisterAddress04 { get; set; }
        public string textBoxInputRegisterNumber { get; set; }
        public string textBoxInputRegisterRange_A { get; set; }
        public string textBoxInputRegisterRange_B { get; set; }
        public string textBoxGoToInputRegisterAddress { get; set; }

        //TabPage4 (holding registers)
        public string textBoxHoldingAddress03 { get; set; }
        public string textBoxHoldingRegisterNumber { get; set; }
        public string textBoxHoldingRange_A { get; set; }
        public string textBoxHoldingRange_B { get; set; }
        public string textBoxHoldingAddress06 { get; set; }
        public string textBoxHoldingValue06 { get; set; }
        public string textBoxHoldingAddress16_A { get; set; }
        public string textBoxHoldingAddress16_B { get; set; }
        public string textBoxHoldingValue16 { get; set; }
        public string textBoxGoToHoldingAddress { get; set; }

        //TabPage5 (diagnostic)

        //TabPage6 (summary)
        public bool statoConsole { get; set; }

        //Altri elementi aggiunti dopo
        public string comboBoxCoilsAddress01_ { get; set; }
        public string comboBoxCoilsRange_A_ { get; set; }
        public string comboBoxCoilsRange_B_ { get; set; }
        public string comboBoxCoilsAddress05_ { get; set; }
        public string comboBoxCoilsValue05_ { get; set; }
        public string comboBoxCoilsAddress15_A_ { get; set; }
        public string comboBoxCoilsAddress15_B_ { get; set; }
        public string comboBoxCoilsValue15_ { get; set; }
        public string comboBoxInputAddress02_ { get; set; }
        public string comboBoxInputRange_A_ { get; set; }
        public string comboBoxInputRange_B_ { get; set; }
        public string comboBoxInputRegisterAddress04_ { get; set; }
        public string comboBoxInputRegisterRange_A_ { get; set; }
        public string comboBoxInputRegisterRange_B_ { get; set; }
        public string comboBoxHoldingAddress03_ { get; set; }
        public string comboBoxHoldingRange_A_ { get; set; }
        public string comboBoxHoldingRange_B_ { get; set; }
        public string comboBoxHoldingAddress06_ { get; set; }
        public string comboBoxHoldingValue06_ { get; set; }
        public string comboBoxHoldingAddress16_A_ { get; set; }
        public string comboBoxHoldingAddress16_B_ { get; set; }
        public string comboBoxHoldingValue16_ { get; set; }

        public string comboBoxCoilsRegistri_ { get; set; }
        public string comboBoxInputRegistri_ { get; set; }
        public string comboBoxInputRegRegistri_ { get; set; }
        public string comboBoxInputRegValori_ { get; set; }
        public string comboBoxHoldingRegistri_ { get; set; }
        public string comboBoxHoldingValori_ { get; set; }

        public string comboBoxCoilsAddress05_b_ { get; set; }
        public string comboBoxCoilsValue05_b_ { get; set; }
        public string comboBoxHoldingAddress06_b_ { get; set; }
        public string comboBoxHoldingValue06_b_ { get; set; }

        public string comboBoxCoilsOffset_ { get; set; }
        public string comboBoxInputOffset_ { get; set; }
        public string comboBoxInputRegOffset_ { get; set; }
        public string comboBoxHoldingOffset_ { get; set; }

        public string textBoxCoilsAddress05_b_ { get; set; }
        public string textBoxCoilsValue05_b_ { get; set; }
        public string textBoxHoldingAddress06_b_ { get; set; }
        public string textBoxHoldingValue06_b_ { get; set; }

        public string textBoxCoilsOffset_ { get; set; }
        public string textBoxInputOffset_ { get; set; }
        public string textBoxInputRegOffset_ { get; set; }
        public string textBoxHoldingOffset_ { get; set; }

        public bool checkBoxUseOffsetInTables_ { get; set; }
        public bool checkBoxUseOffsetInTextBox_ { get; set; }
        public bool checkBoxFollowModbusProtocol_ { get; set; }
        public bool checkBoxSavePackets_ { get; set; }
        public bool checkBoxCloseConsolAfterBoot_ { get; set; }
        public bool checkBoxCellColorMode_ { get; set; }
        public bool checkBoxViewTableWithoutOffset_ { get; set; }

        public string textBoxSaveLogPath_ { get; set; }

        public string comboBoxDiagnosticFunction_ { get; set; }

        public string textBoxDiagnosticFunctionManual_ { get; set; }

        public string colorDefaultReadCell_ { get; set; }
        public string colorDefaultWriteCell_ { get; set; }
        public string colorErrorCell_ { get; set; }

        public string pathToConfiguration_ { get; set; }

        public string TextBoxPollingInterval_ { get; set; }

        public bool? CheckBoxSendValuesOnEditCoillsTable_ { get; set; }
        public bool? CheckBoxSendValuesOnEditHoldingTable_ { get; set; }

        public string language { get; set; }

        public string textBoxReadTimeout { get; set; }
    }

    public class dataGridJson
    {
        public ModBusItem_Save[] items { get; set; }
    }

    public class ModBusItem_Save 
    {
        public string Offset { get; set; }
        public string Register { get; set; }
        public string Value { get; set; }
        public string Notes { get; set; }
    }
}
