

// -------------------------------------------------------------------------------------------

// Copyright (c) 2023 Federico Turco

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



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;

using System.IO;

using System.Collections.ObjectModel;

// Json lib
using System.Web.Script.Serialization;
using LanguageLib;

// Classe con funzioni di conversione DEC-HEX
using Raccolta_funzioni_parser;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for TemplateEditor.xaml
    /// </summary>
    public partial class TemplateEditor : Window
    {
        MainWindow main;
        String pathToConfiguration = "";

        ObservableCollection<ModBus_Item> list_coilsTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_inputsTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_inputRegistersTable = new ObservableCollection<ModBus_Item>();
        ObservableCollection<ModBus_Item> list_holdingRegistersTable = new ObservableCollection<ModBus_Item>();

        ObservableCollection<Group_Item> list_groups = new ObservableCollection<Group_Item>();

        Language lang;

        public TemplateEditor(MainWindow main_)
        {
            InitializeComponent();

            main = main_;
            pathToConfiguration = main.pathToConfiguration;

            dataGridViewCoils.ItemsSource = list_coilsTable;
            dataGridViewInput.ItemsSource = list_inputsTable;
            dataGridViewInputRegister.ItemsSource = list_inputRegistersTable;
            dataGridViewHolding.ItemsSource = list_holdingRegistersTable;

            dataGridViewGroups.ItemsSource = list_groups;

            main.Dispatcher.Invoke((Action)delegate
            {
                if (main.tabControlMain.SelectedIndex > 0 && main.tabControlMain.SelectedIndex < 5)
                {
                    tabControlTemplate.SelectedIndex = main.tabControlMain.SelectedIndex - 1;
                }
            });

            lang = new Language(this);
            lang.loadLanguageTemplate(main.language);

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);

            this.Title = main_.Title + " - Template Editor";

            // Predisposizione DarkMode
            /*if ((bool)main.CheckBoxDarkMode.IsChecked)
            {
                GridMain.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                textBoxCoilsOffset.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                textBoxCoilsOffset.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                textBoxInputOffset.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                textBoxInputOffset.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                textBoxInputRegOffset.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                textBoxInputRegOffset.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                textBoxHoldingOffset.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                textBoxHoldingOffset.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                TextBoxTemplateLabel.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                TextBoxTemplateLabel.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                TextBoxTemplateNotes.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                TextBoxTemplateNotes.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                LabelInfoMappings.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                LabelInfoMappings.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                LabelRegCoils.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                LabelOffCoil.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;

                LabelRegInput.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                LabelOffInput.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;

                LabelRegInputReg.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                LabelOffInputReg.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;

                LabelRegHolding.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                LabelOffHolding.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;

                LabelNotesLabel.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                LabelNotesNotes.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                LabelNotesGroups.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight; ;

                LabelImportExport.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight; ;

                if (main.darkMode)
                {
                    GridCoils.Background = main.BackGroundDark;
                    GridInputs.Background = main.BackGroundDark;
                    GridInputRegister.Background = main.BackGroundDark;
                    GridHoldingRegister.Background = main.BackGroundDark;

                    GridNotes.Background = main.BackGroundDark;
                    GridInfoMappings.Background = main.BackGroundDark;
                }

                // BorderBrush
                Setter SetBorderBrush = new Setter();
                SetBorderBrush.Property = BorderBrushProperty;
                SetBorderBrush.Value = main.darkMode ? main.BackGroundDark : main.BackGroundLight2;

                // Background
                Setter SetBackgroundProperty = new Setter();
                SetBackgroundProperty.Property = BackgroundProperty;
                SetBackgroundProperty.Value = main.darkMode ? main.BackGroundDark : main.BackGroundLight2;

                // Stile custom per cella standard
                Style NewStyle = new Style();
                NewStyle.Setters.Add(SetBorderBrush);
                NewStyle.Setters.Add(SetBackgroundProperty);

                dataGridViewCoils.CellStyle = NewStyle;
                dataGridViewInput.CellStyle = NewStyle;
                dataGridViewInputRegister.CellStyle = NewStyle;
                dataGridViewHolding.CellStyle = NewStyle;

                dataGridViewGroups.CellStyle = NewStyle;
            }*/
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            salvaTemplate();

            main.SaveConfiguration_v2(false);

            // Se esiste una nuova versione del file di configurazione uso l'ultima, altrimenti carico il modello precedente
            if (File.Exists(main.localPath + "\\Json\\" + pathToConfiguration + "\\Config.json"))
            {
                main.LoadConfiguration_v2();
            }
            else
            {
                main.LoadConfiguration_v1();
            }

        }

        private bool salvaTemplate()
        {
            try
            {
                var template = new TEMPLATE();

                template.comboBoxCoilsOffset_ = comboBoxCoilsOffset.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxInputOffset_ = comboBoxInputOffset.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxInputRegOffset_ = comboBoxInputRegOffset.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxHoldingOffset_ = comboBoxHoldingOffset.SelectedIndex == 1 ? "HEX" : "DEC";

                template.comboBoxCoilsRegistri_ = comboBoxCoilsRegistri.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxInputRegistri_ = comboBoxInputRegistri.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxInputRegRegistri_ = comboBoxInputRegRegistri.SelectedIndex == 1 ? "HEX" : "DEC";
                template.comboBoxHoldingRegistri_ = comboBoxHoldingRegistri.SelectedIndex == 1 ? "HEX" : "DEC";

                template.textBoxCoilsOffset_ = textBoxCoilsOffset.Text;
                template.textBoxInputOffset_ = textBoxInputOffset.Text;
                template.textBoxInputRegOffset_ = textBoxInputRegOffset.Text;
                template.textBoxHoldingOffset_ = textBoxHoldingOffset.Text;

                template.dataGridViewCoils = list_coilsTable;
                template.dataGridViewInput = list_inputsTable;
                template.dataGridViewInputRegister = list_inputRegistersTable;
                template.dataGridViewHolding = list_holdingRegistersTable;

                template.Groups = list_groups;

                template.textBoxTemplateLabel_ = TextBoxTemplateLabel.Text;
                template.textBoxTemplateNotes_ = TextBoxTemplateNotes.Text;

                JavaScriptSerializer jss = new JavaScriptSerializer();
                jss.MaxJsonLength = main.MaxJsonLength;
                string file_content = jss.Serialize(template);

                File.WriteAllText(main.localPath + "/Json/" + pathToConfiguration + "/Template.json", file_content);

                Console.WriteLine("Caricata configurazione precedente\n");

                return true;
            }
            catch (Exception err)
            {
                Console.WriteLine("Errore caricamento configurazione\n");
                Console.WriteLine(err);

                return false;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        { 
            try
            {
                string file_content = File.ReadAllText(main.localPath + "/Json/" + pathToConfiguration + "/Template.json");

                JavaScriptSerializer jss = new JavaScriptSerializer();
                jss.MaxJsonLength = main.MaxJsonLength;
                TEMPLATE template = jss.Deserialize<TEMPLATE>(file_content);

                comboBoxCoilsOffset.SelectedIndex = template.comboBoxCoilsOffset_ == "HEX" ? 1 : 0;
                comboBoxInputOffset.SelectedIndex = template.comboBoxInputOffset_ == "HEX" ? 1 : 0;
                comboBoxInputRegOffset.SelectedIndex = template.comboBoxInputRegOffset_ == "HEX" ? 1 : 0;
                comboBoxHoldingOffset.SelectedIndex = template.comboBoxHoldingOffset_ == "HEX" ? 1 : 0;

                comboBoxCoilsRegistri.SelectedIndex = template.comboBoxCoilsRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegistri.SelectedIndex = template.comboBoxInputRegistri_ == "HEX" ? 1 : 0;
                comboBoxInputRegRegistri.SelectedIndex = template.comboBoxInputRegRegistri_ == "HEX" ? 1 : 0;
                comboBoxHoldingRegistri.SelectedIndex = template.comboBoxHoldingRegistri_ == "HEX" ? 1 : 0;

                textBoxCoilsOffset.Text = template.textBoxCoilsOffset_;
                textBoxInputOffset.Text = template.textBoxInputOffset_;
                textBoxInputRegOffset.Text = template.textBoxInputRegOffset_;
                textBoxHoldingOffset.Text = template.textBoxHoldingOffset_;

                if(template.textBoxTemplateLabel_ != null)
                {
                    TextBoxTemplateLabel.Text = template.textBoxTemplateLabel_;
                }
                else
                {
                    TextBoxTemplateLabel.Text = "";
                }

                if(template.textBoxTemplateNotes_ != null)
                {
                    TextBoxTemplateNotes.Text = template.textBoxTemplateNotes_;
                }
                else
                {
                    TextBoxTemplateNotes.Text = "";
                }

                // Tabella coils
                foreach (ModBus_Item item in template.dataGridViewCoils)
                {
                    list_coilsTable.Add(item);
                }

                // Tabella inputs
                foreach (ModBus_Item item in template.dataGridViewInput)
                {
                    list_inputsTable.Add(item);
                }

                // Tabella input registers
                foreach (ModBus_Item item in template.dataGridViewInputRegister)
                {
                    list_inputRegistersTable.Add(item);
                }

                // Tabella holdings
                foreach (ModBus_Item item in template.dataGridViewHolding)
                {
                    list_holdingRegistersTable.Add(item);
                }

                // Tabella groups
                if (template.Groups != null)
                {
                    foreach (Group_Item item in template.Groups)
                    {
                        list_groups.Add(item);
                    }
                }

                Console.WriteLine("Caricata configurazione precedente\n");
            }
            catch(Exception err)
            {
                Console.WriteLine("Errore caricamento configurazione\n");
                Console.WriteLine(err);
            }
        }

        private void salvaToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (salvaTemplate())
            {
                MessageBox.Show("Template salvato correttamente", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Errore salvataggio template", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape)
            {
                this.Close();
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if(Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                switch (e.Key)
                {
                    case Key.D1:
                        tabControlTemplate.SelectedIndex = 0;
                        break;

                    case Key.D2:
                        tabControlTemplate.SelectedIndex = 1;
                        break;

                    case Key.D3:
                        tabControlTemplate.SelectedIndex = 2;
                        break;

                    case Key.D4:
                        tabControlTemplate.SelectedIndex = 3;
                        break;

                    case Key.D5:
                        tabControlTemplate.SelectedIndex = 4;
                        break;
                    
                    case Key.S:
                        salvaTemplate();
                        break;
                }
            }
        }

        private void ButtonImportCsvCoils_Click(object sender, RoutedEventArgs e)
        {
            importCsv(list_coilsTable, comboBoxCoilsOffset, textBoxCoilsOffset, comboBoxCoilsRegistri);
        }

        private void ButtonImportCsvInputs_Click(object sender, RoutedEventArgs e)
        {
            importCsv(list_inputsTable, comboBoxInputOffset, textBoxInputOffset, comboBoxInputRegistri);
        }

        private void ButtonImportCsvInputRegisters_Click(object sender, RoutedEventArgs e)
        {
            importCsv(list_inputRegistersTable, comboBoxInputRegOffset, textBoxInputRegOffset, comboBoxInputRegRegistri);
        }

        private void ButtonImportCsvHoldingRegisters_Click(object sender, RoutedEventArgs e)
        {
            importCsv(list_holdingRegistersTable, comboBoxHoldingOffset, textBoxHoldingOffset, comboBoxHoldingRegistri);
        }

        private void ButtonExportCsvCoils_Click(object sender, RoutedEventArgs e)
        {
            string offset = comboBoxCoilsOffset.SelectedIndex == 1 ? "0x" : "" + textBoxCoilsOffset.Text;
            exportCsv(list_coilsTable, "_Coils", offset, comboBoxCoilsRegistri.SelectedIndex == 1);
        }

        private void ButtonExportCsvInputs_Click(object sender, RoutedEventArgs e)
        {
            string offset = comboBoxInputOffset.SelectedIndex == 1 ? "0x" : "" + textBoxInputOffset.Text;
            exportCsv(list_inputsTable, "_Inputs", offset, comboBoxInputRegistri.SelectedIndex == 1);
        }

        private void ButtonExportCsvInputRegisters_Click(object sender, RoutedEventArgs e)
        {
            string offset = comboBoxInputRegOffset.SelectedIndex == 1 ? "0x" : "" + textBoxInputRegOffset.Text;
            exportCsv(list_inputRegistersTable, "_InputRegisters", offset, comboBoxInputRegRegistri.SelectedIndex == 1);
        }

        private void ButtonExportCsvHoldingRegisters_Click(object sender, RoutedEventArgs e)
        {
            string offset = comboBoxHoldingOffset.SelectedIndex == 1 ? "0x": "" + textBoxHoldingOffset.Text;
            exportCsv(list_holdingRegistersTable, "_HoldingRegisters", offset, comboBoxHoldingRegistri.SelectedIndex == 1);
        }

        public void exportCsv(ObservableCollection<ModBus_Item> collection, String append, String offset, bool registerHex)
        {
            SaveFileDialog window = new SaveFileDialog();

            window.Filter = "csv Files | *.csv";
            window.DefaultExt = ".csv";
            window.FileName = main.pathToConfiguration + append + ".csv";

            if ((bool)window.ShowDialog())
            {
                if (window.FileName.IndexOf("csv") != -1)
                {
                    String content = "Offset,Register,Value,Notes,Mappings\n";

                    foreach (ModBus_Item item in collection)
                    {
                        if (item != null)
                        {
                            content += offset + (registerHex ? ",0x" : ",") + item.Register + "," + item.Value + "," + item.Notes + "," + item.Mappings + "," + item.Group + "\n";
                        }
                    }

                    File.WriteAllText(window.FileName, content);
                }
            }
        }

        public void importCsv(ObservableCollection<ModBus_Item> collection, ComboBox comboBox, TextBox textBox, ComboBox comboBoxReg)
        {
            collection.Clear();

            OpenFileDialog window = new OpenFileDialog();

            window.Filter = "csv Files | *.csv";
            window.DefaultExt = ".csv";

            if ((bool)window.ShowDialog())
            {
                string content = File.ReadAllText(window.FileName);
                string[] splitted = content.Split('\n');

                for (int i = 1; i < splitted.Count(); i++)
                {
                    ModBus_Item item = new ModBus_Item();

                    try
                    {
                        if (splitted[i].Length > 2)
                        {
                            string offset = splitted[i].Split(',')[0];

                            if (offset.IndexOf("0x") != -1)
                            {
                                comboBox.SelectedIndex = 1;
                                offset = offset.Substring(2);
                            }

                            textBox.Text = offset;

                            item.Register = splitted[i].Split(',')[1];

                            if (item.Register.IndexOf("0x") != -1)
                            {
                                comboBoxReg.SelectedIndex = 1;
                                item.Register = item.Register.Substring(2);
                                item.RegisterUInt = UInt16.Parse(item.Register, System.Globalization.NumberStyles.HexNumber);
                            }
                            else
                            {
                                item.RegisterUInt = UInt16.Parse(item.Register, System.Globalization.NumberStyles.Integer);
                            }

                            item.Notes = splitted[i].Split(',')[3];
                            item.Mappings = splitted[i].Split(',')[4];

                            if (splitted[i].Length > 5)
                                item.Group = splitted[i].Split(',')[5];
                            else
                                item.Group = "";

                            collection.Add(item);
                        }
                    }
                    catch
                    {
                    }
                }
            }
        }

        private void importProfileModbusSimulatorMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(lang.languageTemplate["strings"]["importProfileWarning"], "Alert", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                // Cancello le liste attuali
                list_coilsTable.Clear();
                list_inputsTable.Clear();
                list_inputRegistersTable.Clear();
                list_holdingRegistersTable.Clear();

                list_groups.Clear();

                // Resetto le varie combo a decimal
                comboBoxCoilsRegistri.SelectedIndex = 0;
                comboBoxCoilsOffset.SelectedIndex = 0;

                comboBoxInputRegistri.SelectedIndex = 0;
                comboBoxInputOffset.SelectedIndex = 0;

                comboBoxInputRegRegistri.SelectedIndex = 0;
                comboBoxInputRegOffset.SelectedIndex = 0;

                comboBoxHoldingRegistri.SelectedIndex = 0;
                comboBoxHoldingOffset.SelectedIndex = 0;

                // Riporto gli offset a 0
                textBoxCoilsOffset.Text = "0";
                textBoxInputOffset.Text = "0";
                textBoxInputRegOffset.Text = "0";
                textBoxHoldingOffset.Text = "0";

                OpenFileDialog window = new OpenFileDialog();

                window.Filter = "json Files | *.json";
                window.DefaultExt = ".json";

                if ((bool)window.ShowDialog())
                {
                    JavaScriptSerializer jss = new JavaScriptSerializer();
                    jss.MaxJsonLength = main.MaxJsonLength;
                    dynamic profile = jss.DeserializeObject(File.ReadAllText(window.FileName));

                    if (!profile.ContainsKey("type"))
                    {
                        MessageBox.Show("Profile file is not a valid json for this application", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    if (profile["type"] != "ModBusSlave")
                    {
                        MessageBox.Show("Profile file is not a valid json for this application", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    main.Dispatcher.Invoke((Action)delegate
                    {
                        try
                        {
                            main.textBoxModbusAddress.Text = profile["slave_id"][0].ToString();
                        }
                        catch
                        {
                            main.textBoxModbusAddress.Text = "255";
                        }
                    });

                    TextBoxTemplateLabel.Text = profile["label"];
                    TextBoxTemplateNotes.Text = profile["notes"];

                    // profile["label"];
                    // profile["slave_id"];

                    string[] keys = new string[] { "di", "co", "ir", "hr" };

                    foreach (string key in keys)
                    {
                        // profile[key]["len"];

                        foreach (dynamic regConfig in profile[key]["data"])
                        {
                            ModBus_Item item = new ModBus_Item();

                            item.Register = regConfig["reg"];
                            item.Notes = regConfig["label"];
                            item.Mappings = "";

                            if (regConfig.ContainsKey("mappings"))
                                item.Mappings = regConfig["mappings"];

                            if (regConfig.ContainsKey("group"))
                                item.Group = regConfig["group"];

                            switch (key)
                            {
                                case "di":
                                    list_inputsTable.Add(item);
                                    break;

                                case "co":
                                    list_coilsTable.Add(item);
                                    break;

                                case "ir":
                                    list_inputRegistersTable.Add(item);
                                    break;

                                case "hr":
                                    list_holdingRegistersTable.Add(item);
                                    break;
                            }
                        }
                    }

                    foreach(dynamic group in profile["groups"])
                    {
                        Group_Item newGroup = (Group_Item)group;
                        list_groups.Add(newGroup);
                    }
                }
            }
        }

        private void exportProfileModbusSimulatorMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Parser P = new Parser();
                SaveFileDialog saveFileDialog = new SaveFileDialog();

                saveFileDialog.Filter = "json files (*.json)|*.json";

                if (TextBoxTemplateLabel.Text.Length > 0)
                {
                    saveFileDialog.FileName = TextBoxTemplateLabel.Text + ".json";

                }
                else if (main.pathToConfiguration != null)
                {
                    if(main.pathToConfiguration.Length > 0)
                    {
                        saveFileDialog.FileName = main.pathToConfiguration;
                    }
                }

                // saveFileDialog.RestoreDirectory = true;

                if ((bool)saveFileDialog.ShowDialog())
                {
                    string profilePath = saveFileDialog.FileName;

                    MOD_SlaveProfile slave = new MOD_SlaveProfile();

                    slave.slave_id = new List<int>();

                    main.Dispatcher.Invoke((Action)delegate
                    {
                        slave.slave_id.Add(int.Parse(main.textBoxModbusAddress.Text));
                    });

                    slave.label = TextBoxTemplateLabel.Text;
                    slave.notes = TextBoxTemplateNotes.Text;
                    slave.type = "ModBusSlave";

                    slave.di = new MOD_RegProfile();
                    slave.co = new MOD_RegProfile();
                    slave.ir = new MOD_RegProfile();
                    slave.hr = new MOD_RegProfile();

                    slave.di.len = 0;
                    slave.co.len = 0;
                    slave.ir.len = 0;
                    slave.hr.len = 0;

                    slave.di.data = new List<MOD_Reg>();
                    slave.co.data = new List<MOD_Reg>();
                    slave.ir.data = new List<MOD_Reg>();
                    slave.hr.data = new List<MOD_Reg>();

                    // Inputs table
                    foreach (ModBus_Item item in list_inputsTable)
                    {
                        MOD_Reg md = new MOD_Reg();

                        md.reg = (P.uint_parser(item.Register, comboBoxInputRegistri) + P.uint_parser(textBoxInputOffset.Text, comboBoxInputOffset)).ToString();
                        md.label = item.Notes;
                        md.value = 0;
                        md.mappings = "";
                        md.group = item.Group;

                        if ((int.Parse(md.reg) + 1) > slave.di.len)
                            slave.di.len = int.Parse(md.reg) + 1;

                        slave.di.data.Add(md);
                    }

                    // Coils table
                    foreach (ModBus_Item item in list_coilsTable)
                    {
                        MOD_Reg md = new MOD_Reg();

                        md.reg = (P.uint_parser(item.Register, comboBoxCoilsRegistri) + P.uint_parser(textBoxCoilsOffset.Text, comboBoxCoilsOffset)).ToString();
                        md.label = item.Notes;
                        md.value = 0;
                        md.mappings = "";
                        md.group = item.Group;

                        if ((int.Parse(md.reg) + 1) > slave.co.len)
                            slave.co.len = int.Parse(md.reg) + 1;

                        slave.co.data.Add(md);
                    }

                    // Input registers table
                    foreach (ModBus_Item item in list_inputRegistersTable)
                    {
                        MOD_Reg md = new MOD_Reg();

                        md.reg = (P.uint_parser(item.Register, comboBoxInputRegRegistri) + P.uint_parser(textBoxInputRegOffset.Text, comboBoxInputRegOffset)).ToString();
                        md.label = item.Notes;
                        md.value = 0;
                        md.mappings = item.Mappings;
                        md.group = item.Group;

                        if ((int.Parse(md.reg) + 1) > slave.ir.len)
                            slave.ir.len = int.Parse(md.reg) + 1;

                        slave.ir.data.Add(md);
                    }

                    // Holding registers table
                    foreach (ModBus_Item item in list_holdingRegistersTable)
                    {
                        MOD_Reg md = new MOD_Reg();

                        md.reg = (P.uint_parser(item.Register, comboBoxHoldingRegistri) + P.uint_parser(textBoxHoldingOffset.Text, comboBoxHoldingOffset)).ToString();
                        md.label = item.Notes;
                        md.value = 0;
                        md.mappings = item.Mappings;
                        md.group = item.Group;

                        if ((int.Parse(md.reg) + 1) > slave.hr.len)
                            slave.hr.len = int.Parse(md.reg) + 1;

                        slave.hr.data.Add(md);
                    }

                    slave.groups = new List<Group_Item>();

                    // Groups
                    foreach (Group_Item group in list_groups)
                        slave.groups.Add(group);

                    JavaScriptSerializer jss = new JavaScriptSerializer();
                    jss.MaxJsonLength = main.MaxJsonLength;
                    File.WriteAllText(profilePath, jss.Serialize(slave));
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
                MessageBox.Show("Errore salvataggio configurazione", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ButtonImportCsvGroups_Click(object sender, RoutedEventArgs e)
        {
            list_groups.Clear();

            OpenFileDialog window = new OpenFileDialog();

            window.Filter = "csv Files | *.csv";
            window.DefaultExt = ".csv";

            if ((bool)window.ShowDialog())
            {
                string content = File.ReadAllText(window.FileName);
                string[] splitted = content.Split('\n');

                for (int i = 1; i < splitted.Count(); i++)
                {
                    ModBus_Item item = new ModBus_Item();

                    try
                    {
                        if (splitted[i].Length > 1)
                        {
                            Group_Item group = new Group_Item();

                            group.Group = splitted[i].Split(',')[0];
                            group.Label = splitted[i].Split(',')[1];

                            list_groups.Add(group);
                        }
                    }
                    catch
                    {
                    }
                }
            }
        }

        private void ButtonExportCsvGroups_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog window = new SaveFileDialog();

            window.Filter = "csv Files | *.csv";
            window.DefaultExt = ".csv";
            window.FileName = main.pathToConfiguration + "_Groups.csv";

            if ((bool)window.ShowDialog())
            {
                if (window.FileName.IndexOf("csv") != -1)
                {
                    String content = "Group,Label,\n";

                    foreach (Group_Item group in list_groups)
                    {
                        if (group != null)
                        {
                            content += group.Group + "," + group.Label + ",\n";
                        }
                    }

                    File.WriteAllText(window.FileName, content);
                }
            }
        }
    }

    // Classe per caricare dati dal file di configurazione json
    public class TEMPLATE
    {
        public string comboBoxCoilsOffset_ { get; set; }
        public string comboBoxInputOffset_ { get; set; }
        public string comboBoxInputRegOffset_ { get; set; }
        public string comboBoxHoldingOffset_ { get; set; }

        public string comboBoxCoilsRegistri_ { get; set; }
        public string comboBoxInputRegistri_ { get; set; }
        public string comboBoxInputRegRegistri_ { get; set; }
        public string comboBoxHoldingRegistri_ { get; set; }

        public string textBoxCoilsOffset_ { get; set; }
        public string textBoxInputOffset_ { get; set; }
        public string textBoxInputRegOffset_ { get; set; }
        public string textBoxHoldingOffset_ { get; set; }

        public ObservableCollection<ModBus_Item> dataGridViewCoils { get; set; }
        public ObservableCollection<ModBus_Item> dataGridViewInput { get; set; }
        public ObservableCollection<ModBus_Item> dataGridViewInputRegister { get; set; }
        public ObservableCollection<ModBus_Item> dataGridViewHolding { get; set; }
        public string textBoxTemplateLabel_ { get; set; }
        public string textBoxTemplateNotes_ { get; set; }

        public ObservableCollection<Group_Item> Groups { get; set; }
    }

    public class ModBus_Item
    {
        public string Register { get; set; }
        public UInt16 RegisterUInt { get; set; }
        public string Value { get; set; }
        public string ValueBin { get; set; }
        public string ValueConverted { get; set; }
        public string Notes { get; set; }
        public string Mappings { get; set; }
        public string Group { get; set; }
        public string Foreground { get; set; }
        public string Background { get; set; }
    }

    public class Group_Item
    {
        public string Group { get; set; }
        public string Label { get; set; }
    }
    public class MOD_SlaveProfile
    {
        public List<int> slave_id { get; set; }
        public string label { get; set; }
        public string notes { get; set; }
        public string type { get; set; }

        public MOD_RegProfile di { get; set; }
        public MOD_RegProfile co { get; set; }
        public MOD_RegProfile ir { get; set; }
        public MOD_RegProfile hr { get; set; }

        public List<Group_Item> groups { get; set; }
    }
    public class MOD_RegProfile
    {
        public int len { get; set; }
        public List<MOD_Reg> data { get; set; }
    }

    public class MOD_Reg
    {
        public string reg { get; set; }
        public Int32 value { get; set; }
        public string label { get; set; }
        public string mappings { get; set; }
        public string group { get; set; }
    }
}
