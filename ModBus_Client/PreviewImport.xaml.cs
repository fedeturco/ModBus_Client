

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
using System.IO;
using System.Collections.ObjectModel;
using System.Threading;
using System.Windows.Shell;
using Microsoft.Win32;

// Custom Modbus Library
using ModBusMaster_Chicco;

// Json LIBs
using System.Web.Script.Serialization;

// Classe con funzioni di conversione DEC-HEX
using Raccolta_funzioni_parser;
using System.Security.Cryptography;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for PreviewImport.xaml
    /// </summary>
    public partial class PreviewImport : Window
    {
        MainWindow main;
        String[] filePaths;
        byte FC = 0;

        ObservableCollection<ModBus_Import> list_import = new ObservableCollection<ModBus_Import>();

        bool useMultiple = false; 
        bool closeAfterWrite = true;
        bool abortOnError = true;

        bool importRunning = false;     // True while sending modbus write commands passing through registers list
        bool abortExit = false;         // True for stopping window from sending commands qhile exiting

        Parser P = new Parser();

        int stepScrollIntoView = 1;
        public PreviewImport(MainWindow main_, String[] filePaths_, byte FC_)
        {
            InitializeComponent();

            main = main_;
            filePaths = filePaths_;
            FC = FC_;

            if (FC == 5)
            {
                CheckBoxWriteMultiple.Content = main.lang.languageTemplate["strings"]["writeMultipleCoils"];
                LabelNrOf.Content = main.lang.languageTemplate["strings"]["labelMultipleCoils"];
            }
            if (FC == 6)
            {
                CheckBoxWriteMultiple.Content = main.lang.languageTemplate["strings"]["writeMultipleRegisters"];
                LabelNrOf.Content = main.lang.languageTemplate["strings"]["labelMultipleRegisters"];
            }

            if ((bool)main.CheckBoxDarkMode.IsChecked)
            {
                GridMain.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                CheckBoxWriteMultiple.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                // CheckBoxWriteMultiple.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                CheckBoxCloseWindowAfterWrite.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                // CheckBoxCloseWindowAfterWrite.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                CheckBoxAbortOnError.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;

                dataGridPreview.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                LabelNrOf.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                LabelNrOf.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

                TextBoxNrOfRegs.Foreground = main.darkMode ? main.ForeGroundDark : main.ForeGroundLight;
                TextBoxNrOfRegs.Background = main.darkMode ? main.BackGroundDark : main.BackGroundLight;

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

                dataGridPreviewOffset.CellStyle = NewStyle;
                dataGridPreviewRegister.CellStyle = NewStyle;
                dataGridPreviewValue.CellStyle = NewStyle;
                dataGridPreviewNotes.CellStyle = NewStyle;
            }

            TaskbarItemInfo = new TaskbarItemInfo();

            //CheckBoxWriteMultiple.Content = main.CheckBoxPreviewimportWriteMultipleRegisters.Content;
            CheckBoxWriteMultiple.IsChecked = main.CheckBoxPreviewimportWriteMultipleRegisters.IsChecked;
            CheckBoxCloseWindowAfterWrite.Content = main.CheckBoxPreviewImportCloseWindowAfterWrite.Content;
            CheckBoxCloseWindowAfterWrite.IsChecked = main.CheckBoxPreviewImportCloseWindowAfterWrite.IsChecked;
            CheckBoxAbortOnError.Content = main.CheckBoxPreviewImportAbortOnError.Content;
            CheckBoxAbortOnError.IsChecked = main.CheckBoxPreviewImportAbortOnError.IsChecked;
        }

        public void ImportClipboard()
        {
            bool errorParsing = false;

            try
            {
                string clipContent = Clipboard.GetText(TextDataFormat.Text);
                string[] registers = clipContent.Replace("\r", "").Replace("\t", ",").Split('\n');

                for (int i = 0; i < registers.Length; i++)
                {
                    string[] row = registers[i].Split(',');

                    if (row.Length < 3)
                        continue;

                    if (row[0].ToLower().IndexOf("offset") != -1)
                        continue;

                    ModBus_Import item = new ModBus_Import();

                    UInt16 parsed = 0;

                    if (!UInt16.TryParse(row[0].Replace("0x", ""), row[0].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out parsed))
                    {
                        errorParsing = true;
                        continue;
                    }

                    item.Offset = row[0];
                    item.OffsetUInt = parsed;

                    if (!UInt16.TryParse(row[1].Replace("0x", ""), row[1].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out parsed))
                    {
                        errorParsing = true;
                        continue;
                    }

                    item.Register = row[1];
                    item.RegisterUInt = parsed;

                    if (!UInt16.TryParse(row[2].Replace("0x", ""), row[2].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out parsed))
                    {
                        errorParsing = true;
                        continue;
                    }

                    item.Value = row[2];
                    item.ValueUInt = parsed;

                    if(row.Length >= 4)
                        item.Notes = row[3];
                    else
                        item.Notes = "";

                    if ((bool)main.CheckBoxDarkMode.IsChecked)
                    {
                        item.Foreground = main.ForeGroundDarkStr;
                        item.Background = main.BackGroundDarkStr;
                        item.ForegroundDef = main.ForeGroundDarkStr;
                        item.BackgroundDef = main.BackGroundDarkStr;
                    }
                    else
                    {
                        item.Foreground = main.ForeGroundLightStr;
                        item.Background = main.BackGroundLight2Str;
                        item.ForegroundDef = main.ForeGroundLightStr;
                        item.BackgroundDef = main.BackGroundLight2Str;
                    }

                    list_import.Add(item);
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridPreview.ScrollIntoView(list_import.LastOrDefault<ModBus_Import>());
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            if (errorParsing)
                MessageBox.Show(main.lang.languageTemplate["strings"]["importingError"], "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public void ExportClipboard()
        {
            try
            {
                string clipContent = "";
                foreach (ModBus_Import item in list_import)
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

        public void ImportFiles(String[] filePaths)
        {
            bool errorParsing = false;

            try
            {
                if (filePaths != null)
                {
                    String newTitle = "PreviewImport - Files:";

                    foreach (String filePath in filePaths)
                    {
                        newTitle += " " + System.IO.Path.GetFileName(filePath);
                    }

                    if (FC == 5)
                        newTitle += " - FC05/FC15";
                    else
                        newTitle += " - FC06/FC16";

                    this.Title = newTitle;

                    foreach (String filePath in filePaths)
                    {
                        // CSV File
                        if (filePath.IndexOf(".csv") != -1)
                        {
                            string[] registers = File.ReadAllText(filePath).Replace("\r", "").Split('\n');
                            for (int i = 0; i < registers.Length; i++)
                            {
                                string[] row = registers[i].Split(',');

                                if (row.Length < 3)
                                    continue;

                                if (row[0].ToLower().IndexOf("offset") != -1)
                                    continue;

                                ModBus_Import item = new ModBus_Import();

                                UInt16 parsed = 0;

                                if(!UInt16.TryParse(row[0].Replace("0x", ""), row[0].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out parsed))
                                {
                                    errorParsing = true;
                                    continue;
                                }

                                item.Offset = row[0];
                                item.OffsetUInt = parsed;

                                if(!UInt16.TryParse(row[1].Replace("0x", ""), row[1].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out parsed))
                                {
                                    errorParsing = true;
                                    continue;
                                }

                                item.Register = row[1];
                                item.RegisterUInt = parsed;

                                if (!UInt16.TryParse(row[2].Replace("0x", ""), row[2].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer, null, out parsed))
                                {
                                    errorParsing = true;
                                    continue;
                                }

                                item.Value = row[2];
                                item.ValueUInt = parsed;

                                if(row.Length >= 4)
                                    item.Notes = row[3];
                                else
                                    item.Notes = "";

                                if ((bool)main.CheckBoxDarkMode.IsChecked)
                                {
                                    item.Foreground = main.ForeGroundDarkStr;
                                    item.Background = main.BackGroundDarkStr;
                                    item.ForegroundDef = main.ForeGroundDarkStr;
                                    item.BackgroundDef = main.BackGroundDarkStr;
                                }
                                else
                                {
                                    item.Foreground = main.ForeGroundLightStr;
                                    item.Background = main.BackGroundLight2Str;
                                    item.ForegroundDef = main.ForeGroundLightStr;
                                    item.BackgroundDef = main.BackGroundLight2Str;
                                }

                                list_import.Add(item);
                            }
                        }

                        // JSON File
                        if (filePath.IndexOf(".json") != -1)
                        {
                            JavaScriptSerializer jss = new JavaScriptSerializer();
                            jss.MaxJsonLength = main.MaxJsonLength;
                            dynamic obj = jss.DeserializeObject(File.ReadAllText(filePath));

                            foreach (dynamic reg in obj["items"])
                            {
                                ModBus_Import item = new ModBus_Import();

                                item.Offset = reg["Offset"];
                                item.OffsetUInt = UInt16.Parse(reg["Offset"], reg["Offset"].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                                item.Register = reg["Register"];
                                item.RegisterUInt = UInt16.Parse(reg["Register"], reg["Register"].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                                item.Value = reg["Value"];
                                item.ValueUInt = UInt16.Parse(reg["Value"], reg["Value"].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                                item.Notes = reg["Notes"] is null ? "" : reg["Notes"];

                                if ((bool)main.CheckBoxDarkMode.IsChecked)
                                {
                                    item.Foreground = main.ForeGroundDarkStr;
                                    item.Background = main.BackGroundDarkStr;
                                    item.ForegroundDef = main.ForeGroundDarkStr;
                                    item.BackgroundDef = main.BackGroundDarkStr;
                                }
                                else
                                {
                                    item.Foreground = main.ForeGroundLightStr;
                                    item.Background = main.BackGroundLight2Str;
                                    item.ForegroundDef = main.ForeGroundLightStr;
                                    item.BackgroundDef = main.BackGroundLight2Str;
                                }

                                list_import.Add(item);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            if (errorParsing)
                MessageBox.Show(main.lang.languageTemplate["strings"]["importingError"], "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            List<ModBus_Item> content = new List<ModBus_Item>();

            try
            {
                if (filePaths != null)
                {
                    ImportFiles(filePaths);
                }
                else
                {
                    this.Title = "PreviewImport";
                    if (FC == 5)
                        this.Title += " - FC05/FC15";
                    else
                        this.Title += " - FC06/FC16";
                    ImportClipboard();
                }
            }
            catch(Exception err)
            {
                MessageBox.Show(main.lang.languageTemplate["strings"]["importingErrorFiles"], "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                Console.WriteLine(err);
            }

            dataGridPreview.ItemsSource = list_import;

            buttonImportPreview.Focus();
        }

        private void buttonImportPreview_Click(object sender, RoutedEventArgs e)
        {
            buttonImportPreview.IsEnabled = false;
            TextBoxNrOfRegs.IsEnabled = false;
            CheckBoxWriteMultiple.IsEnabled = false;

            Thread t = new Thread(new ThreadStart(ImportPreview));
            t.Start();
        }

        public void ImportPreview()
        {
            importRunning = true;

            try
            {
                // Aggiorno stato registri prima di inviarli
                for (int i = 0; i < list_import.Count; i++)
                {
                    list_import[i].OffsetUInt = UInt16.Parse(list_import[i].Offset.Replace("0x", ""), list_import[i].Offset.ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);
                    list_import[i].RegisterUInt = UInt16.Parse(list_import[i].Register.Replace("0x", ""), list_import[i].Register.ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);
                    list_import[i].ValueUInt = UInt16.Parse(list_import[i].Value.Replace("0x", ""), list_import[i].Value.ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                    if (main.darkMode)
                    {
                        list_import[i].Foreground = main.ForeGroundDarkStr;
                        list_import[i].Background = main.BackGroundDarkStr;
                        list_import[i].ForegroundDef = main.ForeGroundDarkStr;
                        list_import[i].BackgroundDef = main.BackGroundDarkStr;
                    }
                    else
                    {
                        list_import[i].Foreground = main.ForeGroundLightStr;
                        list_import[i].Background = main.BackGroundLight2Str;
                        list_import[i].ForegroundDef = main.ForeGroundLightStr;
                        list_import[i].BackgroundDef = main.BackGroundLight2Str;
                    }
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    dataGridPreview.ItemsSource = null;
                    dataGridPreview.ItemsSource = list_import;

                    dataGridPreview.ScrollIntoView(list_import[0]);
                });

                // Coils
                if (FC == 5)
                {
                    // Svuoto tabella coils
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        main.list_coilsTable.Clear();

                        TaskbarItemInfo.ProgressValue = 0;
                        TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
                    });

                    if (!useMultiple)
                    {
                        for (int i = 0; i < list_import.Count; i++)
                        {
                            if (abortExit)
                                break;

                            try
                            {
                                bool? result = main.ModBus.forceSingleCoil_05(
                                    byte.Parse(main.textBoxModbusAddress_),
                                    (UInt16)(list_import[i].OffsetUInt + list_import[i].RegisterUInt),
                                    list_import[i].ValueUInt,
                                    main.readTimeout
                                    );

                                if (result != null)
                                {
                                    main.insertRowsTable(
                                        main.list_coilsTable,
                                        main.list_template_coilsTable,
                                        P.uint_parser(main.textBoxCoilsOffset_, main.comboBoxCoilsOffset_),
                                        (UInt16)(list_import[i].OffsetUInt + list_import[i].RegisterUInt),
                                        new UInt16[] { list_import[i].ValueUInt },
                                        result == true ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                        main.comboBoxCoilsOffset_,
                                        main.comboBoxCoilsRegistri_,
                                        "DEC",
                                        false,
                                        true);

                                    list_import[i].Background = main.colorDefaultWriteCellStr;
                                }
                                else
                                {
                                    list_import[i].Background = main.colorErrorCellStr;
                                }

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    TaskbarItemInfo.ProgressValue = (double)(i) / (double)(list_import.Count);

                                    dataGridPreview.ItemsSource = null;
                                    dataGridPreview.ItemsSource = list_import;

                                    if (i + stepScrollIntoView < list_import.Count)
                                        dataGridPreview.ScrollIntoView(list_import[i + stepScrollIntoView]);
                                    else
                                        dataGridPreview.ScrollIntoView(list_import[i]);
                                });
                            }
                            catch (InvalidOperationException err)
                            {
                                if (err.Message.IndexOf("non-connected socket") != -1)
                                {
                                    main.SetTableDisconnectError(main.list_coilsTable, true);

                                    if (main.ModBus.type == main.ModBus_Def.TYPE_RTU || main.ModBus.type == main.ModBus_Def.TYPE_ASCII)
                                    {
                                        this.Dispatcher.Invoke((Action)delegate
                                        {
                                            main.buttonSerialActive_Click(null, null);
                                        });
                                    }
                                    else
                                    {
                                        this.Dispatcher.Invoke((Action)delegate
                                        {
                                            main.buttonTcpActive_Click(null, null);
                                        });
                                    }
                                }

                                Console.WriteLine(err);
                            }
                            catch (ModbusException err)
                            {
                                if (err.Message.IndexOf("Timed out") != -1)
                                {
                                    main.SetTableTimeoutError(main.list_coilsTable, false);
                                }
                                if (err.Message.IndexOf("ModBus ErrCode") != -1)
                                {
                                    main.SetTableModBusError(main.list_coilsTable, err, false);
                                }
                                if (err.Message.IndexOf("CRC Error") != -1)
                                {
                                    main.SetTableCrcError(main.list_coilsTable, false);
                                }

                                Console.WriteLine(err);

                                if (abortOnError)
                                    break;
                            }
                            catch (Exception err)
                            {
                                main.SetTableInternalError(main.list_coilsTable, true);
                                Console.WriteLine(err);

                                break;
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            main.dataGridViewCoils.ItemsSource = null;
                            main.dataGridViewCoils.ItemsSource = main.list_coilsTable;
                        });
                    }
                    else
                    {
                        UInt16 maxNumRegister = 0;

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (!UInt16.TryParse(TextBoxNrOfRegs.Text, out maxNumRegister))
                            {
                                MessageBox.Show(main.lang.languageTemplate["strings"]["importingErrorParse"], "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        });

                        UInt16 startAddr = 0;   // Start address FC15
                        int startIndex = 0;     // Index start address
                        int countRegs = 0;      // Number of registers aggregated
                        int currAddr = -1;      // Current address
                        int oldCurrAddr = -1;   // Previous address
                        bool[] toSendBuffer = new bool[maxNumRegister];

                        for (int i = 0; i <= list_import.Count; i++)
                        {
                            if (abortExit)
                                break;

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                TaskbarItemInfo.ProgressValue = (double)(i) / (double)(list_import.Count);
                            });

                            if (i < list_import.Count)
                                currAddr = list_import[i].RegisterUInt + list_import[i].OffsetUInt;

                            if (oldCurrAddr == -1)
                                startAddr = (UInt16)currAddr;

                            if ((currAddr != (oldCurrAddr + 1) || countRegs >= maxNumRegister || countRegs >= list_import.Count) && oldCurrAddr != -1)
                            {
                                // Invio
                                bool[] toSend = new bool[countRegs];
                                UInt16[] toSendUInt16 = new UInt16[countRegs];
                                for (int a = 0; a < countRegs; a++)
                                {
                                    toSend[a] = toSendBuffer[a];
                                    toSendUInt16[a] = (UInt16)(toSendBuffer[a] ? 1 : 0);
                                }

                                try
                                {
                                    bool? result = main.ModBus.forceMultipleCoils_15(
                                        byte.Parse(main.textBoxModbusAddress_),
                                        startAddr,
                                        toSend,
                                        main.readTimeout
                                        );

                                    if (result != null)
                                    {
                                        main.insertRowsTable(
                                                main.list_coilsTable,
                                                main.list_template_coilsTable,
                                                P.uint_parser(main.textBoxCoilsOffset_, main.comboBoxCoilsOffset_),
                                                startAddr,
                                                toSendUInt16,
                                                result == true ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                                main.comboBoxHoldingOffset_,
                                                main.comboBoxHoldingRegistri_,
                                                "DEC",
                                                false,
                                                true);

                                        for(int a = startIndex; a < (startIndex + countRegs); a++)
                                        {
                                            list_import[a].Background = main.colorDefaultWriteCellStr;
                                        }
                                    }
                                    else
                                    {
                                        for (int a = startIndex; a < (startIndex + countRegs); a++)
                                        {
                                            list_import[a].Background = main.colorErrorCellStr;
                                        }
                                    }

                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        TaskbarItemInfo.ProgressValue = (double)(i) / (double)(list_import.Count);

                                        dataGridPreview.ItemsSource = null;
                                        dataGridPreview.ItemsSource = list_import;

                                        if (i + stepScrollIntoView < list_import.Count)
                                        {
                                            dataGridPreview.ScrollIntoView(list_import[i + stepScrollIntoView]);
                                        }
                                        else
                                        {
                                            if (i < list_import.Count)
                                                dataGridPreview.ScrollIntoView(list_import[i]);
                                        }
                                    });

                                    countRegs = 0;
                                    startAddr = (UInt16)currAddr;
                                    startIndex = i;
                                }
                                catch (InvalidOperationException err)
                                {
                                    if (err.Message.IndexOf("non-connected socket") != -1)
                                    {
                                        main.SetTableDisconnectError(main.list_coilsTable, true);

                                        if (main.ModBus.type == main.ModBus_Def.TYPE_RTU || main.ModBus.type == main.ModBus_Def.TYPE_ASCII)
                                        {
                                            this.Dispatcher.Invoke((Action)delegate
                                            {
                                                main.buttonSerialActive_Click(null, null);
                                            });
                                        }
                                        else
                                        {
                                            this.Dispatcher.Invoke((Action)delegate
                                            {
                                                main.buttonTcpActive_Click(null, null);
                                            });
                                        }
                                    }

                                    Console.WriteLine(err);
                                }
                                catch (ModbusException err)
                                {
                                    if (err.Message.IndexOf("Timed out") != -1)
                                    {
                                        main.SetTableTimeoutError(main.list_coilsTable, false);
                                    }
                                    if (err.Message.IndexOf("ModBus ErrCode") != -1)
                                    {
                                        main.SetTableModBusError(main.list_coilsTable, err, false);
                                    }
                                    if (err.Message.IndexOf("CRC Error") != -1)
                                    {
                                        main.SetTableCrcError(main.list_coilsTable, false);
                                    }

                                    Console.WriteLine(err);

                                    if (abortOnError)
                                        break;
                                }
                                catch (Exception err)
                                {
                                    main.SetTableInternalError(main.list_coilsTable, true);
                                    Console.WriteLine(err);

                                    break;
                                }
                            }

                            if (i < list_import.Count)
                                toSendBuffer[countRegs] = list_import[i].ValueUInt > 0;

                            countRegs++;
                            oldCurrAddr = currAddr;
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            main.dataGridViewCoils.ItemsSource = null;
                            main.dataGridViewCoils.ItemsSource = main.list_coilsTable;
                        });
                    }
                }

                // Registers
                if (FC == 6)
                {
                    // Svuoto tabella coils
                    this.Dispatcher.Invoke((Action)delegate
                    {
                        main.list_holdingRegistersTable.Clear();

                        TaskbarItemInfo.ProgressValue = 0;
                        TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
                    });

                    if (!useMultiple)
                    {
                        for (int i = 0; i < list_import.Count; i++)
                        {
                            if (abortExit)
                                break;

                            try
                            {
                                bool? result = main.ModBus.presetSingleRegister_06(
                                    byte.Parse(main.textBoxModbusAddress_),
                                    (UInt16)(list_import[i].OffsetUInt + list_import[i].RegisterUInt),
                                    list_import[i].ValueUInt,
                                    main.readTimeout
                                    );

                                if (result != null)
                                {
                                        main.insertRowsTable(
                                        main.list_holdingRegistersTable,
                                        main.list_template_holdingRegistersTable,
                                        P.uint_parser(main.textBoxHoldingOffset_, main.comboBoxHoldingOffset_),
                                        (UInt16)(list_import[i].OffsetUInt + list_import[i].RegisterUInt),
                                        new UInt16[] { list_import[i].ValueUInt },
                                        result == true ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                        main.comboBoxHoldingOffset_,
                                        main.comboBoxHoldingRegistri_,
                                        main.comboBoxHoldingValori_,
                                        false,
                                        true);

                                    list_import[i].Background = main.colorDefaultWriteCellStr;
                                }
                                else
                                {
                                    list_import[i].Background = main.colorErrorCellStr;
                                }

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    TaskbarItemInfo.ProgressValue = (double)(i) / (double)(list_import.Count);

                                    dataGridPreview.ItemsSource = null;
                                    dataGridPreview.ItemsSource = list_import;

                                    if (i + stepScrollIntoView < list_import.Count)
                                        dataGridPreview.ScrollIntoView(list_import[i + stepScrollIntoView]);
                                    else
                                        dataGridPreview.ScrollIntoView(list_import[i]);
                                });
                            }
                            catch (InvalidOperationException err)
                            {
                                if (err.Message.IndexOf("non-connected socket") != -1)
                                {
                                    main.SetTableDisconnectError(main.list_coilsTable, true);
                                    
                                    if (main.ModBus.type == main.ModBus_Def.TYPE_RTU || main.ModBus.type == main.ModBus_Def.TYPE_ASCII)
                                    {
                                        this.Dispatcher.Invoke((Action)delegate
                                        {
                                            main.buttonSerialActive_Click(null, null);
                                        });
                                    }
                                    else
                                    {
                                        this.Dispatcher.Invoke((Action)delegate
                                        {
                                            main.buttonTcpActive_Click(null, null);
                                        });
                                    }
                                }

                                Console.WriteLine(err);
                            }
                            catch (ModbusException err)
                            {
                                if (err.Message.IndexOf("Timed out") != -1)
                                {
                                    main.SetTableTimeoutError(main.list_holdingRegistersTable, false);
                                }
                                if (err.Message.IndexOf("ModBus ErrCode") != -1)
                                {
                                    main.SetTableModBusError(main.list_holdingRegistersTable, err, false);
                                }
                                if (err.Message.IndexOf("CRC Error") != -1)
                                {
                                    main.SetTableCrcError(main.list_holdingRegistersTable, false);
                                }

                                Console.WriteLine(err);

                                if (abortOnError)
                                    break;
                            }
                            catch (Exception err)
                            {
                                main.SetTableInternalError(main.list_holdingRegistersTable, true);
                                Console.WriteLine(err);
                            }
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            main.dataGridViewHolding.ItemsSource = null;
                            main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                        });
                    }
                    else
                    {
                        UInt16 maxNumRegister = 0;

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            if (!UInt16.TryParse(TextBoxNrOfRegs.Text, out maxNumRegister))
                            {
                                MessageBox.Show(main.lang.languageTemplate["strings"]["importingErrorParse"], "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                        });

                        UInt16 startAddr = 0;   // Start address FC16
                        int startIndex = 0;     // Index start address
                        int countRegs = 0;      // Number of aggregated registers
                        int currAddr = -1;      // Current address
                        int oldCurrAddr = -1;   // Previous address
                        UInt16[] toSendBuffer = new UInt16[maxNumRegister];

                        for (int i = 0; i <= list_import.Count; i++)
                        {
                            if (abortExit)
                                break;

                            if (i < list_import.Count)
                                currAddr = list_import[i].RegisterUInt + list_import[i].OffsetUInt;

                            if (oldCurrAddr == -1)
                                startAddr = (UInt16)currAddr;

                            if ((currAddr != (oldCurrAddr + 1) || countRegs >= maxNumRegister || countRegs >= list_import.Count) && oldCurrAddr != -1)
                            {
                                // Invio
                                UInt16[] toSend = new UInt16[countRegs];
                                for (int a = 0; a < countRegs; a++)
                                {
                                    toSend[a] = toSendBuffer[a];
                                }

                                try
                                {
                                    // UInt16[] writtenRegs = null;

                                    UInt16[] writtenRegs = main.ModBus.presetMultipleRegisters_16(
                                        byte.Parse(main.textBoxModbusAddress_),
                                        startAddr,
                                        toSend,
                                        main.readTimeout
                                        );

                                    if (writtenRegs != null)
                                    {
                                        main.insertRowsTable(
                                            main.list_holdingRegistersTable,
                                            main.list_template_holdingRegistersTable,
                                            P.uint_parser(main.textBoxHoldingOffset_, main.comboBoxHoldingOffset_),
                                            startAddr,
                                            writtenRegs,
                                            writtenRegs.Length == toSend.Length ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                            main.comboBoxHoldingOffset_,
                                            main.comboBoxHoldingRegistri_,
                                            main.comboBoxHoldingValori_,
                                            false,
                                            true);

                                        for (int a = startIndex; a < (startIndex + countRegs); a++)
                                        {
                                            list_import[a].Background = main.colorDefaultWriteCellStr;
                                        }
                                    }
                                    else
                                    {
                                        for (int a = startIndex; a < (startIndex + countRegs); a++)
                                        {
                                            list_import[a].Background = main.colorErrorCellStr;
                                        }
                                    }

                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        TaskbarItemInfo.ProgressValue = (double)(i) / (double)(list_import.Count);

                                        dataGridPreview.ItemsSource = null;
                                        dataGridPreview.ItemsSource = list_import;

                                        if (i + stepScrollIntoView < list_import.Count)
                                        {
                                            dataGridPreview.ScrollIntoView(list_import[i + stepScrollIntoView]);
                                        }
                                        else
                                        {
                                            if (i < list_import.Count)
                                                dataGridPreview.ScrollIntoView(list_import[i]);
                                        }
                                    });

                                    countRegs = 0;
                                    startAddr = (UInt16)currAddr;
                                    startIndex = i;
                                }
                                catch (InvalidOperationException err)
                                {
                                    if (err.Message.IndexOf("non-connected socket") != -1)
                                    {
                                        main.SetTableDisconnectError(main.list_coilsTable, true);

                                        if (main.ModBus.type == main.ModBus_Def.TYPE_RTU || main.ModBus.type == main.ModBus_Def.TYPE_ASCII)
                                        {
                                            this.Dispatcher.Invoke((Action)delegate
                                            {
                                                main.buttonSerialActive_Click(null, null);
                                            });
                                        }
                                        else
                                        {
                                            this.Dispatcher.Invoke((Action)delegate
                                            {
                                                main.buttonTcpActive_Click(null, null);
                                            });
                                        }
                                    }

                                    Console.WriteLine(err);
                                }
                                catch (ModbusException err)
                                {
                                    if (err.Message.IndexOf("Timed out") != -1)
                                    {
                                        main.SetTableTimeoutError(main.list_holdingRegistersTable, false);
                                    }
                                    if (err.Message.IndexOf("ModBus ErrCode") != -1)
                                    {
                                        main.SetTableModBusError(main.list_holdingRegistersTable, err, false);
                                    }
                                    if (err.Message.IndexOf("CRC Error") != -1)
                                    {
                                        main.SetTableCrcError(main.list_holdingRegistersTable, false);
                                    }

                                    Console.WriteLine(err);

                                    if (abortOnError)
                                        break;
                                }
                                catch (Exception err)
                                {
                                    main.SetTableInternalError(main.list_holdingRegistersTable, true);
                                    Console.WriteLine(err);

                                    break;
                                }
                            }

                            if (i < list_import.Count)
                                toSendBuffer[countRegs] = list_import[i].ValueUInt;

                            countRegs++;
                            oldCurrAddr = currAddr;
                        }

                        this.Dispatcher.Invoke((Action)delegate
                        {
                            main.dataGridViewHolding.ItemsSource = null;
                            main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                        });
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }

            this.Dispatcher.Invoke((Action)delegate
            {
                TaskbarItemInfo.ProgressState = TaskbarItemProgressState.None;
                buttonImportPreview.IsEnabled = true;
                TextBoxNrOfRegs.IsEnabled = true;
                CheckBoxWriteMultiple.IsEnabled = true;
            });

            importRunning = false;

            if (closeAfterWrite)
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    this.Close();
                });
            }
        }

        private void CheckBoxWriteMultiple_Checked(object sender, RoutedEventArgs e)
        {
            useMultiple = (bool)CheckBoxWriteMultiple.IsChecked;
        }

        private void CheckBoxCloseWindowAfterWrite_Checked(object sender, RoutedEventArgs e)
        {
            closeAfterWrite = (bool)CheckBoxCloseWindowAfterWrite.IsChecked;
        }

        private void CheckBoxAbortOnError_Checked(object sender, RoutedEventArgs e)
        {
            abortOnError = (bool)CheckBoxAbortOnError.IsChecked;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if(importRunning)
            {
                if (MessageBox.Show(main.lang.languageTemplate["strings"]["abortPreviweImport"], "Alert", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    abortExit = true;
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                // Vincolato al ctrl
                if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                {
                    switch (e.Key)
                    {
                        case Key.C:
                            ExportClipboard();
                            break;

                        case Key.X:
                            ExportClipboard();
                            list_import.Clear();
                            break;

                        case Key.V:
                            ImportClipboard();
                            break;
                    }
                }

                if (!Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                {
                    switch (e.Key)
                    {
                        case Key.Delete:
                            list_import.Clear();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private void importToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = ".txt";
            openFileDialog.AddExtension = false;
            openFileDialog.Filter = "CSV|*.csv|JSON|*.json";
            openFileDialog.Title = main.lang.languageTemplate["strings"]["openFileCoils"];
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();

            if (openFileDialog.FileNames.Length > 0)
            {
                ImportFiles(openFileDialog.FileNames);
            }
        }

        private void exportToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialogBox = new SaveFileDialog();

                saveFileDialogBox.DefaultExt = "csv";
                saveFileDialogBox.AddExtension = false;

                if (FC == 5)
                    saveFileDialogBox.FileName = "Coils";
                else
                    saveFileDialogBox.FileName = "HoldingRegisters";

                saveFileDialogBox.Filter = "CSV|*.csv|JSON|*.json";
                saveFileDialogBox.Title = "";

                if ((bool)saveFileDialogBox.ShowDialog())
                {
                    if (saveFileDialogBox.FileName.IndexOf("csv") != -1)
                    {
                        String file_content = "";

                        for (int i = 0; i < (list_import.Count); i++)
                        {
                            file_content += list_import[i].Offset + ",";
                            file_content += list_import[i].Register + ",";
                            file_content += list_import[i].Value + ",";
                            file_content += list_import[i].Notes + "\n";
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

                        save.items = new ModBusItem_Save[list_import.Count];

                        for (int i = 0; i < (list_import.Count); i++)
                        {
                            try
                            {
                                ModBusItem_Save item = new ModBusItem_Save();

                                item.Offset = list_import[i].Offset;
                                item.Register = list_import[i].Register;
                                item.Value = list_import[i].Value;
                                item.Notes = list_import[i].Notes;

                                save.items[i] = item;
                            }
                            catch { }
                        }


                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        jss.MaxJsonLength = main.MaxJsonLength;
                        string file_content = jss.Serialize(save);

                        StreamWriter writer = new StreamWriter(saveFileDialogBox.OpenFile());

                        writer.Write(file_content);
                        writer.Dispose();
                        writer.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private void deleteToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            list_import.Clear();
        }

        private void cutToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ExportClipboard();
            list_import.Clear();
        }

        private void copyToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ExportClipboard();
        }

        private void pasteToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ImportClipboard();
        }

        private void writeSingleMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CheckBoxWriteMultiple.IsChecked = false;
            buttonImportPreview_Click(null, null);
        }

        private void writeMultipleMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CheckBoxWriteMultiple.IsChecked = true;
            buttonImportPreview_Click(null, null);
        }
    }

    class ModBus_Import
    {
        public string Offset { get; set; }
        public UInt16 OffsetUInt { get; set; }
        public string Register { get; set; }
        public UInt16 RegisterUInt { get; set; }
        public string Value { get; set; }

        public UInt16 ValueUInt { get; set; }
        public string Notes { get; set; }
        public string Foreground { get; set; }
        public string Background { get; set; }
        
        public string ForegroundDef { get; set; }
        public string BackgroundDef { get; set; }
    }
}
