

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
using System.IO;
using System.Collections.ObjectModel;

// Custom Modbus Library
using ModBusMaster_Chicco;

// Json LIBs
using System.Web.Script.Serialization;

// Classe con funzioni di conversione DEC-HEX
using Raccolta_funzioni_parser;

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
        bool edited = false;

        Parser P = new Parser();

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
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            List<ModBus_Item> content = new List<ModBus_Item>();

            try
            {
                foreach (String filePath in filePaths)
                {
                    // CSV File
                    if (filePath.IndexOf(".csv") != -1)
                    {
                        string[] registers = File.ReadAllText(filePath).Split('\n');
                        for (int i = 0; i < registers.Length; i++)
                        {
                            string[] row = registers[i].Split(',');

                            if (row.Length < 3)
                                continue;

                            if (row[0].ToLower().IndexOf("offset") != -1)
                                continue;

                            ModBus_Import item = new ModBus_Import();

                            item.Offset = row[0];
                            item.OffsetUInt = UInt16.Parse(row[0].Replace("0x", ""), row[0].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                            item.Register = row[1];
                            item.RegisterUInt = UInt16.Parse(row[1].Replace("0x", ""), row[1].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                            item.Value = row[2];
                            item.ValueUInt = UInt16.Parse(row[2].Replace("0x", ""), row[2].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

                            item.Notes = row[3];

                            if ((bool)main.CheckBoxDarkMode.IsChecked)
                            {
                                item.Foreground = main.ForeGroundDarkStr;
                                item.Background = main.BackGroundDarkStr;
                            }
                            else
                            {
                                item.Foreground = main.ForeGroundLightStr;
                                item.Background = main.BackGroundLight2Str;
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
                            }
                            else
                            {
                                item.Foreground = main.ForeGroundLightStr;
                                item.Background = main.BackGroundLight2Str;
                            }

                            list_import.Add(item);
                        }
                    }
                }
            }
            catch(Exception err)
            {
                MessageBox.Show("Error importing files", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                Console.WriteLine(err);
            }

            dataGridPreview.ItemsSource = list_import;

            buttonImportPreview.Focus();
        }

        private void buttonImportPreview_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Aggiorno stato registri prima di inviarli
                for (int i = 0; i < list_import.Count; i++)
                {
                    list_import[i].OffsetUInt = UInt16.Parse(list_import[i].Offset.Replace("0x", ""), list_import[i].Offset.ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);
                    list_import[i].RegisterUInt = UInt16.Parse(list_import[i].Register.Replace("0x", ""), list_import[i].Register.ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);
                    list_import[i].ValueUInt = UInt16.Parse(list_import[i].Value.Replace("0x", ""), list_import[i].Value.ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);
                }

                // Coils
                if (FC == 5)
                {
                    // Svuoto tabella coils
                    main.list_coilsTable.Clear();

                    if (!useMultiple)
                    {
                        foreach (ModBus_Import item in list_import)
                        {
                            try
                            {
                                bool? result = main.ModBus.forceSingleCoil_05(
                                    byte.Parse(main.textBoxModbusAddress_),
                                    (UInt16)(item.OffsetUInt + item.RegisterUInt),
                                    item.ValueUInt,
                                    main.readTimeout
                                    );

                                if (result != null)
                                {
                                    main.insertRowsTable(
                                        main.list_coilsTable,
                                        main.list_template_coilsTable,
                                        P.uint_parser(main.textBoxCoilsOffset_, main.comboBoxCoilsOffset_),
                                        (UInt16)(item.OffsetUInt + item.RegisterUInt),
                                        new UInt16[] { item.ValueUInt },
                                        result == true ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                        main.comboBoxCoilsRegistri_,
                                        "DEC",
                                        false);
                                }
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

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    main.dataGridViewCoils.ItemsSource = null;
                                    main.dataGridViewCoils.ItemsSource = main.list_coilsTable;
                                });

                                if (abortOnError)
                                    break;
                            }
                            catch (Exception err)
                            {
                                main.SetTableInternalError(main.list_coilsTable, true);
                                Console.WriteLine(err);

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    main.dataGridViewCoils.ItemsSource = null;
                                    main.dataGridViewCoils.ItemsSource = main.list_coilsTable;
                                });

                                break;
                            }
                        }
                    }
                    else
                    {
                        UInt16 maxNumRegister = 0;

                        if (!UInt16.TryParse(TextBoxNrOfRegs.Text, out maxNumRegister))
                        {
                            MessageBox.Show("Error parsing value", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        UInt16 startAddr = 0;   // Start address FC16
                        int countRegs = 0;      // Numero di registri accorpati
                        int currAddr = -1;      // Address corrente
                        int oldCurrAddr = -1;   // Address precedente
                        bool[] toSendBuffer = new bool[maxNumRegister];

                        for (int i = 0; i <= list_import.Count; i++)
                        {
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
                                                main.comboBoxHoldingRegistri_,
                                                "DEC",
                                                false);
                                    }

                                    countRegs = 0;
                                    startAddr = (UInt16)currAddr;
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

                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        main.dataGridViewCoils.ItemsSource = null;
                                        main.dataGridViewCoils.ItemsSource = main.list_coilsTable;
                                    });

                                    if (abortOnError)
                                        break;
                                }
                                catch (Exception err)
                                {
                                    main.SetTableInternalError(main.list_coilsTable, true);
                                    Console.WriteLine(err);

                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        main.dataGridViewCoils.ItemsSource = null;
                                        main.dataGridViewCoils.ItemsSource = main.list_coilsTable;
                                    });

                                    break;
                                }
                            }

                            if (i < list_import.Count)
                                toSendBuffer[countRegs] = list_import[i].ValueUInt > 0;

                            countRegs++;
                            oldCurrAddr = currAddr;
                        }
                    }
                }

                // Registers
                if (FC == 6)
                {
                    // Svuoto tabella coils
                    main.list_holdingRegistersTable.Clear();

                    if (!useMultiple)
                    {
                        foreach (ModBus_Import item in list_import)
                        {
                            try
                            {
                                bool? result = main.ModBus.presetSingleRegister_06(
                                    byte.Parse(main.textBoxModbusAddress_),
                                    (UInt16)(item.OffsetUInt + item.RegisterUInt),
                                    item.ValueUInt,
                                    main.readTimeout
                                    );

                                if (result != null)
                                {
                                    main.insertRowsTable(
                                        main.list_holdingRegistersTable,
                                        main.list_template_holdingRegistersTable,
                                        P.uint_parser(main.textBoxHoldingOffset_, main.comboBoxHoldingOffset_),
                                        (UInt16)(item.OffsetUInt + item.RegisterUInt),
                                        new UInt16[] { item.ValueUInt },
                                        result == true ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                        main.comboBoxHoldingRegistri_,
                                        main.comboBoxHoldingValori_,
                                        false);
                                }
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

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    main.dataGridViewHolding.ItemsSource = null;
                                    main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                                });

                                if (abortOnError)
                                    break;
                            }
                            catch (Exception err)
                            {
                                main.SetTableInternalError(main.list_holdingRegistersTable, true);
                                Console.WriteLine(err);

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    main.dataGridViewHolding.ItemsSource = null;
                                    main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                                });
                            }
                        }
                    }
                    else
                    {
                        UInt16 maxNumRegister = 0;

                        if (!UInt16.TryParse(TextBoxNrOfRegs.Text, out maxNumRegister))
                        {
                            MessageBox.Show("Error parsing value", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        UInt16 startAddr = 0;   // Start address FC16
                        int countRegs = 0;      // Numero di registri accorpati
                        int currAddr = -1;      // Address corrente
                        int oldCurrAddr = -1;   // Address precedente
                        UInt16[] toSendBuffer = new UInt16[maxNumRegister];

                        for (int i = 0; i <= list_import.Count; i++)
                        {
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
                                                main.comboBoxHoldingRegistri_,
                                                main.comboBoxHoldingValori_,
                                                false);
                                    }

                                    countRegs = 0;
                                    startAddr = (UInt16)currAddr;
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

                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        main.dataGridViewHolding.ItemsSource = null;
                                        main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                                    });

                                    if(abortOnError)
                                        break;
                                }
                                catch (Exception err)
                                {
                                    main.SetTableInternalError(main.list_holdingRegistersTable, true);
                                    Console.WriteLine(err);

                                    this.Dispatcher.Invoke((Action)delegate
                                    {
                                        main.dataGridViewHolding.ItemsSource = null;
                                        main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                                    });

                                    break;
                                }
                            }

                            if (i < list_import.Count)
                                toSendBuffer[countRegs] = list_import[i].ValueUInt;

                            countRegs++;
                            oldCurrAddr = currAddr;
                        }
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }

            if (closeAfterWrite)
                this.Close();
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
    }
}
