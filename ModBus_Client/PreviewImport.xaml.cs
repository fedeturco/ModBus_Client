

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
        String filePath = "";
        byte FC = 0;

        ObservableCollection<ModBus_Import> list_import = new ObservableCollection<ModBus_Import>();

        bool useMultiple = false;

        Parser P = new Parser();

        public PreviewImport(MainWindow main_, String filePath_, byte FC_)
        {
            InitializeComponent();

            main = main_;
            filePath = filePath_;
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
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            List<ModBus_Item> content = new List<ModBus_Item>();

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
                    item.OffsetUInt = UInt16.Parse(row[0].Replace("0x",""), row[0].ToLower().IndexOf("0x") != -1 ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer);

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

                foreach(dynamic reg in obj["items"])
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

            dataGridPreview.ItemsSource = list_import;
        }

        private void buttonImportPreview_Click(object sender, RoutedEventArgs e)
        {
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
                        catch (ModbusException err)
                        {
                            if (err.Message.IndexOf("Timed out") != -1)
                            {
                                main.SetTableTimeoutError(main.list_holdingRegistersTable, true);
                            }
                            if (err.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                main.SetTableModBusError(main.list_holdingRegistersTable, err, true);
                            }
                            if (err.Message.IndexOf("CRC Error") != -1)
                            {
                                main.SetTableCrcError(main.list_holdingRegistersTable, true);
                            }

                            Console.WriteLine(err);

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                main.dataGridViewHolding.ItemsSource = null;
                                main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                            });

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
                            catch (ModbusException err)
                            {
                                if (err.Message.IndexOf("Timed out") != -1)
                                {
                                    main.SetTableTimeoutError(main.list_holdingRegistersTable, true);
                                }
                                if (err.Message.IndexOf("ModBus ErrCode") != -1)
                                {
                                    main.SetTableModBusError(main.list_holdingRegistersTable, err, true);
                                }
                                if (err.Message.IndexOf("CRC Error") != -1)
                                {
                                    main.SetTableCrcError(main.list_holdingRegistersTable, true);
                                }

                                Console.WriteLine(err);

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    main.dataGridViewHolding.ItemsSource = null;
                                    main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                                });

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
                                    "DEC",
                                    false);
                            }
                        }
                        catch (ModbusException err)
                        {
                            if (err.Message.IndexOf("Timed out") != -1)
                            {
                                main.SetTableTimeoutError(main.list_holdingRegistersTable, true);
                            }
                            if (err.Message.IndexOf("ModBus ErrCode") != -1)
                            {
                                main.SetTableModBusError(main.list_holdingRegistersTable, err, true);
                            }
                            if (err.Message.IndexOf("CRC Error") != -1)
                            {
                                main.SetTableCrcError(main.list_holdingRegistersTable, true);
                            }

                            Console.WriteLine(err);

                            this.Dispatcher.Invoke((Action)delegate
                            {
                                main.dataGridViewHolding.ItemsSource = null;
                                main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                            });
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
                                            "DEC",
                                            false);
                                }

                                countRegs = 0;
                                startAddr = (UInt16)currAddr;
                            }
                            catch (ModbusException err)
                            {
                                if (err.Message.IndexOf("Timed out") != -1)
                                {
                                    main.SetTableTimeoutError(main.list_holdingRegistersTable, true);
                                }
                                if (err.Message.IndexOf("ModBus ErrCode") != -1)
                                {
                                    main.SetTableModBusError(main.list_holdingRegistersTable, err, true);
                                }
                                if (err.Message.IndexOf("CRC Error") != -1)
                                {
                                    main.SetTableCrcError(main.list_holdingRegistersTable, true);
                                }

                                Console.WriteLine(err);

                                this.Dispatcher.Invoke((Action)delegate
                                {
                                    main.dataGridViewHolding.ItemsSource = null;
                                    main.dataGridViewHolding.ItemsSource = main.list_holdingRegistersTable;
                                });

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

            this.Close();
        }

        private void CheckBoxWriteMultiple_Checked(object sender, RoutedEventArgs e)
        {
            useMultiple = (bool)CheckBoxWriteMultiple.IsChecked;
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
