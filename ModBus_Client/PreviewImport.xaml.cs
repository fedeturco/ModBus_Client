

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
// Json LIBs
using System.Web.Script.Serialization;

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
            if(FC == 5)
            {
                // Svuoto tabella coils
                main.list_coilsTable.Clear();

                if (!useMultiple)
                {
                    foreach(ModBus_Import item in list_import)
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
                                main.template_coilsOffset,
                                item.OffsetUInt,
                                item.RegisterUInt,
                                new UInt16[] { item.ValueUInt },
                                result == true ? main.colorDefaultWriteCellStr : main.colorErrorCellStr, 
                                main.comboBoxCoilsRegistri_, 
                                "DEC",
                                false);
                        }
                    }
                }
                else
                {
                    UInt16 addressStart = 0;
                    UInt16 offsetStart = 0;
                    UInt16 maxNumRegister = 0;
                    
                    if(!UInt16.TryParse(TextBoxNrOfRegs.Text, out maxNumRegister))
                    {
                        MessageBox.Show("Error parsing value", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    UInt16 start = 0;

                    for(int i = 0; i <= list_import.Count; i++)
                    {
                        if((i % maxNumRegister == 0 || i == list_import.Count) && i > 0)
                        {
                            bool[] response = new bool[i % maxNumRegister == 0 ? maxNumRegister :  list_import.Count % maxNumRegister];
                            UInt16[] response_ = new UInt16[i % maxNumRegister == 0 ? maxNumRegister : list_import.Count % maxNumRegister];

                            for (int a = 0; a < i; a++)
                            {
                                response[a] = list_import[start + a].ValueUInt > 0;
                                response_[a] = list_import[start + a].ValueUInt;
                            }

                            start = (UInt16)i;

                            bool? result = main.ModBus.forceMultipleCoils_15(
                            byte.Parse(main.textBoxModbusAddress_),
                            main.useOffsetInTable ? (UInt16)(offsetStart + addressStart) : addressStart,
                            response,
                            main.readTimeout
                            );

                            if (result != null)
                            {
                                main.insertRowsTable(
                                        main.list_coilsTable,
                                        main.list_template_coilsTable,
                                        main.template_coilsOffset,
                                        offsetStart,
                                        addressStart,
                                        response_,
                                        result == true ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                        main.comboBoxCoilsRegistri_,
                                        "DEC",
                                        false);
                            }

                            if (i == list_import.Count)
                                break;
                        }
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
                                main.template_HoldingOffset,
                                item.OffsetUInt,
                                item.RegisterUInt,
                                new UInt16[] { item.ValueUInt },
                                result == true ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                main.comboBoxHoldingRegistri_,
                                "DEC",
                                false);
                        }
                    }
                }
                else
                {
                    UInt16 addressStart = 0;
                    UInt16 offsetStart = 0;
                    UInt16 maxNumRegister = 0;

                    if (!UInt16.TryParse(TextBoxNrOfRegs.Text, out maxNumRegister))
                    {
                        MessageBox.Show("Error parsing value", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    UInt16 start = 0;

                    for (int i = 0; i <= list_import.Count; i++)
                    {
                        if ((i % maxNumRegister == 0 || i == list_import.Count) && i > 0)
                        {
                            UInt16[] response = new UInt16[i % maxNumRegister == 0 ? maxNumRegister : list_import.Count % maxNumRegister];

                            for (int a = 0; a < i; a++)
                            {
                                response[a] = list_import[start + a].ValueUInt;
                            }

                            start = (UInt16)i;

                            UInt16[] writtenRegs = main.ModBus.presetMultipleRegisters_16(
                            byte.Parse(main.textBoxModbusAddress_),
                            main.useOffsetInTable ? (UInt16)(offsetStart + addressStart) : addressStart,
                            response,
                            main.readTimeout
                            );

                            if (writtenRegs != null)
                            {
                                main.insertRowsTable(
                                        main.list_holdingRegistersTable,
                                        main.list_template_holdingRegistersTable,
                                        main.template_HoldingOffset,
                                        offsetStart,
                                        addressStart,
                                        writtenRegs,
                                        writtenRegs.Length == response.Length ? main.colorDefaultWriteCellStr : main.colorErrorCellStr,
                                        main.comboBoxHoldingRegistri_,
                                        "DEC",
                                        false);
                            }

                            if (i == list_import.Count)
                                break;
                        }
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
