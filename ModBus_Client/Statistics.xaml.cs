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

using ModBusMaster_Chicco;

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for Statistics.xaml
    /// </summary>
    public partial class Statistics : Window
    {
        public Statistics(String title, ModBus_Chicco Modbus)
        {
            InitializeComponent();

            UpdateStatistics(Modbus);
        }

        public void UpdateStatistics(ModBus_Chicco Modbus)
        {
            if (Modbus != null)
            {
                LabelFC01.Content = Modbus.Count_FC01.ToString();
                LabelFC02.Content = Modbus.Count_FC02.ToString();
                LabelFC03.Content = Modbus.Count_FC03.ToString();
                LabelFC04.Content = Modbus.Count_FC04.ToString();
                LabelFC05.Content = Modbus.Count_FC05.ToString();
                LabelFC06.Content = Modbus.Count_FC06.ToString();
                LabelFC08.Content = Modbus.Count_FC08.ToString();
                LabelFC15.Content = Modbus.Count_FC15.ToString();
                LabelFC16.Content = Modbus.Count_FC16.ToString();

                LabelTotal.Content =
                    (
                    Modbus.Count_FC01 +
                    Modbus.Count_FC02 +
                    Modbus.Count_FC03 +
                    Modbus.Count_FC04 +
                    Modbus.Count_FC05 +
                    Modbus.Count_FC06 +
                    Modbus.Count_FC08 +
                    Modbus.Count_FC15 +
                    Modbus.Count_FC16).ToString();
            }
        }
    }
}
