

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

// Json lib
using System.Web.Script.Serialization;

// Libreria lingue
using LanguageLib; // Libreria custom per caricare etichette in lingue differenti

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for Carica_impianto.xaml
    /// </summary>
    public partial class Carica_impianto : Window
    {
        public string path { get; set; }

        Language lang;

        public Carica_impianto(String defaultPath, MainWindow main_)
        {
            InitializeComponent();

            lang = new Language(this);

            String[] subFolders = Directory.GetDirectories("Json\\");

            comboBoxImpianto.Items.Clear();

            for (int i = 0; i < subFolders.Length; i++)
            {
                comboBoxImpianto.Items.Add(subFolders[i].Substring(subFolders[i].IndexOf('\\') + 1));
            }

            comboBoxImpianto.SelectedItem = defaultPath;

            // Centro la finestra
            double screenWidth = System.Windows.SystemParameters.PrimaryScreenWidth;
            double screenHeight = System.Windows.SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);

            lang.loadLanguageTemplate(main_.language);

            this.Title = main_.title + " " + main_.version;
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            path = comboBoxImpianto.SelectedItem.ToString();
            this.DialogResult = true;
        }

        private void buttonAnnulla_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }

        private void LodaProfile_Loaded(object sender, RoutedEventArgs e)
        {
            comboBoxImpianto.Focus();
        }

        private void LodaProfile_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                buttonOpen_Click(sender, e);
            }

            if(e.Key == Key.Escape)
            {
                buttonAnnulla_Click(sender, e);
            }
        }
    }
}
