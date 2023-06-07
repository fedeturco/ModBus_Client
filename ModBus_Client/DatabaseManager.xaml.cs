

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
using System.Collections.ObjectModel;
using System.IO;
using System.Diagnostics;
using System.IO.Compression;
using Microsoft.Win32;

// Libreria lingue
using LanguageLib; // Libreria custom per caricare etichette in lingue differenti

namespace ModBus_Client
{
    /// <summary>
    /// Interaction logic for DatabaseManager.xaml
    /// </summary>
    public partial class DatabaseManager : Window
    {
        ObservableCollection<profile> db = new ObservableCollection<profile>();

        Language lang;
        public string SelectedProfile = "";

        public DatabaseManager(MainWindow main_)
        {
            InitializeComponent();

            lang = new Language(this);

            DataGridDb.ItemsSource = db;

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

         void LoadDb()
        {
            String[] subFolders = Directory.GetDirectories(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\");

            db.Clear();

            for (int i = 0; i < subFolders.Length; i++)
            {
                profile tmp = new profile();

                tmp.name = subFolders[i].Split('\\')[subFolders[i].Split('\\').Length - 1];

                db.Add(tmp);
            }
        }

        private void ButtonRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadDb();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDb();
        }

        private void ButtonOpenFileLocation_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("explorer.exe", "Json");
            }
            catch
            {
                MessageBox.Show("Imposssibile aprire la cartella di configurazione Database", "Alert");
                Console.WriteLine("Imposssibile aprire la cartella di configurazione Database");
            }
        }

        private void ButtonExportZip_Click(object sender, RoutedEventArgs e)
        {
            profile currItem = (profile)DataGridDb.SelectedItem;
            
            SaveFileDialog window = new SaveFileDialog();

            window.Filter = "Zip Files | *.zip";
            window.DefaultExt = ".zip";
            window.FileName = currItem.name + ".zip";

            if ((bool)window.ShowDialog())
            {
                // Se il file lo esiste gia' lo elimino
                if (File.Exists(window.FileName))
                    File.Delete(window.FileName);

                ZipFile.CreateFromDirectory("Json\\" + currItem.name, window.FileName);
            }
        }

        private void ButtonImportZip_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog window = new OpenFileDialog();

            window.Filter = "Zip Files | *.zip";
            window.DefaultExt = ".zip";

            if ((bool)window.ShowDialog())
            {
                if(Directory.Exists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + System.IO.Path.GetFileNameWithoutExtension(window.FileName)))
                {
                    if(MessageBox.Show(lang.languageTemplate["strings"]["overwriteProfile"], "Info", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
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

            LoadDb();
        }

        private void DataGridDb_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            try
            {
                if (DataGridDb.SelectedItem != null)
                {
                    profile selected = (profile)DataGridDb.SelectedItem;

                    labelProfileSelected.Content = labelProfileSelected.Content.ToString().Split(':')[0] + ": " + selected.name;
                    labelProfileSelected.Visibility = Visibility.Visible;

                    ButtonExportZip.IsEnabled = true;
                    //ButtonImportZip.IsEnabled = true;
                    ButtonDeleteProfile.IsEnabled = true;
                }
            }
            catch
            {
                ButtonExportZip.IsEnabled = false;
                //ButtonImportZip.IsEnabled = false;
                ButtonDeleteProfile.IsEnabled = false;
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape)
            {
                this.Close();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.DialogResult = false;

            if (DataGridDb.SelectedItem != null)
            {
                profile selected = (profile)DataGridDb.SelectedItem;
                this.SelectedProfile = selected.name;
                this.DialogResult = true;

                // debug
                Console.WriteLine("Selected: {0}", selected.name);
            }
        }

        private void ButtonDeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (DataGridDb.SelectedItem != null)
                {
                    profile selected = (profile)DataGridDb.SelectedItem;

                    labelProfileSelected.Content = labelProfileSelected.Content.ToString().Split(':')[0] + ": " + selected.name;
                    labelProfileSelected.Visibility = Visibility.Visible;

                    if (Directory.Exists(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + selected.name))
                    {
                        if (MessageBox.Show(lang.languageTemplate["strings"]["deleteProfile"], "Info", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                        {
                            Directory.Delete(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Json\\" + selected.name, true);

                            ButtonExportZip.IsEnabled = false;
                            ButtonDeleteProfile.IsEnabled = false;

                            LoadDb();
                        }
                        else
                        {
                            return;
                        }
                    }
                }
            }
            catch
            {
                ButtonDeleteProfile.IsEnabled = false;
            }
        }
    }

    public class profile
    {
        public String name { get; set; }
    }
}
