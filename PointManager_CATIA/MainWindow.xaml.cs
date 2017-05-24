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
using System.Windows.Navigation;
using System.Windows.Shapes;
using INFITF;
using MECMOD;
using HybridShapeTypeLib;
using NPOI;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.HSSF;
using NPOI.SS;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using System.Reflection;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using AnnotationTypeLib;
using System.ComponentModel;

namespace PointManager_CATIA
{


    public partial class MainWindow : System.Windows.Window
    {
        public static INFITF.Application CATIA = null;

        public MainWindow()
        {
            InitializeComponent();
            try
            {
                var CATIA = (INFITF.Application)Marshal.GetActiveObject("Catia.Application");
            }
            catch
            {
                MessageBox.Show("CATIA V5 Не запущена! Сначала запустите CATIA, а затем Point Manager.", "Упс!", MessageBoxButton.OK, MessageBoxImage.Information);
                this.Close();
            }
            
        }

        private void ViewTypeCombo_Loaded(object sender, RoutedEventArgs e)
        {
            var ViewList = new List<string>();
            ViewList.Add("Windshield");
            ViewList.Add("Backlite");
            ViewList.Add("Left Door");
            ViewList.Add("Right Door");
            ViewTypeCombo.ItemsSource = ViewList;
            ViewTypeCombo.SelectedIndex = 0;

        }

        private void GraphTypeCombo_Loaded(object sender, RoutedEventArgs e)
        {
            var TypeList = new List<string>();
            TypeList.Add("Curvature");
            TypeList.Add("Gap or Size");
            TypeList.Add("Random Points");
            GraphTypeCombo.ItemsSource = TypeList;
            GraphTypeCombo.SelectedIndex = 0;
        }

        private void checkBox_Checked(object sender, RoutedEventArgs e)
        {
            ViewTypeCombo.IsEnabled = true;
        }

        private void checkBox_Unchecked(object sender, RoutedEventArgs e)
        {
            ViewTypeCombo.IsEnabled = false;
        }
    }
}
