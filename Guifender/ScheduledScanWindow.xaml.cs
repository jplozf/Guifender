using System;
using System.Collections.Generic;
using System.Windows;

namespace Guifender
{
    public partial class ScheduledScanWindow : Window
    {
        public ScheduledScan Scan { get; private set; }
        public List<string> ScanTypes { get; } = new List<string> { "QuickScan", "FullScan", "CustomScan" };

        public ScheduledScanWindow(ScheduledScan scan)
        {
            InitializeComponent();
            Scan = scan;
            DataContext = this;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Scan.Path = dialog.SelectedPath;
                }
            }
        }
    }
}
