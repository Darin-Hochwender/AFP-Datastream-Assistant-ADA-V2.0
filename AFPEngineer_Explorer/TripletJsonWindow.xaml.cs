using System.Diagnostics;
using System.Windows;

namespace AFPEngineer_Explorer
{
    public partial class TripletJsonWindow : Window
    {
        private string _url;

        public TripletJsonWindow(string heading, string rawJson, string url)
        {
            InitializeComponent();
            HeaderTxt.Text = heading;
            JsonContentTxt.Text = string.IsNullOrWhiteSpace(rawJson) ? "No JSON data found." : rawJson;
            _url = url;

            if (string.IsNullOrWhiteSpace(_url))
            {
                OpenUrlBtn.Visibility = Visibility.Collapsed;
            }
        }

        private void OpenUrl_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(_url))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _url,
                    UseShellExecute = true
                });
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}