using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace TVCOMNewString
{
    public partial class ColorPickerWindow : Window, INotifyPropertyChanged
    {
        private int _red = 255;
        private int _green = 0;
        private int _blue = 0;

        public int Red
        {
            get => _red;
            set
            {
                _red = Math.Max(0, Math.Min(255, value));
                OnPropertyChanged(nameof(Red));
                OnPropertyChanged(nameof(SelectedColor));
            }
        }

        public int Green
        {
            get => _green;
            set
            {
                _green = Math.Max(0, Math.Min(255, value));
                OnPropertyChanged(nameof(Green));
                OnPropertyChanged(nameof(SelectedColor));
            }
        }

        public int Blue
        {
            get => _blue;
            set
            {
                _blue = Math.Max(0, Math.Min(255, value));
                OnPropertyChanged(nameof(Blue));
                OnPropertyChanged(nameof(SelectedColor));
            }
        }

        public SolidColorBrush SelectedColor
        {
            get => new SolidColorBrush(Color.FromRgb((byte)Red, (byte)Green, (byte)Blue));
        }

        public Color SelectedWpfColor => Color.FromRgb((byte)Red, (byte)Green, (byte)Blue);

        public ColorPickerWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void ColorButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string colorTag)
            {
                string[] rgb = colorTag.Split(',');
                if (rgb.Length == 3)
                {
                    Red = int.Parse(rgb[0]);
                    Green = int.Parse(rgb[1]);
                    Blue = int.Parse(rgb[2]);
                }
            }
        }

        private void RgbSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            OnPropertyChanged(nameof(SelectedColor));
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}