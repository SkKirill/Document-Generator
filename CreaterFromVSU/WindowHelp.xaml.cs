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

namespace CreaterFromVSU
{
    /// <summary>
    /// Логика взаимодействия для WindowHelp.xaml
    /// </summary>
    public partial class WindowHelp : Window
    {
        public WindowHelp()
        {
            InitializeComponent();
        }

        private void ButtonExsit_MouseLeave(object sender, MouseEventArgs e)
        {
            ButtonExsit.Background = new SolidColorBrush(Color.FromRgb(255, 0, 0));
        }

        private void ButtonExsit_MouseEnter(object sender, MouseEventArgs e)
        {
            ButtonExsit.Background = new SolidColorBrush(Color.FromRgb(255, 80, 80));
        }

        private void ButtonExsit_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Close();
        }
        private bool isDragging = false;
        private Point lastPosition;

        private void TopPanel_MouseUp(object sender, MouseButtonEventArgs e)
        {
            isDragging = false;
        }
    

        private void TopPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point currentPosition = e.GetPosition(this);
                Left = Left - (lastPosition.X - currentPosition.X);
                Top = Top - (lastPosition.Y - currentPosition.Y);
            }
        }

        private void TopPanel_MouseDown(object sender, MouseButtonEventArgs e)
        {
            isDragging = true;
            lastPosition = e.GetPosition(this);
        }
    }
}
