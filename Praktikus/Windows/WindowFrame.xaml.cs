using Praktikus.Misc;
using Praktikus.Pages;
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

namespace Praktikus.Windows
{
    /// <summary>
    /// Логика взаимодействия для WindowFrame.xaml
    /// </summary>
    public partial class WindowFrame : Window
    {
        public WindowFrame()
        {
            InitializeComponent();
            Nav.MainFrame = MainFrame;
            Nav.MainFrame.Navigate(new PageLogin());
        }
    }
}
