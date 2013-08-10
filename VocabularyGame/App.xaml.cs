using System.Globalization;
using System.Threading;
using System.Windows;

namespace VocabularyGame
{
    public partial class App : Application
    {
        public App()
        {
            CultureInfo cInfo = new CultureInfo(VocabularyGame.Properties.Settings.Default.Language);
            Thread.CurrentThread.CurrentUICulture = cInfo;
            Thread.CurrentThread.CurrentCulture = cInfo;
        }
    }
}