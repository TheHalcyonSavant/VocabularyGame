using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace VocabularyGame
{
    public partial class MainWindow : Window
    {

#region Properties

        private static int _seconds;
        private static BitmapImage _biSound = new BitmapImage(new Uri("images/Sound.png", UriKind.Relative));
        private static Properties.Settings _s = Properties.Settings.Default;

        private bool _isInitialized = false;
        private bool _isInternetOk;
        private bool _isReadingWrongs = false;
        private bool _isToolTipRevealed = false;
        private bool[] _answerTypes = new bool[3];
        private int _iPlayWords;
        private int _missedODictIdx = 0;
        private int _repeatingLimit;
        private string _formattedTitle;
        private string _xlsmSafeFileName = "";
        private BackgroundWorker _bgWorker = new BackgroundWorker();
        private CorrectClass _correct = new CorrectClass();
        private Dictionary<string, byte> _dictRepeats = new Dictionary<string, byte>();
        private DispatcherTimer _timerAfterChoice = new DispatcherTimer();
        private DispatcherTimer _timerCountdown = new DispatcherTimer();
        private List<Uri> _lPlayWords = new List<Uri>();
        private LoadingWindow _wLoading;
        private OrderedDictionary _odict = new OrderedDictionary();
        private OrderedDictionary _odictWrongs = new OrderedDictionary();
        private RecordsWindow _wRecords;
        private Regex _rgxNumKeys = new Regex("^(D|NumPad)(?<num>[1-5])");
        
        public int points;
        public string xlsmSafeFileNameNoExt = "";
        public BinaryFormatter binFormatter = new BinaryFormatter();

#endregion

#region Constructor & Window Events

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            if (_isInitialized) return;

            _formattedTitle = Title + " ({0})";
            _bgWorker.ProgressChanged += (_, evnt) =>
            {
                if (evnt.ProgressPercentage == 100)
                {
                    _wLoading.Hide();
                    return;
                }
                _wLoading.lblMain.Content = evnt.UserState;
            };
            _bgWorker.WorkerReportsProgress = true;
            _bgWorker.WorkerSupportsCancellation = true;
            _bgWorker.DoWork += Worker_Init;
            _bgWorker.RunWorkerCompleted += Worker_InitComplete;
            _isInitialized = true;
            _wLoading = new LoadingWindow();
            _bgWorker.RunWorkerAsync();
            _wLoading.ShowDialog();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!Directory.Exists("dat")) Directory.CreateDirectory("dat");

            using (FileStream fs = new FileStream("dat/" + xlsmSafeFileNameNoExt + _s.RepeatsSuffix, FileMode.Create))
                binFormatter.Serialize(fs, _dictRepeats);

            using (FileStream fs = new FileStream("dat/" + xlsmSafeFileNameNoExt + _s.WrongsSuffix, FileMode.Create))
                binFormatter.Serialize(fs, _odictWrongs);

            saveRecord();

            _s.Save();
            Environment.Exit(0);
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            Match m = _rgxNumKeys.Match(e.Key.ToString());
            if (!m.Success) return;

            RadioButton rb = spRbs.Children[int.Parse(m.Result("${num}")) - 1] as RadioButton;
            rb.IsChecked = true;
            rb.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }

#endregion

#region Control Events

        #region MenuItem Events

        private void miAnswerTypes_Click(object sender, RoutedEventArgs e)
        {
            MenuItem mi = sender as MenuItem;
            if (!mi.IsChecked)
            {
                bool isOk = false;
                foreach (MenuItem otherMi in miAnswerTypes.Items)
                    if (otherMi.IsChecked) isOk = true;
                if (!isOk)
                {
                    mi.IsChecked = true;
                    return;
                }
            }
            _s["AT" + mi.Name.Substring(2)] = _answerTypes[int.Parse(mi.Tag.ToString())] = mi.IsChecked;
        }

        private void miAutoPronounce_Click(object sender, RoutedEventArgs e)
        {
            _s.AutoPronounce = (sender as MenuItem).IsChecked;
        }

        private void miCopy_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(lblQuestion.Content.ToString());
        }

        private void miCountdown_Click(object sender, RoutedEventArgs e)
        {
            if (miCountdown.IsChecked)
            {
                lblCountdown.Content = null;
                spCountdown.ToolTip = miCountdown.ToolTip;
                spCountdown.Visibility = Visibility.Visible;
            }
            else
            {
                spCountdown.Visibility = Visibility.Hidden;
                _timerCountdown.Stop();
            }
            _s.Countdown = miCountdown.IsChecked;
        }

        private void miExit_Click(object sender, RoutedEventArgs e)
        {
            Window_Closing(sender, new CancelEventArgs());
        }

        private void miLoadXlsm_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = ".xlsm";
            ofd.FileName = _xlsmSafeFileName;
            ofd.Filter = "Excel Macro-Enabled (*.xlsm)|*.xlsm";
            ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (ofd.ShowDialog() == true)
            {
                _s.DictionaryPath = ofd.FileName;
                gameInit();
            }
        }

        private void miOpenXlsm_Click(object sender, RoutedEventArgs e)
        {
            showInExcel(_missedODictIdx);
        }

        private void miReveal_Click(object sender, RoutedEventArgs e)
        {
            chooseWrong("revealed");
            showInExcel(_correct.correctODictIdx);
        }

        private void mirbLangueage_Checked(object sender, RoutedEventArgs e)
        {
            _s.Language = (sender as RadioButton).Tag.ToString();
            miSettings.IsSubmenuOpen = false;
            MessageBox.Show(t("msgNeedRestart"));
        }

        private void mirbRepeatingLimit_Checked(object sender, RoutedEventArgs e)
        {
            _repeatingLimit = _s.RepeatingsLimit = int.Parse((sender as RadioButton).Tag.ToString());
            miSettings.IsSubmenuOpen = false;
        }

        private void miRecords_Click(object sender, RoutedEventArgs e)
        {
            _wRecords.ShowDialog();
        }

        private void miResetSettings_Click(object sender, RoutedEventArgs e)
        {
            int iTag;

            for (int i = 0; i < miAnswerTypes.Items.Count; i++)
            {
                MenuItem mi = miAnswerTypes.Items[i] as MenuItem;
                string t = "AT" + mi.Name.Substring(2);
                _answerTypes[i] = mi.IsChecked = bool.Parse(_s.Properties["AT" + mi.Name.Substring(2)].DefaultValue.ToString());
            }
            miAutoPronounce.IsChecked = bool.Parse(_s.Properties["AutoPronounce"].DefaultValue.ToString());
            miCountdown.IsChecked = bool.Parse(_s.Properties["Countdown"].DefaultValue.ToString());
            foreach (RadioButton mirb in miRepeatingLimit.Items)
            {
                iTag = int.Parse(mirb.Tag.ToString());
                if (iTag == int.Parse(_s.Properties["RepeatingsLimit"].DefaultValue.ToString()))
                {
                    _repeatingLimit = iTag;
                    mirb.IsChecked = true;
                }
            }
            mirbEnglish.IsChecked = true;
            miRepeatWrongs.IsChecked = bool.Parse(_s.Properties["miRepeatWrongs"].DefaultValue.ToString());
        }

        #endregion

        #region Worker Events

        private void Worker_Init(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(0, t("init"));

            _timerAfterChoice.Tick += (_, __) =>
            {
                Dispatcher.Invoke(new Action(() =>
                {
                    _timerAfterChoice.Stop();
                    if (lblCorrect.Visibility == Visibility.Visible)
                    {
                        if (_isToolTipRevealed)
                        {
                            points -= 5;
                            if (points < 0) points = 0;
                        }
                        else points += 5;
                    }
                    else if (lblWrong.Tag.ToString() == "timeout")
                    {
                        points -= 15;
                        if (points < 0) points = 0;
                    }
                    else points = 0;
                    lblPoints.Content = points;
                    askQuestion();
                }));
            };
            _timerAfterChoice.Interval = TimeSpan.FromSeconds(_s.TimeAfterChoice);

            _timerCountdown.Tick += (_, __) =>
            {
                Dispatcher.Invoke(new Action(() =>
                {
                    lblCountdown.Content = --_seconds;
                    if (_seconds == 0)
                    {
                        _timerCountdown.Stop();
                        chooseWrong("timeout");
                    }
                }));
            };
            _timerCountdown.Interval = new TimeSpan(0, 0, 1);
        }

        private void Worker_InitComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            bool isWin8OrAbove = Environment.OSVersion.Platform == PlatformID.Win32NT
                && Environment.OSVersion.Version >= new Version(6, 2, 9200, 0);
            int i;
            Thickness mirbMargin = new Thickness(-24, 0, -50, 0);
            Thickness mirbPadding = new Thickness(15, 0, 0, 0);
            Thickness rbMargin = new Thickness(0, 10, 0, 0);

            _wRecords = new RecordsWindow(this);
            MinHeight = ActualHeight;
            MinWidth = ActualWidth;

            if (!Directory.Exists("sounds")) Directory.CreateDirectory("sounds");

            for (i = 0; i < miAnswerTypes.Items.Count; i++)
            {
                MenuItem mi = miAnswerTypes.Items[i] as MenuItem;
                mi.Click += miAnswerTypes_Click;
                mi.IsCheckable = true;
                _answerTypes[i] = mi.IsChecked = (bool)_s["AT" + mi.Name.Substring(2)];
            }
            miAutoPronounce.IsChecked = _s.AutoPronounce;
            if (_s.Countdown)
            {
                miCountdown.IsChecked = true;
                miCountdown_Click(null, null);
            }
            foreach (RadioButton mirb in miRepeatingLimit.Items)
            {
                mirb.Margin = mirbMargin;
                mirb.Padding = mirbPadding;
                int iTag = int.Parse(mirb.Tag.ToString());
                if (iTag == _s.RepeatingsLimit)
                {
                    _repeatingLimit = iTag;
                    mirb.IsChecked = true;
                }
                mirb.Checked += mirbRepeatingLimit_Checked;
            }
            if (WinAPI.GetSystemDefaultLCID() != 1071)
                mirbMacedonian.IsEnabled = false;
            else if (_s.Language == "mk-MK") mirbMacedonian.IsChecked = true;
            foreach (RadioButton mirb in miLanguage.Items)
            {
                mirb.Checked += mirbLangueage_Checked;
                mirb.Margin = mirbMargin;
                mirb.Padding = mirbPadding;
            }
            miRepeatWrongs.Header = String.Format(t("miRepeatWrongs"), _s.WrongsLimit);
            miRepeatWrongs.IsChecked = _s.RepeatWrongs;
            
            for (i = 0; i < 5; i++)
            {
                RadioButton rb = new RadioButton();
                rb.Height = 50;
                rb.Width = spRbs.ActualWidth;
                rb.Margin = rbMargin;
                rb.Click += rb_Click;
                rb.MouseRightButtonUp += rb_MouseRightButtonUp;
                rb.Tag = i;
                if (isWin8OrAbove) rb.VerticalContentAlignment = VerticalAlignment.Center;

                TextBlock tb = new TextBlock();
                tb.FontSize = 16;
                tb.TextWrapping = TextWrapping.Wrap;
                tb.ToolTip = new ToolTip { FontSize = 15, Visibility = Visibility.Hidden };

                rb.Content = tb;
                spRbs.Children.Add(rb);
            }

            _bgWorker.DoWork -= Worker_Init;
            _bgWorker.RunWorkerCompleted -= Worker_InitComplete;

            if (!File.Exists(_s.DictionaryPath))
            {
                miLoadXlsm_Click(null, null);
                if (!File.Exists(_s.DictionaryPath))
                {
                    _s.DictionaryPath = "";
                    _s.Save();
                    Environment.Exit(0);
                }
            }
            else gameInit();
        }

        private void Worker_Sound(object sender, DoWorkEventArgs e)
        {
            string sName = e.Argument.ToString();
            Uri soundFile = new Uri(Path.GetFullPath("sounds/" + sName.ToLower() + ".mp3"));

            if (_isInternetOk && !File.Exists(soundFile.LocalPath))
            {
                Dispatcher.Invoke(new Action(() =>
                {
                    gifImage.ClearValue(GifImage.CursorProperty);
                    gifImage.GifSource = "/images/Loader.gif";
                    gifImage.MouseUp -= gifImage_MouseUp;
                    gifImage.StartAnimation();
                    media.Source = null;
                }));

                WebClient wc = new WebClient();
                Action<string, string> aDownload = (fileName, filePath) => { wc.DownloadFile(_s.GStaticLink + fileName + ".mp3", filePath); };
                if (sName.Contains(" "))
                    try { aDownload(sName.Replace(" ", "_"), soundFile.LocalPath); }
                    catch (WebException)
                    {
                        string[] words = Regex.Split(sName, " ");
                        foreach (string word in words)
                            try { aDownload(word, "sounds/" + word + ".mp3"); }
                            catch (WebException) { break; }
                    }
                else
                    try { aDownload(sName, soundFile.LocalPath); }
                    catch (WebException)
                    {
                        try { aDownload(sName + "@1", soundFile.LocalPath); }
                        catch (WebException) { }
                    }
            }
            e.Result = sName;
        }

        private void Worker_SoundComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            string sName = e.Result.ToString();
            Action<Uri> aPrepareToListen = uri =>
            {
                gifImage.Cursor = Cursors.Hand;
                gifImage.Source = _biSound;
                gifImage.MouseUp += gifImage_MouseUp;
                if (_s.AutoPronounce) gifImage_MouseUp(null, null);
            };
            Uri soundFile = new Uri(Path.GetFullPath("sounds/" + sName + ".mp3"));

            gifImage.StopAnimation();
            gifImage.ClearValue(GifImage.CursorProperty);
            gifImage.ClearValue(GifImage.GifSourceProperty);
            gifImage.ClearValue(GifImage.SourceProperty);
            gifImage.MouseUp -= gifImage_MouseUp;

            _lPlayWords.Clear();
            if (File.Exists(soundFile.LocalPath))
            {
                _lPlayWords.Add(soundFile);
                aPrepareToListen(soundFile);
            }
            else if (sName.Contains(" "))
            {
                bool isOk = true;
                string[] words = Regex.Split(sName, " ");
                foreach (string word in words)
                {
                    soundFile = new Uri(Path.GetFullPath("sounds/" + word + ".mp3"));
                    if (!File.Exists(soundFile.LocalPath))
                    {
                        isOk = false;
                        return;
                    }
                    _lPlayWords.Add(soundFile);
                }
                if (isOk) aPrepareToListen(_lPlayWords[0]);
            }
        }

        private void Worker_Startup(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(0, String.Format(t("dictGeneration"), _xlsmSafeFileName));

            bool hasDuplicates = false;
            int flags, i = 2;
            Excel.App excel = new Excel.App(_s.DictionaryPath);
            string fn = "dat/" + xlsmSafeFileNameNoExt + _s.RepeatsSuffix, key = excel.getString("A" + i), prevKey = "";

            _odict.Clear();
            while (!String.IsNullOrEmpty(key == null ? null : key.Trim()))
            {
                if (prevKey == key)
                {
                    hasDuplicates = true;
                    break;
                }
                prevKey = key;
                _odict[key] = new Translation(excel, i, key);
                key = excel.getString("A" + ++i);
            }
            excel.Close();

            if (hasDuplicates || _odict.Count < 5)
            {
                _bgWorker.ReportProgress(100);
                MessageBox.Show(String.Format(t("msgErrorOdict"), _xlsmSafeFileName));
                if (hasDuplicates) MessageBox.Show(String.Format(t("msgDuplicateFound"), key));
                _s.DictionaryPath = "";
                _s.Save();
                Environment.Exit(0);
            }

            _bgWorker.ReportProgress(60, String.Format(t("loadingRepetitions"), fn));
            if (File.Exists(fn))
                using (FileStream fs = new FileStream(fn, FileMode.Open))
                    _dictRepeats = binFormatter.Deserialize(fs) as Dictionary<string, byte>;
            //foreach (var kvp in _dictRepeats) Console.WriteLine("{0}, {1}", kvp.Key, kvp.Value);

            fn = "dat/" + xlsmSafeFileNameNoExt + _s.WrongsSuffix;
            _bgWorker.ReportProgress(70, String.Format(t("loadingWrongs"), fn));
            if (File.Exists(fn))
                using (FileStream fs = new FileStream(fn, FileMode.Open))
                    _odictWrongs = binFormatter.Deserialize(fs) as OrderedDictionary;

            fn = "dat/" + xlsmSafeFileNameNoExt + _s.RecordsSuffix;
            _bgWorker.ReportProgress(80, String.Format(t("deserializeRecords"), fn));
            _wRecords.deserializeOC(fn);

            _bgWorker.ReportProgress(90, t("checkConnection"));
            _isInternetOk = WinAPI.InternetGetConnectedState(out flags, 0);
            if (_isInternetOk)
                _isInternetOk = WinAPI.InternetCheckConnection("http://www.google.com/", 1, 0);
        }

        private void Worker_StartupComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            _bgWorker.DoWork -= Worker_Startup;
            _bgWorker.RunWorkerCompleted -= Worker_StartupComplete;
            _bgWorker.DoWork += Worker_Sound;
            _bgWorker.RunWorkerCompleted += Worker_SoundComplete;

            _wLoading.Hide();
            dockMain.Visibility = Visibility.Visible;
            Title = String.Format(_formattedTitle, _xlsmSafeFileName);
            askQuestion();
        }

        #endregion

        #region Other Control Events

        private void gifImage_MouseUp(object sender, MouseButtonEventArgs e)
        {
            media.Source = _lPlayWords[_iPlayWords = 0];
            media.Play();
        }

        private void media_MediaEnded(object sender, RoutedEventArgs e)
        {
            media.Stop();
            media.Position = TimeSpan.Zero;
            if (++_iPlayWords < _lPlayWords.Count)
            {
                media.Source = _lPlayWords[_iPlayWords];
                media.Play();
            }
        }

        private void rb_Click(object sender, RoutedEventArgs e)
        {
            if ((int)(sender as RadioButton).Tag == _correct.correctRbIdx)
            {
                lblCorrect.Content = t("lblCorrect");
                if (_wRecords.ocRecordsList.Count > 0)
                {
                    int tempPoints = points;
                    if (_isToolTipRevealed)
                    {
                        tempPoints -= 5;
                        if (tempPoints < 0) tempPoints = 0;
                    }
                    int diffToRecord = _wRecords.ocRecordsList.First().Score - tempPoints + 5;
                    if (diffToRecord > 0 && diffToRecord <= 15)
                        lblCorrect.Content = String.Format(t("lblCorrectBeforeRecord"), diffToRecord);
                }
                lblWrong.Visibility = Visibility.Collapsed;
                lblCorrect.Visibility = Visibility.Visible;
                changeRepeats(1);
                _odictWrongs.Remove(_correct.keyEnglish);
                if (_odictWrongs.Count == 0) _isReadingWrongs = false;
            }
            else
            {
                lblCorrect.Visibility = Visibility.Collapsed;
                lblWrong.Tag = e.Source is string ? e.Source : "";
                switch (lblWrong.Tag.ToString())
                {
                    case "revealed": lblWrong.Content = t("lblRevealed"); break;
                    case "timeout": lblWrong.Content = t("lblTimeUp"); break;
                    default: lblWrong.Content = t("lblWrong"); break;
                }
                lblWrong.Visibility = Visibility.Visible;
                TextBlock tb = (spRbs.Children[_correct.correctRbIdx] as RadioButton).Content as TextBlock;
                tb.Background = Brushes.LightSkyBlue;
                tb.Foreground = Brushes.Black;
                changeRepeats(-1);
                _missedODictIdx = _correct.correctODictIdx;
                _odictWrongs[_correct.keyEnglish] = _correct.answer;
                saveRecord();
            }
            spRbs.IsEnabled = false;
            if (_timerCountdown.IsEnabled)
                _timerCountdown.Stop();
            _timerAfterChoice.Start();
        }

        private void rb_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            TextBlock tb = (sender as RadioButton).Content as TextBlock;
            Translation trans = tb.Tag as Translation;
            ToolTip tt = tb.ToolTip as ToolTip;
            string ttContent = tt.Content as string;

            if (tb.FontWeight != FontWeights.Bold || trans.llMacedonian.Count == 0) return;
            if (String.IsNullOrEmpty(ttContent))
                tt.Content = trans.llMacedonian.First().First();
            else
            {
                List<string> lAll = new List<string>();
                foreach (List<string> l in trans.llMacedonian)
                    foreach (string s in l)
                        lAll.Add(s);
                tt.Content = lAll[lAll.Count == 1 ? 0 : (lAll.IndexOf(ttContent) + 1) % lAll.Count];
            }
            tt.HorizontalOffset = 10;
            tt.Placement = PlacementMode.Right;
            tt.PlacementTarget = tb;
            tt.IsOpen = true;
            tt.Visibility = Visibility.Visible;

            _isToolTipRevealed = true;
            changeRepeats(-1);
        }

        #endregion

#endregion

#region Private Methods

        private void askQuestion()
        {
            int answerType = -1;
            string str;
            List<int> lIdxs = Enumerable.Range(0, _odict.Count).ToList();
            RadioButton rb;
            Random rnd = new Random();
            StringBuilder sbHash = new StringBuilder();
            TextBlock tb;
            Translation correctTrans = null, wrongTrans;
            ToolTip tt;

            _correct.answer = null;
            _correct.correctRbIdx = rnd.Next(5);
            rb = spRbs.Children[_correct.correctRbIdx] as RadioButton;
            tb = rb.Content as TextBlock;

            if (!_isReadingWrongs && _odictWrongs.Count >= _s.WrongsLimit) _isReadingWrongs = true;
            if (_isReadingWrongs && miRepeatWrongs.IsChecked)
            {
                _correct.keyEnglish = _odictWrongs.Keys.OfType<string>().First();
                if (_odict.Contains(_correct.keyEnglish))
                {
                    _correct.answer = _odictWrongs[0] as string;
                    correctTrans = _odict[_correct.keyEnglish] as Translation;
                    tb.ClearValue(TextBlock.FontStyleProperty);
                    tb.ClearValue(TextBlock.FontWeightProperty);

                    if (correctTrans.llLexicon.Contains(_correct.answer))
                    {
                        tb.FontStyle = FontStyles.Italic;
                        answerType = 0;
                    }
                    else if (correctTrans.findSynonym(_correct.answer))
                    {
                        answerType = 1;
                        tb.FontWeight = FontWeights.Bold;
                    }
                    else if (correctTrans.findMacedonian(_correct.answer))
                        answerType = 2;

                    if (answerType > -1)
                    {
                        lblQuestion.Content = correctTrans.keyEnglish;
                        lIdxs.Remove(correctTrans.oIdx);
                        _correct.correctODictIdx = correctTrans.oIdx;
                        tb.Tag = correctTrans;
                        tb.Text = _correct.answer;
                    }
                    else
                    {
                        _odictWrongs.Remove(_correct.keyEnglish);
                        _correct.answer = null;
                        tb.Tag = null;
                    }
                }
            }

            Queue<int> qUniqueIdxs = new Queue<int>(lIdxs.OrderBy(x => rnd.Next()));
            if (String.IsNullOrEmpty(_correct.answer))
            {
                while (qUniqueIdxs.Count > 0)
                {
                    _correct.correctODictIdx = qUniqueIdxs.Dequeue();
                    correctTrans = _odict[_correct.correctODictIdx] as Translation;
                    lblQuestion.Content = correctTrans.keyEnglish;
                    _correct.answer = correctTrans.getRandomTranslation(tb, _answerTypes);
                    if (_correct.answer != "")
                    {
                        _correct.keyEnglish = correctTrans.keyEnglish;
                        if (!_dictRepeats.ContainsKey(_correct.repeatsKey) || _dictRepeats[_correct.repeatsKey] < _repeatingLimit)
                            break;
                    }
                }
                if (qUniqueIdxs.Count < 5)
                {
                    MessageBox.Show(String.Format(t("msgMaster"), _s.RepeatsSuffix));
                    return;
                }
            }

            for (int i = 0; i < 5; i++)
            {
                rb = spRbs.Children[i] as RadioButton;
                rb.IsChecked = false;
                tb = rb.Content as TextBlock;
                tb.ClearValue(TextBlock.BackgroundProperty);
                tb.ClearValue(TextBlock.ForegroundProperty);
                tt = tb.ToolTip as ToolTip;
                tt.Content = "";
                tt.IsOpen = false;
                tt.Visibility = Visibility.Hidden;
                if (i != _correct.correctRbIdx)
                {
                    if (qUniqueIdxs.Count < 5 - i)
                    {
                        MessageBox.Show(String.Format(t("msgMaster"), _s.RepeatsSuffix));
                        return;
                    }
                    while (qUniqueIdxs.Count > 0)
                    {
                        wrongTrans = _odict[qUniqueIdxs.Dequeue()] as Translation;
                        if (correctTrans.findSynonym(wrongTrans.keyEnglish)) continue;
                        str = wrongTrans.getRandomTranslation(tb, _answerTypes);
                        if (!correctTrans.findMacedonian(str) && str != "") break;
                    }
                }
            }

            _isToolTipRevealed = false;
            if (!_bgWorker.IsBusy)
                _bgWorker.RunWorkerAsync(Regex.Replace(lblQuestion.Content as string, @" \(\w+\)", ""));
            lblCorrect.Visibility = Visibility.Hidden;
            lblWrong.Visibility = Visibility.Hidden;
            if (miCountdown.IsChecked)
            {
                lblCountdown.Content = _seconds = _s.CountdownSeconds;
                spCountdown.ClearValue(StackPanel.ToolTipProperty);
                _timerCountdown.Start();
            }
            spRbs.IsEnabled = true;
        }

        private void changeRepeats(int direction)
        {
            if (_dictRepeats.ContainsKey(_correct.repeatsKey))
            {
                _dictRepeats[_correct.repeatsKey] += (byte)direction;
                if (_dictRepeats[_correct.repeatsKey] > 100) _dictRepeats[_correct.repeatsKey] = 100;
                else if (_dictRepeats[_correct.repeatsKey] == 0) _dictRepeats.Remove(_correct.repeatsKey);
            }
            else if (direction == 1) _dictRepeats[_correct.repeatsKey] = 1;
        }

        private void chooseWrong(string reason)
        {
            rb_Click(
                spRbs.Children.OfType<RadioButton>().First(x => (int)x.Tag != _correct.correctRbIdx),
                new RoutedEventArgs(null, reason)
            );
        }

        private void gameInit()
        {
            _xlsmSafeFileName = Path.GetFileName(_s.DictionaryPath);
            xlsmSafeFileNameNoExt = Path.GetFileNameWithoutExtension(_s.DictionaryPath);
            points = 0;
            dockMain.Visibility = Visibility.Hidden;
            _bgWorker.DoWork -= Worker_Sound;
            _bgWorker.RunWorkerCompleted -= Worker_SoundComplete;
            _bgWorker.DoWork += Worker_Startup;
            _bgWorker.RunWorkerCompleted += Worker_StartupComplete;
            _wLoading.Show();
            _bgWorker.RunWorkerAsync();
        }

        private void saveRecord()
        {
            if (points < 30) return;
            if (_wRecords.ocRecordsList.Count == 0 || points > _wRecords.ocRecordsList.First().Score)
            {
                _wRecords.isSaving = true;
                _wRecords.ShowDialog();
            }
        }

        private void showInExcel(int oDictIdx)
        {
            int row = oDictIdx + 2;

            Excel.App excel = new Excel.App(_s.DictionaryPath, true);
            excel.selectRange("A" + row + ":D" + row);
            WinAPI.ShowWindow(excel.MainWindowHandle, WinAPI.SW_SHOWMAXIMIZED);
        }

        private string t(string key)
        {
            return (string)Application.Current.FindResource(key);
        }

#endregion

    }
}