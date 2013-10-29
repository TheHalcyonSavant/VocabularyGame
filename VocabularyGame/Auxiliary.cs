using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows;
using System.Runtime.InteropServices;
using System.Net;

namespace VocabularyGame
{
    [Serializable]
    public class Record
    {
        public int Score { get; set; }
        public string Name { get; set; }
        public DateTime Time { get; set; }

        public Record(string name, int score)
        {
            Name = name;
            Score = score;
            Time = DateTime.Now;
        }
    }

    public class CorrectClass
    {
        public int correctODictIdx;
        public int correctRbIdx;
        public string answer;
        public string keyEnglish;
        public string repeatsKey { get { return keyEnglish + "=" + answer; } }
    }

    public class Translation
    {
        private const string LEXICON_COL = "B";
        private const string MACEDONIAN_COL = "D";
        private const string SYNONYMS_COL = "C";

        private int _row;
        private Excel.App _excel;

        public int oIdx { get { return _row - 2; } }
        public string keyEnglish;
        public List<string> llLexicon;
        public List<List<string>> llMacedonian;
        public List<List<string>> llSynonyms;

        public Translation(Excel.App excel, int row, string english)
        {
            _excel = excel;
            _row = row;
            keyEnglish = english;

            string[] subL = getCellStrings(LEXICON_COL);
            llLexicon = new List<string>();
            foreach (string defns in subL)
                llLexicon.Add(defns.Replace("\n", ""));

            fillLL(MACEDONIAN_COL, ref llMacedonian);
            fillLL(SYNONYMS_COL, ref llSynonyms);
        }

        public bool findMacedonian(string needle)
        {
            foreach (List<string> l in llMacedonian)
                if (l.Contains(needle)) return true;
            return false;
        }

        public bool findSynonym(string needle)
        {
            foreach (List<string> l in llSynonyms)
                if (l.Contains(needle)) return true;
            return false;
        }

        public string getRandomTranslation(TextBlock tb, bool[] answerTypes)
        {
            List<string> subL;
            Random rnd = new Random();

            tb.ClearValue(TextBlock.FontStyleProperty);
            tb.ClearValue(TextBlock.FontWeightProperty);
            int[] options = Enumerable.Range(0, 3).OrderBy(x => rnd.Next()).ToArray();
            for (int i = 0; i < options.Length; i++)
            {
                switch (options[i])
                {
                    case 0:
                        if (!answerTypes[0] || llLexicon.Count == 0) continue;
                        tb.Text = llLexicon[rnd.Next(llLexicon.Count)];
                        tb.FontStyle = FontStyles.Italic;
                        return tb.Text;
                    case 1:
                        if (!answerTypes[1] || llSynonyms.Count == 0) continue;
                        subL = llSynonyms[rnd.Next(llSynonyms.Count)];
                        tb.Text = subL[rnd.Next(subL.Count)];
                        tb.FontWeight = FontWeights.Bold;
                        return tb.Text;
                    case 2:
                        if (!answerTypes[2] || llMacedonian.Count == 0) continue;
                        subL = llMacedonian[rnd.Next(llMacedonian.Count)];
                        tb.Text = subL[rnd.Next(subL.Count)];
                        return tb.Text;
                }
            }
            return "";
        }

        private string[] getCellStrings(string col)
        {
            string str = _excel.getString(col + _row);
            if (String.IsNullOrEmpty(str)) return new string[0];
            return str.Split(';');
        }

        private void fillLL(string col, ref List<List<string>> ll)
        {
            string[] subSubL, subL = getCellStrings(col);
            ll = new List<List<string>>();
            foreach (string strSubL in subL)
            {
                ll.Add(new List<string>());
                subSubL = strSubL.Split(',');
                foreach (string word in subSubL)
                    ll.Last().Add(word.Trim());
            }
        }
    }

    public class WinAPI
    {
        public const int SW_SHOWMAXIMIZED = 3;

        [DllImport("wininet.dll", SetLastError = true)]
        public static extern bool InternetCheckConnection(string lpszUrl, int dwFlags, int dwReserved);

        [DllImport("wininet.dll", SetLastError = true)]
        public static extern bool InternetGetConnectedState(out int lpdwFlags, int dwReserved);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("kernel32.dll")]
        public static extern uint GetSystemDefaultLCID();
    }
}