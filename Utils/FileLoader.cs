using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;

namespace ExcelAddIn1.Utils
{
    class FileLoader
    {
        private OpenFileDialog openDlg;
        private System.Windows.Forms.Button selectBnt;

        public DialogResult OpenFileDialogForm()
        {
            openDlg = new OpenFileDialog()
            {
                FileName = "Select a LUA file",
                Filter = "LUA file (*.lua)|*.lua",
                Title = "Open LUA file"
            };

            selectBnt = new System.Windows.Forms.Button()
            {
                Size = new Size(100, 20),
                Location = new System.Drawing.Point(15, 15),
                Text = "Select file"
            };
            //selectBnt.Click += new EventHandler(selectBnt_Click);
            //Controls.Add(selectBnt);
            //return selectBnt;
            return openDlg.ShowDialog();
        }
        public void loadLua(Excel.Application app)
        {
            Excel.Worksheet sheet = app.ActiveSheet;
            try
            {
                var filePath = openDlg.FileName;
                using (Stream str = openDlg.OpenFile())
                {
                    TextReader tr = new StreamReader(str);
                    string text;
                    int lineNum = 1;
                    while ((text = tr.ReadLine()) != null)
                    {
                        string rangName = "A" + (lineNum++).ToString();//:A100";
                        Range cells = sheet.Range[rangName].Cells;
                        foreach (Range cell in cells)
                        {
                            cell.Value2 = text;
                        }
                        app.StatusBar = string.Format("loading {0}", lineNum);
                    }
                }
            }
            catch (SecurityException ex)
            {
                MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }
        }

        public void loadDataToMem(Excel.Application app)
        {
            try
            {
                var filePath = openDlg.FileName;
                using (Stream str = openDlg.OpenFile())
                {
                    TextReader tr = new StreamReader(str);
                    string text;
                    int lineNum = 1;
                    bool bFound = false;
                    string groupData = "";
                    string titleName = "";
                    while ((text = tr.ReadLine()) != null)
                    {
                        if (!bFound && text.Contains("Add("))
                        {
                            bFound = true;
                            string pattern = @"\[\[([^\]]+)\]";                            
                            Match m = Regex.Match(text, pattern);
                            if (m.Success)
                            {
                                titleName = m.Value.Replace("[","").Replace("]","");
                                continue;
                            }
                        }

                        if (bFound && text.Contains(");"))
                        {
                            bFound = false;                            
                        }

                        if (bFound)
                        {
                            groupData += text;
                        }


                        app.StatusBar = string.Format("loading {0}", lineNum);
                    }
                }
            }
            catch (SecurityException ex)
            {
                MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }

        }

        public void loadDataToMemByRegex(Excel.Application app)
        {
            try
            {
                var filePath = openDlg.FileName;
                using (Stream str = openDlg.OpenFile())
                {
                    TextReader tr = new StreamReader(str);
                    string text;
                    int lineNum = 1;                    
                    //string groupData = "";
                    string titleName = "";
                    string wholeString = "";
                    Dictionary<string, string> _dicEle = new Dictionary<string, string>();

                    wholeString = tr.ReadToEnd();

                    string patternSkill = @"Add\(\s*\[\[(\S+)\]\],(\s*--.*)*\r?\n\s*\{\r?\n(\s*((?!\s*\}).*\r?\n)*)\s*\}\s*\);\r?\n";
                    //스킬 단위로 분리
                    Regex r = new Regex(patternSkill, RegexOptions.IgnoreCase);
                    Match mSkill = r.Match(wholeString);
                    int matchCount = 0;
                    while(mSkill.Success)
                    {
                        //for (int n = 0; n < mSkill.Groups.Count; n++)
                        {
                            Group gD = mSkill.Groups[0];
                            Debug.WriteLine("Match Count = " + (++matchCount));
                            Debug.WriteLine(gD);
                            CaptureCollection cc = gD.Captures;
                            for (int j = 0; j < cc.Count; j++)
                            {
                                Capture cData = cc[j];
                                string groupData = cData.ToString();
                                string patternName = @"\[\[([^\]]+)\]";
                                Match mName = Regex.Match(groupData, patternName);
                                if (mName.Success)
                                {
                                    titleName = mName.Value.Replace("[", "").Replace("]", "");
                                }
                                string patternKeyValue = @"(\S+)\s*=\s*(.*),(\s*--.*)*\r?\n";
                                Regex element = new Regex(patternKeyValue, RegexOptions.IgnoreCase);
                                Match eleM = element.Match(groupData);                                
                                int eleCount = 0;
                                while(eleM.Success)
                                {
                                    Group eD = eleM.Groups[0];
                                    Debug.WriteLine("Ele Match Count = " + (++eleCount));
                                    Debug.WriteLine(eD);
                                    eleM = eleM.NextMatch();
                                }
                                //Debug.WriteLine(Regex.Replace(groupData, patternKeyValue, m => string.Format("{0}{1}", m.Value, m.Value)));
                            }
                        }
                        mSkill = mSkill.NextMatch();
                        app.StatusBar = string.Format("loading {0}", matchCount);
                    }                    
                }
            }
            catch (SecurityException ex)
            {
                MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }

        }
    }
}
