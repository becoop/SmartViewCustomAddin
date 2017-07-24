using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        Dictionary<string, string> smartviewData =
            new Dictionary<string, string>();
        string[] searchTerms;
        string[] replaceTerms;


        public void WorkWithDocument(Microsoft.Office.Interop.Word.Document Doc)
        {
            searchTerms = new string[] { "P0", "P1", "Priority Rating", "ImageDate", "Equipment", "Location"};

            foreach (string i in searchTerms)
            {
                buildDictionary(i);
            }

            replaceTerms = new string[] { "P1Temperature", "P2Temperature", "DT1", "Photo 1."};

            foreach (string i in replaceTerms)
            {
                buildDoc(i);
            }
            buildTable();
            deleteMarkerTable();
            fixFooter();
        }

        private void buildDictionary(object searchVal)
        {
            int intFound = 0;
            int maxRows = Application.ActiveDocument.InlineShapes.Count / 2;
            bool skip = false;
            int dateSkip = 3;
            int dateCounter = 0;
            Application.Selection.Find.ClearFormatting();
            Application.Selection.Find.Forward = true;
            Application.ActiveDocument.Words[1].Select();
            Application.Selection.Find.Execute(searchVal, true, true, false, false, false, true, Word.WdFindWrap.wdFindContinue, false, "", false, false, false, false, false);
            while (Application.Application.Selection.Find.Found && intFound < maxRows)
            {
                if (Application.Selection.Information[Word.WdInformation.wdWithInTable])
                {

                   
                        Application.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1);
                  

                    if (Application.Selection.Text != "Dt1")
                    {

                       

                        if (searchVal.ToString() == "Priority Rating")
                        {
                            if (Application.Selection.Text.Trim() == "Location")
                            {
                                skip = true;
                            }
                            else
                            {
                                skip = false;
                            }
                        }
                        if (searchVal.ToString() == "Location")
                        {
                            if (Application.Selection.Text.Trim() == "Equipment")
                            {
                                skip = true;
                            }
                            else
                            {
                                skip = false;
                            }
                        }
                        if (searchVal.ToString() == "Equipment")
                        {
                            if (Application.Selection.Text.Trim() == "Sp1 Temperature")
                            {
                                skip = true;
                            }
                            else
                            {
                                skip = false;
                            }
                        }

                        if (!skip)
                        { 
                        intFound++;
                        smartviewData.Add(searchVal + intFound.ToString(), Application.Selection.Text);
                         }
                            
                    }
              
                    Application.Selection.Find.Execute();
                }
            } 
        }

        private void buildDoc(object searchVal)
        {
            int intFound = 0;
            int maxRows = Application.ActiveDocument.InlineShapes.Count / 2;
            Application.Selection.Find.ClearFormatting();
            Application.Selection.Find.Forward = true;
            Application.ActiveDocument.Words[1].Select();
            Application.Selection.Find.Execute(searchVal, true, true, false, false, false, true, Word.WdFindWrap.wdFindContinue, false, "", false, false, false, false, false);
            while (Application.Application.Selection.Find.Found && intFound < maxRows)
            {
                if (Application.Selection.Information[Word.WdInformation.wdWithInTable])
                {

                        intFound++;
                        string key = "";
                        string keyData = "";
                        switch (searchVal)
                        {
                            case "P1Temperature":
                                key = "P0" + intFound;
                                Application.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1);
                                keyData = tryDictionaryKey(key);
                                break;
                            case "P2Temperature":
                                key = "P1" + intFound;
                                Application.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1);
                                keyData = tryDictionaryKey(key);
                                break;
                           
                            case "DT1":
                                 Application.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1);
                                 string d1 = tryDictionaryKey(("P0" + intFound).ToString());
                                 double d1d = Convert.ToDouble(d1.Remove(d1.Length - 2).Trim());
                                 string d2 = tryDictionaryKey(("P1" + intFound).ToString());
                                double d2d = Convert.ToDouble(d2.Remove(d2.Length - 2).Trim());
                                 double delta1 = d1d - d2d;
                                keyData = delta1.ToString();
                                break;
                            case "Photo 1.":
                                keyData = "Photo " + intFound + ".";
                                break;
                           // case "Location":
                                //Application.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1);
                                //key = "Location" + intFound;
                                //keyData = tryDictionaryKey(key);
                                //break;
                           // case "Equipment":
                                //Application.Selection.MoveRight(Microsoft.Office.Interop.Word.WdUnits.wdCell, 1);
                                //key = "Equipment" + intFound;
                                //keyData = tryDictionaryKey(key);
                                //break;
                        }
                        Application.Selection.Text = keyData;
                    

                    Application.Selection.Find.Execute();
                }
            }
        }

        public  string tryDictionaryKey(string key)
        {
            if (smartviewData.ContainsKey(key))
            {
                string value = smartviewData[key];
                return value;
            }
            else
            {
                return null;
            }
        }

        private  void buildTable()
        {
            int rowcnt;
            int t = 0;
            int p = 0;
            int maxRows = Application.ActiveDocument.InlineShapes.Count / 2;

            Word.Table oTable =  Application.ActiveDocument.Bookmarks["summaryTable"].Range.Tables[1];
            if(oTable.Rows.Count > 1)
            {
                for(p=oTable.Rows.Count; p>1; p--)
                {
                    if(p != 1){
                        oTable.Rows[p].Delete();
                    }
                }
            }
            for (rowcnt = 1; rowcnt <= maxRows; rowcnt++)
            {

                oTable.Rows.Add();
                Word.Row oNewRow = oTable.Rows[oTable.Rows.Count];
                oNewRow.Cells[1].Range.Text = tryDictionaryKey("ImageDate" + rowcnt);
                oNewRow.Cells[2].Range.Text = rowcnt.ToString();
                oNewRow.Cells[3].Range.Text = tryDictionaryKey("Priority Rating" + rowcnt);
                oNewRow.Cells[4].Range.Text = tryDictionaryKey("Location" + rowcnt);
                oNewRow.Cells[5].Range.Text = tryDictionaryKey("Equipment" + rowcnt);
                oNewRow.Cells[6].Range.Text = tryDictionaryKey("P0" + rowcnt);
                string d1 = tryDictionaryKey(("P0" + rowcnt).ToString());
                double d1d = Convert.ToDouble(d1.Remove(d1.Length - 2).Trim());
                string d2 = tryDictionaryKey(("P1" + rowcnt).ToString());
                double d2d = Convert.ToDouble(d2.Remove(d2.Length - 2).Trim());
                double delta1 = d1d - d2d;
                oNewRow.Cells[7].Range.Text = delta1.ToString();
            }
                     
        }

        public static void clickedIt()
        {
            Globals.ThisAddIn.WorkWithDocument(Globals.ThisAddIn.Application.ActiveDocument);
            }

        private void deleteMarkerTable()
        {
            int t = 0;
            int maxRows = Application.ActiveDocument.InlineShapes.Count / 2;
            t = maxRows - 2;
            t = (t < 0 ? 0 : t);
            int maxRowCount = maxRows;
            int p = 0;
            
            bool once = false;
            int intFound = 0;

            for(int i =0; i<maxRows; i++)
            {
                int loc = 5*maxRowCount;
                if(!once)
                {   p = (6 + loc);
                    p = (p > Application.ActiveDocument.Tables.Count ? Application.ActiveDocument.Tables.Count : p);
                }
                else
                {
                  
                    p = (6 + loc) + i - t;
                }
                
                once = true;
                maxRowCount--;

                Application.ActiveDocument.Tables[p].Delete();
            }
            Application.ActiveDocument.Words[1].Select();
            Application.Selection.Find.ClearFormatting();
            Application.Selection.Find.Forward = true;
            Application.ActiveDocument.Words[1].Select();
            Application.Selection.Find.Execute("Main Image Markers", true, true, false, false, false, true, Word.WdFindWrap.wdFindContinue, false, "", false, false, false, false, false);

            while (Application.Selection.Find.Found && intFound < maxRows)
            {
                Application.Selection.Text = "";
                intFound++;
                Application.Selection.Find.Execute();
            }

        }

        private void fixFooter()
        {





            foreach (Word.Section wordSection in this.Application.ActiveDocument.Sections)
            {

                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
               // footerRange.Delete(Word.WdUnits.wdCharacter, footerRange.Characters.Count);
               // footerRange.Text = "";
                footerRange.Tables[1].Delete();
                
                //Word.Table oTable = Application.ActiveDocument.Bookmarks["footerTable"].Range.Tables[1]; oTable.Rows.Add();
                //Word.Row oNewRow = oTable.Rows[oTable.Rows.Count];
                //oNewRow.Cells[1].Range.Text = "21312#";

                //footerRange.Tables.Add(footerRange, 1, 3);
                // footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                //footerRange.Font.Size = 20;
                //footerRange.Text = "Confidential";
            }
        }

        public static void main(string[] args)
        {
            var wordApp = new Word.Application();
            wordApp.Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
        }

        void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
