using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelMoodleAddin
{
    public partial class Ribbon
    {
        private Excel.Style HeadingStyle { get; set; }
        private Excel.Style DescStyle { get; set; }

        private void HelpButton_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                text: "First click on the \"Create Question Sheet\" item \nAnd after finish, click on \"Save questions\" to save as GIFT format \nThen enter Moodle and import the file there",
                caption: "Extension Help",
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Question,
                defaultButton: MessageBoxDefaultButton.Button1,
                options: MessageBoxOptions.RightAlign,
                displayHelpButton: false
                );
        }

        private void AboutButton_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                text: "Excel Moodle Extension",
                caption: "Help",
                buttons: MessageBoxButtons.OK,
                icon: MessageBoxIcon.Information,
                options: MessageBoxOptions.RightAlign,
                defaultButton: MessageBoxDefaultButton.Button1,
                displayHelpButton: false
                );
        }

        private void BuildSheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet newWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            newWorksheet.Name = $"Moodle Question Editor {Globals.ThisAddIn.Counter++}";

            Excel.Range range1 = newWorksheet.get_Range("A1", "J3");
            range1.Style = GetDescriptionStyle();

            Excel.Range range2 = newWorksheet.get_Range("J1");
            range2.Value2 = "This sheet is for adding Moodle questions. Read Help menu for more info.";

            Excel.Range range3 = newWorksheet.get_Range("A4", "J4");
            range3.Style = "Heading 1";

            newWorksheet.get_Range("J4").Value2 = "Question Title";
            newWorksheet.get_Range("I4").Value2 = "Question Body";
            newWorksheet.get_Range("H4").Value2 = "Correct answer";
            newWorksheet.get_Range("G4").Value2 = "Feedbacl";
            newWorksheet.get_Range("F4").Value2 = "Wrong answer";
            newWorksheet.get_Range("E4").Value2 = "Feedback";
            newWorksheet.get_Range("D4").Value2 = "Wrong answer";
            newWorksheet.get_Range("C4").Value2 = "Feedback";
            newWorksheet.get_Range("B4").Value2 = "Wrong Answer";
            newWorksheet.get_Range("A4").Value2 = "Feedback";

            range3.Columns.ColumnWidth = 15;
        }

        public Excel.Style GetHeadingStyle()
        {
            if (HeadingStyle == null)
            {
                Excel.Style style = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("Moodle Title Style");

                style.Font.Size = 15;
                style.Font.Bold = true;
                style.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DodgerBlue);
                style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                style.Interior.Pattern = Excel.XlPattern.xlPatternSolid;

                style.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                style.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                style.Borders[XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DodgerBlue);

                HeadingStyle = style;

                return style;
            }
            else
            {
                return HeadingStyle;
            }
        }

        public Excel.Style GetDescriptionStyle()
        {
            if (DescStyle == null)
            {
                Excel.Style style = Globals.ThisAddIn.Application.ActiveWorkbook.Styles.Add("Moodle Desc Style");

                style.Font.Size = 12;
                style.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                style.Interior.Pattern = Excel.XlPattern.xlPatternSolid;

                DescStyle = style;

                return style;
            }
            else
            {
                return DescStyle;
            }
        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            string gift = "";

            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

            for (int i = 5; ; i++)
            {
                string questionTitle  = activeSheet.get_Range($"J{i}").Value2;
                string questionText   = activeSheet.get_Range($"I{i}").Value2;
                string rightAnswer    = activeSheet.get_Range($"H{i}").Value2;
                string rightFeedback  = activeSheet.get_Range($"G{i}").Value2;
                string wrongAnswer1   = activeSheet.get_Range($"F{i}").Value2;
                string wrong1Feedback = activeSheet.get_Range($"E{i}").Value2;
                string wrongAnswer2   = activeSheet.get_Range($"D{i}").Value2;
                string wrong2Feedback = activeSheet.get_Range($"C{i}").Value2;
                string wrongAnswer3   = activeSheet.get_Range($"B{i}").Value2;
                string wrong3Feedback = activeSheet.get_Range($"A{i}").Value2;

                if (string.IsNullOrWhiteSpace(questionText))
                    break;

                gift += $"::{questionTitle}:: {questionText} {Environment.NewLine} {{ = {rightAnswer} # {rightFeedback} ~{wrongAnswer1} # {wrong1Feedback} ~{wrongAnswer2} # {wrong2Feedback} ~{wrongAnswer3} # {wrong3Feedback} }}{Environment.NewLine}{Environment.NewLine}";
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                FileName = $"{activeSheet.Name} - GIFT Export.txt",
                Title = "Moodle output"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    File.WriteAllText(saveFileDialog.FileName, gift, Encoding.UTF8);
                } catch
                {
                    MessageBox.Show(
                       text: "Error in saving",
                       caption: "Error",
                       buttons: MessageBoxButtons.OK,
                       icon: MessageBoxIcon.Exclamation,
                       defaultButton: MessageBoxDefaultButton.Button1,
                       options: MessageBoxOptions.RightAlign,
                       displayHelpButton: false
                       );
                }
            }
        }
    }
}
