using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace StaffoutputAddIn
{
    public partial class Ribbon1
    {
        
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnCSVselect_Click(object sender, RibbonControlEventArgs e)
        {
            getRowCount("A");
            divideColomns();
        }

        private Int32 getRowCount(string column)
        {
            //Anzahl Reihen ermitteln. Dies wird ermittelt
            Int32 count;
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            for (count = 1; count < 65535; count++)
            {
                if (currentSheet.Cells[count, column].Text.Equals("") && currentSheet.Cells[count + 1, column].Text.Equals(""))
                {
                    //System.Windows.Forms.MessageBox.Show("Count: " + (count-1).ToString());
                    break;
                }

            }
            return count - 1;
        }
        private Int32 getRowCount(int column_index)
        {
            //Anzahl Reihen ermitteln. Dies wird ermittelt
            Int32 count;
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            for (count = 1; count < 65535; count++)
            {
                if (currentSheet.Cells[count, column_index].Text.Equals("") && currentSheet.Cells[count + 1, column_index].Text.Equals(""))
                {
                    //System.Windows.Forms.MessageBox.Show("Count: " + (count - 1).ToString());
                    break;
                }

            }
            return count - 1;
        }

        private void divideColomns()
        {
            string rowText;
            string[] rowTextArray;
            int textCounter = 0;
            
            Worksheet currentSheet = Globals.ThisAddIn.getActiveWorksheet();

            for (int i = 1; i <= getRowCount("A"); i++)
            {
                rowText = currentSheet.Cells[i,"A"] .Text;
                rowTextArray = rowText.Split(new Char[] { ',' });

                if(i==3)
                {
                    //Lass die Reihe so wie sie ist.
                }
                else if(i==5) /*Reihe 5 spezial*/
                {
                    int counter = 3;
                    currentSheet.Cells[i, 1].Value = rowTextArray[0]; 
                    currentSheet.Cells[i, 2].Value = rowTextArray[1];

                    while (2*counter-3 < rowTextArray.Count())
                    {
                        currentSheet.Cells[i, counter].Value = rowTextArray[2*(counter)-4]+", "+ rowTextArray[2*(counter)-3];
                        counter++;
                    }

                }
                else /*Alle anderen Reihen*/
                {
                    foreach (var item in rowTextArray)
                    {
                        textCounter++;
                        currentSheet.Cells[i, textCounter].Value = item;
                    }
                    textCounter = 0;
                }
                

            }
        }

    }
}
