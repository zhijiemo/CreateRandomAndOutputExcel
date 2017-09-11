using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace CreateRandomAndOutputExcel
{
    class CreateRandomTable
    {
        public void createRandomTable()
        {
            int i, j;
            //double fValue;
            Random rand = new Random();
            double[] doubledata = new double[] { rand.Next(10), rand.NextDouble(), 2.3, 4.5, 6.9 };

            using (SLDocument sl = new SLDocument())
            {
                sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Random");

                for (i = 1; i <= 16; ++i)
                {
                    for (j = 1; j <= 6; ++j)
                    {
                        switch (rand.Next(5))
                        {
                            case 0:
                            case 1:
                                sl.SetCellValue(i, j, doubledata[rand.Next(doubledata.Length)]);
                                break;
                            case 2:
                            case 3:
                                sl.SetCellValue(i, j, rand.NextDouble() * 1000.0 + 350.0);
                                break;

                            case 4:
                                if (rand.NextDouble() < 0.5)
                                {
                                    sl.SetCellValueNumeric(i, j, "3.1415926535898");
                                }
                                else
                                {
                                    sl.SetCellValueNumeric(i, j, "2.7182818284590");
                                }
                                break;
                        }
                    }
                }
                sl.SaveAs("Miscellaneous.xlsx");
            }
        }
    }
}
