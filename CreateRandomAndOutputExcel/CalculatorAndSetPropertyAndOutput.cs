using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System.IO;


namespace CreateRandomAndOutputExcel
{
    class CalculatorAndSetPropertyAndOutput
    {
        public SLDocument Save()
        {
            SLDocument slPass = new SLDocument("Miscellaneous.xlsx", "Sheet");
            slPass.SaveAs("Miscellaneous.xlsx");
            return slPass;
        }
        

        //MemoryStream msPass = new MemoryStream();


        public void SetCellDocument()
        {
            this.Save();
            SLDocument slPass = new SLDocument("Miscellaneous.xlsx", "Sheet");
            slPass.SetCellValue("A17", "So this is the random number table");
            SLStyle style = slPass.CreateStyle();
            style.SetFont("Impact", 24);
            style.Font.Underline = UnderlineValues.Single;
            slPass.SetCellStyle("A17", style);
            slPass.SaveAs("Miscellaneous.xlsx");
            //slPass.SaveAs(msPass);

            //tl.SaveAs(msFirstPass);
        }
        public void GetPresentTime()
        {
            SLDocument slPass = new SLDocument("Miscellaneous.xlsx", "Sheet");
            slPass.SetCellValue("G7", "The time is ");
            //private Columns AutoFit(SheetData sheetData)
            //slPass.SetCellStyle("G7");
            //slPass.ActiveSheet.Range("A2").ColumnWith = 100
            slPass.SetCellValue("H7", DateTime.Now.ToString());//获取当前日期和时间
            //slPass.SaveAs(msPass);
            slPass.SaveAs("Miscellaneous.xlsx");
            //slSecondPass.SaveAs(msThirdPass);

        }
        public void Sum()
        {
            SLDocument slPass = new SLDocument("Miscellaneous.xlsx", "Sheet");
            slPass.SetCellValue("G1", "The sum is");
            //两个方法实现功能一致
            slPass.SetCellValue("H1", "=SUM(A1:F2)");
            slPass.SetCellValue(SLConvert.ToCellReference(2, 8), string.Format("=SUM({0})", SLConvert.ToCellRange(1, 1, 2, 6)));
            slPass.SaveAs("Miscellaneous.xlsx");
            //slPass.SaveAs(msPass);

        }
        public void Average()
        {
            SLDocument slPass = new SLDocument("Miscellaneous.xlsx", "Sheet");
            //以下两个实现相同功能
            slPass.SetCellValue("G4", "The average number is");
            slPass.SetCellValue("H4", StringValue.ToString("=AVERAGE(A1:F2)"));
            slPass.SetCellValue("H5", String.Format("=AVERAGE({0})", SLConvert.ToCellRange(1, 1, 2, 6)));
            slPass.SaveAs("Miscellaneous.xlsx");
            //slPass.SaveAs(msPass);
        }
        public void Subtract()
        {
            SLDocument slPass = new SLDocument("Miscellaneous.xlsx", "Sheet");
            slPass.SetCellValue("G3", "The difference is");
            slPass.SetCellValue("H3", "=A2-A3");
            //slPass.SaveAs(msPass);
            slPass.SaveAs("Miscellaneous.xlsx");
        }
        public void SetPropertyAndOutput()
        {
            SLDocument slPass = new SLDocument("Miscellaneous.xlsx", "Sheet");
            slPass.DocumentProperties.Creator = "ZhouL";
            slPass.DocumentProperties.ContentStatus = "Secret";
            slPass.DocumentProperties.Title = "Random number table";
            slPass.DocumentProperties.Description = "Get data and manipulate it and export it";
            slPass.SaveAs("Miscellaneous.xlsx");

        }
    }
}
