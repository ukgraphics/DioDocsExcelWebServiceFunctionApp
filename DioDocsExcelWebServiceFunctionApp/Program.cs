using GrapeCity.Documents.Excel;
using System;

namespace DioDocsExcelWebServiceFunctionApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("DioDocs for ExcelでExcelのワークシートにWEBSERVICE、FILTERXML関数を追加します。");

            // トライアル版か製品版のライセンスキーを設定
            //Workbook.SetLicenseKey("");

            Workbook workbook = new Workbook();
            IWorksheet ws = workbook.Worksheets[0];

            // WEBSERVICE関数でデータを取得するURL
            ws.Range["A1"].Value = "URL";
            ws.Range["A2"].Value = "WEBSERVICE関数でデータを取得するURL";
            ws.Range["A2"].VerticalAlignment = VerticalAlignment.Center;
            ws.Range["B1"].Value = "https://rss-weather.yahoo.co.jp/rss/days/3410.xml"; // 仙台
            //ws.Range["B1"].Value = "https://rss-weather.yahoo.co.jp/rss/days/4410.xml"; // 東京

            // WEBSERVICE関数
            ws.Range["B2"].Formula = "=WEBSERVICE(B1)";

            // WEBSERVICEから返されたXMLを折り返し
            ws.Range["B2"].WrapText = true;
            ws.Range["B2"].RowHeight = 200;

            // 動的配列数式を有効にする（二つ目のFILTERXML関数でスピルするため）
            workbook.AllowDynamicArray = true;

            // FILTERXML関数
            ws.Range["A4"].Formula = "=FILTERXML(B2, \"//channel/title\")";

            // FILTERXML関数
            ws.Range["B4"].Formula = "=FILTERXML(B2, \"//channel/item/title\")";

            // Excelファイルに保存
            workbook.Save("result.xlsx");
        }
    }
}
