using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCreateTestData_Click(object sender, EventArgs e)
        {
            var lst = new List<SelectDataEntity>();
            for (int i = 1; i <= 25; i++)
            {
                lst.Add(new SelectDataEntity(i));
            }
            this.dataGridView1.DataSource = lst;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK.Equals(this.saveFileDialog1.ShowDialog()))
            {
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                var lst = (List<SelectDataEntity>)this.dataGridView1.DataSource;

                for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cell(1, i).Value = dataGridView1.Columns[i - 1].Name;
                }
                worksheet.Cell(2, 1).Value = lst.Select<SelectDataEntity, string>(record => record.ID.ToString());
                worksheet.Cell(2, 2).Value = lst.Select<SelectDataEntity, string>(record => record.Name);
                worksheet.Cell(2, 3).Value = lst.Select<SelectDataEntity, string>(record => record.Address);
                worksheet.Cell(2, 4).Value = lst.Select<SelectDataEntity, string>(record => record.Price.ToString());
                worksheet.Cell(2, 5).Value = lst.Select<SelectDataEntity, string>(record => record.RegistDate.ToString("yyyy/MM/dd"));

                //全行上下中央
                //ヘッダ左右中央
                worksheet.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                worksheet.Row(1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                //列幅調整
                worksheet.Columns().AdjustToContents();
                for (int i = 1; i <= 5; i++)
                {
                    //カラムインデックスは1オリジン
                    //Widthの単位はpt
                    if (worksheet.Column(i).Width > 5)
                    {
                        worksheet.Column(i).Width = 5;
                    }
                }
                //折り返し設定
                worksheet.Style.Alignment.WrapText = true;

                //印刷方向：横
                //Portrait＝肖像画＝縦の意味
                //Landscape＝風景画＝横の意味
                worksheet.PageSetup.SetPageOrientation(XLPageOrientation.Landscape);

                //印刷用紙：A4
                worksheet.PageSetup.SetPaperSize(XLPaperSize.A4Paper);

                //印刷設定：左余白を1cm
                //印刷設定：右余白を0.6cm
                //印刷設定：上余白を1.6cm
                //印刷設定：下余白を1.6cm
                //印刷設定：ヘッダ余白を0.8cm
                //印刷設定：フッタ余白を0.8cm
                //印刷設定の単位はインチ指定
                worksheet.PageSetup.Margins.SetLeft(UnitConvertCentimeter.ToInch(1));
                worksheet.PageSetup.Margins.SetRight(UnitConvertCentimeter.ToInch(0.6));
                worksheet.PageSetup.Margins.SetTop(UnitConvertCentimeter.ToInch(1.6));
                worksheet.PageSetup.Margins.SetBottom(UnitConvertCentimeter.ToInch(1.6));
                worksheet.PageSetup.Margins.SetHeader(UnitConvertCentimeter.ToInch(0.8));
                worksheet.PageSetup.Margins.SetFooter(UnitConvertCentimeter.ToInch(0.8));

                //印刷範囲：全ての列を1シートに印刷
                //指定したくない方向は0で良い様子
                worksheet.PageSetup.FitToPages(1, 0);

                //印刷設定：枠線あり
                //罫線を引かなくても枠線付きで印刷できる
                worksheet.PageSetup.SetShowGridlines(true);

                //ヘッダ・フッタ設定
                //予約語はEnumで指定可能
                worksheet.PageSetup.Header.Left.AddText(XLHFPredefinedText.File);
                worksheet.PageSetup.Header.Left.AddNewLine().AddText(DateTime.Now.ToString("yyyy/MM/dd"));
                worksheet.PageSetup.Header.Left.AddNewLine().AddText(DateTime.Now.AddYears(12).ToString("平成yy年"));
                worksheet.PageSetup.Header.Right.AddText(XLHFPredefinedText.Date);
                worksheet.PageSetup.Footer.Center.AddText(XLHFPredefinedText.PageNumber);
                worksheet.PageSetup.Footer.Center.AddText("/");
                worksheet.PageSetup.Footer.Center.AddText(XLHFPredefinedText.NumberOfPages);

                //印刷時ヘッダ繰り返し行
                worksheet.PageSetup.SetRowsToRepeatAtTop(1, 1);

                workbook.SaveAs(this.saveFileDialog1.FileName);
            }
        }
    }

    public class UnitConvertCentimeter
    {
        public static double ToInch(double value)
        {
            return value / 2.54;
        }
    }

    public class SelectDataEntity
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public decimal Price { get; set; }
        public DateTime RegistDate { get; set; }

        public SelectDataEntity(int registNumber)
        {
            this.ID = registNumber;
            this.Name = "名前" + registNumber.ToString();
            this.Address = "住所" + registNumber.ToString();
            this.Price = registNumber * 100;
            this.RegistDate = DateTime.Now.AddDays(-registNumber);
        }
    }
}
