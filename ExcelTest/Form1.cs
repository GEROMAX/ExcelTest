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
            for (int i = 1; i <= 10; i++)
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
                
                worksheet.Cell(1, 1).Value = lst.Select<SelectDataEntity, string>(record => record.ID.ToString());
                worksheet.Cell(1, 2).Value = lst.Select<SelectDataEntity, string>(record => record.Name);
                worksheet.Cell(1, 3).Value = lst.Select<SelectDataEntity, string>(record => record.Address);
                worksheet.Cell(1, 4).Value = lst.Select<SelectDataEntity, string>(record => record.Price.ToString());
                worksheet.Cell(1, 5).Value = lst.Select<SelectDataEntity, string>(record => record.RegistDate.ToString("yyyy/MM/dd"));

                //列幅調整
                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(this.saveFileDialog1.FileName);
            }
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
