using System;
using System.Diagnostics.SymbolStore;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace DotNetEPPlusPlay
{
    public partial class Form1 : Form
    {
        // 使用完畢記得釋放資源
        private ExcelPackage excelPackage;

        public Form1()
        {
            InitializeComponent();

            // 使用EPPlus要先設定版權
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var fileOpenDialog = new OpenFileDialog
            {
                Filter = "僅限Excel類型檔案 (*.xlsx)|*.xlsx"
            };

            if (fileOpenDialog.ShowDialog() == DialogResult.OK)
            {
                this.label1.Text = fileOpenDialog.FileName;

                excelPackage = new ExcelPackage(fileOpenDialog.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (excelPackage == null)
            {
                var yesNoDialog = MessageBox.Show("未發現任何已開啟的excel檔案，是否現在建立一個?", "按下是選擇路徑後創立檔案", MessageBoxButtons.YesNo);

                if (yesNoDialog == DialogResult.Yes)
                {
                    var saveFileDialog = new SaveFileDialog
                    {
                        Filter = "僅限Excel類型檔案 (*.xlsx)|*.xlsx"
                    };

                    // 儲存成功後透過EPPlus開啟
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        this.label1.Text = saveFileDialog.FileName; // 更換畫面上的路徑

                        excelPackage = new ExcelPackage(saveFileDialog.FileName);

                        // 首次建立必須至少包含一張sheet
                        string defaultName = "sheet-test";

                        if (!excelPackage.Workbook.Worksheets.Any(p => p.Name == defaultName))
                        {
                            var sheet = excelPackage.Workbook.Worksheets.Add(defaultName);
                            sheet.Cells["A1"].Value = "Hello World!";
                        }
                    }
                }
            }

            // 存檔+釋放資源
            excelPackage.Save();
            excelPackage.Dispose();
            excelPackage = null;
            this.label1.Text = "檔案已儲存成功並關閉";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(excelPackage == null) 
            {
                MessageBox.Show("請選擇檔案後再操作");
                return;
            }

            var sheetName = this.textBox1.Text;

            if (string.IsNullOrEmpty(sheetName))
            {
                MessageBox.Show("需輸入sheet名稱後才能操作後續");
                return;
            }

            // 轉小寫後再處理
            var sheet = this.excelPackage.Workbook.Worksheets.FirstOrDefault(p => p.Name.ToLower() == sheetName.ToLower());

            // 不存在就建立
            if (sheet == null)
            {
                sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                sheet.Cells["A1"].Value = "Hello World!";
            }

            var listData = sheet.Cells.Select(x => new
            {
                Address = x.Address,
                Value = x.Value
            }).ToList();

            this.listView1.Items.Clear();

            foreach (var item in listData)
            {
                this.listView1.Items.Add(new ListViewItem 
                {
                    Text = $"{item.Address} {item.Value}",
                    Tag = item
                });
            }
        }
    }
}
