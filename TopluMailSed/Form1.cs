using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using TopluMailSed.Models;

namespace TopluMailSed
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Dosyaları|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    LoadExcelData(filePath);

                }
            }
        }

        private void LoadExcelData(string filePath)
        {
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // İlk çalışma sayfasını seçin

                // DataTable oluşturun ve sütun başlıklarını ekleyin
                DataTable dataTable = new DataTable();
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    dataTable.Columns.Add(worksheet.Cells[1, col].Text);
                }

                // Excel verilerini DataTable'a aktarın
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTable.Rows.Add(dataRow);
                }

                // DataGridView'e verileri bağlayın

                dataGridView1.DataSource = dataTable;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool enAzBirTextBoxBos = false;

            foreach (Control control in groupBox2.Controls)
            {
                if (control is System.Windows.Forms.TextBox textBox)
                {
                    if (string.IsNullOrWhiteSpace(textBox.Text))
                    {
                        enAzBirTextBoxBos = true;
                        break; // En az bir boş TextBox bulundu, döngüyü sonlandır.
                    }
                }
            }

            if (enAzBirTextBoxBos)
            {
                MessageBox.Show("Boş kutular var doldurmadan işlem yapamazsınız.");
            }
            else
            {
                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("DataGridView boş. Veri ekleyin.");
                }
                else
                {
                    SendMail sendMail = new SendMail();
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow) // Yeni bir satırı hariç tut
                        {
                            // İlgili hücre değerlerini alın
                            string email = row.Cells[0].Value.ToString(); // Örnek olarak sütun 1'den (0 indeksli) e-posta adresini alıyoruz.

                            sendMail.Microsoft(
                                Convert.ToInt32(SmtpPt.Text),
                                SmtpHost.Text,
                                GondericiAdSoyad.Text,
                                GondericiMail.Text,
                                GondericiPass.Text,
                                email, // İlgili hücreden aldığımız e-posta adresini kullanıyoruz
                                Baslik.Text,
                                txtDetay.Text
                            );
                        }
                    }
                    MessageBox.Show("Tüm mailler atıldı.");
                }
            }
        }

        private void excelSendMailAdd_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Dosyaları|*.xlsx";
            openFileDialog.Title = "Excel Dosyası Seçiniz";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelDosyaYolu = openFileDialog.FileName;

                using (var package = new ExcelPackage(new FileInfo(excelDosyaYolu)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // İlk çalışma sayfasını alın.

                    // Excel'deki verileri TextBox'lara ekleyin.
                    SmtpPt.Text = worksheet.Cells[1, 1].Text; // İlk hücre
                    SmtpHost.Text = worksheet.Cells[1, 2].Text; // İkinci hücre
                    comboBox1.Text = worksheet.Cells[1, 3].Text; // Üçüncü hücre
                    GondericiMail.Text = worksheet.Cells[1, 4].Text; // Üçüncü hücre
                    GondericiPass.Text = worksheet.Cells[1, 5].Text; // Üçüncü hücre
                    GondericiAdSoyad.Text = worksheet.Cells[1, 6].Text; // Üçüncü hücre

                    // Diğer TextBox'lara aynı şekilde devam edebilirsiniz.

                    MessageBox.Show("Excel verileri TextBox'lara aktarıldı.");
                }
            }


        }
    }
}