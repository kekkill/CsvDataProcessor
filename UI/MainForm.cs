using CsvDataProcess.Services.Import;
using CsvDataProcess.Data;
using CsvDataProcess.Services.Database;
using CsvDataProcess.Services.Export;

namespace WinFormsApp1
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            Load += Form1_Load;
        }

        private void UpdateRecordCount()
        {
            using (var context = new AppDbContext())
            {
                var service = new PersonService(context);
                int count = service.GetTotalCount();
                lblTotalRecords.Text = $"Всего записей в базе: {count}";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            using (var context = new AppDbContext())
            {
                context.InitializeDatabase();
                UpdateRecordCount();
            }
        }

        private async void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV Files|*.csv|All Files|*.*";
            openFileDialog.Title = "Выберите CSV-файл";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string filePath = openFileDialog.FileName;

            if (!File.Exists(filePath))
            {
                MessageBox.Show("Файл не найден!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            progressBar.Visible = true;
            progressBar.Value = 0;
            btnImport.Enabled = false;

            try
            {
                Progress<int> progress = new Progress<int>(percent =>
                {
                    progressBar.Value = percent;
                    lblStatus.Text = $"Прогресс: {percent}%";
                });
                using (var context = new AppDbContext())
                {
                    var result = await CsvImporter.ImportCsvToDatabaseAsync(filePath, context, progress);
                    MessageBox.Show($"Импорт завершён!\nУспешно: {result.successCount}\nОшибок: {result.errorCount}", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateRecordCount();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка импорта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                btnImport.Enabled = true;
            }
        }

        private async void btnExportExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            saveFileDialog.Title = "Сохранить как Excel";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FileName = $"export_{DateTime.Now:yyyyMMddHHmmss}.xlsx";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string filePath = saveFileDialog.FileName;

            progressBar.Visible = true;
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;
            btnExportExcel.Enabled = false;

            try
            {
                bool useStartDate = chkStartDate.Checked;
                bool useEndDate = chkEndDate.Checked;
                DateTime startDate = dtpStartDate.Value;
                DateTime endDate = dtpEndDate.Value;

                using (var context = new AppDbContext())
                {
                    var personService = new PersonService(context);
                    var exportService = new ExportService();

                    var query = personService.GetFilteredPeople(useStartDate, startDate, useEndDate, endDate, txtFirstName.Text, txtLastName.Text, txtSurName.Text, txtCity.Text, txtCountry.Text);

                    await exportService.ExportToExcelAsync(query, filePath);
                    MessageBox.Show($"Данные экспортированы в:\n{filePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                btnExportExcel.Enabled = true;
            }
        }

        private async void btnExportXml_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML Files|*.xml|All Files|*.*";
            saveFileDialog.Title = "Сохранить как XML";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FileName = "export.xml";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string filePath = saveFileDialog.FileName;

            progressBar.Visible = true;
            progressBar.Value = 0;
            btnExportXml.Enabled = false;

            try
            {
                bool useStartDate = chkStartDate.Checked;
                bool useEndDate = chkEndDate.Checked;
                DateTime startDate = dtpStartDate.Value;
                DateTime endDate = dtpEndDate.Value;

                using (var context = new AppDbContext())
                {
                    var personService = new PersonService(context);
                    var exportService = new ExportService();

                    var query = personService.GetFilteredPeople(useStartDate, startDate, useEndDate, endDate, txtFirstName.Text, txtLastName.Text, txtSurName.Text, txtCity.Text, txtCountry.Text);

                    await exportService.ExportToXmlAsync(query, filePath);
                    MessageBox.Show($"Данные экспортированы в:\n{filePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Visible = false;
                btnExportXml.Enabled = true;
            }
        }

        private void btnClearData_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить ВСЕ данные?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
            {
                return;
            }

            try
            {
                using (var context = new AppDbContext())
                {
                    var service = new PersonService(context);
                    service.ClearAllData();
                    UpdateRecordCount();
                    MessageBox.Show("Все данные удалены!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка очистки: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkStartDate_CheckedChanged(object sender, EventArgs e)
        {
            dtpStartDate.Enabled = chkStartDate.Checked;
        }

        private void chkEndDate_CheckedChanged(object sender, EventArgs e)
        {
            dtpEndDate.Enabled = chkEndDate.Checked;
        }
    }
}