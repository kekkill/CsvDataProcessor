using CsvDataProcessor.Data;
using CsvDataProcessor.Services.Database;
using CsvDataProcessor.Services.Export;
using CsvDataProcessor.Services.Import;

namespace CsvDataProcessor.UI
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
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;
            btnImport.Enabled = false;
            lblStatus.Text = "Начало импорта...";

            try
            {
                Progress<string> progress = new Progress<string>(status =>
                {
                    lblStatus.Text = status;
                });

                using (var context = new AppDbContext())
                {
                    var result = await CsvImporter.ImportCsvToDatabaseAsync(filePath, context, progress);
                    MessageBox.Show($"Импорт завершён!\nУспешно: {result.successCount:N0}\nОшибок: {result.errorCount:N0}", "Результат", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            lblStatus.Text = "Подготовка данных...";

            try
            {
                if (!ValidateFilters())
                {
                    return;
                }

                var parameters = new ExportParameters
                {
                    UseStartDate = chkStartDate.Checked,
                    StartDate = dtpStartDate.Value,
                    UseEndDate = chkEndDate.Checked,
                    EndDate = dtpEndDate.Value,
                    FirstName = txtFirstName.Text,
                    LastName = txtLastName.Text,
                    SurName = txtSurName.Text,
                    City = txtCity.Text,
                    Country = txtCountry.Text
                };

                var exportService = new ExportService();
                await exportService.ExportToExcelAsync(parameters, filePath);
                MessageBox.Show($"Данные экспортированы в:\n{filePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                lblStatus.Text = "";
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
            saveFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
            saveFileDialog.Title = "Сохранить как XML";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FileName = $"export_{DateTime.Now:yyyyMMddHHmmss}.xml";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string filePath = saveFileDialog.FileName;

            progressBar.Visible = true;
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;
            btnExportXml.Enabled = false;
            lblStatus.Text = "Подготовка данных...";

            try
            {
                if (!ValidateFilters())
                {
                    return;
                }

                var parameters = new ExportParameters
                {
                    UseStartDate = chkStartDate.Checked,
                    StartDate = dtpStartDate.Value,
                    UseEndDate = chkEndDate.Checked,
                    EndDate = dtpEndDate.Value,
                    FirstName = txtFirstName.Text,
                    LastName = txtLastName.Text,
                    SurName = txtSurName.Text,
                    City = txtCity.Text,
                    Country = txtCountry.Text
                };

                var exportService = new ExportService();
                await exportService.ExportToXmlAsync(parameters, filePath);
                MessageBox.Show($"Данные экспортированы в:\n{filePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private bool ValidateFilters()
        {
            if (chkStartDate.Checked && chkEndDate.Checked &&
                dtpStartDate.Value > dtpEndDate.Value)
            {
                MessageBox.Show("Начальная дата не может быть больше конечной!\n" +
                                "Пожалуйста, исправьте даты.",
                    "Ошибка валидации",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                dtpEndDate.Value = dtpStartDate.Value.AddDays(1);
                chkEndDate.Checked = true;

                lblStatus.Text = "Исправлено: конечная дата установлена позже начальной";
                lblStatus.ForeColor = Color.Red;
                lblStatus.Visible = true;

                return false;
            }
            return true;
        }
    }
}