using System;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SupportDesk.forms
{
  
    public partial class DataTicketForm : Form
    {
        private string requestKod;
        #region Настройка формы
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        #endregion
        #region Локальный переменные
        private SQLiteConnection DB;
        private string openedRequestId;
        private readonly DateTime thisDay = DateTime.Now;
        #endregion
        public DataTicketForm(string kod)
        {
            InitializeComponent();
            requestKod = kod;
        }

        private async void DataTicketForm_Load(object sender, EventArgs e)
        {
            DB = new SQLiteConnection(Database.connectionString);
            await DB.OpenAsync();
            label11.Text = $"Описание задачи № {DataRequests.Kod}";
            openedRequestId = DataRequests.Kod;
            labelMain.Text = $"Тема задачи: {DataRequests.Tema}\n\nДата создания: {DataRequests.TheDate}\n\nТип проблемы: {DataRequests.Type}\n\nПользователь: {DataRequests.User}\n\nНомер кабинета: {DataRequests.Kabinet}\n\nВремя ожидания специалиста: {DataRequests.Time}\n\nКраткое описание проблемы:";
            textBoxDiscription.Text = DataRequests.Discription;
            CheckUserGroup();
            CheckDocs();
            buttonAcceptRequest.Visible = DataUsers.GroupUser == "WORKER" &&
                DataRequests.Executor == DataUsers.Initials && DataRequests.Status == "Новая";
        }
        #region Управление формой
        private void ButtonExit_Click(object sender, EventArgs e)
        {
            Close();
            DataRequests.MarkRequestClosed(openedRequestId);
        }
        private void FormControl(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _ = ReleaseCapture();
                _ = SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Close();
        }
        #endregion
        #region Загрузка БД
        private void CheckDocs()
        {
            if (DataRequests.PathTo == null || DataRequests.PathTo == "")
            {
                return;
            }
            else
            {
                string extension = Path.GetExtension(DataRequests.PathTo);
                expUIButtonFile.Visible = true;
                expUIButtonFile.Text = DataRequests.NameFile;
                expUIButtonFile.Image = Image.FromFile(FileImage.PathFile(extension));
            }
        }
        private void CheckUserGroup()
        {
            if (DataUsers.GroupUser == "ADMIN")
            {
                LoadingExecutors();
                buttonExecutor.Enabled = true;
                buttonSaveData.Enabled = true;
            }
            else if (DataUsers.GroupUser == "WORKER")
            {
                _ = comboBoxExecutors.Items.Add(DataRequests.Executor);
                _ = comboBoxExecutor2.Items.Add(DataRequests.Executor2);
                textBoxDiscription.ReadOnly = true;
            }
            else if (DataUsers.GroupUser == "USER")
            {
                _ = comboBoxExecutors.Items.Add(DataRequests.Executor);
                _ = comboBoxExecutor2.Items.Add(DataRequests.Executor2);
                textBoxDiscription.ReadOnly = true;
            }
            comboBoxExecutors.SelectedItem = DataRequests.Executor;
            comboBoxExecutor2.SelectedItem = DataRequests.Executor2 == "" || DataRequests.Executor2 == null ? "Не назначен" : (object)DataRequests.Executor2;
        }

        private async void ButtonAcceptRequest_Click(object sender, EventArgs e)
        {
            if (DataUsers.GroupUser != "WORKER" || DataRequests.Executor != DataUsers.Initials)
            {
                return;
            }

            using (SQLiteCommand command = new SQLiteCommand(
                $"UPDATE {Table_requests.main} SET [{Table_requests.status}]=@status WHERE [id]=@id AND [{Table_requests.status}]=@oldStatus", DB))
            {
                command.Parameters.AddWithValue("status", "В работе");
                command.Parameters.AddWithValue("id", Convert.ToInt32(DataRequests.Kod));
                command.Parameters.AddWithValue("oldStatus", "Новая");
                await command.ExecuteNonQueryAsync();
            }

            DataRequests.Status = "В работе";
            buttonAcceptRequest.Visible = false;
            MessageBoxUI.Show("Задача принята!", "Принять задачу", "Accept");
        }
        private async void LoadingExecutors()
        {
            SQLiteDataReader sqlReader = null;
            SQLiteCommand command = new SQLiteCommand($"SELECT * FROM [{Table_executor.main}]", DB);
            try
            {
                sqlReader = (SQLiteDataReader)await command.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    string initials = Convert.ToString($"{sqlReader[$"{Table_executor.initials}"]}");
                    _ = comboBoxExecutors.Items.Add(initials);
                    _ = comboBoxExecutor2.Items.Add(initials);
                }
            }
            catch (Exception ex)
            {
                MessageBoxUI.Show($"{ex.Message}", $"{ex.Source}", "Error");
            }
            finally
            {
                sqlReader?.Close();
            }
        }
        #endregion
        #region Назначить исполнителя
        private async void ButtonExecutor_Click(object sender, EventArgs e)
        {
            string date = DataRequests.TheDate;
            string user = DataRequests.User;
            string executor = comboBoxExecutors.Text;
            string executor2 = comboBoxExecutor2.Text;
            string cabinet = DataRequests.Kabinet;
            string time = DataRequests.Time;
            string disc = DataRequests.Discription;
            if (executor == executor2)
            {
                MessageBoxUI.Show("Исполнитель и дублер должны быть разные люди!", "Назначение исполнителей", "Error");
            }
            else
            {
                SQLiteCommand commandUpdate = new SQLiteCommand($"UPDATE {Table_requests.main} SET [{Table_requests.executor}]=@executor, [{Table_requests.executor2}]=@executor2 WHERE [{Table_requests.theDate}]=@date AND [{Table_requests.user}]=@user AND [{Table_requests.kabinet}]=@cabinet", DB);
                _ = commandUpdate.Parameters.AddWithValue("executor", executor);
                _ = commandUpdate.Parameters.AddWithValue("executor2", executor2);
                _ = commandUpdate.Parameters.AddWithValue("date", date);
                _ = commandUpdate.Parameters.AddWithValue("user", user);
                _ = commandUpdate.Parameters.AddWithValue("cabinet", cabinet);
                _ = await commandUpdate.ExecuteNonQueryAsync();
                if (executor2 == "Не назначен")
                {
                    TelegramForm.AddExecutor(DataUsers.Initials, user, cabinet, time, executor, disc);
                }
                else
                {
                    TelegramForm.AddExecutor(DataUsers.Initials, user, cabinet, time, executor, disc);
                    TelegramForm.AddExecutor(DataUsers.Initials, user, cabinet, time, executor2, disc);
                }
                File.AppendAllText(Logs.file, $"[Исполнитель]{DataUsers.Initials} закрепил заявку за {executor}, дублер {executor2} ({thisDay:dd.MM.yyyy HH:mm})");
                File.AppendAllText(Logs.file, Environment.NewLine);
                Close();
            }
        }
        #endregion
        #region Сохранить описание
        private async void ButtonSaveData_Click(object sender, EventArgs e)
        {
            string date = DataRequests.TheDate;
            string user = DataRequests.User;
            string cabinet = DataRequests.Kabinet;
            string time = DataRequests.Time;
            string disc = textBoxDiscription.Text;
            SQLiteCommand commandUpdate = new SQLiteCommand($"UPDATE {Table_requests.main} SET [{Table_requests.discription}]=@disc WHERE [{Table_requests.theDate}]=@date AND [{Table_requests.user}]=@user AND [{Table_requests.kabinet}]=@cabinet", DB);
            _ = commandUpdate.Parameters.AddWithValue("date", date);
            _ = commandUpdate.Parameters.AddWithValue("user", user);
            _ = commandUpdate.Parameters.AddWithValue("cabinet", cabinet);
            _ = commandUpdate.Parameters.AddWithValue("disc", disc);
            _ = await commandUpdate.ExecuteNonQueryAsync();
            File.AppendAllText(Logs.file, $"[Редактирование заявки]{DataUsers.Initials} отредактировал(а) описание заявки {date} {user} {cabinet} {DataRequests.Type}({thisDay:dd.MM.yyyy HH:mm})");
            File.AppendAllText(Logs.file, Environment.NewLine);
            TelegramForm.EditRequest(DataUsers.Initials, DataRequests.Tema, DataRequests.Type, user, cabinet, time, disc);
            Close();
        }
        #endregion
        #region Открыть файл
        private void ExpUIButtonFile_Click(object sender, EventArgs e)
        {
            _ = Process.Start(DataRequests.PathTo);
        }

        #endregion

        private void btnSeveNotes_Click(object sender, EventArgs e)
        {

        }
    }
}
