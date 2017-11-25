using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using СurriculumParse;
using СurriculumParse.ExcelParsers;

namespace WindowsInterface
{
    public partial class Form1 : Form
    {
        private readonly Api _api;

        public Form1()
        {
            InitializeComponent();

            openFileDialog1.Filter = "Text files(*.xlsx)|*.xlsx";
            openFileDialog2.Filter = "Text files(*.xlsx)|*.xlsx";
            openFileDialog1.FileName = "";
            openFileDialog2.FileName = "";
            Closing += this.Form_Close;
            this.Name = "Программа";

            try
            {
                _api = new Api();
            }
            catch (Exception)
            {
                MessageBox.Show(@"Ошибка старта приложения",
                    @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void Form_Close(object sender, EventArgs e)
        {
            _api.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            var filename = openFileDialog1.FileName;
            // читаем файл в строку
            switch (_api.ParsePps(filename))
            {
                case PpsReadStatus.Success:
                    MessageBox.Show("Выполнено успешно", "ППС", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                case PpsReadStatus.CurriculumNotFound:
                    MessageBox.Show(@"Ошибка. Не найден соответствующий учебный план. Возможно имеются различия профиле уп, или же уп не был загружен в базу.",
                        @"ППС", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case PpsReadStatus.PpsReadError:
                    MessageBox.Show(@"Ошибка. Не удалось прочитать файл",
                        @"ППС", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case PpsReadStatus.FileOpenException:
                    MessageBox.Show(@"Ошибка. Закройте файл перед началом работы.",
                        @"ППС", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                default:
                    MessageBox.Show("Неведомая ошибка", "УП", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            var filename = openFileDialog1.FileName;
            // читаем файл в строку
            if (_api.ParseCurriculum(filename))
            {
                MessageBox.Show("Файл успешно загружен", "УП", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Произошла ошибка при чтении файла", "УП", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            var dir = folderBrowserDialog1.SelectedPath;
            // читаем файл в строку
            _api.ParseCurriculumsDirrectory(dir);
            MessageBox.Show(
                "Запись в базу завершена. В файле Ошибки.txt можете ознаокмиться с тем, какие файлы не удалось прочитать.",
                "УП", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

    }
}
