using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace TextRedactor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void шрифтToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt|Word Documents (*.docx)|*.docx|All Files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(saveFileDialog.FileName).ToLower() == ".txt")
                {
                    File.WriteAllText(saveFileDialog.FileName, richTextBox1.Text);
                }
                else if (Path.GetExtension(saveFileDialog.FileName).ToLower() == ".docx")
                {
                    // Сохранение в формате .docx
                    // Необходимо использовать библиотеку для работы с файлами Word, например, Microsoft.Office.Interop.Word
                    // Для этого потребуется установка библиотеки и наличие Microsoft Word на компьютере
                    // Пример:
                    // SaveAsDocx(saveFileDialog.FileName);
                }
                else
                {
                    MessageBox.Show("Unsupported file format", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text Files (*.txt)|*.txt|Word Documents (*.docx)|*.docx|All Files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileExtension = Path.GetExtension(openFileDialog.FileName).ToLower();

                if (fileExtension == ".txt" || fileExtension == ".docx")
                {
                    if (fileExtension == ".txt")
                    {
                        richTextBox1.Text = File.ReadAllText(openFileDialog.FileName);
                    }
                    else if (fileExtension == ".docx")
                    {
                        try
                        {
                            using (DocX document = DocX.Load(openFileDialog.FileName))
                            {
                                richTextBox1.Text = document.Text;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error opening Word document: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Unsupported file format", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                string content = richTextBox1.Text;

                try
                {
                    // Сохранение содержимого в файл
                    File.WriteAllText(filePath, content);

                    MessageBox.Show("Файл успешно сохранен", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при сохранении файла: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void копироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void вставитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();

        }

        private void вырезатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void нToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Создаем новый экземпляр диалогового окна выбора шрифта
            FontDialog fontDialog = new FontDialog();

            // Устанавливаем текущий шрифт и цвет текста в диалоговом окне на шрифт и цвет текущего выделенного текста
            fontDialog.Font = richTextBox1.SelectionFont;
            fontDialog.Color = richTextBox1.SelectionColor;

            // Открываем диалоговое окно выбора шрифта
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                // Применяем выбранные настройки шрифта к текущему выделенному тексту в RichTextBox
                richTextBox1.SelectionFont = fontDialog.Font;
                richTextBox1.SelectionColor = fontDialog.Color;
            }
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog colorDialog = new ColorDialog();

            // Устанавливаем текущий цвет фона текста в диалоговом окне на цвет текущего фона
            colorDialog.Color = richTextBox1.BackColor;

            // Открываем диалоговое окно выбора цвета
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                // Применяем выбранный цвет фона к тексту в RichTextBox
                richTextBox1.BackColor = colorDialog.Color;
            }
        }



        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Left; // Выравнивание текста по левому краю
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Center; // Выравнивание текста по центру
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Right; // Выравнивание текста по правому краю
        }


        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void поискToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Отображаем диалоговое окно для ввода текста для поиска
            string searchText = Microsoft.VisualBasic.Interaction.InputBox("Введите текст для поиска:", "Поиск");

            // Проверяем, что введенный текст не пустой
            if (!string.IsNullOrEmpty(searchText))
            {
                // Находим первое вхождение текста
                int index = richTextBox1.Find(searchText);

                // Если текст найден, выделяем его и перемещаем курсор к нему
                if (index >= 0)
                {
                    richTextBox1.SelectionStart = index;
                    richTextBox1.SelectionLength = searchText.Length;
                    richTextBox1.ScrollToCaret();
                }
                else
                {
                    // Если текст не найден, выводим сообщение
                    MessageBox.Show("Текст не найден", "Поиск", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void заменаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string searchText = Microsoft.VisualBasic.Interaction.InputBox("Введите текст для поиска:", "Поиск");
            string replaceText = Microsoft.VisualBasic.Interaction.InputBox("Введите текст для замены:", "Замена");

            // Проверяем, что введенные тексты не пустые
            if (!string.IsNullOrEmpty(searchText) && !string.IsNullOrEmpty(replaceText))
            {
                // Выполняем замену текста
                richTextBox1.Text = richTextBox1.Text.Replace(searchText, replaceText);
            }
        }
        private void ApplySyntaxHighlighting(string[] keywords, Color color)
        {
            // Сохраняем текущую позицию и стиль текста
            int selectionStart = richTextBox1.SelectionStart;
            int selectionLength = richTextBox1.SelectionLength;
            FontStyle selectionFontStyle = richTextBox1.SelectionFont.Style;

            // Применяем форматирование к каждому ключевому слову
            foreach (string keyword in keywords)
            {
                int index = 0;
                while (index < richTextBox1.Text.Length)
                {
                    int wordStart = richTextBox1.Find(keyword, index, RichTextBoxFinds.WholeWord);
                    if (wordStart != -1)
                    {
                        richTextBox1.Select(wordStart, keyword.Length);
                        richTextBox1.SelectionColor = color;
                        richTextBox1.SelectionFont = new Font(richTextBox1.Font, richTextBox1.SelectionFont.Style); // Устанавливаем текущий стиль
                        index = wordStart + keyword.Length;
                    }
                    else
                    {
                        break;
                    }
                }
            }

            // Возвращаем предыдущую позицию и стиль текста
            richTextBox1.Select(selectionStart, selectionLength);
            richTextBox1.SelectionFont = new Font(richTextBox1.Font, selectionFontStyle);
        }


        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            int selectionStart = richTextBox1.SelectionStart;
            int selectionLength = richTextBox1.SelectionLength;

            // Очищаем форматирование текста перед применением нового
            richTextBox1.SelectAll();
            richTextBox1.SelectionColor = Color.Black;

            // Применяем подсветку синтаксиса для каждого языка программирования
            ApplyCSharpSyntaxHighlighting();
            ApplyJavaSyntaxHighlighting();
            ApplyPythonSyntaxHighlighting();

            // Восстанавливаем предыдущее положение курсора и выделение
            richTextBox1.Select(selectionStart, selectionLength);
        }

        private void ApplyCSharpSyntaxHighlighting()
        {
            string[] keywords = { "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char", "checked", "class", "const", "continue", "decimal", "default", "delegate", "do", "double", "else", "enum", "event", "explicit", "extern", "false", "finally", "fixed", "float", "for", "foreach", "goto", "if", "implicit", "in", "int", "interface", "internal", "is", "lock", "long", "namespace", "new", "null", "object", "operator", "out", "override", "params", "private", "protected", "public", "readonly", "ref", "return", "sbyte", "sealed", "short", "sizeof", "stackalloc", "static", "string", "struct", "switch", "this", "throw", "true", "try", "typeof", "uint", "ulong", "unchecked", "unsafe", "ushort", "using", "var", "virtual", "void", "volatile", "while" };

            foreach (string keyword in keywords)
            {
                int index = 0;
                while (index < richTextBox1.Text.Length)
                {
                    index = richTextBox1.Find(keyword, index, RichTextBoxFinds.WholeWord);
                    if (index == -1)
                        break;

                    richTextBox1.Select(index, keyword.Length);
                    richTextBox1.SelectionColor = Color.Green;
                    index += keyword.Length;
                }
            }
        }

        private void ApplyJavaSyntaxHighlighting()
        {
            string[] keywords = { "protected", "short", "super", "throw", "while", "yield", "super" };

            foreach (string keyword in keywords)
            {
                int index = 0;
                while (index < richTextBox1.Text.Length)
                {
                    index = richTextBox1.Find(keyword, index, RichTextBoxFinds.WholeWord);
                    if (index == -1)
                        break;

                    richTextBox1.Select(index, keyword.Length);
                    richTextBox1.SelectionColor = Color.Red;
                    index += keyword.Length;
                }
            }
        }

        private void ApplyPythonSyntaxHighlighting()
        {
            string[] keywords = { "False", "None", "True", "and", "as", "assert", "break", "class", "continue", "def", "del", "elif", "else", "except", "finally", "for", "from", "global", "if", "import", "in", "is", "lambda", "nonlocal", "not", "or", "pass", "raise", "return", "try", "while", "with", "yield" };

            foreach (string keyword in keywords)
            {
                int index = 0;
                while (index < richTextBox1.Text.Length)
                {
                    index = richTextBox1.Find(keyword, index, RichTextBoxFinds.WholeWord);
                    if (index == -1)
                        break;

                    richTextBox1.Select(index, keyword.Length);
                    richTextBox1.SelectionColor = Color.Blue;
                    index += keyword.Length;
                }
            }
        }
    }
}
