using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using TestNetFrameworkAPI;
using System.Threading.Tasks;
using System.Net.Http;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualBasic;
using System.Diagnostics;
using System.Drawing;

namespace word_test
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        public static string recieved_req_message;
        private static string currentSelectedText;
        public static bool is_installed_cert = false;

        private static string default_token = "MTk4NGVhNDMtNzdmZi00MjYwLTg1ODQtOTFlZWRkNzZkYjRlOmU5NjhhZWQ4LTFhNTItNDYwOS1iM2Q2LWZhNzhlMzQ5NjM4OA==";
        public static string GetCurrentSelected()
        {
            return currentSelectedText;
        }

        private Office.IRibbonUI ribbon;

        public Ribbon()
        {

        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("word_test.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public void GetButtonID(Office.IRibbonControl control)
        {
            if (TestRequestAPI.GetToken() == "")
            {
                MessageBox.Show("Сначала введите токен");
                return;
            }
            if (TestRequestAPI.GetToken().Length != 100)
            {
                MessageBox.Show("Формат токена неверный");
                return;
            }

            Microsoft.Office.Interop.Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;

            currentSelectedText = currentRange.Text;
            string result = "";
            if (control.Id == "context_button1")
            {
                result = Task.Run(async () => await TestRequestAPI.Run("Перепеши следующий текст по другому\n" + currentSelectedText)).GetAwaiter().GetResult();
            }
            else if (control.Id == "context_button2")
            {
                result = Task.Run(async () => await TestRequestAPI.Run("Дополни следущий текст\n" + currentSelectedText)).GetAwaiter().GetResult();
            }
            else if (control.Id == "context_button3")
            {
                result = Task.Run(async () => await TestRequestAPI.Run("Дай краткое содержание следующего текста\n" + currentSelectedText)).GetAwaiter().GetResult();
            }
            else if (control.Id == "context_button4")
            {
                result = Task.Run(async () => await TestRequestAPI.Run("Переведи следующий текст на английский\n" + currentSelectedText)).GetAwaiter().GetResult();
            }
            else if (control.Id == "context_button5")
            {
                result = Task.Run(async () => await TestRequestAPI.Run("Переведи следующий текст на русский\n" + currentSelectedText)).GetAwaiter().GetResult();
            }
            else if (control.Id == "submenu_button_1")
            {
                result = Task.Run(async () => await TestRequestAPI.Run("Перепиши следующий текст в деловом стиле\n" + currentSelectedText)).GetAwaiter().GetResult();
            }
            else if (control.Id == "submenu_button_2")
            {
                result = Task.Run(async () => await TestRequestAPI.Run("Перепиши следующий текст в научном стиле\n" + currentSelectedText)).GetAwaiter().GetResult();
            }
            else if (control.Id == "submenu_button_3")
            {
                result = Task.Run(async () => await TestRequestAPI.Run("Перепиши следующий текст в разговорном стиле\n" + currentSelectedText)).GetAwaiter().GetResult();
            }

            if (result != "error")
                currentRange.Text = result;
            else
            {
                MessageBox.Show("Неверный токен, попробуйте другой");
            }
        }

        public void InsertToken(Office.IRibbonControl control)
        {
            string input = Interaction.InputBox("Введите свой токен", "Ввод токена", "Длина токена должна быть 100 символов", 500, 700);
            
            if (input.Length == 100)
                TestRequestAPI.ChangeToken(input);
            else
                MessageBox.Show("Формат токена неверный");
        }

        public void SendRequest(Office.IRibbonControl control)
        {
            if (TestRequestAPI.GetToken() == "")
            {
                MessageBox.Show("Сначала введите токен");
                return;
            }
            if (TestRequestAPI.GetToken().Length != 100)
            {
                MessageBox.Show("Формат токена неверный");
                return;
            }

            string input = Interaction.InputBox("Введите свой запрос", "Ввод запроса", "Например: Какие спутники есть у Сатурна", 500, 700);
            string result = Task.Run(async () => await TestRequestAPI.Run(input)).GetAwaiter().GetResult();

            if (result != "error")
                MessageBox.Show(result);
            else
                MessageBox.Show("Неверный токен, попробуйте другой");

        }

        public class CustomMessage : Form
        {
            private Label messageLabel;
            private Button buttonToSite;
            public CustomMessage(string message)
            {
                
                this.Text = "Пользовательская инструкция по поиску авторизационного ключа";
                this.Size = new System.Drawing.Size(670, 200);
                this.StartPosition = FormStartPosition.CenterScreen;

                
                messageLabel = new Label();
                messageLabel.Text = message;
                messageLabel.Location = new System.Drawing.Point(10, 10);
                messageLabel.Size = new System.Drawing.Size(700, 100);
                this.Controls.Add(messageLabel);

                buttonToSite = new Button();
                buttonToSite.Text = "Получить ключ";
                buttonToSite.Location = new System.Drawing.Point(270, 120);
                buttonToSite.Size = new Size(75, 35);
                buttonToSite.Click += new EventHandler(buttonToSite_Click);
                this.Controls.Add(buttonToSite);
            }

            private void buttonToSite_Click(object sender, EventArgs e)
            {
                Process.Start("https://developers.sber.ru/studio/workspaces/");
            }
            public static void Show(string message)
            {
                CustomMessage inst = new CustomMessage(message);
                inst.ShowDialog();
            }
        }

        public void ShowInstruction(Office.IRibbonControl control)
        {

            CustomMessage.Show("Пользовательская инструкция по поиску ключа авторизации: \n" +
            "1.\tНужно пройти регистрацию через телефон или SberID по ссылке - https://developers.sber.ru/studio/registration\n" +
            "2.\tВ меню, в левой части экрана необходимо нажать кнопку «Создать проект» и заполнить все поля для создания.\n" +
            "3.\tТеперь нужно зайти в созданный проект.\n" +
            "4.\tВ открытом проекте смотрим на правую колонку, где присутствует надпись \n\t«Используйте ключи для подключения сервиса», здесь нам нужно нажать «Сгенерировать»\n" +
            "5.\tВ открытом окне «Новый Client Secret» нас интересует поле «Авторизационные данные».\r\n\tДанные оттуда нужно сохранить к себе. Этот набор символов и есть Ваш ключ для использования нашего плагина.");
        }

        public void ShowSertificates(Office.IRibbonControl control)
        {
            var result = MessageBox.Show("Перейти на сайт https://www.gosuslugi.ru/crt чтобы скачать сертификаты минцифры?",
                "", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                Process.Start("https://www.gosuslugi.ru/crt");
            }
        }

        public void ShowHelp(Office.IRibbonControl control)
        {
            MessageBox.Show("Помощь: ");
        }

        public void ShowAbout(Office.IRibbonControl control)
        {
            MessageBox.Show("AI плагин для Microsoft Word на основе Gigachat API\nВерсия: 1.0.0.2\n(C) 2024 Plague-in corp");
        }

        public void InsertDefaultToken(Office.IRibbonControl control)
        {
            var result = MessageBox.Show("Вставить токен по умолчанию?\n(Он доступен всем пользователям поэтому может закончится в любое время)",
                "", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                TestRequestAPI.ChangeToken(default_token);
            }
        }

        public System.Drawing.Image GetImage(string imageName)
        {
            return (System.Drawing.Image)Properties.Resources.ResourceManager.GetObject(imageName);
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
