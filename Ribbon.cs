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

namespace word_test
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        public static string recieved_req_message;
        private static string currentSelectedText;
        public static bool is_installed_cert = false;

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

            currentRange.Text = result;
        }

        public void InsertToken(Office.IRibbonControl control)
        {
            string input = Interaction.InputBox("Введите свой токен", "Ввод токена", "Длина токена должна быть 100 символов", 500, 700);
            // MTk4NGVhNDMtNzdmZi00MjYwLTg1ODQtOTFlZWRkNzZkYjRlOmE2N2FhZDA1LTRjM2EtNDg2Ni04M2U0LWRiYjM3NWZiY2Y3Yw==
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

            MessageBox.Show(result);
        }

        public void ShowInstruction(Office.IRibbonControl control)
        {
            MessageBox.Show("Инструкция по поиску токена GigaChat:\n1) ...\n2) ...\n3) ...");
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
            MessageBox.Show("AI плагин для Microsoft Word на основе Gigachat API\nВерсия: 1.0.0.0\n(C) 2024 Plague-in corp");
        }

        public void InsertDefaultToken(Office.IRibbonControl control)
        {
            var result = MessageBox.Show("Вставить токен по умолчанию?\n(Он доступен всем пользователям поэтому может закончится в любое время)",
                "", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                TestRequestAPI.ChangeToken("MTk4NGVhNDMtNzdmZi00MjYwLTg1ODQtOTFlZWRkNzZkYjRlOmE2N2FhZDA1LTRjM2EtNDg2Ni04M2U0LWRiYjM3NWZiY2Y3Yw==");
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
