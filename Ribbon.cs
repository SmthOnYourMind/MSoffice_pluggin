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

namespace word_test
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        public static string recieved_req_message;

        private static string currentSelectedText;

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
                //currentRange.Text = "Дополненный текст (Было: " + currentSelectedText + ")";
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
                result = Task.Run(async () => await TestRequestAPI.Run("Переведи следующий текст на английский\n" + currentSelectedText)).GetAwaiter().GetResult();
            }

            currentRange.Text = result;
        }

        public void GetToken(Office.IRibbonControl control)
        {
            MessageBox.Show("Введите свой токен");

            string result;
            result = Task.Run(async () => await TestRequestAPI.Run("кто такой абоба")).GetAwaiter().GetResult();

            MessageBox.Show(result);
        }

        public void GetRequest(Office.IRibbonControl control)
        {
            MessageBox.Show("Введите запрос");
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
