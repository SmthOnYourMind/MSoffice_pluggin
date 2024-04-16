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
using GigaChatAdapterNetFramework;
using TestNetFrameworkAPI;
using System.Threading.Tasks;
using System.Net.Http;
using Task = System.Threading.Tasks.Task;

namespace word_test
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
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

            if (control.Id == "context_button1")
            {
                currentRange.Text = "Переписанный текст (Было: " + currentSelectedText + ")";
            }
            else if (control.Id == "context_button2")
            {
                currentRange.Text = "Дополненный текст (Было: " + currentSelectedText + ")";
            }
        }

        public async Task CallNet6ApiAsync()
        {
            using (var client = new HttpClient())
            {
                // Replace with your .NET 6 Web API URL
                var url = "http://localhost:5000/api/your-endpoint";
                var response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadAsStringAsync();
                }
            }
        }


        public void GetToken(Office.IRibbonControl control)
        {
            Task.Run(async () => await TestNetFrameworkAPI.TestNetFrameworkAPI.Main()).GetAwaiter().GetResult();
            MessageBox.Show("Введите токен");
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
