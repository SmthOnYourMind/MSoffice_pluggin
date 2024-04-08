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
using System.Threading.Tasks;
using System.Net.Http;
using RestSharp;
using System.Net.Http.Headers;

namespace word_test
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
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

            string current_selected = currentRange.Text;

            if (control.Id == "context_button1")
            {
                currentRange.Text = "Переписанный текст (Было: " + current_selected + ")";
            }
            else if (control.Id == "context_button2")
            {
                currentRange.Text = "Дополненный текст (Было: " + current_selected + ")";
            }
        }

        public async void SendPostRequest()
        {
            //3f05354f-1346-44c2-8f6c-1d34a03fc054      client id
            //8f791d9d-10d6-4a86-b90c-7efed0d41240      client secret
            //edbd3eac-c185-4112-880b-f8b1351c7f10      uuid
            //M2YwNTM1NGYtMTM0Ni00NGMyLThmNmMtMWQzNGEwM2ZjMDU0OjhmNzkxZDlkLTEwZDYtNGE4Ni1iOTBjLTdlZmVkMGQ0MTI0MA==   Авторизационные данные
            var client = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Post, "https://ngw.devices.sberbank.ru:9443/api/v2/oauth");
            request.Headers.Add("Accept", "application/json");
            request.Headers.Add("Authorization", "Basic 6a98f861-99b2-4931-9e60-9ed6fdc56f95");
            var content = new StringContent(string.Empty);
            content.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");
            request.Content = content;
            var response = await client.SendAsync(request);
            response.EnsureSuccessStatusCode();
            Console.WriteLine(await response.Content.ReadAsStringAsync());
        }

        public void GetToken(Office.IRibbonControl control)
        {
            MessageBox.Show("Введите токен");
        }

        public void GetRequest(Office.IRibbonControl control)
        {
            //SendPostRequest();

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
