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
using GigaChatAdapter;
using System.Threading.Tasks;

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

        public void SendPostRequest()
        {
            System.Threading.Tasks.Task.Run(async () => await RunGigaChat()).GetAwaiter().GetResult();
        }

        static async System.Threading.Tasks.Task RunGigaChat()
        {
            /*
            // Ваш код здесь
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.InputEncoding = Encoding.GetEncoding(1251);
            Console.OutputEncoding = Encoding.GetEncoding(1251);

            string authData = "MTk4NGVhNDMtNzdmZi00MjYwLTg1ODQtOTFlZWRkNzZkYjRlOmE2N2FhZDA1LTRjM2EtNDg2Ni04M2U0LWRiYjM3NWZiY2Y3Yw==";

            GigaChatAdapter.Authorization auth = new GigaChatAdapter.Authorization(authData, GigaChatAdapter.Auth.RateScope.GIGACHAT_API_PERS);
            var authResult = await auth.SendRequest();

            if (authResult.AuthorizationSuccess)
            {
                Completion completion = new Completion();
                Console.WriteLine("Напишите запрос к модели. В ином случае закройте окно, если дальнейшую работу с чатботом необходимо прекратить.");

                while (true)
                {
                    var prompt = Console.ReadLine();
                    await auth.UpdateToken(reserveTime: new TimeSpan(0, 1, 0));
                    CompletionSettings settings = new CompletionSettings("GigaChat:latest", 2, null, 4, 1);
                    var result = await completion.SendRequest(auth.LastResponse.GigaChatAuthorizationResponse?.AccessToken, prompt, true, settings);

                    if (result.RequestSuccessed)
                    {
                        foreach (var it in result.GigaChatCompletionResponse.Choices)
                        {
                            Console.WriteLine(it.Message.Content);
                        }
                    }
                    else
                    {
                        Console.WriteLine(result.ErrorTextIfFailed);
                    }
                }
            }
            else
            {
                Console.WriteLine(authResult.ErrorTextIfFailed);
            }
            */
        }

        public void GetToken(Office.IRibbonControl control)
        {
            MessageBox.Show("Введите токен");
        }

        public void GetRequest(Office.IRibbonControl control)
        {
            //SendPostRequest();
            //MessageBox.Show("Введите запрос");
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
