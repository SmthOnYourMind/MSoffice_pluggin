using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net.Http;
using Task = System.Threading.Tasks.Task;
using TestNetFrameworkAPI;

namespace word_test
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string result;
            TestRequestAPI.ChangeToken(Ribbon.GetDefaultToken());
            result = Task.Run(async () => await TestRequestAPI.Run("Привет")).GetAwaiter().GetResult();

            if (result == "error")
            {
                MessageBox.Show("Вас приветствует Расширение Gigachat для Word. " +
                "Для дальнейшей работы с плагином вам потребуются сертификаты минцифры, " +
                "скачать и установить которые можно по следующей ссылке " +
                "https://www.gosuslugi.ru/crt . " +
                "Если у вас уже они есть, можете продолжить работу.");
                var result_box = MessageBox.Show("Вы хотите перейти на сайт https://www.gosuslugi.ru/crt ?", "", MessageBoxButtons.YesNo);
                if (result_box == DialogResult.Yes)
                {
                    Process.Start("https://www.gosuslugi.ru/crt");
                }
            }

            TestRequestAPI.ChangeToken(null);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //MessageBox.Show("Всего хорошего");
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
