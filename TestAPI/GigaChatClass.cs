using GigaChatAdapter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Authorization = GigaChatAdapter.Authorization;

namespace TestAPI
{
    public class testRequ
    {

        public static async Task Run(string text)
        {
            Console.WriteLine("анлаки");
            //Настройка для работы консоли с кириллицей
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.InputEncoding = Encoding.GetEncoding(1251);
            Console.OutputEncoding = Encoding.GetEncoding(1251);

            //аутентификационные данные из личного кабинета
            string authData = "MTk4NGVhNDMtNzdmZi00MjYwLTg1ODQtOTFlZWRkNzZkYjRlOmE2N2FhZDA1LTRjM2EtNDg2Ni04M2U0LWRiYjM3NWZiY2Y3Yw==";

            //Запуск авторизации в гигачате
            Authorization auth = new Authorization(authData, GigaChatAdapter.Auth.RateScope.GIGACHAT_API_PERS);
            var authResult = await auth.SendRequest();

            if (authResult.AuthorizationSuccess)
            {
                Console.WriteLine("анлаки");
                Completion completion = new Completion();
                Console.WriteLine("запрос"); //RU

                while (true)
                {
                    //Чтение промпта с консоли
                    var prompt = text;

                    //Обновление токена, если он просрочился (запас времени - 1 минута до просрочки)
                    await auth.UpdateToken(reserveTime: new TimeSpan(0, 1, 0));

                    //++установка доп.настроек
                    CompletionSettings settings = new CompletionSettings("GigaChat:latest", 2, null, 1, 1);

                    //request / отправка промпта
                    var result = await completion.SendRequest(auth.LastResponse.GigaChatAuthorizationResponse?.AccessToken, prompt, true);
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
        }
    }
}