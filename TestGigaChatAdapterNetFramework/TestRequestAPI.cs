using GigaChatAdapterNetFramework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Authorization = GigaChatAdapterNetFramework.Authorization;

namespace TestNetFrameworkAPI
{
    public class TestRequestAPI
    {
        // MTk4NGVhNDMtNzdmZi00MjYwLTg1ODQtOTFlZWRkNzZkYjRlOmE2N2FhZDA1LTRjM2EtNDg2Ni04M2U0LWRiYjM3NWZiY2Y3Yw==

        //аутентификационные данные из личного кабинета
        static string authData = null;

        public static string GetToken()
        {
            return authData;
        }

        public static void ChangeToken(string new_token)
        {
            authData = new_token;
        }

        public static async Task<string> Run(string text)
        {
            //Запуск авторизации в гигачате
            Authorization auth = new Authorization(authData, GigaChatAdapterNetFramework.Auth.RateScope.GIGACHAT_API_PERS);
            var authResult = await auth.SendRequest();

            string recieved_message = "";
            if (authResult.AuthorizationSuccess)
            {
                Completion completion = new Completion();

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
                            recieved_message = it.Message.Content;
                        }
                    }
                    else
                    {
                        recieved_message = result.ErrorTextIfFailed;
                    }
                    return recieved_message;
                }
            }
            else
            {
                recieved_message = authResult.ErrorTextIfFailed;
            }

            return "error";
        }
    }
}