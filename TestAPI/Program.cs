using GigaChatAdapter;
using System;
using System.Text;
using System.Threading.Tasks;

namespace TestAPI
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string str = Ribbon.GetCurrentSelected();
            Console.WriteLine(str);
            await testRequ.Run("привет");
        }
    }
}