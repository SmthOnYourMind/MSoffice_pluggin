using GigaChatAdapterNetFramework;
using System;
using System.Text;
using System.Threading.Tasks;

namespace TestNetFrameworkAPI
{
    public static class TestNetFrameworkAPI
    {
        public static async Task Main()
        {
            await TestRequestAPI.Run("привет");
        }
    }
}