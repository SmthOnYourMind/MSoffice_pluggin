using System;
using System.Text;
using System.Threading.Tasks;

namespace TestNetFrameworkAPI
{
    public static class TestAPI
    {
        public static async Task Start()
        {
            await TestRequestAPI.Run("привет");
        }
    }
}