using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RSA加密解密
{
    class Program
    {
        /// <summary>
        /// Gpromote
        /// RSA加密解密工具类使用示例
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            Console.WriteLine("===============RSA 加密/解密 小程序===============");
            //展示公钥
            Console.WriteLine($"RSA-公钥-base64格式：{RSAUtil.GetRSAPulibcKey(1)}");
            //输入待加密文本
            Console.Write($"请输入待加密文本（文本最大长度：{RSAUtil.GetMaxEncryptBlockSize()}，回车结束）:");
            string str = Console.ReadLine();
            //展示加密后文本
            string encryptStr = RSAUtil.RSAEncrypt(str);
            Console.WriteLine($"加密后文本：{encryptStr}");
            //解密文本
            Console.WriteLine($"解密后文本：{RSAUtil.RSADecrypt(encryptStr)}");


            Console.ReadLine();
        }
    }
}
