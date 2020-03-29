using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace RSA加密解密
{
    public static class RSAUtil
    {
        /// <summary>
        /// RSA密钥文件生成后存储路径
        /// （可以修改代码，使用数据库获取其他方式缓存密钥）
        /// </summary>
        private static readonly string m_S_rsaFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"RSAKey.rsa");

        /// <summary>
        /// 缓存RSA对象
        /// </summary>
        private static RSA rSA;

        /// <summary>
        /// 私有属性 全局缓存 RSA对象并按规则更新
        /// 更新规则：23后获取的RSA对象都是重新创建的最新的。其他时间都是使用的缓存对象。
        /// 应为RSA密钥是一天一创建，为了提高效率，其他时间就没必要重新创建对象。
        /// </summary>
        private static RSA m_RSA
        {
            get
            {
                if (rSA == null)
                {
                    rSA = new RSA();
                    rSA.InitRSAKey(m_S_rsaFilePath);
                }

                return rSA;
            }
        }

        /// <summary>
        /// 获取允许最大密钥加密长度
        /// </summary>
        /// <returns></returns>
        public static int GetMaxEncryptBlockSize()
        {
            return m_RSA.m_MaxEncryptBlockSize;
        }

        /// <summary>
        /// 获取RSA的公钥
        /// </summary>
        /// <param name="keyType">0:xml格式 1:base64格式</param>
        /// <returns></returns>
        public static string GetRSAPulibcKey(int keyType)
        {
            lock (m_RSA)
            {
                return m_RSA.GetRSAPulibcKey(keyType);
            }
        }

        /// <summary>
        /// RSA加密
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        public static string RSAEncrypt(string content)
        {
            lock (m_RSA)
            {
                if (content.Length > m_RSA.m_MaxEncryptBlockSize)
                {
                    throw new ArgumentException($"待加密文本过长，无法加密，请缩短待加密文本后再尝试。允许最大加密文本长度：{ m_RSA.m_MaxEncryptBlockSize}");
                }
                return m_RSA.RSAEncrypt(content);
            }
        }

        /// <summary>
        /// RSA解密
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        public static string RSADecrypt(string content)
        {
            lock (m_RSA)
            {
                return m_RSA.RSADecrypt(content);
            }
        }
    }
}
