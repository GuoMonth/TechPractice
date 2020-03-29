using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace RSA加密解密
{
    class RSA
    {
        /// <summary>
        /// RSA长度 固定设置为2048
        /// </summary>
        private static readonly int m_S_RSAlength = 2048;

        /// <summary>
        /// RSA密钥对象
        /// </summary>
        private RSASecretKey m_rSASecretKey = null;

        /// <summary>
        /// 公开属性 公钥 xml格式
        /// </summary>
        public string m_PublicKeyXML
        {
            get
            {
                return m_rSASecretKey.PublicKey;
            }
        }

        /// <summary>
        /// 公开属性 公钥 base64格式
        /// </summary>
        public string m_PublicKeyBase64
        {
            get
            {
                return m_rSASecretKey.PublicKeyBase64;
            }
        }

        /// <summary>
        /// 允许最大密钥加密长度
        /// </summary>
        public int m_MaxEncryptBlockSize
        {
            get
            {
                return m_rSASecretKey.MaxEncryptBlockSize;
            }
        }


        /// <summary>
        /// 无参构造函数
        /// </summary>
        public RSA()
        {
        }

        /// <summary>
        /// 初始化
        /// 1、获取 有没有RSA密钥
        /// 2、如果有则完成初化始变量 <see cref="m_rSASecretKey"/>
        /// 3、如果没有 则生成密钥，并缓存到文件中 并完成初始化变量 <see cref="m_rSASecretKey"/>
        /// </summary>
        /// <param name="rsafilePath">存放RSA密钥的文件位置（包括公钥和私钥）</param>
        public void InitRSAKey(string rsafilePath)
        {
            lock (this)
            {
                if (File.Exists(rsafilePath))
                {
                    //读取文件内容 并加载密钥
                    using (StreamReader sr = new StreamReader(rsafilePath, Encoding.UTF8))
                    {
                        string rsaStr = sr.ReadToEnd().Trim();
                        m_rSASecretKey = new RSASecretKey(rsaStr);
                        sr.Close();
                    }
                }
                else
                { //文件不存在 则重新生成密钥 并存入此文件路径
                    m_rSASecretKey = GenerateRSAKey();
                    using (StreamWriter sw = new StreamWriter(rsafilePath, false, Encoding.UTF8)) //存在同名文件则进行覆盖（工程上使用为保险起见应该是备份后再覆盖）
                    {
                        sw.Write(m_rSASecretKey.ToString());
                        sw.Close();
                    }
                }
            }
        }


        /// <summary>
        /// 获取RSA的公钥
        /// </summary>
        /// <param name="keyType">0:xml格式 1:base64格式</param>
        /// <returns></returns>
        public string GetRSAPulibcKey(int keyType)
        {
            return keyType == 0 ? m_PublicKeyXML : m_PublicKeyBase64;
        }

        /// <summary>
        /// 生成RSA密钥对
        /// </summary>
        /// <returns></returns>
        private RSASecretKey GenerateRSAKey()
        {
            RSASecretKey rsaKey = null;
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider(m_S_RSAlength))
            {
                rsaKey = new RSASecretKey(rsa.ToXmlString(true), rsa.ToXmlString(false), m_S_RSAlength); //m_S_RSAlength rsa.KeySize
            }
            return rsaKey;
        }

        /// <summary>
        /// RSA加密
        /// </summary>
        /// <param name="xmlPublicKey"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        public string RSAEncrypt(string content)
        {
            string xmlPublicKey = m_PublicKeyXML; //获取公钥 xml格式
            string encryptedContent = string.Empty;
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(xmlPublicKey);
                byte[] encryptedData = rsa.Encrypt(Encoding.UTF8.GetBytes(content), false);
                encryptedContent = Convert.ToBase64String(encryptedData);
            }
            return encryptedContent;
        }

        /// <summary>
        /// RSA解密
        /// </summary>
        /// <param name="xmlPrivateKey"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        public string RSADecrypt(string content)
        {
            string xmlPrivateKey = m_rSASecretKey.PrivateKey; //获取私钥 xml格式
            string decryptedContent = string.Empty;
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(xmlPrivateKey);
                byte[] decryptedData = rsa.Decrypt(Convert.FromBase64String(content), false);
                decryptedContent = Encoding.GetEncoding("utf-8").GetString(decryptedData); //utf-8 gb2312
            }
            return decryptedContent;
        }
    }

}
