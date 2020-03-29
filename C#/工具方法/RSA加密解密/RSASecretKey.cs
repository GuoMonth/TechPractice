using Org.BouncyCastle.Asn1.Pkcs;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace RSA加密解密
{
    /// <summary>
    /// gpromote
    /// 日期：2020年3月29日
    /// </summary>
    public class RSASecretKey
    {
        #region 私有变量

        private string m_publicKey;

        private string m_privateKey;

        private string m_publicKeyBase64;

        private string m_privateBase64;

        private int m_keySize;

        #endregion

        #region 属性

        
        /// <summary>
        /// 密钥大小
        /// </summary>
        public int KeySize
        {
            get
            {
                return m_keySize;
            }
        }

        /// <summary>
        /// .NET Framework 中提供的 RSA 算法规定 待加密的字节数不能超过密钥的长度值除以 8 再减去 11
        /// </summary>
        public int MaxEncryptBlockSize
        {
            get
            {
                return m_keySize / 8 - 11;
            }
        }

        /// <summary>
        /// 公钥 XML格式
        /// </summary>
        public string PublicKey
        {
            get
            {
                return m_publicKey;
            }
        }

        /// <summary>
        /// 私钥 XML格式
        /// </summary>
        public string PrivateKey
        {
            get
            {
                return m_privateKey;
            }
        }

        /// <summary>
        /// 公钥 base64格式
        /// </summary>
        public string PublicKeyBase64
        {
            get
            {
                if (string.IsNullOrEmpty(m_publicKeyBase64) && !string.IsNullOrEmpty(m_publicKey))
                {
                    m_publicKeyBase64 = PublicKeyToBase64String();
                }
                return m_publicKeyBase64;
            }
        }

        /// <summary>
        /// 私钥 base64格式
        /// </summary>
        public string PrivateKeyBase64
        {
            get
            {
                if (string.IsNullOrEmpty(m_privateBase64) && !string.IsNullOrEmpty(m_privateKey))
                {
                    m_privateBase64 = PrivateKeyToBase64String();
                }
                return m_privateBase64;
            }
        }

        #endregion

        #region 有参构造函数

        /// <summary>
        /// 有参构造函数
        /// </summary>
        /// <param name="privateKey"></param>
        /// <param name="publicKey"></param>
        /// <param name="keySize"></param>
        public RSASecretKey(string privateKey, string publicKey, int keySize)
        {
            m_privateKey = privateKey;
            m_publicKey = publicKey;
            m_keySize = keySize;
        }

        /// <summary>
        /// RSASecretKey.ToString生成的格式化字符串
        /// </summary>
        /// <param name="rsaSecretKeyStr"></param>
        public RSASecretKey(string rsaSecretKeyStr)
        {
            string[] strArray = rsaSecretKeyStr.Split('\r');
            if (strArray.Length != 3)
            {
                throw new ArgumentException($"RSA密钥加载错误：{rsaSecretKeyStr}");
            }
            m_privateKey = strArray[0].Split(':')[1].Trim();
            m_publicKey = strArray[1].Split(':')[1].Trim();
            m_keySize = Convert.ToInt32(strArray[2].Split(':')[1].Trim());
        }

        #endregion

        /// <summary>
        /// 重写父类方法ToString
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return string.Format(
                "PrivateKey: {0}\r\nPublicKey: {1}\r\nKeySize:{2}", PrivateKey, PublicKey, KeySize);
        }

        /// <summary>
        /// 方便将密钥给第三方系统。第三方系统一般都只接收base64的格式。
        /// xml private key -> base64 private key string
        /// </summary>
        /// <returns></returns>
        private string PrivateKeyToBase64String()
        {
            string result = string.Empty;
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(PrivateKey);
                RSAParameters param = rsa.ExportParameters(true);
                RsaPrivateCrtKeyParameters privateKeyParam = new RsaPrivateCrtKeyParameters(
                    new BigInteger(1, param.Modulus), new BigInteger(1, param.Exponent),
                    new BigInteger(1, param.D), new BigInteger(1, param.P),
                    new BigInteger(1, param.Q), new BigInteger(1, param.DP),
                    new BigInteger(1, param.DQ), new BigInteger(1, param.InverseQ));
                PrivateKeyInfo privateKey = PrivateKeyInfoFactory.CreatePrivateKeyInfo(privateKeyParam);

                result = Convert.ToBase64String(privateKey.ToAsn1Object().GetEncoded());
            }
            return result;
        }

        /// <summary>
        /// xml public key -> base64 public key string
        /// </summary>
        /// <returns></returns>
        private string PublicKeyToBase64String()
        {
            string result = string.Empty;
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(PublicKey);
                RSAParameters p = rsa.ExportParameters(false);
                RsaKeyParameters keyParams = new RsaKeyParameters(
                    false, new BigInteger(1, p.Modulus), new BigInteger(1, p.Exponent));
                SubjectPublicKeyInfo publicKeyInfo = SubjectPublicKeyInfoFactory.CreateSubjectPublicKeyInfo(keyParams);
                result = Convert.ToBase64String(publicKeyInfo.ToAsn1Object().GetEncoded());
            }
            return result;
        }

        /// <summary>
        /// base64 private key string -> xml private key
        /// </summary>
        /// <param name="privateKeyBase64Str"></param>
        /// <returns></returns>
        private string ToXmlPrivateKey(string privateKeyBase64Str)
        {
            RsaPrivateCrtKeyParameters privateKeyParams =
                PrivateKeyFactory.CreateKey(Convert.FromBase64String(privateKeyBase64Str)) as RsaPrivateCrtKeyParameters;
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                RSAParameters rsaParams = new RSAParameters()
                {
                    Modulus = privateKeyParams.Modulus.ToByteArrayUnsigned(),
                    Exponent = privateKeyParams.PublicExponent.ToByteArrayUnsigned(),
                    D = privateKeyParams.Exponent.ToByteArrayUnsigned(),
                    DP = privateKeyParams.DP.ToByteArrayUnsigned(),
                    DQ = privateKeyParams.DQ.ToByteArrayUnsigned(),
                    P = privateKeyParams.P.ToByteArrayUnsigned(),
                    Q = privateKeyParams.Q.ToByteArrayUnsigned(),
                    InverseQ = privateKeyParams.QInv.ToByteArrayUnsigned()
                };
                rsa.ImportParameters(rsaParams);
                return rsa.ToXmlString(true);
            }
        }

        /// <summary>
        /// base64 public key string -> xml public key
        /// </summary>
        /// <param name="pubilcKeyBase64Str"></param>
        /// <returns></returns>
        private string ToXmlPublicKey(string pubilcKeyBase64Str)
        {
            RsaKeyParameters p =
                PublicKeyFactory.CreateKey(Convert.FromBase64String(pubilcKeyBase64Str)) as RsaKeyParameters;
            using (RSACryptoServiceProvider rsa = new RSACryptoServiceProvider())
            {
                RSAParameters rsaParams = new RSAParameters
                {
                    Modulus = p.Modulus.ToByteArrayUnsigned(),
                    Exponent = p.Exponent.ToByteArrayUnsigned()
                };
                rsa.ImportParameters(rsaParams);
                return rsa.ToXmlString(false);
            }
        }
    }
}
