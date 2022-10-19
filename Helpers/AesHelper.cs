using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Coginov.GraphApi.Library.Helpers
{
    public static class AesHelper
    {
        private static int keySize = 256;
        private static string keyValue = "AH6ptEHe.CoginovInc2022.eDd5wM5B";

        public static string EncryptToString(string text, string keyString = null)
        {
            if (string.IsNullOrWhiteSpace(keyString))
            {
                keyString = keyValue;
            }

            var encryptedByteArray = Encrypt(text, keyString);
            return Convert.ToBase64String(encryptedByteArray);
        }

        public static string DecryptToString(string cipherText, string keyString = null)
        {
            if (string.IsNullOrWhiteSpace(cipherText))
            {
                return string.Empty;
            }

            if (string.IsNullOrWhiteSpace(keyString))
            {
                keyString = keyValue;
            }

            var fullCipher = Convert.FromBase64String(cipherText);
            return Decrypt(fullCipher, keyString);
        }

        #region Private Methods

        private static byte[] Encrypt(string text, string keyString)
        {
            var key = Encoding.UTF8.GetBytes(keyString);

            using (var aesAlg = Aes.Create())
            {
                aesAlg.KeySize = keySize;
                using (var encryptor = aesAlg.CreateEncryptor(key, aesAlg.IV))
                {
                    using (var msEncrypt = new MemoryStream())
                    {
                        using (var csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                        {

                            using (var swEncrypt = new StreamWriter(csEncrypt))
                            {
                                swEncrypt.Write(text);
                            }

                            var iv = aesAlg.IV;
                            var decryptedContent = msEncrypt.ToArray();

                            var result = new byte[iv.Length + decryptedContent.Length];

                            Buffer.BlockCopy(iv, 0, result, 0, iv.Length);
                            Buffer.BlockCopy(decryptedContent, 0, result, iv.Length, decryptedContent.Length);

                            return result;
                        }
                    }
                }
            }
        }

        private static string Decrypt(byte[] fullCipher, string keyString)
        {
            var key = Encoding.UTF8.GetBytes(keyString);

            var iv = new byte[16];
            var cipher = new byte[fullCipher.Length - iv.Length];

            Buffer.BlockCopy(fullCipher, 0, iv, 0, iv.Length);
            Buffer.BlockCopy(fullCipher, iv.Length, cipher, 0, fullCipher.Length - iv.Length);

            using (var aesAlg = Aes.Create())
            {
                aesAlg.KeySize = keySize;
                using (var decryptor = aesAlg.CreateDecryptor(key, iv))
                {
                    using (var msDecrypt = new MemoryStream(cipher))
                    {
                        using (var csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                        {
                            using (var srDecrypt = new StreamReader(csDecrypt))
                            {
                                string result = srDecrypt.ReadToEnd();
                                return result;
                            }
                        }
                    }
                }
            }
        }
    }

    #endregion
}