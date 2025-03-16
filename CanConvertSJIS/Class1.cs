using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CanConvertSJIS
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)] // これは必須
    [Guid("E0BA9F7B-2198-46DD-A569-B329B46F80BF")]
    public class CanConvertSJIS
    {
        // Shift_JIS エンコーディングを取得
        Encoding sjisEncoding = Encoding.GetEncoding("Shift_JIS");

        public string TargetText
        {
            set
            {
                CanEncodeSjis(value);
            }
        }

        public bool CanEncode
        {
            get
            {
                return lastCanEncode;
            }
        }

        private bool lastCanEncode = false;

        private bool CanEncodeSjis(string text)
        {
            try
            {
                // 文字列を SJIS にエンコード（変換）
                byte[] encodedBytes = sjisEncoding.GetBytes(text);

                // デコードできるかどうかをチェック (エンコードとデコードで情報損失がないか)
                string decodedText = sjisEncoding.GetString(encodedBytes);
                // 文字列のエンコードとデコードで変化がある場合、SJISに変換できない文字が含まれている
                lastCanEncode = text.Equals(decodedText);

                return lastCanEncode;
            }
            catch (EncoderFallbackException)
            {
                lastCanEncode = false;
                // エンコードできない文字が含まれている場合、EncoderFallbackExceptionが発生
                return false;
            }
            catch (ArgumentNullException) // textがnullだった場合
            {
                lastCanEncode = false;
                return false;
            }

            lastCanEncode = false;
            return false;
        }
    }
}
