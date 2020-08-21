using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web;

namespace OutlookAddinMicrosoftGraphASPNET.Helpers
{
    public class Upload
    {
        // <summary>
        // OneDrive upload helper for files > 4MB
        //</summary>
        // <returns> Reponse from upload session </returns>
        public HttpResponseMessage UploadFileBySession(string url, byte[] file)
        {
            int fragSize = 1024 * 1024 * 4;
            var arrayBatches = ByteArrayIntoBatches(file, fragSize);
            int start = 0;
            HttpResponseMessage response = new HttpResponseMessage();

            foreach (var byteArray in arrayBatches)
            {
                int byteArrayLength = byteArray.Length;
                var contentRange = " bytes " + start + "-" + (start + (byteArrayLength - 1)) + "/" + file.Length;

                using (var client = new HttpClient())
                {
                    var content = new ByteArrayContent(byteArray);
                    content.Headers.Add("Content-Length", byteArrayLength.ToString());
                    content.Headers.Add("Content-Range", contentRange);

                    response = client.PutAsync(url, content).Result;
                }

                start = start + byteArrayLength;
            }
            return response;
        }

        internal IEnumerable<byte[]> ByteArrayIntoBatches(byte[] bArray, int intBufforLengt)
        {
            int bArrayLenght = bArray.Length;
            byte[] bReturn = null;

            int i = 0;
            for (; bArrayLenght > (i + 1) * intBufforLengt; i++)
            {
                bReturn = new byte[intBufforLengt];
                Array.Copy(bArray, i * intBufforLengt, bReturn, 0, intBufforLengt);
                yield return bReturn;
            }

            int intBufforLeft = bArrayLenght - i * intBufforLengt;
            if (intBufforLeft > 0)
            {
                bReturn = new byte[intBufforLeft];
                Array.Copy(bArray, i * intBufforLengt, bReturn, 0, intBufforLeft);
                yield return bReturn;
            }
        }

    }
}