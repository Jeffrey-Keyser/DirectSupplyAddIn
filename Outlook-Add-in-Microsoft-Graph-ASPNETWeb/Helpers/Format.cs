﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace OutlookAddinMicrosoftGraphASPNET.Helpers
{
    // <summary>
    // For string formatting
    //</summary>
    public class Format
    {
        // <summary>
        // Formats filenames for OneDrive upload
        //</summary>
        // <returns> Valid filename </returns>
        public static string MakeFileNameValid(string originalFileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return string.Join("_", originalFileName.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
        }


        // <summary>
        // Returns the string between strStart and strEnd within strSource. Exclusive
        //</summary>
        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }

            return "";
        }
    }
}