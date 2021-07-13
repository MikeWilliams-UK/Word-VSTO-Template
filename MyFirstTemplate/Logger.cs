﻿using System;
using System.IO;

namespace MyFirstTemplate
{
    public static class Logger
    {
        private const string FileName = @"C:\Temp\VSTO.log";

        public static void Info(string message)
        {
            WriteToFile($"{TimeStamp()} - {message}");
        }

        private static string TimeStamp()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        }

        private static void WriteToFile(string message)
        {
            using (StreamWriter w = File.AppendText(FileName))
            {
                w.WriteLine(message);
            }
        }
    }
}