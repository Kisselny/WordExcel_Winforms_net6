using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
//using System.Collections.Generic;
using OfficeOpenXml;
//using Paragraph = Xceed.Document.NET.Paragraph;
using System.Text.RegularExpressions;
using System.Linq;
using System.Windows.Forms;
//using System.IO.Packaging;
//using System.Text.RegularExpressions;
//using Xceed.Document.NET;
using Xceed.Words.NET;


namespace WordExcel_Winforms_net6
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]

        
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            Application.Run(new Form1());
		}
    }
}