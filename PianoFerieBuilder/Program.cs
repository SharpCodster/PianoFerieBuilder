using CommandLine;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using PianoFerieBuilder.Models;
using PianoFerieBuilder.Services;

namespace PianoFerieBuilder
{
    public static class Program
    {
        static int Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            return Parser.Default.ParseArguments<ProgramOptions>(args)
                .MapResult(opts => RunOptionsAndReturnExitCode(opts), errs => HandleParseError(errs));
        }

        public class ProgramOptions
        {
            [Option('y', "year", Required = true, HelpText = "Anno di riferimento")]
            public int Year { get; set; }

            [Option('o', "out", Required = true, HelpText = "Path dell'Excel che verrà creato")]
            public string ExcelFileOutPath { get; set; }
        }

        private static int RunOptionsAndReturnExitCode(ProgramOptions opts)
        {
            var result = 0;
            Console.WriteLine($"Anno: {opts.Year}");
            Console.WriteLine($"Path: {opts.ExcelFileOutPath}");

            Calendar cal = new Calendar(opts.Year);

            ExcelBuilder builder = new ExcelBuilder();
            result = builder.CreateFile(opts.ExcelFileOutPath, opts.Year, cal.Days);

            Console.WriteLine($"Exit code {result}");
            return result;
        }

        static int HandleParseError(IEnumerable<Error> errs)
        {
            var result = -2;
            Console.WriteLine($"Numero errori: {errs.Count()}");
            if (errs.Any(x => x is HelpRequestedError || x is VersionRequestedError))
            {
                result = -1;
            }
            Console.WriteLine($"Exit code {result}");
            return result;
        }
    }
}
