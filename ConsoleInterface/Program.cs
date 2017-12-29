using System.Collections.Generic;
using System.Diagnostics;
using СurriculumParse;
using СurriculumParse.ExcelParsers;
using СurriculumParse.Logger;

namespace ConsoleInterface
{
    public class Program
    {
        private static Process _mongod;

        private static string GetPath(IReadOnlyList<string> args)
        {
            if (args.Count == 0)
            {
                System.Console.WriteLine("Please enter root path.");
                return null;
            }

            return args[0];
        }

        private static void Main(string[] args)
        {
            try
            {
                StartMongod();
                var logger = new SerilogFileLogger();
                var dbmanager = new MongoDbManager();
                var parser = new CurriculumReader(logger);
                logger.Info($"Application started");
                var path = GetPath(args);

                //ParseCurriculumsDirrectory(path, dbmanager, parser, logger);

                var pps = new PPSReader(logger, dbmanager);
                pps.WorkWithPPS(
                    "C:\\Users\\baski\\Documents\\База ОПОП\\09.03.01\\Вычислительные машины, комплексы, системы и сети_очная\\2015\\ППС_09.03.01_Вычисл.машины,компл,сис. и сети_2015_оН.xlsx");

                //ParseCurriculum(parser, dbmanager);
                logger.Info($"Application finished");
                //stopping the mongod server (when app is closing)
            }
            finally
            {
                _mongod?.Kill();
            }
            
        }

        private static void ParseCurriculum(CurriculumReader parser, MongoDbManager dbmanager)
        {
            var obj = parser.ParseDocumenr(
                "C:\\Users\\baski\\Documents\\База ОПОП\\09.03.01\\Вычислительные машины, комплексы, системы и сети_очная\\2015\\УП2015_09.03.01_Вычислительные машины, комплексы, системы и сети_для_о.xlsx");
            dbmanager.InsertCurriculumAsync(obj).GetAwaiter().GetResult();
        }

        private static void ParsePPS(string path, ILogger logger, IDBManager dbManager)
        {
            var pps = new PPSReader(logger, dbManager);
            pps.WorkWithPPS(path);
        }

        private static void ParseCurriculumsDirrectory(string path, IDBManager dbmanager, IDocumentParser parser,
            ILogger logger)
        {
            var fileManager = new FilesManager(path, dbmanager, parser, logger);
            fileManager.ProcessAllProgamms();
        }

        private static void StartMongod()
        {
            var start = new ProcessStartInfo
            {
                FileName = "mongod.exe",
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = @"--config c:\mongod.conf"
            };

            _mongod = Process.Start(start);
        }

    }
}
