using СurriculumParse.ExcelParsers;
using СurriculumParse.Logger;

namespace СurriculumParse
{
    public class Program
    {

        static string GetPath(string[] args)
        {
            if (args.Length == 0)
            {
                System.Console.WriteLine("Please enter root path.");
                return null;
            }

            return args[0];
        }
        static void Main(string[] args)
        {
            var logger = new SerilogFileLogger();
            var dbmanager = new MongoDbManager();
            var parser = new CurriculumReader(logger);
            logger.Info($"Application started");
            var path = GetPath(args);
            var fileManager = new FilesManager(path, dbmanager, parser, logger);
            fileManager.ProcessAllProgamms();

            var pps = new PPSReader(logger, dbmanager);
            //pps.WorkWithPPS("C:\\Users\\baski\\Documents\\База ОПОП\\09.03.01\\Вычислительные машины, комплексы, системы и сети_очная\\2015\\ППС_09.03.01_Вычисл.машины,компл,сис. и сети_2015_оН.xlsx");
            //var obj = parse.ParseDocumenr("C:\\Users\\baski\\Documents\\Visual Studio 2015\\Projects\\СurriculumParse\\СurriculumParse\\bin\\Debug\\УП_09.03.04_ФГОС_ВО.xlsx");
            //dbmanager.InsertCurriculumAsync(obj).GetAwaiter().GetResult();
            logger.Info($"Application finished");
        }
    }
}
