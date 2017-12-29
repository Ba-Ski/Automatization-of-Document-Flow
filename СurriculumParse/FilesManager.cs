using System.Collections.Generic;
using System.IO;
using System.Linq;
using СurriculumParse.Logger;

namespace СurriculumParse
{
    public class FilesManager
    {
        private readonly string _rootPath;
        private readonly IDBManager _dbManager;
        private readonly ILogger _logger;
        private readonly IDocumentParser _parser;

        public struct ParseResult
        {
            public IEnumerable<string> Errors;
            public IEnumerable<string> Successes;
        }

        public FilesManager(string path, IDBManager dbManager, IDocumentParser parser, ILogger logger)
        {
            _rootPath = path;
            _dbManager = dbManager;
            _parser = parser;
            _logger = logger;
        }

        private IEnumerable<DirectoryInfo> GetSubdirs()
        {
            var dirInfo = new DirectoryInfo(_rootPath);
            return !dirInfo.Exists ? null : dirInfo.GetDirectories();
        }

        public ParseResult ProcessAllProgamms()
        {
            var dirs = GetSubdirs();
            var errorsList = new List<string>();
            var succesList = new List<string>();
            foreach (var dir in dirs)
            {
                ProcessProgramm(dir, errorsList, succesList);
            }

            ParseResult result;
            result.Errors = errorsList;
            result.Successes = succesList; 
            return result;
        }

        public void ProcessProgramm(DirectoryInfo dirInfo, List<string> errorsList, List<string> succesList)
        {
            if (!dirInfo.Exists)
            {
                _logger.Info("File manager: broken directory");
                return;
            }

            var programNumber = dirInfo.Name;
            _logger.Info($"File manager: steped in programm number {programNumber}");

            foreach (var programNameDir in dirInfo.GetDirectories())
            {
                var progarmmName = programNameDir.Name;
                _logger.Info($"File manager: steped in programm with name {progarmmName}");

                var yearDir = programNameDir.GetDirectories("2015").FirstOrDefault();
                if (yearDir == null)
                {
                    errorsList.Add(programNameDir.FullName + "\\" + " -- нет директории для 2015 года");
                    _logger.Info("File Manager: no directory for 2015 year");
                    continue;
                }

                var files = GetProgrammDocs(yearDir);

                var file = files?.FirstOrDefault(f => f.Name.StartsWith("УП2015"));

                if (file != null)
                {
                    if (file.Extension == ".xls")
                    {
                        _logger.Info("File is with obsolete format change it to xlsx");
                        errorsList.Add(file.Name + " -- старый формат файла. Нужно переделать в xlsx");
                        continue;
                    }
                    var curriculum = _parser.ParseDocumenr(file.FullName);
                    if (curriculum != null)
                    {
                        _dbManager.ReplaceCurriculumAsync(curriculum);
                        succesList.Add(file.Name);
                    }
                    else
                    {
                        errorsList.Add(file.Name);
                    }
                }
                else
                {
                    errorsList.Add(yearDir.FullName + "\\" + " -- нет учебного плана");
                }

            }
        }

        private IEnumerable<FileInfo> GetProgrammDocs(DirectoryInfo dir)
        {
            if (!dir.Exists)
            {
                _logger.Info("File manager: can't get files from broken directory");
                return null;
            }
            
            var files = dir.GetFiles();
            if (files.Length != 0)
            {
                return files;
            }

            var internalDirs = dir.GetDirectories();
            if (internalDirs.Length != 0)
            {
                _logger.Info("File Manager: unsupported file path");
                return null;
            }
            else
            {
                _logger.Info("File Manager: empty direcetory");
                return null;
            }
        }
    }
}
