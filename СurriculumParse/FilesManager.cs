using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using СurriculumParse.ExcelParsers;
using СurriculumParse.Logger;

namespace СurriculumParse
{
    internal class FilesManager
    {
        private readonly string _rootPath;
        private readonly IDBManager _dbManager;
        private readonly ILogger _logger;
        private readonly IDocumentParser _parser;

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

        public void ProcessAllProgamms()
        {
            var dirs = GetSubdirs();
            foreach (var dir in dirs)
            {
                ProcessProgramm(dir);
            }
        }

        public void ProcessProgramm(DirectoryInfo dirInfo)
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
                        continue;
                    }
                    _dbManager.InsertCurriculumAsync(_parser.ParseDocumenr(file.FullName));
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
