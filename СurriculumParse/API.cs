﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using СurriculumParse.ExcelParsers;
using СurriculumParse.Logger;

namespace СurriculumParse
{
    public class Api:IDisposable
    {
        private static Process _mongod;
        private readonly ILogger _logger;
        private readonly IDBManager _dbManager;
        private readonly IDocumentParser _parser;

        public Api()
        {
            StartMongod();
            _logger = new SerilogFileLogger();
            _dbManager = new MongoDbManager();
            _parser = new CurriculumReader(_logger);
        }

        public bool ParseCurriculum(string path)
        {
            var obj = _parser.ParseDocumenr(path);
            if (obj == null)
            {
                return false;
            }
            try
            {
                _dbManager.InsertCurriculumAsync(obj).GetAwaiter().GetResult();
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        public PpsReadStatus ParsePps(string path)
        {
            var pps = new PPSReader(_logger, _dbManager);
            return pps.WorkWithPPS(path);
        }

        public void ParseCurriculumsDirrectory(string path)
        {
            var fileManager = new FilesManager(path, _dbManager, _parser, _logger);
            var files = fileManager.ProcessAllProgamms();
            File.WriteAllLines("Ошибки.txt", files);
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

        public void Dispose()
        {
            try
            {
                _mongod?.Kill();
            }
            catch (Exception)
            {
                
            }
        }
    }
}
