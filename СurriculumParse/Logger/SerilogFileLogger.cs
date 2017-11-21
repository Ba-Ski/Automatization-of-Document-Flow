using System;
using Serilog;

namespace СurriculumParse.Logger
{
    internal class SerilogFileLogger : ILogger
    {
        private readonly Serilog.Core.Logger _logger;

        //[Conditional("DEBUG")]

        public SerilogFileLogger()
        {
            _logger = ConfigureLogger();
        }

        public void Info(string prefix)
        {
            _logger.Information(prefix);
        }

        public void Error(string prefix, Exception ex)
        {
            _logger.Error(prefix, ex);
        }

        private static Serilog.Core.Logger ConfigureLogger()
        {
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            return new LoggerConfiguration()
                .WriteTo.LiterateConsole()
                .WriteTo.Async(i => i.RollingFile(baseDirectory + "Logs/log-{Date}.txt"))
                .CreateLogger();
        }
    }
}
