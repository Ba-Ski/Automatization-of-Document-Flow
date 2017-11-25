using System;

namespace СurriculumParse.Logger
{
    public interface ILogger
    {
        void Info(string prefix);

        void Error(string prefix, Exception ex);

    }
}
