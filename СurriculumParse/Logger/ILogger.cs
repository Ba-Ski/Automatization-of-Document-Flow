using System;

namespace СurriculumParse.Logger
{
    internal interface ILogger
    {
        void Info(string prefix);

        void Error(string prefix, Exception ex);

    }
}
