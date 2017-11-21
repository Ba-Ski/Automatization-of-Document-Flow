using СurriculumParse.Structures;

namespace СurriculumParse
{
    internal interface IDocumentParser
    {
        Curriculum ParseDocumenr(string path);
    }
}
