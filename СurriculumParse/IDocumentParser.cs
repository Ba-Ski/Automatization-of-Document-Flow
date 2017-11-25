using СurriculumParse.Structures;

namespace СurriculumParse
{
    public interface IDocumentParser
    {
        Curriculum ParseDocumenr(string path);
    }
}
