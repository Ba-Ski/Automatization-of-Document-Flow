using System;
using System.Threading.Tasks;
using СurriculumParse.Structures;

namespace СurriculumParse
{
    public interface IDBManager
    {
        Task InsertCurriculumAsync(Curriculum curriculum);
        Task ReplaceCurriculumAsync(Curriculum curriculum);
        Task<Curriculum> GetCurriculumByIdAsync(Guid id);
        Curriculum GetCurriculumAsync(string specialityNumber, string profile, int year, int edForm);
    }
}
