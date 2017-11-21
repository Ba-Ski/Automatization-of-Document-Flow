using System;
using System.Threading.Tasks;
using СurriculumParse.Structures;

namespace СurriculumParse
{
    internal interface IDBManager
    {
        Task InsertCurriculumAsync(Curriculum curriculum);
        Task<Curriculum> GetCurriculumByIdAsync(Guid id);
        Task<Curriculum> GetCurriculumAsync(string specialityNumber, string profile, int year, int edForm);
    }
}
