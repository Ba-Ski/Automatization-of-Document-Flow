using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СurriculumParse.ExcelParsers
{
    public static class Constants
    {
        public const string WeeksPerSemesterName = "Кол-во недель обучения";
        public const string SpecialityName = "НАПРАВЛЕНИЕ ПОДГОТОВКИ";
        public const string BasePartName = "Б.1.1 БАЗОВАЯ ЧАСТЬ";
        public const string VariativePartName = "Б.1.2 ВАРИАТИВНАЯ ЧАСТЬ";
        public const string PhysicalEdPartName = "Б.1.3. ФИЗИЧЕСКАЯ КУЛЬТУРА (ЭЛЕКТИВНАЯ ДИСЦИПЛИНА)";
        public const string PracticePartName = "Б.2. ПРАКТИКИ";
        public const string AttestationPartName = "Б.3. ГОСУДАРСТВЕННАЯ ИТОГОВАЯ АТТЕСТАЦИЯ";

        public const int SubjectRows = 6;
        public const int DefaultSemesterCont = 8;
        public const int SemestersInYear = 2;
        public const int ActivitiesTypesCount = 4;
        
    }

    public enum XlsxSectionsRows2015
    {
        SpecialityField = 20,
        ProfileField = 22,
        DurationField = 24,
        WeeksPerSemester = 40,
        BasePart = 49,
        VariativePart = 438,
        PhysicalEdPart = 837,
        PracticePart = 847,
        AttestationPart = 855
    }

    public enum SubjectType2015
    {
        BaseSubject,
        VariativeSubject,
        PhysicalEd,
        Practice,
        Attestation
    }

    public enum ActivitiesRows
    {
        Lections = 0,
        Labs,
        Practices,
        SelfStudy
    };

    public enum SubjectColumn
    {
        Index = 1,
        Name,
        ExamsSemesters,
        CreditSemesters,
        TotalScore,
        TotalAuditoryHours,
        HoursByType,    
        HoursTable,
        Department = 16
    };

    public enum PpsReadStatus
    {
        Success,
        CurriculumNotFound,
        PpsReadError,
        FileOpenException
    }

    public enum EducationalForm
    {
        Intramural = 0,     //очная форма
        Extramural = 1,     //заочная
        IntraExtramural = 2 //очно-заочная
    }
}
