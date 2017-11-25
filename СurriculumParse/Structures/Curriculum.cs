using System;
using System.Collections.Generic;
using System.Runtime.Remoting;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Office.Interop.Excel;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using СurriculumParse.ExcelParsers;

namespace СurriculumParse.Structures
{
    public class Curriculum
    {
        [BsonId]
        public  Guid Id { get; set; }

        /// <summary>
        /// Номер специальности
        /// Primary key
        /// </summary>
        public string SpecialityNumber { get; private set; }

        /// <summary>
        /// Специальность
        /// </summary>
        public string SpecialityName { get; private set; }

        /// <summary>
        /// Профиль
        /// Primary key
        /// </summary>
        public string Profile { get; private set; }

        /// <summary>
        /// Форма обучения
        /// Primary key
        /// </summary>
        public int EdForm { get; private set; }

        /// <summary>
        /// Год учебного плана
        /// Primary key
        /// </summary>
        public int Year { get; private set; }

        /// <summary>
        /// срок получения образования
        /// </summary>
        public double StudyPeriod { get; private set; }

        /// <summary>
        /// Базовая часть
        /// </summary>
        public IEnumerable<Subject> BaseSubjects { get; private set; }

        /// <summary>
        /// Количество недель в семестр. Берётся из шапки УП.
        /// </summary>
        public WeeksPerSemesterPair[] WeeksPerSemester { get; private set; }

        public Curriculum(string specialityNumber, string specialityName, string profile, double studyPeriod, IEnumerable<Subject> baseSubjects,
            int year, EducationalForm edForm, WeeksPerSemesterPair[] weeksPerSemester)
        {
            SpecialityNumber = specialityNumber;
            SpecialityName = specialityName;
            Profile = profile;
            StudyPeriod = studyPeriod;
            BaseSubjects = baseSubjects;
            Year = year;
            EdForm = (int)edForm;
            WeeksPerSemester = weeksPerSemester;

            using (var md5 = MD5.Create())
            {
                var hash = md5.ComputeHash(Encoding.Default.GetBytes(SpecialityNumber + Profile.ToLower() + Year + EdForm));
                Id = new Guid(hash);
            }
        }
    }

    public class WeeksPerSemesterPair
    {
        public int Weeks { get; private set; }
        public int AttestationWeeks { get; private set; }

        public WeeksPerSemesterPair(int weeks, int attWeeks)
        {
            Weeks = weeks;
            AttestationWeeks = attWeeks;
        }
    }
}
