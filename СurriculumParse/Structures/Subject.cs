using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;
using СurriculumParse.ExcelParsers;

namespace СurriculumParse.Structures
{
    public class Subject
    {
        public string Index { get; private set; }
        public string Name { get; private set; }

        /// <summary>
        /// список модулей с экзаменами
        /// </summary>
        public  IEnumerable<int> Exams { get; private set; }

        /// <summary>
        /// список модулей с экзаменами
        /// </summary>
        public IEnumerable<int> Credits { get; private set; }

        /// <summary>
        /// общая трудоемкость в часах
        /// </summary>
        [BsonIgnoreIfNull]
        public double? TotalComplexityHours { get; set; }

        /// <summary>
        /// общая трудоемкость в з.е.
        /// </summary>
        [BsonIgnore]
        public double TotalComplexityUnits { get; private set; }

        /// <summary>
        /// Аудиторные занятия час.
        /// </summary>
        [BsonIgnore]
        public double AuditoryClassesHours { get; private set; }

        [BsonIgnoreIfNull]
        public IEnumerable<Complexity> LectureHours { get; private set; }

        /// <summary>
        /// Часы практических работ
        /// </summary>
        [BsonIgnoreIfNull]
        public IEnumerable<Complexity> PracticeHours { get; private set; }

        /// <summary>
        /// Часы лабораторных работ
        /// </summary>
        [BsonIgnoreIfNull]
        public IEnumerable<Complexity> LabStudiesHours { get; private set; }

        /// <summary>
        /// Часы срс
        /// </summary>
        [BsonIgnoreIfNull]
        public IEnumerable<Complexity> StudentSelfStudyHours { get; private set; }

        /// <summary>
        /// Кафедра
        /// </summary>
        [BsonIgnoreIfNull]
        public IEnumerable<string> Department { get; private set; } //TODO enum?

        public SubjectType2015 SubjectTypeIdTypeId { get; private set; }

        public Subject(string index,
            string name,
            IEnumerable<int> exams,
            IEnumerable<int> credits,
            IEnumerable<Complexity> lectureHours,
            IEnumerable<Complexity> practiceHours,
            IEnumerable<Complexity> labStudiesHours,
            IEnumerable<Complexity> studentSelfStudyHours,
            IEnumerable<string> department,
            SubjectType2015 subjectTypeId)
        {
            Index = index;
            Name = name;
            Exams = exams;
            Credits = credits;
            LectureHours = lectureHours;
            PracticeHours = practiceHours;
            LabStudiesHours = labStudiesHours;
            StudentSelfStudyHours = studentSelfStudyHours;
            Department = department;
            SubjectTypeIdTypeId = subjectTypeId;
        }

        public Subject(string index,
            string name,
            IEnumerable<int> exams,
            IEnumerable<int> credits,
            double? totalTotalHours,
            SubjectType2015 subjectTypeId
            )
        {
            Index = index;
            Name = name;
            Exams = exams;
            Credits = credits;
            TotalComplexityHours = totalTotalHours;
            SubjectTypeIdTypeId = subjectTypeId;
        }
    }
}
