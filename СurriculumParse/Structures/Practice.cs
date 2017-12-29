using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace СurriculumParse.Structures
{

    /// <summary>
    /// Решил описывать практику в классе Subject единообразно как и предметы, аттестации и физическую культуру. 
    /// Не уверен, будет ли это нужно далее.
    /// </summary>
    public class Practice
    {
        public string Index { get; private set; }
        public IEnumerable<string> Name { get; private set; }
        public IEnumerable<int> Exams { get; private set; }
        public IEnumerable<int> Credits { get; private set; }
        public double? Hours { get; set; }

        public Practice(string practiceIndex, IEnumerable<string> name, IEnumerable<int> exams,
            IEnumerable<int> credits, double? hours)
        {
            Index = practiceIndex;
            Name = name;
            Exams = exams;
            Credits = credits;
            Hours = hours;
        }
    }
}
