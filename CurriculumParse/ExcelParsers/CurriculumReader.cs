using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using СurriculumParse.Structures;
using СurriculumParse.Logger;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace СurriculumParse.ExcelParsers
{
    public class CurriculumReader : IDocumentParser
    {

        private readonly ILogger _logger;
        private ExcelWorksheet _ws;
        private int _semestersCount;
        private FileInfo _fileInfo;

        private const int CurriculumYear = 2015;

        private delegate IEnumerable<int> GetSemestersMethod(int row, int column);

        private struct CommonSubjectInfo
        {
            public string Index;
            public IEnumerable<string> Names;
            public IEnumerable<int> Exams;
            public IEnumerable<int> Credits;
            public double? Hours;
        };

        public CurriculumReader(ILogger logger)
        {
            _logger = logger;
            _semestersCount = Constants.DefaultSemesterCont;
        }

        public Curriculum ParseDocumenr(string path)
        {
            _fileInfo = new FileInfo(path);
            try
            {
                using (var package = new ExcelPackage(_fileInfo))
                {
                    _ws = package.Workbook.Worksheets[1];

                    var edForm = GetEdForm();

                    var speciality = ReadDescriptionField((int) XlsxSectionsRows2015.SpecialityField);
                    var parts = speciality.Split(new[] {' '}, StringSplitOptions.RemoveEmptyEntries);
                    var specialiyNumber = parts[0]; //ToDo validation
                    var specialityName = parts[1];
                    var profile = ReadDescriptionField((int) XlsxSectionsRows2015.ProfileField);
                    var durationStr = ReadDescriptionField((int) XlsxSectionsRows2015.DurationField);
                    var years = GetDuration(durationStr);
                    _semestersCount = (int) years * 2;

                    var weeksPerSemesterRow = CheckOrGetStartRow(Constants.WeeksPerSemesterName, 1,
                        (int) XlsxSectionsRows2015.WeeksPerSemester, (int) SubjectColumn.HoursTable);
                    if (weeksPerSemesterRow == -1)
                    {
                        throw new ApplicationException($"Wrong document structure. Can't find weeks per semester");
                    }

                    var basePartRow = CheckOrGetStartRow(Constants.BasePartName, weeksPerSemesterRow,
                        (int) XlsxSectionsRows2015.BasePart);
                    if (basePartRow == -1)
                    {
                        throw new ApplicationException($"Wrong document structure. Can't find base part");
                    }
                    var variativePartRow = CheckOrGetStartRow(Constants.VariativePartName, basePartRow,
                        (int) XlsxSectionsRows2015.VariativePart);
                    if (variativePartRow == -1)
                    {
                        throw new ApplicationException($"Wrong document structure. Can't find variative part");
                    }
                    var phisicalPartRow = CheckOrGetStartRow(Constants.PhysicalEdPartName, variativePartRow,
                        (int) XlsxSectionsRows2015.PhysicalEdPart);
                    if (phisicalPartRow == -1)
                    {
                        throw new ApplicationException($"Wrong document structure. Can't find phisical part");
                    }
                    var practicePartRow = CheckOrGetStartRow(Constants.PracticePartName, phisicalPartRow,
                        (int) XlsxSectionsRows2015.PracticePart);
                    if (practicePartRow == -1)
                    {
                        throw new ApplicationException($"Wrong document structure. Can't find practice part");
                    }
                    var attestaionPartRow = CheckOrGetStartRow(Constants.AttestationPartName, practicePartRow,
                        (int) XlsxSectionsRows2015.AttestationPart);
                    if (attestaionPartRow == -1)
                    {
                        throw new ApplicationException($"Wrong document structure. Can't find attestation part");
                    }

                    var weeksPerSemester = GetWeeksPerSemester(weeksPerSemesterRow);
                    var baseSubjects = ParseSubjects(basePartRow, variativePartRow - 5, SubjectType2015.BaseSubject);
                    var variativeSubjects = ParseSubjects(variativePartRow, phisicalPartRow - 9,
                        SubjectType2015.VariativeSubject);
                    var physicalSubjects =
                        ParseSubjects(phisicalPartRow, practicePartRow - 4, SubjectType2015.PhysicalEd);
                    var practiceSection = ParsePractice(practicePartRow);
                    var attestationSection = ParseAttestaion(attestaionPartRow);

                    var allSubjects = new List<Subject>();

                    allSubjects.AddRange(baseSubjects);
                    allSubjects.AddRange(variativeSubjects);
                    allSubjects.AddRange(physicalSubjects);
                    allSubjects.AddRange(practiceSection);
                    allSubjects.AddRange(attestationSection);

                    var curriculum = new Curriculum(specialiyNumber, specialityName, profile, years, allSubjects,
                        CurriculumYear, edForm, weeksPerSemester);
                    return curriculum;
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Parsing curriculum {_fileInfo.Name}. Error during parsing: {ex.Message}", ex);
                return null;
            }
        }

        private string ReadDescriptionField(int row)
        {
            return _ws.GetValue<string>(row, 2) ??
                   _ws.GetValue<string>(row, 1);
        }

        private WeeksPerSemesterPair[] GetWeeksPerSemester(int row)
        {
            var pairs = new List<WeeksPerSemesterPair>();
            var column = (int) SubjectColumn.HoursTable;
            for (var i = 0; i < _semestersCount; i++)
            {
                var weeks = ReadIntCellSafe(row, column + i);
                var attWeeks = ReadIntCellSafe(row + 1, column + i);
                if (weeks < 0 || attWeeks < 0)
                {
                    throw new ApplicationException("Can't read weeks per semester");
                }
                pairs.Add(new WeeksPerSemesterPair(weeks, attWeeks));
            }
            return pairs.ToArray();
        }

        private EducationalForm GetEdForm()
        {
            var l = _fileInfo.Name.Length - _fileInfo.Extension.Length;
            var parts = _fileInfo.Name.Substring(0, l).ToLower()
                .Split(new[] {'_'}, StringSplitOptions.RemoveEmptyEntries);
            var s = parts[parts.Length - 1];
            if (s == "о")
                return EducationalForm.Intramural;
            else if (s == "оз")
            {
                return EducationalForm.IntraExtramural;
            }
            else if (s == "з")
            {
                return EducationalForm.Extramural;
            }
            else
            {
                throw new ApplicationException($"Can't define education form {s}");
            }
        }

        private int CheckOrGetStartRow(string headerStr, int startRow, int row, int column = 1)
        {

            var header = _ws.GetValue<string>(row - 2, column); //А если название блока будет не на две строки выше?
            if (header == null)
            {
                return -1;
            }
            if (header == headerStr)
            {
                return row;
            }

             while (_ws.GetValue<string>(startRow, column) != headerStr)
            {
                if (_ws.Row(startRow).Hidden)
                {
                    startRow += 6;
                    continue;
                }
                startRow++;
            }

            if (startRow == 1000)
            {
                return -1;
            }

            return startRow + 2;
        }

        private static double GetDuration(string str)
        {
            var words = str.Split(new[] {' '}, StringSplitOptions.RemoveEmptyEntries);
            var l = words.Length;
            var number = words[l - 2]; //the penultimate word, where number should be located
            var duration = 0d;
            if (!double.TryParse(number.Replace('.', ','), out duration))
            {
                throw new ArgumentException("Can't parse duration information in row " + (int)XlsxSectionsRows2015.DurationField);
            }

            return duration;
        }

        private IEnumerable<Subject> ParseSubjects(int firstRow, int lastRow, SubjectType2015 type)
        {
            var row = firstRow;
            var subjects = new List<Subject>();

            do
            {
                if (_ws.Row(row).Hidden)
                {
                    row += 6;
                    continue;
                }

                var info = GetCommonSubjectInfo(row, GetSemesters);

                if (info.Names.All(string.IsNullOrEmpty) || string.IsNullOrEmpty(info.Index))
                {
                    row += 6;
                    continue;
                }

                var lections = new List<Complexity>();
                var labs = new List<Complexity>();
                var practices = new List<Complexity>();
                var selfStudy = new List<Complexity>();

                for (var i = 0; i < _semestersCount; i++)
                {
                    ReadActivityHours(row + (int) ActivitiesRows.Lections, i, lections);
                    ReadActivityHours(row + (int) ActivitiesRows.Labs, i, labs);
                    ReadActivityHours(row + (int) ActivitiesRows.Practices, i, practices);
                    ReadActivityHours(row + (int) ActivitiesRows.SelfStudy, i, selfStudy);
                }
                var department = GetFullString(row, (int) SubjectColumn.Department, Constants.SubjectRows);

                subjects.AddRange(info.Names.Select(name =>
                    new Subject(info.Index, name, info.Exams, info.Credits, lections, practices, labs, selfStudy, department, type)));

                row += 6;
            } while (row < lastRow);

            return subjects;
        }

        private IEnumerable<string> GetFullString(int row, int column, int height)
        {
            try
            {
                var strs = new List<string>();
                var res = new List<string>();
                for (var i = 0; i < height; i++)
                {
                    var cell = _ws.Cells[row + i, column];

                    var str = _ws.Cells[row + i, column].GetValue<string>();

                    var borderBottom = cell.Style.Border.Bottom;

                    if (string.IsNullOrEmpty(str) && borderBottom.Style == ExcelBorderStyle.None)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(str) && !string.IsNullOrWhiteSpace(str))
                    {
                        str = str.Trim();
                        strs.Add(str);
                    }


                    if (borderBottom.Style != ExcelBorderStyle.None)
                    {
                        res.Add(string.Join(" ", strs));
                        strs.Clear();
                    }
                }

                if (strs.Count > 0)
                {
                    res.Add(string.Join(" ", strs));
                }

                return res;
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Can't read string in {row}", ex);
            }
        }

        private IEnumerable<int> GetSemesters(int row, int column)
        {
            var exams = new List<int>(); // ToDo maybe return null?
            for (var i = 0; i < Constants.SubjectRows; i++)
            {
                var values = GetCommaSeparatedInt(row + i, column);
                if (values != null)
                {
                    exams.AddRange(values);
                }
            }
            return exams;
        }

        private int ReadIntCellSafe(int row, int column)
        {
            var value = _ws.GetValue<string>(row, column);
            int res;
            if (int.TryParse(value, out res))
            {
                return res;
            }
            {
                _logger.Info($"Parsing curriculum {_fileInfo.Name}. Can't parse int in {row}:{column}");
            }
            return -1;
        }

        private double? ReadDoubleCellSafe(int row, int column)
        {
            try
            {
                var value = _ws.GetValue<string>(row, column);
                if (string.IsNullOrEmpty(value))
                {
                    return null;
                }
                double res;
                if (double.TryParse(value, out res))
                {
                    return res;
                }
                {
                    _logger.Info($"Parsing curriculum {_fileInfo.Name}. Can't parse double in {row}:{column}");
                    throw new ApplicationException();
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Can't read double in {row}:{column}", ex);
            }
        }

        private IEnumerable<Subject> ParsePractice(int row)
        {
            var list = new List<Subject>();

            do
            {
                var info = GetCommonSubjectInfo(row, GetCommaSeparatedInt);

                list.Add(new Subject(info.Index, info.Names.FirstOrDefault(), info.Exams, info.Credits, info.Hours, SubjectType2015.Practice));
                row++;

            } while (!string.IsNullOrEmpty(_ws.GetValue<string>(row, (int)SubjectColumn.Index)));

            return list;
        }

        private CommonSubjectInfo GetCommonSubjectInfo(int row, GetSemestersMethod method)
        {
            try
            {
                var index = _ws.GetValue<string>(row, (int) SubjectColumn.Index);
                var name = GetFullString(row, (int) SubjectColumn.Name, 1);
                var exams = method(row, (int) SubjectColumn.ExamsSemesters);
                var credits = method(row, (int) SubjectColumn.CreditSemesters);
                var hours = ReadDoubleCellSafe(row, (int) SubjectColumn.TotalScore);
                if (!hours.HasValue)
                {
                    _logger.Info(
                        $"Parsing curriculum {_fileInfo.Name}. Can't read activity hours in {row}:{(int) SubjectColumn.TotalScore}");
                }

                return new CommonSubjectInfo()
                {
                    Index = index,
                    Names = name,
                    Exams = exams,
                    Credits = credits,
                    Hours = hours
                };
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"Can't read subject in {row}", ex);
            }
        }

        private IEnumerable<Subject> ParseAttestaion(int row)
        {
            var list = new List<Subject>();

            do
            {
                var info = GetCommonSubjectInfo(row, GetCommaSeparatedInt);

                list.Add(new Subject(info.Index, info.Names.FirstOrDefault(), info.Exams, info.Credits, info.Hours, SubjectType2015.Attestation));
                row++;

            } while (!string.IsNullOrEmpty(_ws.GetValue<string>(row, (int)SubjectColumn.Index)));

            if (list.Count == 1)
            {
                var totalHours = ReadDoubleCellSafe(row, (int) SubjectColumn.TotalScore);
                if (totalHours.HasValue)
                {
                    list.ForEach(m => m.TotalComplexityHours = totalHours);
                }
                else
                {
                    _logger.Info($"Parsing curriculum {_fileInfo.Name}. Attestation doesn't have complexity hours");
                }
            }
            else
            {
                _logger.Info($"Parsing curriculum {_fileInfo.Name}. Attestation has more then one row in attestation table");
            }
            
            return list;
        }

        private IEnumerable<int> GetCommaSeparatedInt(int row, int column)
        {
            try
            {
                var s = _ws.GetValue<string>(row, column);

                if (string.IsNullOrEmpty(s))
                {
                    return null;
                }

                var semesters = s.Split(new[] {' ', ','}, StringSplitOptions.RemoveEmptyEntries);

                var lst = new List<int>(s.Length);

                foreach (var semester in semesters)
                {
                    int sem;
                    if (int.TryParse(semester, out sem))
                    {
                        lst.Add(sem);
                    }
                    else
                    {
                        _logger.Info($"Parsing curriculum {_fileInfo.Name}. Can't read numbers in {row}:{column}");
                    }
                }
                return lst;
            }
            catch (Exception e)
            {
                throw new ApplicationException($"Can't read string value {row}:{column}", e);
            }
            
        }

        private void ReadActivityHours(int row, int i, ICollection<Complexity> lections)
        {
            var column = (int) SubjectColumn.HoursTable + i;
            var lectionHours = ReadDoubleCellSafe(row, column);

            if (lectionHours.HasValue && lectionHours.Value > 0)
            {
                lections.Add(new Complexity(i + 1, lectionHours.Value));
            }
        }
    }
}
