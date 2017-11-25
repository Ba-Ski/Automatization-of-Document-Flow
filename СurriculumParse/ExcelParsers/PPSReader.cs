using System;
using System.Drawing;
using System.IO;
using System.Linq;
using MongoDB.Driver;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using СurriculumParse.Logger;
using СurriculumParse.Structures;

namespace СurriculumParse.ExcelParsers
{
    public class PPSReader
    {
        private readonly ILogger _logger;
        private ExcelWorksheet _ws;
        private FileInfo _fileInfo;
        private IDBManager _dbManager;

        private const int CurriculumYear = 2015;
        private const int StartRow = 6;
        private const int TableHeaderRow = 5;
        private const int IndexColumn = 2;
        private const int SubjectNameColumn = 3;
        private const int SubjectActivityTypeColumn = 4;
        private const int SemesterColumn = 7;
        private const int StepenColumn = 11;
        private const int ZwanieColumn = 12;
        private const int RateColumn = 32;
        private const string RateColumnLetters = "AF";
        private const int RateConst = 900;

        private const string LabName = "лабораторные занятия";
        private const string PracticeName = "практические занятия";
        private const string LectionName = "лекционные занятия";
        private const string StudentSelfWorkName = "срс";
        private const string RateColumnName = "доля ставки по дисциплине";

        public PPSReader(ILogger logger, IDBManager dbManager)
        {
            _logger = logger;
            _dbManager = dbManager;
        }

        public PpsReadStatus WorkWithPPS(string filePath)
        {
            _fileInfo = new FileInfo(filePath);
            var startRow = StartRow;

            try
            {
                using (var package = new ExcelPackage(_fileInfo))
                {
                    _ws = package.Workbook.Worksheets[1];

                    var specialityPair = _ws.GetValue<string>(2, 1);
                    var specialityNumber = specialityPair.Split(new[] {' '})[0];
                    var values = _ws.GetValue<string>(3, 1);
                    if (values.Contains('('))
                    {
                        var bracket = values.IndexOf('(');
                        values = values.Substring(0, bracket).Trim();
                    }
                    //var profile = values.Substring(0, values.IndexOf('(') - 1);
                    //var tuple = values.Substring(values.IndexOf('('), values.IndexOf(')'));
                    //var yeraAndForm = tuple.Split(new[] {' ', ','}, StringSplitOptions.RemoveEmptyEntries);
                    //var year = yeraAndForm[0];
                    var l = _fileInfo.Name.Length - _fileInfo.Extension.Length;
                    var parts = _fileInfo.Name.Substring(0, l).ToLower()
                        .Split(new[] {'_'}, StringSplitOptions.RemoveEmptyEntries);
                    var s = parts[parts.Length - 1];
                    var year = parts[parts.Length - 2];


                    int trueYear;
                    if (!int.TryParse(year, out trueYear))
                    {
                        _logger.Info($"Bad year {year}");
                        trueYear = 2015;
                    }
                    s = "о";
                    var edForm = ParseEdForm(s);

                    var profile = values; //ToDo toLower
                    //var key = specialityNumber + values.ToLower() + year + (int) edForm;

                    //using (var md5 = MD5.Create())
                    //{
                    //    var hash = md5.ComputeHash(Encoding.Default.GetBytes(key));
                    //    var guid = new Guid(hash);
                    //    curriculum = _dbManager.GetCurriculumAsync(guid).Result;
                    //}

                    var curriculum = _dbManager.GetCurriculumAsync(specialityNumber, profile, trueYear, (int) edForm);
                    if (curriculum == null)
                    {
                        throw new MongoException(
                            $"Can't find curriculum in data base: {specialityNumber}, {profile}, {trueYear}, {edForm}");
                    }
                    _ws.SetValue(TableHeaderRow, RateColumn, RateColumnName);

                    var ostepenionnost = 0.0d;
                    var isNotEmpty = true;
                    do
                    {
                        var index = _ws.GetValue<string>(startRow, IndexColumn);
                        //var subjName = _ws.GetValue<string>(startRow, SubjectNameColumn);
                        var activityForm = _ws.GetValue<string>(startRow, SubjectActivityTypeColumn);

                        if (string.IsNullOrEmpty(index) &&
                            string.IsNullOrEmpty(activityForm))
                        {
                            isNotEmpty = false;
                            continue;
                        }

                        if (string.IsNullOrEmpty(index))
                        {
                            _logger.Info($"Empty entry in {startRow}:{IndexColumn}");
                            startRow++;
                            continue;
                        }

                        index = index.Trim();

                        //if (string.IsNullOrEmpty(subjName))
                        //{
                        //    _logger.Info($"Empty entry in {startRow}:{SubjectNameColumn}");
                        //    startRow++;
                        //    continue;
                        //}
                        //subjName = subjName.ToLower().Trim();

                        if (string.IsNullOrEmpty(activityForm))
                        {
                            _logger.Info($"Empty entry in {startRow}:{SubjectActivityTypeColumn}");
                            startRow++;
                            continue;
                        }
                        activityForm = activityForm.ToLower().Trim();
                        var semester = ReadIntCellSafe(startRow, SemesterColumn);
                        if (semester == -1)
                        {
                            startRow++;
                            continue;
                        }

                        var subject =
                            curriculum.BaseSubjects.FirstOrDefault(
                                c => c.Index == index); // Может класть в базу сразу строчные?

                        if (subject == null)
                        {
                            _logger.Info($"Curriculum subject hasn't found: index {index}");
                            startRow++;
                            continue;
                        }

                        var ratio = 0d;
                        var hours = 0d;
                        Complexity complexity;
                        switch (activityForm)
                        {
                            case LabName:
                                complexity = subject.LabStudiesHours.FirstOrDefault(c => c.Semester == semester);
                                break;

                            case PracticeName:
                                complexity = subject.PracticeHours.FirstOrDefault(c => c.Semester == semester);
                                break;

                            case LectionName:
                                complexity = subject.LectureHours.FirstOrDefault(c => c.Semester == semester);
                                break;

                            case StudentSelfWorkName:
                                complexity = subject.StudentSelfStudyHours.FirstOrDefault(c => c.Semester == semester);
                                break;

                            default:
                                _logger.Info($"Bad activity form {activityForm}");
                                continue;
                        }

                        hours = complexity?.Hours ?? 0;
                        var weeks = curriculum.WeeksPerSemester[semester - 1].Weeks;
                        ratio = hours * weeks / 900;
                        _ws.Cells[startRow, RateColumn].Style.Numberformat.Format = "#,#####0.00000";
                        _ws.SetValue(startRow, RateColumn, ratio);

                        var stepen = _ws.GetValue<string>(startRow, StepenColumn);
                        var zwanie = _ws.GetValue<string>(startRow, ZwanieColumn);

                        if (!string.IsNullOrEmpty(stepen) || !string.IsNullOrEmpty(zwanie))
                        {
                            ostepenionnost += ratio;
                        }

                        startRow++;
                    } while (isNotEmpty);

                    _ws.Cells[startRow, RateColumn].Style.Numberformat.Format = "#,#####0.00000";
                    _ws.Cells[startRow, RateColumn].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _ws.Cells[startRow, RateColumn].Style.Fill.BackgroundColor.SetColor(Color.Red);
                    //_ws.Cells[startRow, RateColumn].Formula = $"SUM({RateColumnLetters}{StartRow}:{RateColumnLetters}{startRow - 1})";
                    _ws.Cells[startRow, RateColumn].Formula = $"SUMIF($K{StartRow}:$L{startRow - 1}, \"<>\", $AF{StartRow}:$AF{startRow - 1})";
                    //_ws.Cells[startRow, RateColumn].Value = ostepenionnost;

                    package.Workbook.Calculate();
                    package.Save();
                }

                return PpsReadStatus.Success;
            }
            catch (MongoException ex)
            {
                _logger.Error(ex.Message, ex);
                return PpsReadStatus.CurriculumNotFound;
            }
            catch (IOException ex)
            {
                return PpsReadStatus.FileOpenException;
            }
            catch (Exception ex)
            {
                _logger.Error($"Reading pps {_fileInfo.Name}. Error {startRow}: {ex.Message}", ex);
                return PpsReadStatus.PpsReadError;
            }

        }

        private EducationalForm ParseEdForm(string str)
        {
            if (str == "о")
                return EducationalForm.Intramural;
            else if (str == "оз")
            {
                return EducationalForm.IntraExtramural;
            }
            else if (str == "з")
            {
                return EducationalForm.Extramural;
            }
            else
            {
                throw new ApplicationException($"Can't define education form {str}");
            }
        }

        private int ReadIntCellSafe(int row, int cell)
        {
            var value = _ws.GetValue<string>(row, cell);
            int res;
            if (int.TryParse(value, out res))
            {
                return res;
            }
            else
            {
                _logger.Info($"Can't read int {row}:{cell}");
            }
            return -1;
        }
    }
}
