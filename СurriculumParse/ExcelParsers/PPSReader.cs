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
        private readonly IDBManager _dbManager;

        private const int StartRow = 6;
        private const int TableHeaderRow = 5;

        private const int RateConst = 900;

        private const string LabName = "лабораторные занятия";
        private const string PracticeName = "практические занятия";
        private const string LectionName = "лекционные занятия";
        private const string StudentSelfWorkName = "срс";
        private const string RateColumnName = "Доля ставки по дисциплине";

        public PPSReader(ILogger logger, IDBManager dbManager)
        {
            _logger = logger;
            _dbManager = dbManager;
        }
        /// <summary>
        /// Производит работу с ППС по расчёту доли ставки и остепенённости
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public PpsReadStatus WorkWithPPS(string filePath)
        {
            _fileInfo = new FileInfo(filePath);
            var startRow = StartRow;

            try
            {
                using (var package = new ExcelPackage(_fileInfo))
                {
                    _ws = package.Workbook.Worksheets[1];

                    var curriculum = GetCurriculum();
                    _ws.SetValue(TableHeaderRow, (int) PPSColumnsNumbers.Rate, RateColumnName);

                    if (curriculum == null)
                    {
                        throw new ArgumentNullException(
                            $"Can't find curriculum in data base");
                    }

                    startRow = WriteRate(startRow, curriculum);

                    WriteOstepienionnost(startRow);

                    package.Workbook.Calculate();
                    package.Save();
                }

                return PpsReadStatus.Success;
            }
            catch (ArgumentNullException ex)
            {
                _logger.Error(ex.Message, ex);
                return PpsReadStatus.CurriculumNotFound;
            }
            catch (IOException ex)
            {
                _logger.Error(ex.Message, ex);
                return PpsReadStatus.FileOpenException;
            }
            catch (Exception ex)
            {
                _logger.Error($"Reading pps {_fileInfo.Name}. Error {startRow}: {ex.Message}", ex);
                return PpsReadStatus.PpsReadError;
            }

        }

        /// <summary>
        /// Производит рассчёт доли ставки для каждой строки в ППС, для которой удалось извлечь учебный план из базы.
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="curriculum"></param>
        /// <returns></returns>
        private int WriteRate(int startRow, Curriculum curriculum)
        {
            if (curriculum == null)
            {
                return startRow;
            }

            var isNotEmpty = true;
            do
            {
                var index = _ws.GetValue<string>(startRow, (int) PPSColumnsNumbers.Index);
                var activityForm = _ws.GetValue<string>(startRow, (int) PPSColumnsNumbers.SubjectActivityType);

                //Если в обрабатываемой строке и индекс и вид заняти отсутствуют,
                //значит строка пустая, и обработку таблицы можно закончить
                if (string.IsNullOrEmpty(index) &&
                    string.IsNullOrEmpty(activityForm))
                {
                    isNotEmpty = false;
                    continue;
                }

                //обработка ошибок
                if (string.IsNullOrEmpty(index))
                {
                    _logger.Info($"Empty entry in {startRow}:{(int) PPSColumnsNumbers.Index}");
                    startRow++;
                    continue;
                }

                index = index.Trim();

                if (string.IsNullOrEmpty(activityForm))
                {
                    _logger.Info($"Empty entry in {startRow}:{(int) PPSColumnsNumbers.SubjectActivityType}");
                    startRow++;
                    continue;
                }

                var semester = ReadIntCellSafe(startRow, (int) PPSColumnsNumbers.Semester);
                if (semester == -1)
                {
                    startRow++;
                    continue;
                }

                var subject =
                    curriculum.BaseSubjects.FirstOrDefault(
                        c => c.Index == index);

                if (subject == null)
                {
                    _logger.Info($"Curriculum subject hasn't found: index {index}");
                    startRow++;
                    continue;
                }

                //Расчёт доли ставки для предмета, по виду занятий, за определённый семестр 
                var ratio = 0d;
                var hours = 0d;
                var complexity = GetComplexity(activityForm, subject, semester);
                
                hours = complexity?.Hours ?? 0;
                var weeks = curriculum.WeeksPerSemester[semester - 1].Weeks;
                ratio = hours * weeks / RateConst; //ToDo: только для УП 2015 года. Нужно юудет учесть и другие шаблоны.
                _ws.Cells[startRow, (int) PPSColumnsNumbers.Rate].Style.Numberformat.Format = "#,#####0.00000";
                _ws.SetValue(startRow, (int) PPSColumnsNumbers.Rate, ratio);

                startRow++;
            } while (isNotEmpty);
            return startRow;
        }

        /// <summary>
        /// Вставляет в конец таблицы, в столбец с долей ставки, excel-формулу для рассчёта остепенённости. 
        /// </summary>
        /// <param name="startRow"></param>
        private void WriteOstepienionnost(int startRow)
        {
            _ws.Cells[startRow, (int) PPSColumnsNumbers.Rate].Style.Numberformat.Format = "#,#####0.00000";
            _ws.Cells[startRow, (int) PPSColumnsNumbers.Rate].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _ws.Cells[startRow, (int) PPSColumnsNumbers.Rate].Style.Fill.BackgroundColor.SetColor(Color.Red);
            _ws.Cells[startRow, (int) PPSColumnsNumbers.Rate].Formula =
                $"SUMIF($K{StartRow}:$L{startRow - 1}, \"<>\", $AF{StartRow}:$AF{startRow - 1})";
        }

        /// <summary>
        /// Возвращает трудоемкость дисциплины.
        /// </summary>
        /// <param name="activityForm">Вид занятий: лабы, практики, лекции, срс.</param>
        /// <param name="subject"> Предмет.</param>
        /// <param name="semester">Семестр, в котором осуществляется преподавание.</param>
        /// <returns></returns>
        private Complexity GetComplexity(string activityForm, Subject subject, int semester)
        {
            Complexity complexity;
            activityForm = activityForm.ToLower().Trim();
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
                    return null;
            }
            return complexity;
        }

        /// <summary>
        /// Получает УП для ППС, с которым идёт работа.
        /// </summary>
        /// <returns></returns>
        private Curriculum GetCurriculum()
        {
            //Получаем профиль и номер специальности из шапки.
            var specialityPair = _ws.GetValue<string>(2, 1);
            var specialityNumber = specialityPair.Split(new[] {' '})[0];
            var profile = _ws.GetValue<string>(3, 1);
            if (profile.Contains('('))
            {
                var bracket = profile.IndexOf('(');
                profile = profile.Substring(0, bracket).Trim();
            }

            //Получаем год и форму обучения из имени файла.
            var l = _fileInfo.Name.Length - _fileInfo.Extension.Length;
            var parts = _fileInfo.Name.Substring(0, l).ToLower()
                .Split(new[] {'_'}, StringSplitOptions.RemoveEmptyEntries);
            var s = parts[parts.Length - 1];
            var yearStr = parts[parts.Length - 2];

            int year;
            if (!int.TryParse(yearStr, out year))
            {
                _logger.Info($"Bad year {yearStr}");
                return null;
            }
            //s = "о";
            var edForm = ParseEdForm(s);

            var curriculum = _dbManager.GetCurriculumAsync(specialityNumber, profile, year, (int) edForm);
            return curriculum;
        }

        /// <summary>
        /// Возвращает форму обучения исходя из сокращения, используемого в именовании ППС-файлов.
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private static EducationalForm ParseEdForm(string str)
        {
            if (str == "о")
            {
                return EducationalForm.Intramural;
            }
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
