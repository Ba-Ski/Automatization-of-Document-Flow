using System;
using System.Collections.Generic;
using System.Linq;
using СurriculumParse.Logger;
using СurriculumParse.Structures;
using Excel = Microsoft.Office.Interop.Excel;

namespace СurriculumParse.ExcelParsers
{
    //only for xls files (excel 2003)
    //unsupported
    //internal class InteropExcelParser : IDocumentParser, IDisposable
    //{
    //    private readonly Excel.Application _xlApp;
    //    private readonly ILogger _logger;
    //    private Excel.Worksheet _worksheet;

    //    public InteropExcelParser(ILogger logger)
    //    {
    //        _logger = logger;
    //        _xlApp = new Excel.Application();
    //    }

    //    public Curriculum ParseDocumenr(string path) // ToDo return Cirriculum
    //    {

    //        Excel.Workbook workbook = null;
    //        try
    //        {
    //            if (string.IsNullOrEmpty(path))
    //            {
    //                return null;
    //            }

    //            workbook = _xlApp.Workbooks.Open(path); //ToDo complete path if it is not absolute
    //            var worksheet = workbook.Worksheets[1] as Excel.Worksheet;

    //            if (worksheet == null)
    //            {
    //                _logger.Info("Wrong file name");
    //                return null;
    //            }

    //            _worksheet = worksheet;
    //            var usedRange = worksheet.UsedRange;

    //            var specialityField = usedRange.Find(Constants.SpecialityName, Type.Missing, // in every documents, cells are the same
    //                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,                  // that's why find is not needed
    //                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
    //                Type.Missing, Type.Missing);


    //            var speciality =
    //                Convert.ToString(usedRange.Cells[specialityField.Row + 1, specialityField.Column].Value2);
    //            var profile = Convert.ToString(usedRange.Cells[specialityField.Row + 3, specialityField.Column].Value2);

    //            var basePart = usedRange.Find(Constants.BasePartName, Type.Missing,
    //                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
    //                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
    //                Type.Missing, Type.Missing);

    //            var variativePart = usedRange.Find(Constants.VariativePartName, Type.Missing,
    //                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
    //                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
    //                Type.Missing, Type.Missing);

    //            var row = basePart.Row + 2;
    //            var lastRow = variativePart.Row - 3; //ToDo Try to find block end another way


    //            var baseS = ParseSubjects(row, lastRow, usedRange, SubjectType2015.BaseSubject);

    //            var physicalPart = usedRange.Find(Constants.PhysicalEdPartName, Type.Missing,
    //                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
    //                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
    //                Type.Missing, Type.Missing);

    //            row = variativePart.Row + 2;
    //            lastRow = physicalPart.Row - 7; //ToDo Try to find block end another way

    //            var variativeS = ParseSubjects(row, lastRow, usedRange, SubjectType2015.VariativeSubject);

    //            var practicalPart = usedRange.Find(Constants.PracticePartName, Type.Missing,
    //                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
    //                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
    //                Type.Missing, Type.Missing);

    //            row = physicalPart.Row + 2;
    //            lastRow = practicalPart.Row - 2;

    //            var physicalS = ParseSubjects(row, lastRow, usedRange, SubjectType2015.PhysicalEd);

    //            row = practicalPart.Row + 2;

    //            var practiceS = ParsePracticeAttestation(row, usedRange);

    //            var attestaionPart = usedRange.Find(Constants.AttestationPartName, Type.Missing,
    //                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
    //                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
    //                Type.Missing, Type.Missing);

    //            row = attestaionPart.Row + 2;

    //            var attestationS = ParsePracticeAttestation(row, usedRange);

    //            var curriculum = new Curriculum(speciality, speciality, profile, 4, baseS, 2015, EducationalForm.Intramural); //it doesn't work

    //            ReleaseRcm(usedRange);

    //            return curriculum;
    //        }
    //        catch (Exception ex)
    //        {
    //            _logger.Error($"[EXCEL PARSING ERROR. FILE {path}]", ex);
    //            return null;
    //        }
    //        finally
    //        {
    //            if (workbook != null)
    //                ReleaseRcm(workbook);
    //        }
    //    }


    //    private static IEnumerable<Subject> ParseSubjects(int firstRow, int lastRow, Excel.Range range, SubjectType2015 type)
    //    {
    //        var row = firstRow;
    //        var subjects = new List<Subject>();

    //        do
    //        {
    //            var index = Convert.ToString(range.Cells[row, SubjectColumn.Index].Value2);
    //            var names = GetFullString(range, row, (int) SubjectColumn.Name, Constants.SubjectRows);
    //            if (names.All(string.IsNullOrEmpty))
    //            {
    //                row += 6;
    //                continue;
    //            }
    //            var exams = GetSemesters(range, row, (int) SubjectColumn.ExamsSemesters);
    //            var credits = GetSemesters(range, row, (int) SubjectColumn.CreditSemesters);

    //            var lections = new List<Complexity>();
    //            var labs = new List<Complexity>();
    //            var practices = new List<Complexity>();
    //            var selfStudy = new List<Complexity>();

    //            for (var i = 0; i < Constants.DefaultSemesterCont; i++)
    //            {
    //                ReadActivityHours(range, row + (int) ActivitiesRows.Lections, i, lections);
    //                ReadActivityHours(range, row + (int) ActivitiesRows.Labs, i, labs);
    //                ReadActivityHours(range, row + (int) ActivitiesRows.Practices, i, practices);
    //                ReadActivityHours(range, row + (int) ActivitiesRows.SelfStudy, i, selfStudy);
    //            }
    //            var department = GetFullString(range, row, (int) SubjectColumn.Department, Constants.SubjectRows);

    //            subjects.AddRange(names.Select(name =>
    //                new Subject(index, name, exams, credits, lections, practices, labs, selfStudy, department, type)));

    //            row += 6;
    //        } while (row < lastRow);

    //        return subjects;
    //    }

    //    private static IEnumerable<Practice> ParsePracticeAttestation(int row, Excel.Range range)
    //    {

    //        var list = new List<Practice>();

    //        do
    //        {
    //            var index = Convert.ToString(range.Cells[row, SubjectColumn.Index].Value2);
    //            var name = GetFullString(range, row, (int) SubjectColumn.Name, 1);
    //            var exams = new List<int>();
    //            var exam = ReadIntCell(range, row, (int) SubjectColumn.ExamsSemesters); //ToDo another function to paarse numbers with zapyataya delimiter
    //            if (exam.HasValue)
    //            {
    //                exams.Add(exam.Value);
    //            }
    //            var credits = new List<int>();
    //            var credit = ReadIntCell(range, row, (int) SubjectColumn.CreditSemesters);
    //            if (credit.HasValue)
    //            {
    //                credits.Add(credit.Value);
    //            }

    //            var hoursNullable = ReadDoubleCell(range, row + 3, (int) SubjectColumn.TotalScore);
    //            var hours = 0d;
    //            if (hoursNullable.HasValue && Math.Abs(hoursNullable.Value) > 0)
    //            {
    //                hours = hoursNullable.Value;
    //            }
    //            {
    //                //ToDo log it
    //            }

    //            list.Add(new Practice(index, name, exams, credits, hours));
    //            row++;

    //        } while (!string.IsNullOrEmpty(Convert.ToString(range.Cells[row, SubjectColumn.Index].Value2)));

    //        return list;
    //    }

    //    private static void ReadActivityHours(Excel.Range range, int row, int i, ICollection<Complexity> lections)
    //    {
    //        var lectionHours = ReadDoubleCell(range, row, (int) SubjectColumn.HoursTable + i);
    //        //var lectionHours = (double?)(range.Cells[row, SubjectColumn.HoursTable + i] as Excel.Range)?.Value2;
    //        if (lectionHours.HasValue && Math.Abs(lectionHours.Value) > 0)
    //        {
    //            lections.Add(new Complexity(i + 1, lectionHours.Value));
    //        }
    //    }

    //    private static double? ReadDoubleCell(Excel.Range range, int row, int cell)
    //    {
    //        var value = Convert.ToString(range.Cells[row, cell]?.Value2);
    //        double res;
    //        if (double.TryParse(value, out res))
    //        {
    //            return res;
    //        }
    //        return null;
    //    }

    //    private static int? ReadIntCell(Excel.Range range, int row, int cell)
    //    {
    //        var value = Convert.ToString(range.Cells[row, cell]?.Value2);
    //        int res;
    //        if (int.TryParse(value, out res))
    //        {
    //            return res;
    //        }
    //        return null;
    //    }

    //    private static IEnumerable<int> GetSemesters(Excel.Range range, int row, int column)
    //    {
    //        var exams = new List<int>(); // ToDo may be return null?
    //        for (var i = 0; i < Constants.SubjectRows; i++)
    //        {
    //            var value = ReadIntCell(range, row + i, column);
    //            //var value2 = (range.Cells[row + i, column] as Excel.Range)?.Value2;
    //            if (value.HasValue)
    //            {
    //                exams.Add(value.Value);
    //            }
    //        }
    //        return exams;
    //    }

    //    private static IEnumerable<string> GetFullString(Excel.Range range, int row, int column, int height)
    //    {
    //        var strs = new List<string>();
    //        var res = new List<string>();
    //        for (var i = 0; i < height; i++)
    //        {
    //            var cell = range.Cells[row + i, column];
    //            var str = Convert.ToString(cell.Value2);

    //            var border = cell.Borders[Excel.XlBordersIndex.xlEdgeBottom];

    //            if (string.IsNullOrEmpty(str) && border.LineStyle == (int)Excel.XlLineStyle.xlLineStyleNone)
    //            {
    //                continue;
    //            }

    //            strs.Add(str);

    //            if (border.LineStyle != (int) Excel.XlLineStyle.xlLineStyleNone)
    //            {
    //                res.Add(string.Join(" ", strs));
    //                strs.Clear();
    //            }
    //        }

    //        return res;
    //    }

    //    private static void ReleaseRcm(object o)
    //    {
    //        try
    //        {
    //            System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
    //        }
    //        catch
    //        {
    //            // ignored
    //        }
    //        finally
    //        {
    //            o = null;
    //        }
    //    }

    //    private void ReleaseUnmanagedResources()
    //    {
    //        ReleaseRcm(_xlApp);
    //    }

    //    public void Dispose()
    //    {
    //        ReleaseUnmanagedResources();
    //        GC.SuppressFinalize(this);
    //    }

    //    ~InteropExcelParser()
    //    {
    //        ReleaseUnmanagedResources();
    //    }
    //}

}
