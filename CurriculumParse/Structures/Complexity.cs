namespace СurriculumParse.Structures
{
    public class Complexity
    {
        /// <summary>
        /// Номер семестра
        /// </summary>
        public int Semester { get; private set; }

        /// <summary>
        /// Трудоемкость в часах. Для шаблона 2015 года трудоемкость выражается в часах в неделю.
        /// </summary>
        public double Hours { get; private set; }

        public Complexity(int semester, double hours)
        {
            Semester = semester;
            Hours = hours;
        }
    }
}
