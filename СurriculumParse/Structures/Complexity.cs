namespace СurriculumParse.Structures
{
    public class Complexity
    {
        public int Semester { get; private set; }
        public double Hours { get; private set; }

        public Complexity(int semester, double hours)
        {
            Semester = semester;
            Hours = hours;
        }
    }
}
