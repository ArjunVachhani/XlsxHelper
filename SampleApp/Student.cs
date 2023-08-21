namespace SampleApp;

public class Student
{
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public string? Grade { get; set; }
    public Marks? Marks { get; set; }
}

public class Marks
{
    public int Mathematics { get; set; }
    public int Physics { get; set; }
    public int Biology { get; set; }
    public int Chemistry { get; set; }
}
