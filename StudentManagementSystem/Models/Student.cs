using System.ComponentModel.DataAnnotations;

namespace StudentManagementSystem.Models
{
    public class Student
    {
        public int Id { get; set; }

        public string Name { get; set; } = "";
        public string Email { get; set; } = "";
        public string Course { get; set; } = "";

    }
}
