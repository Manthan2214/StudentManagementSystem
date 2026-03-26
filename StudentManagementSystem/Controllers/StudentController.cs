using Microsoft.AspNetCore.Mvc;
using StudentManagementSystem.Data;
using StudentManagementSystem.Models;
using System.Linq;
using ClosedXML.Excel;
using System.IO;

namespace StudentManagementSystem.Controllers
{
    public class StudentController : Controller
    {
        private readonly AppDbContext _context;

        public StudentController(AppDbContext context)
        {
            _context = context;
        }
        // SEARCH INDEX
        public IActionResult Index(string search)
        {
            var students = _context.Students.AsQueryable();
            if (!string.IsNullOrEmpty(search))
            {
                students = students.Where(s =>
                s.Name.Contains(search) ||
                s.Email.Contains(search) ||
                s.Course.Contains(search));
            }
            return View(students.ToList());
        }

        // CREATE (GET)
        public IActionResult Create()
        {
            return View();
        }

        // CREATE (POST)
        [HttpPost]
        public IActionResult Create(Student student)
        {
            if (ModelState.IsValid)
            {
                _context.Students.Add(student);
                _context.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(student);
        }

        // UPDATE (GET)
        public IActionResult Edit(int id)
        {
            var student = _context.Students.Find(id);
            if (student == null) return NotFound();
            return View(student);
        }

        // UPDATE (POST)
        [HttpPost]
        public IActionResult Edit(Student student)
        {
            _context.Students.Update(student);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }

        // DELETE (GET)
        public IActionResult Delete(int id)
        {
            var student = _context.Students.Find(id);
            if (student == null) return NotFound();
            return View(student);
        }

        // DELETE (POST)
        [HttpPost]
        public IActionResult DeleteConfirmed(int id)
        {
            var student = _context.Students.Find(id);
            if (student == null) return NotFound();
            _context.Students.Remove(student);
            _context.SaveChanges();
            return RedirectToAction("Index");
        }


        // EXPORT METHOD
        public IActionResult ExportToExcel()
        {
            var students = _context.Students.ToList();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Students");

                worksheet.Cell(1, 1).Value = "Name";
                worksheet.Cell(1, 2).Value = "Email";
                worksheet.Cell(1, 3).Value = "Course";

                for (int i = 0; i < students.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = students[i].Name;
                    worksheet.Cell(i + 2, 2).Value = students[i].Email;
                    worksheet.Cell(i + 2, 3).Value = students[i].Course;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Students.xlsx");
                }
            }
        }
    }
}
