using EasyExcelGenerator.Models;
using System.Drawing;

namespace ApiApp;


public class User
{
    public int Id { get; set; }

    public string? FullName { get; set; }

    public string? PersonnelCode { get; set; }

    public string? Nationality { get; set; }
}