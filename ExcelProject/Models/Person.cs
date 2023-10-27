using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ExcelProject.Models
{
    public class Person
    {
        [Required]
        [Display (Name= "Id değerini girmek zorunludur")]
        public int Id { get; set; }
        public string Name { get; set; }
    }
}