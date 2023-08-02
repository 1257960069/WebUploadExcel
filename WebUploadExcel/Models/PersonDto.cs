using EPPlus.Core.Extensions.Attributes;
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations;

namespace WebUploadExcel
{
    [ExcelWorksheet(nameof(PersonDto))]
    public class PersonDto
    {
        [ExcelTableColumn("First name")]
        [Required(ErrorMessage = "First name cannot be empty.")]
        [MaxLength(50, ErrorMessage = "First name cannot be more than {1} characters.")]
        public string FirstName { get; set; }

        [ExcelTableColumn(ColumnName = "Last name", IsOptional = true)]
        public string LastName { get; set; }

        [ExcelTableColumn(3)]
        [Range(1900, 2050, ErrorMessage = "Please enter a value bigger than {1}")]
        public int YearBorn { get; set; }

        public decimal NotMapped { get; set; }

        [ExcelTableColumn(IsOptional = true)]
        public decimal OptionalColumn1 { get; set; }

        [ExcelTableColumn(ColumnIndex = 999, IsOptional = true)]
        public decimal OptionalColumn2 { get; set; }
    }
}