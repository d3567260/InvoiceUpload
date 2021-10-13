using System.ComponentModel.DataAnnotations;

namespace InvoiceUpload.Dtos
{
    public class LoginDto
    {
        [Required(ErrorMessage = "手機號碼為必填")]
        [RegularExpression(@"^[0-9]{10}$", ErrorMessage = "手機號碼格式不正確")]
        public string Phone { get; set; }
        [Required(ErrorMessage = "密碼為必填")]
        [StringLength(20, MinimumLength = 8, ErrorMessage = "密碼長度為8至20個字元")]
        public string Password { get; set; }
    }
}