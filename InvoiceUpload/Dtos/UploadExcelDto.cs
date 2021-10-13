using Microsoft.AspNetCore.Http;

namespace InvoiceUpload.Dtos
{
    public class UploadExcelDto
    {
        public IFormFile FormFile { get; set; }
    }
}