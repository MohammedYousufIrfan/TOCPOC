using Aspose.Words;
using Litera.WebAddin.API.Models;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Text;

namespace Litera.WebAddin.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TOCController : Controller
    {
        [HttpPost("addtoc")]
        public JsonResult AddToc([FromBody] Word word)
        {
            if (word.OOXML == null)
            {
                return Json(new { data = string.Empty });
            }
           
            MemoryStream mStrm = new MemoryStream(Encoding.UTF8.GetBytes(word.OOXML));
            Document doc = new Document(mStrm);
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            doc.UpdateFields();
            string filepath = Path.GetTempPath() + "output2.xml";
            doc.Save(filepath);
            var output = System.IO.File.ReadAllText(filepath);
            return Json(new { data = output });
        }
    }

}
