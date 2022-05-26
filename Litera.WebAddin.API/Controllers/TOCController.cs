using Aspose.Words;
using Litera.WebAddin.API.Models;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Text;

namespace Litera.WebAddin.API.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class TOCController : Controller
    {
        [HttpPost("addtoc")]
        public JsonResult AddToc([FromBody] Word word)
        {
            if (word.OOXML == null)
            {
                return Json(new { ooxml = string.Empty });
            }
            MemoryStream mStrm = new MemoryStream(Encoding.UTF8.GetBytes(word.OOXML));
            Document doc = new Document(mStrm);
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            doc.UpdateFields();
            string filepath = Path.GetTempPath() + "output2.xml";
            doc.Save(filepath);
            var output = System.IO.File.ReadAllText(filepath);
            if (System.IO.File.Exists(filepath))
                System.IO.File.Delete(filepath);
            return Json(new { ooxml = output });
        }
    }

}
