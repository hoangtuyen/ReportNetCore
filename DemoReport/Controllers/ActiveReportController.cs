using Microsoft.AspNetCore.Mvc;

namespace DemoReport.Controllers
{
    public class ActiveReportController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
