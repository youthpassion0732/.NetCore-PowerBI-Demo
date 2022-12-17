using Microsoft.AspNetCore.Mvc;
namespace PowerBi.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
