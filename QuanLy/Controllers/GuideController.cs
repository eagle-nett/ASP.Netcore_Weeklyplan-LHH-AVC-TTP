using Microsoft.AspNetCore.Mvc;

namespace QuanLy.Controllers
{
    public class GuideController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Help()
        {
            return View();
        }
    }
}
