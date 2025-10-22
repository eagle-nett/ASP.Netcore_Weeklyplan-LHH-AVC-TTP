using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using QuanLy.Data;
using QuanLy.Models;
using System.Linq;
using System.Threading.Tasks;

namespace QuanLy.Controllers
{
    [Authorize(Roles = "Admin")]
    public class PhongBanController : Controller
    {
        private readonly ApplicationDbContext _context;

        public PhongBanController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: PhongBan
        public async Task<IActionResult> Index()
        {
            var phongBans = await _context.PhongBans.ToListAsync();
            return View(phongBans);
        }

        // GET: PhongBan/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: PhongBan/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(PhongBan phongBan)
        {
            if (ModelState.IsValid)
            {
                _context.Add(phongBan);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(phongBan);
        }

        // GET: PhongBan/Edit/{id}
        public async Task<IActionResult> Edit(string id)
        {
            if (id == null)
                return NotFound();

            var phongBan = await _context.PhongBans.FindAsync(id);
            if (phongBan == null)
                return NotFound();

            return View(phongBan);
        }

        // POST: PhongBan/Edit/{id}
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(string id, [Bind("MaPhongBan,TenPhongBan,TenCongTy")] PhongBan phongBan)
        {
            if (id != phongBan.MaPhongBan)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                _context.Update(phongBan);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(phongBan);
        }


        // GET: PhongBan/Delete/{id}
        public async Task<IActionResult> Delete(string id)
        {
            var phongBan = await _context.PhongBans
                .FirstOrDefaultAsync(m => m.MaPhongBan == id);
            if (phongBan == null)
                return NotFound();

            return View(phongBan);
        }

        // POST: PhongBan/Delete/{id}
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(string id)
        {
            var phongBan = await _context.PhongBans.FindAsync(id);
            _context.PhongBans.Remove(phongBan);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }
    }
}
