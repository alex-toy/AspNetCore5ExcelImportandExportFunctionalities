using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using SampleExcel.Models;
using System.Collections.Generic;

namespace SampleExcel.Controllers
{
    public class UserController : Controller
    {
        // GET: UserController
        public ActionResult Index()
        {
            var users = GetUsers();
            return View(users);
        }

        private List<User> GetUsers()
        {
            var users = new List<User>()
            {
                new User(){ Name = "alex", Email = "alex@test.fr", Phone = "8676875976"},
                new User(){ Name = "seb", Email = "seb@test.fr", Phone = "8676959867"},
                new User(){ Name = "kate", Email = "kate@test.fr", Phone = "345677808"},
                new User(){ Name = "jule", Email = "jule@test.fr", Phone = "776554433"},
            };

            return users;
        }

        // GET: UserController/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: UserController/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: UserController/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: UserController/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: UserController/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: UserController/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: UserController/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }
    }
}
