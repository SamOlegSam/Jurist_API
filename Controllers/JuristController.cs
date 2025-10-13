using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Jurist.Models;
using Jurist.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Diagnostics.Metrics;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq.Dynamic.Core;
using System.Security.Cryptography;
using System.Text;
using static System.Runtime.InteropServices.JavaScript.JSType;
using ClosedXML.Excel;
using Microsoft.DotNet.Scaffolding.Shared.Messaging;

namespace Jurist.Controllers
{
    public class md5
    {
        public static string hashPassword(string password)
        {
            MD5 md5 = MD5.Create();
            byte[] b = Encoding.ASCII.GetBytes(password);
            byte[] hash = md5.ComputeHash(b);

            StringBuilder sb = new StringBuilder();
            foreach (var a in hash)
            {
                sb.Append(a.ToString("X2"));
            }
            return Convert.ToString(sb);
        }
    }

    [Route("api/[controller]")]
    [ApiController]
    public class JuristController : Controller
    {
        [HttpPost]
        [Route("[action]")]
        public IActionResult Index()
        {
            return View();
        }

        public JuristContext db;
        private readonly TokenService _tokenService;

        // Объединенный конструктор
        public JuristController(JuristContext context, TokenService tokenService)
        {
            db = context;
            _tokenService = tokenService;
        }
        //--Здесь будут методы для справочников------------------------------------------------
        [HttpPost]
        [Route("[action]")]
        public IActionResult Sotrudniki()
        {            
            List<Employee> listsotrudniki = new List<Employee>();
            listsotrudniki = db.Employees.OrderBy(l => l.LastName).ToList();
            //--------------------------------------------------------------
            List<Doljnost> listdoljnost = new List<Doljnost>();
            listdoljnost = db.Doljnosts.OrderBy(l => l.Name).ToList();
            //--------------------------------------------------------------
            List<Filial> listfilial = new List<Filial>();
            listfilial = db.Filials.OrderBy(l => l.Name).ToList();

            //-----------Объединяем списки филиалов и отделов----------------------------------------------------
            var listEmployee = from emp in listsotrudniki
                         join dolj in listdoljnost on emp.DoljId equals dolj.DoljId into doljGroup
                         from doljnost in doljGroup.DefaultIfEmpty()

                         join fil in listfilial on emp.FilId equals fil.FilId into filGroup
                         from filial in filGroup.DefaultIfEmpty()

                         select new
                         {
                             EmplId = emp.EmplId,
                             LastName = emp.LastName,
                             FirstName = emp.FirstName,
                             MiddleName = emp.MiddleName,
                             DoljId = emp.DoljId,
                             Doljnost = doljnost?.Name ?? "",
                             FilId = emp.FilId,
                             Filial = filial?.Name ?? "",
                             Primechanie = emp.Primechanie,
                             UserMod = emp.UserMod,
                             DateMod = emp.DateMod
                         };
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            //---Проверим Admin == 1 или нет и выведем соответствующие списки сотрудников------
            if (admin == 1)
            {
                return Ok(listEmployee);
            }
            else
            {
                return Ok(listEmployee.Where(h => h.FilId == filialId));
            }                             
        }
        //--Список филиалов------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult Filials()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            List<Filial> listfilial = new List<Filial>();
            if (admin == 1)
            {
                listfilial = db.Filials.OrderBy(l => l.Name).ToList();
            }
            else
            {
                listfilial = db.Filials.Where(i => i.FilId == filialId).OrderBy(l => l.Name).ToList();
            }
            ;

            return Ok(listfilial);
        }
        //--Список должностей------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult Doljnost()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            List<Doljnost> listdoljnost = new List<Doljnost>();           
            listdoljnost = db.Doljnosts.OrderBy(l => l.Name).ToList();  
            return Ok(listdoljnost);
        }
        //--Добавление сотрудника------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddSotrudnik([FromBody] Employee responsible)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Employee res = new Employee();
                res.LastName = responsible.LastName;
                res.FirstName = responsible.FirstName;
                res.MiddleName = responsible.MiddleName;
                res.DoljId = responsible.DoljId;
                res.FilId = responsible.FilId;
                res.UserMod = username;
                res.DateMod = DateTime.Now;
                db.Employees.Add(res);
                db.SaveChanges();
                return Ok(res);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Редактирование сотрудника------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditSotrudnik([FromBody] Employee responsible)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Employee resp = new Employee();
                resp = db.Employees.FirstOrDefault(s => s.EmplId == responsible.EmplId);
                
                resp.LastName = responsible.LastName;
                resp.FirstName = responsible.FirstName;
                resp.MiddleName = responsible.MiddleName;
                resp.DoljId = responsible.DoljId;
                resp.FilId = responsible.FilId;
                resp.UserMod = username;
                resp.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(resp);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление сотрудника------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteSotrudnik([FromBody] Employee responsible)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Employee respon = new Employee();
                respon = db.Employees.FirstOrDefault(s => s.EmplId == responsible.EmplId);
                db.Employees.Remove(respon);
                db.SaveChanges();
                return Ok(respon);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Добавление филиала------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddFilial([FromBody] Filial filial)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Filial fil = new Filial();
                fil.Name = filial.Name;
                fil.NameFull = filial.NameFull;                
                fil.UserMod = username;
                fil.DateMod = DateTime.Now;
                db.Filials.Add(fil);
                db.SaveChanges();
                return Ok(fil);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Редактирование филиала------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditFilial([FromBody] Filial filial)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Filial fil = new Filial();
                fil = db.Filials.FirstOrDefault(s => s.FilId == filial.FilId);

                fil.Name = filial.Name;
                fil.NameFull = filial.NameFull;                
                fil.UserMod = username;
                fil.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(fil);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление филиала------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteFilial([FromBody] Filial filial)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Filial fil = new Filial();
                fil = db.Filials.FirstOrDefault(s => s.FilId == filial.FilId);
                db.Filials.Remove(fil);
                db.SaveChanges();
                return Ok(fil);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Добавление должности------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddDoljnost([FromBody] Doljnost doljnost)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Doljnost dolj = new Doljnost();
                dolj.Name = doljnost.Name;
                dolj.Primechanie = doljnost.Primechanie;
                dolj.UserMod = username;
                dolj.DateMod = DateTime.Now;
                db.Doljnosts.Add(dolj);
                db.SaveChanges();
                return Ok(dolj);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Редактирование должности------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditDoljnost([FromBody] Doljnost doljnost)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Doljnost dolj = new Doljnost();
                dolj = db.Doljnosts.FirstOrDefault(s => s.DoljId == doljnost.DoljId);

                dolj.Name = doljnost.Name;
                dolj.Primechanie = doljnost.Primechanie;
                dolj.UserMod = username;
                dolj.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(dolj);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление должности------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteDoljnost([FromBody] Doljnost doljnost)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Doljnost dolj = new Doljnost();
                dolj = db.Doljnosts.FirstOrDefault(s => s.DoljId == doljnost.DoljId);
                db.Doljnosts.Remove(dolj);
                db.SaveChanges();
                return Ok(dolj);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--Список валют------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult Valuta()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            List<Valutum> listval = new List<Valutum>();
            listval = db.Valuta.OrderBy(l => l.Name).ToList();
            return Ok(listval);
        }
        //--------------------------------------------------------------------------------------------------------
        //--Добавление валюты------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddValuta([FromBody] Valutum valuta)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Valutum val = new Valutum();
                val.Name = valuta.Name;
                val.CodeVal = valuta.CodeVal;
                val.NameFull = valuta.NameFull;
                val.UserMod = username;
                val.DateMod = DateTime.Now;
                db.Valuta.Add(val);
                db.SaveChanges();
                return Ok(val);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }
        
        //--Редактирование валюты------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditValuta([FromBody] Valutum valuta)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Valutum val = new Valutum();
                val = db.Valuta.FirstOrDefault(s => s.ValId == valuta.ValId);

                val.Name = valuta.Name;
                val.CodeVal = valuta.CodeVal;
                val.NameFull = valuta.NameFull;
                val.UserMod = username;
                val.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(val);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление валюты------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteValuta([FromBody] Valutum valuta)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Valutum val = new Valutum();
                val = db.Valuta.FirstOrDefault(s => s.ValId == valuta.ValId);
                db.Valuta.Remove(val);
                db.SaveChanges();
                return Ok(val);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //-------------------------------------------------------------------------------------------------------
        //--Список стран------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult Country()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            List<Country> listcountry = new List<Country>();
            listcountry = db.Countries.OrderBy(l => l.Name).ToList();
            return Ok(listcountry);
        }
        //--------------------------------------------------------------------------------------------------------
        //--Добавление страны------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddCountry([FromBody] Country country)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Country count = new Country();
                count.Name = country.Name;
                count.NameFull = country.NameFull;
                count.UserMod = username;
                count.DateMod = DateTime.Now;
                db.Countries.Add(count);
                db.SaveChanges();
                return Ok(count);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }

        //--Редактирование страны------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditCountry([FromBody] Country country)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Country count = new Country();
                count = db.Countries.FirstOrDefault(s => s.CountryId == country.CountryId);

                count.Name = country.Name;
                count.NameFull = country.NameFull;
                count.UserMod = username;
                count.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(count);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление страны------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteCountry([FromBody] Country country)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Country count = new Country();
                count = db.Countries.FirstOrDefault(s => s.CountryId == country.CountryId);
                db.Countries.Remove(count);
                db.SaveChanges();
                return Ok(count);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--------------------------------------------------------------------------------------------------
        //--Список городов------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult City()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            List<Country> listcountry = new List<Country>();
            listcountry = db.Countries.ToList();

            List<Models.City> listcity = new List<Models.City>();
            listcity = db.Cities.OrderBy(l => l.Name).ToList();            

            //Объеденим эти списки
            var citiesCountries = (from city in listcity
                                       join country in listcountry
                                       on city.CountryId equals country.CountryId
                                       select new
                                       {
                                           CityId = city.CityId,
                                           Name = city.Name,
                                           CountryId = city.CountryId,
                                           CountryName = country.Name,
                                           CountryFullName = country.NameFull,                                           
                                       }).OrderBy(x => x.Name).ToList();
            return Ok(citiesCountries);
        }
        //--------------------------------------------------------------------------------------------------------
        //--Добавление города------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddCity([FromBody] Models.City city)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Models.City cit = new Models.City();
                cit.Name = city.Name;
                cit.CountryId = city.CountryId;
                cit.Primechanie = city.Primechanie;
                cit.UserMod = username;
                cit.DateMod = DateTime.Now;
                db.Cities.Add(cit);
                db.SaveChanges();
                return Ok(cit);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }

        //--Редактирование города------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditCity([FromBody] Models.City city)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Models.City cit = new Models.City();
                cit = db.Cities.FirstOrDefault(s => s.CityId == city.CityId);

                cit.Name = city.Name;
                cit.CountryId = city.CountryId;
                cit.Primechanie = city.Primechanie;
                cit.UserMod = username;
                cit.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(cit);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление города------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteCity([FromBody] Models.City city)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Models.City cit = new Models.City();
                cit = db.Cities.FirstOrDefault(s => s.CityId == city.CityId);
                db.Cities.Remove(cit);
                db.SaveChanges();
                return Ok(cit);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--Список предметов------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult Predmet()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            List<Predmet> listpredmet = new List<Predmet>();
            listpredmet = db.Predmets.OrderBy(l => l.Predmet1).ToList();
            return Ok(listpredmet);
        }
        //--------------------------------------------------------------------------------------------------------
        //--Добавление предмета------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddPredmet([FromBody] Predmet predmet)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Predmet pred = new Predmet();
                pred.Predmet1 = predmet.Predmet1;
                pred.Primechanie1 = predmet.Primechanie1;
                pred.Primechanie2 = predmet.Primechanie2;
                pred.Primechanie3 = predmet.Primechanie3;
                pred.UserMod = username;
                pred.DateMod = DateTime.Now;
                db.Predmets.Add(pred);
                db.SaveChanges();
                return Ok(pred);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }

        //--Редактирование предмета------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditPredmet([FromBody] Predmet predmet)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Predmet pred = new Predmet();
                pred = db.Predmets.FirstOrDefault(s => s.PredmetId == predmet.PredmetId);

                pred.Predmet1 = predmet.Predmet1;
                pred.Primechanie1 = predmet.Primechanie1;
                pred.Primechanie2 = predmet.Primechanie2;
                pred.Primechanie3 = predmet.Primechanie3;
                pred.UserMod = username;
                pred.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(pred);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление предмета------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeletePredmet([FromBody] Predmet predmet)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Predmet pred = new Predmet();
                pred = db.Predmets.FirstOrDefault(s => s.PredmetId == predmet.PredmetId);
                db.Predmets.Remove(pred);
                db.SaveChanges();
                return Ok(pred);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }        
        //--Список статусов------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult Status()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            List<Status> liststatus = new List<Status>();
            liststatus = db.Statuses.OrderBy(l => l.Name).ToList();
            return Ok(liststatus);
        }
        //--------------------------------------------------------------------------------------------------------
        //--Добавление статуса------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddStatus([FromBody] Status status)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Status stat = new Status();
                stat.Name = status.Name;
                stat.Primechanie = status.Primechanie;
                stat.UserMod = username;
                stat.DateMod = DateTime.Now;
                db.Statuses.Add(stat);
                db.SaveChanges();
                return Ok(stat);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }

        //--Редактирование статуса------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditStatus([FromBody] Status status)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Status stat = new Status();
                stat = db.Statuses.FirstOrDefault(s => s.StatusId == status.StatusId);

                stat.Name = status.Name;
                stat.Primechanie = status.Primechanie;
                stat.UserMod = username;
                stat.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(stat);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление статуса------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteStatus([FromBody] Status status)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Status stat = new Status();
                stat = db.Statuses.FirstOrDefault(s => s.StatusId == status.StatusId);
                db.Statuses.Remove(stat);
                db.SaveChanges();
                return Ok(stat);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--------------------------------------------------------------------------------------------------
        //--Список логинов, паролей и пользователей------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult LoginPassword()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            List<Employee> listEmployee = new List<Employee>();
            listEmployee = db.Employees.ToList();

            List<LoginPassword> listslogin = new List<LoginPassword>();
            listslogin = db.LoginPasswords.OrderBy(l => l.Login).ToList();

            //---Сформируем список с Join---
            var loginEmployee = (from log in listslogin
                                   join empl in listEmployee
                                   on log.EmplId equals empl.EmplId
                                   select new
                                   {
                                       Id = log.Id,
                                       Login = log.Login,
                                       Password = log.Password,
                                       Admin = log.Admin,
                                       EmplId = log.EmplId,
                                       DateMod = log.DateMod,
                                       UserMod = log.UserMod,
                                       LastName = empl.LastName,
                                       FirstName = empl.FirstName,
                                       MiddleName = empl.MiddleName,
                                       ThemeColor = log.ThemeColor,
                                       FIO = $"{empl.LastName} {empl.FirstName[0]}. {empl.MiddleName[0]}."
                                   }).OrderBy(x => x.Login).ToList();
            return Ok(loginEmployee);
        }        
        //--Добавление логина, пароль и пользователя------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddLoginPassword([FromBody] LoginPassword login)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                LoginPassword log = new LoginPassword();
                log.Login = login.Login;
                log.Password = BCrypt.Net.BCrypt.HashPassword(login.Password);
                log.EmplId = login.EmplId;
                log.Admin = login.Admin;
                log.ThemeColor = login.ThemeColor;
                log.UserMod = username;
                log.DateMod = DateTime.Now;
                db.LoginPasswords.Add(log);
                db.SaveChanges();
                return Ok(log);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }

        //--Редактирование логина, пароля и пользователя------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditLoginPassword([FromBody] LoginPassword login)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                LoginPassword log = new LoginPassword();
                log = db.LoginPasswords.FirstOrDefault(s => s.Id == login.Id);

                log.Login = login.Login;
                log.Password = BCrypt.Net.BCrypt.HashPassword(login.Password);
                log.EmplId = login.EmplId;
                log.Admin = login.Admin;
                log.ThemeColor = login.ThemeColor;
                log.UserMod = username;
                log.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(log);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление логина, пароля и пользователя------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteLoginPassword([FromBody] LoginPassword login)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                LoginPassword log = new LoginPassword();
                log = db.LoginPasswords.FirstOrDefault(s => s.Id == login.Id);
                db.LoginPasswords.Remove(log);
                db.SaveChanges();
                return Ok(log);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }        
        //----------------Список организаций------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult Organization()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            List<Organization> listorganization = new List<Organization>();
            listorganization = db.Organizations.ToList();

            List<Country> listCountry = new List<Country>();
            listCountry = db.Countries.ToList();

            List<Models.City> listcity = new List<Models.City>();
            listcity = db.Cities.OrderBy(l => l.Name).ToList();

            //Объеденим эти списки
            var organizationCity = (from organization in listorganization
                                    join city in listcity
                                    on organization.CityId equals city.CityId
                                    join country in listCountry
                                    on city.CountryId equals country.CountryId
                                    select new
                                    {
                                        OrgId = organization.OrgId,
                                        Name = organization.Name,
                                        UNP = organization.Unp,
                                        Address = organization.Address,
                                        CityId = city.CityId,
                                        CityName = city.Name,
                                        CountryId = country.CountryId,
                                        CountryName = country.Name,
                                        CountryFullName = country.NameFull,
                                        UserMod = organization.UserMod,
                                        DateMod = organization.DateMod,
                                    }).OrderBy(x => x.Name).ToList();
            return Ok(organizationCity);
        }
        //--------------------------------------------------------------------------------------------------------
        //--Добавление организации------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddOrganization([FromBody] Organization organization)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Organization org = new Organization();
                org.Name = organization.Name;
                org.Unp = organization.Unp;
                org.Address = organization.Address;
                org.CityId = organization.CityId;
                org.UserMod = username;
                org.DateMod = DateTime.Now;
                db.Organizations.Add(org);
                db.SaveChanges();
                return Ok(org);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }

        //--Редактирование организации------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditOrganization([FromBody] Organization organization)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Organization org = new Organization();
                org = db.Organizations.FirstOrDefault(s => s.OrgId == organization.OrgId);

                org.Name = organization.Name;
                org.Unp = organization.Unp;
                org.Address = organization.Address;
                org.CityId = organization.CityId;
                org.UserMod = username;
                org.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(org);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление организации------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteOrganization([FromBody] Organization organization)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Organization org = new Organization();
                org = db.Organizations.FirstOrDefault(s => s.OrgId == organization.OrgId);
                db.Organizations.Remove(org);
                db.SaveChanges();
                return Ok(org);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //----------------Список претензий------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult Pretense()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            List<Pretense> listpretense = new List<Pretense>();
            listpretense = db.Pretenses.ToList();

            List<Organization> listorganization = new List<Organization>();
            listorganization = db.Organizations.ToList();

            List<Valutum> listvaluta = new List<Valutum>();
            listvaluta = db.Valuta.OrderBy(l => l.Name).ToList();

            List<Filial> listfilial = new List<Filial>();
            listfilial = db.Filials.ToList();
            
            List<Predmet> listpredmet = new List<Predmet>();
            listpredmet = db.Predmets.ToList();

            List<Status> liststatus = new List<Status>();
            liststatus = db.Statuses.ToList();

            //Объеденим эти списки
            //var listpretenseJoin = (from pretense in listpretense
            //                        join organization in listorganization
            //                        on pretense.OrgId equals organization.OrgId
            //                        join valuta in listvaluta
            //                        on pretense.ValId equals valuta.ValId
            //                        join filial in listfilial
            //                        on pretense.FilId equals filial.FilId
            //                        join predmet in listpredmet
            //                        on pretense.PredmetId equals predmet.PredmetId
            //                        select new
            //                        {
            //                            PretId = pretense.PretId,
            //                            OrgId = pretense.OrgId,
            //                            OrgName = organization.Name,
            //                            UNP = organization.Unp,
            //                            Address = organization.Address,
            //                            NumberPret = pretense.NumberPret,
            //                            DatePret = pretense.DatePret,
            //                            SummaDolg = pretense.SummaDolg,
            //                            SummaPeny = pretense.SummaPeny,
            //                            SummaProc = pretense.SummaProc,
            //                            SummaItog = pretense.SummaDolg + pretense.SummaPeny + pretense.SummaProc,
            //                            ValId = pretense.ValId,
            //                            Inout = pretense.Inout,
            //                            Visible = pretense.Visible,
            //                            Arhiv = pretense.Arhiv,
            //                            ValName = valuta.Name,
            //                            ValFullName = valuta.NameFull,
            //                            FilId = pretense.FilId,
            //                            FilName = filial.Name,
            //                            PredmetId = pretense.PredmetId,
            //                            PredmetName = predmet.Predmet1,
            //                            UserMod = pretense.UserMod,
            //                            DateMod = pretense.DateMod,
            //                        }).OrderBy(x => x.FilName).ThenBy(u => u.OrgName).Where(i =>i.Visible !=0 && i.Arhiv !=1).ToList();            

            //return Ok(listpretenseJoin);

            List<TablePretense> listTablePretense = new List<TablePretense>();
            listTablePretense = db.TablePretenses.Where(o =>o.Delet != 1).ToList();

            var listpretenseJoin = (
    from pretense in listpretense
    join organization in listorganization on pretense.OrgId equals organization.OrgId
    join valuta in listvaluta on pretense.ValId equals valuta.ValId
    join filial in listfilial on pretense.FilId equals filial.FilId
    join predmet in listpredmet on pretense.PredmetId equals predmet.PredmetId
    select new
    {
        PretId = pretense.PretId,
        OrgId = pretense.OrgId,
        OrgName = organization.Name,
        UNP = organization.Unp,
        Address = organization.Address,
        NumberPret = pretense.NumberPret,
        DatePret = pretense.DatePret,
        SummaDolg = pretense.SummaDolg,
        SummaPeny = pretense.SummaPeny,
        SummaProc = pretense.SummaProc,
        SummaPoshlina = pretense.SummaPoshlina,
        SummaItog = pretense.SummaDolg + pretense.SummaPeny + pretense.SummaProc + pretense.SummaPoshlina,
        ValId = pretense.ValId,
        Inout = pretense.Inout,
        Visible = pretense.Visible,
        Arhiv = pretense.Arhiv,
        ValName = valuta.Name,
        ValFullName = valuta.NameFull,
        FilId = pretense.FilId,
        FilName = filial.Name,
        PredmetId = pretense.PredmetId,
        PredmetName = predmet.Predmet1,
        UserMod = pretense.UserMod,
        DateMod = pretense.DateMod,
        TablePretenses = (
            from tp in listTablePretense
            join status in liststatus on tp.StatusId equals status.StatusId
            where tp.PretId == pretense.PretId
            select new
            {
                tp.TabPretId,
                tp.DateTabPret,
                tp.SummaDolg,
                tp.SummaPeny,
                tp.SummaProc,
                tp.SummaPoshlina,
                summaItog = tp.SummaDolg + tp.SummaPeny + tp.SummaProc + tp.SummaPoshlina,
                valName = valuta.Name,
                tp.ValId,
                tp.Result,
                tp.Primechanie,
                tp.UserMod,
                tp.DateMod,
                tp.StatusId,
                StatusName = status.Name
            }
        ).ToList()
    })
    .Where(i => i.Visible != 0 && i.Arhiv != 1)
    .OrderBy(x => x.FilName)
    .ThenBy(u => u.OrgName)
    .ToList();

            return Ok(listpretenseJoin);
        }
           //--Добавление претензии------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddPretense([FromBody] Pretense pretense)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Pretense pret = new Pretense();
                pret.NumberPret = pretense.NumberPret;
                pret.DatePret = pretense.DatePret;
                pret.SummaDolg = pretense.SummaDolg;
                pret.SummaPeny = pretense.SummaPeny;
                pret.SummaProc = pretense.SummaProc;
                pret.SummaPoshlina = pretense.SummaPoshlina;
                pret.ValId = pretense.ValId;
                pret.OrgId = pretense.OrgId;
                pret.DateRassmPret = pretense.DateRassmPret;
                pret.FilId = pretense.FilId;
                pret.Inout = pretense.Inout;
                pret.PredmetId = pretense.PredmetId;
                pret.Visible = 1;
                pret.Arhiv = 0;
                pret.UserMod = username;
                pret.DateMod = DateTime.Now;
                db.Pretenses.Add(pret);
                db.SaveChanges();
                return Ok(pret);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }
        //--Редактирование претензии------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditPretense([FromBody] Pretense pretense)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Pretense pret = new Pretense();
                pret = db.Pretenses.FirstOrDefault(s => s.PretId == pretense.PretId);

                pret.NumberPret = pretense.NumberPret;
                pret.DatePret = pretense.DatePret;
                pret.SummaDolg = pretense.SummaDolg;
                pret.SummaPeny = pretense.SummaPeny;
                pret.SummaProc = pretense.SummaProc;
                pret.SummaPoshlina = pretense.SummaPoshlina;
                pret.ValId = pretense.ValId;
                pret.OrgId = pretense.OrgId;
                pret.DateRassmPret = pretense.DateRassmPret;
                pret.FilId = pretense.FilId;
                pret.Inout = pretense.Inout;
                pret.PredmetId = pretense.PredmetId;
                pret.UserMod = username;
                pret.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(pret);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Удаление претензии------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeletePretense([FromBody] Pretense pretense)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Pretense pret = new Pretense();
                pret = db.Pretenses.FirstOrDefault(s => s.PretId == pretense.PretId);                
                db.SaveChanges();
                pret.Visible = 0;
                pret.UserMod = username;
                pret.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(pret);                

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--Претензию в архив------------------------------------------------
        [HttpPost]
        [Route("[action]")]
        public IActionResult AddArchive([FromBody] Pretense pretense)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Pretense pret = new Pretense();
                pret = db.Pretenses.FirstOrDefault(s => s.PretId == pretense.PretId);
                db.SaveChanges();
                pret.Arhiv = 1;
                pret.UserMod = username;
                pret.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(pret);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--Добавление результата------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddResult([FromBody] TablePretense result)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                TablePretense tab = new TablePretense();
                tab.PretId = result.PretId;
                tab.DateTabPret = result.DateTabPret;
                tab.SummaDolg = result.SummaDolg;
                tab.SummaPeny = result.SummaPeny;
                tab.SummaProc = result.SummaProc;
                tab.SummaPoshlina = result.SummaPoshlina;
                tab.ValId = result.ValId;
                tab.Result = result.Result;
                tab.StatusId = result.StatusId;
                tab.UserMod = username;
                tab.DateMod = DateTime.Now;
                db.TablePretenses.Add(tab);
                db.SaveChanges();
                return Ok(tab);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }
        //--Редактирование результата------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditResult([FromBody] TablePretense result)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                TablePretense tab = new TablePretense();
                tab = db.TablePretenses.FirstOrDefault(s => s.TabPretId == result.TabPretId);
                //tab.PretId = result.PretId;
                tab.DateTabPret = result.DateTabPret;
                tab.SummaDolg = result.SummaDolg;
                tab.SummaPeny = result.SummaPeny;
                tab.SummaProc = result.SummaProc;
                tab.SummaPoshlina = result.SummaPoshlina;
                tab.ValId = result.ValId;
                tab.Result = result.Result;
                tab.StatusId = result.StatusId;
                tab.UserMod = username;
                tab.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(tab);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--Удаление результата------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult DeleteResult([FromBody] TablePretense result)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                TablePretense res = new TablePretense();
                res = db.TablePretenses.FirstOrDefault(s => s.TabPretId == result.TabPretId);
                res.Delet = 1;
                res.UserMod = username;
                res.DateMod = DateTime.Now;
                db.SaveChanges();
                return Ok(res);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при удалении записи");
            }
        }
        //--------------------------------------------------------------------------------------
        [HttpPost]
        [Route("[action]")]
        public async Task<IActionResult> ValutaKurs([FromBody] DateTime date)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            List<Valutum> listval = new List<Valutum>();
            listval = db.Valuta.OrderBy(l => l.Name).ToList();

            // Формируем запрос к API НБРБ
            var client = new HttpClient();
            var formattedDate = date.ToString("yyyy-MM-dd");
            var response = await client.GetAsync($"https://www.nbrb.by/api/exrates/rates?ondate={formattedDate}&periodicity=0");

            if (!response.IsSuccessStatusCode)
                return StatusCode((int)response.StatusCode, "Ошибка при получении курсов валют");

            var json = await response.Content.ReadAsStringAsync();
            var nbrbRates = JsonConvert.DeserializeObject<List<NbrbValuta>>(json);

            // Сопоставляем курсы с твоими валютами
            var result = (from val in listval
                          join rate in nbrbRates
                          on val.Name.ToLower() equals rate.Cur_Abbreviation.ToLower()
                          into joinedRates
                          from rate in joinedRates.DefaultIfEmpty()
                          select new
                          {
                              val.ValId,
                              val.Name,
                              val.CodeVal,
                              Rate = rate?.Cur_OfficialRate ?? 0m,
                              Scale = rate?.Cur_Scale ?? 1
                          }).ToList();

            return Ok(result);
        }
        //-----------------------Печатная форма отчета в WORD---------------------------------------------------------------------------------
        
        [HttpPost]
        [Route("[action]")]
        public async Task<IActionResult> ReportSvod([FromBody] string date1)
        {
            try {
                
                if (!DateTime.TryParse(date1, out var date))
                {                    
                    return BadRequest("Неверный формат даты");
                }
                
                var stream = new MemoryStream();

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = new Body();

                    // Альбомная ориентация
                    SectionProperties sectionProps = new SectionProperties(
                        new PageSize
                        {
                            Width = 16838, // 11.69 inches * 1440 twips
                            Height = 11906, // 8.27 inches * 1440 twips
                            Orient = PageOrientationValues.Landscape
                        },
                        new PageMargin
                        {
                            Top = 720,
                            Bottom = 720,
                            Left = 720,
                            Right = 720,
                            Header = 450,
                            Footer = 450,
                            Gutter = 0
                        }
                    );

                    // Заголовок
                    var titleRunProps = new RunProperties
                    {
                        RunFonts = new RunFonts
                        {
                            Ascii = "Times New Roman",
                            HighAnsi = "Times New Roman",
                            EastAsia = "Times New Roman",
                            ComplexScript = "Times New Roman"
                        },
                        Bold = new Bold(),
                        Underline = new Underline { Val = UnderlineValues.Single },
                        FontSize = new FontSize { Val = "32" } // 16pt
                    };
                    //Это для филиала подчеркнутый текст
                    var filialRunProps = new RunProperties
                    {
                        RunFonts = new RunFonts
                        {
                            Ascii = "Times New Roman",
                            HighAnsi = "Times New Roman",
                            EastAsia = "Times New Roman",
                            ComplexScript = "Times New Roman"
                        },
                        Bold = new Bold(),
                        Italic = new Italic(),
                        FontSize = new FontSize { Val = "32" } // 16pt
                    };
                    var tableRunProps = new RunProperties
                    {
                        RunFonts = new RunFonts
                        {
                            Ascii = "Times New Roman",
                            HighAnsi = "Times New Roman",
                            EastAsia = "Times New Roman",
                            ComplexScript = "Times New Roman"
                        },
                        FontSize = new FontSize { Val = "24" } // 12pt
                    };

                    Paragraph titleParagraph = new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                    new Run(titleRunProps.CloneNode(true), new Text("Сведения о проводимой претензионно-исковой работе")),
                    new Run(new Break()),
                    new Run(titleRunProps.CloneNode(true), new Text("в ОАО Гомельтранснефть Дружба на " + date.ToString("g")))
                );
                    body.Append(titleParagraph);

                    //-----Выберем претензии со связанной таблицей и сгруппируем по контрагентам--
                    List<Organization> listorganization = new List<Organization>();
                    listorganization = db.Organizations.ToList();

                    List<Valutum> listvaluta = new List<Valutum>();
                    listvaluta = db.Valuta.OrderBy(l => l.Name).ToList();

                    List<Filial> listfilial = new List<Filial>();
                    listfilial = db.Filials.ToList();

                    List<Predmet> listpredmet = new List<Predmet>();
                    listpredmet = db.Predmets.ToList();

                    List<Status> liststatus = new List<Status>();
                    liststatus = db.Statuses.ToList();

                    List<Pretense> listpretense = new List<Pretense>();
                    //listpretense = db.Pretenses.Where(j => j.DatePret <= date && j.Visible == 1 && j.Arhiv == 0).ToList();
                    listpretense = db.Pretenses.Where(j => j.DatePret <= date && j.Visible == 1 && j.Arhiv == 0).ToList();

                    List<Summa> listsumma = new List<Summa>();
                    listsumma = db.Summas.ToList();

                    List<PretenseSumma> listpretensesumma = new List<PretenseSumma>();
                    listpretensesumma = db.PretenseSummas.ToList();

                    List<TablePretense> listtablepretense = new List<TablePretense>();
                    listtablepretense = db.TablePretenses.Where(g => g.DateTabPret <= date && g.Delet != 1).ToList();

                    List<ResultSumma> listresultsumma = new List<ResultSumma>();
                    listresultsumma = db.ResultSummas.ToList();

                    //Создаем список отсортированыый по убыванию по датам для заполнения информайии о количестве дел, находящихся в производстве
                    List<TablePretense> listtablepretenseOrderBy = new List<TablePretense>();
                    listtablepretenseOrderBy = db.TablePretenses.Where(g => g.DateTabPret <= date && g.Delet != 1).OrderByDescending(o => o.DateTabPret).ToList();

                    //-----------------------------------------------------------------------------
                    var listpretenseJoin = (
        from pretense in listpretense
        join organization in listorganization on pretense.OrgId equals organization.OrgId
        join filial in listfilial on pretense.FilId equals filial.FilId
        join predmet in listpredmet on pretense.PredmetId equals predmet.PredmetId
        let summaItems = new[]
        {
        new { Summa = listpretensesumma.FirstOrDefault(ps => ps.PretId == pretense.PretId && ps.SummaId == 1), Type = "Dolg" },
        new { Summa = listpretensesumma.FirstOrDefault(ps => ps.PretId == pretense.PretId && ps.SummaId == 2), Type = "Peny" },
        new { Summa = listpretensesumma.FirstOrDefault(ps => ps.PretId == pretense.PretId && ps.SummaId == 3), Type = "Proc" },
        new { Summa = listpretensesumma.FirstOrDefault(ps => ps.PretId == pretense.PretId && ps.SummaId == 4), Type = "Poshlina" }
        }
        let currencyGroups = summaItems
            .Where(x => x.Summa != null)
            .Select(x => new
            {
                CurrencyId = x.Summa.ValId,
                CurrencyName = listvaluta.FirstOrDefault(v => v.ValId == x.Summa.ValId)?.Name,
                Type = x.Type,
                Value = x.Summa.Value
            })
            .GroupBy(x => new { x.CurrencyId, x.CurrencyName })
            .Select(g => new
            {
                CurrencyId = g.Key.CurrencyId,
                CurrencyName = g.Key.CurrencyName,
                SummaDolg = g.Where(x => x.Type == "Dolg").Sum(x => x.Value),
                SummaPeny = g.Where(x => x.Type == "Peny").Sum(x => x.Value),
                SummaProc = g.Where(x => x.Type == "Proc").Sum(x => x.Value),
                SummaPoshlina = g.Where(x => x.Type == "Poshlina").Sum(x => x.Value),
                SummaItog = g.Sum(x => x.Value)
            })
            .ToList()

        let tablePretenseList = (
            from tp in listtablepretense
            where tp.PretId == pretense.PretId
            let resultItems = new[]
            {
            new { Summa = listresultsumma.FirstOrDefault(rs => rs.ResultId == tp.TabPretId && rs.SummaId == 1), Type = "Dolg" },
            new { Summa = listresultsumma.FirstOrDefault(rs => rs.ResultId == tp.TabPretId && rs.SummaId == 2), Type = "Peny" },
            new { Summa = listresultsumma.FirstOrDefault(rs => rs.ResultId == tp.TabPretId && rs.SummaId == 3), Type = "Proc" },
            new { Summa = listresultsumma.FirstOrDefault(rs => rs.ResultId == tp.TabPretId && rs.SummaId == 4), Type = "Poshlina" }
            }
            let resultCurrencyGroups = resultItems
                .Where(x => x.Summa != null)
                .Select(x => new
                {
                    CurrencyId = x.Summa.ValId,
                    CurrencyName = listvaluta.FirstOrDefault(v => v.ValId == x.Summa.ValId)?.Name,
                    Type = x.Type,
                    Value = x.Summa.Value
                })
                .GroupBy(x => new { x.CurrencyId, x.CurrencyName })
                .Select(g => new
                {
                    CurrencyId = g.Key.CurrencyId,
                    CurrencyName = g.Key.CurrencyName,
                    SummaDolg = g.Where(x => x.Type == "Dolg").Sum(x => x.Value),
                    SummaPeny = g.Where(x => x.Type == "Peny").Sum(x => x.Value),
                    SummaProc = g.Where(x => x.Type == "Proc").Sum(x => x.Value),
                    SummaPoshlina = g.Where(x => x.Type == "Poshlina").Sum(x => x.Value),
                    SummaItog = g.Sum(x => x.Value)
                })
                .ToList()

            select new
            {
                tp.TabPretId,
                tp.DateTabPret,
                tp.Result,
                tp.Primechanie,
                tp.StatusId,
                tp.UserMod,
                tp.DateMod,
                ResultCurrencyGroups = resultCurrencyGroups
            }
        ).ToList()

        select new
        {
            PretId = pretense.PretId,
            OrgId = pretense.OrgId,
            OrgName = organization.Name,
            UNP = organization.Unp,
            Address = organization.Address,
            NumberPret = pretense.NumberPret,
            DatePret = pretense.DatePret,
            Inout = pretense.Inout,
            Visible = pretense.Visible,
            Arhiv = pretense.Arhiv,
            FilId = pretense.FilId,
            FilName = filial.Name,
            FilNameFull = filial.NameFull,
            PredmetId = pretense.PredmetId,
            PredmetName = predmet.Predmet1,
            UserMod = pretense.UserMod,
            DateMod = pretense.DateMod,
            CurrencyGroups = currencyGroups,
            TablePretenseList = tablePretenseList
        })
    .Where(i => i.Visible != 0 && i.Arhiv != 1)
    .OrderBy(x => x.FilName)
    .ThenBy(u => u.OrgName)
    .ToList();


                    var groupedByFilial = listpretenseJoin
        .GroupBy(x => new { x.FilId, x.FilName, x.FilNameFull })
        .Select(filialGroup => new
        {
            filialGroup.Key.FilId,
            filialGroup.Key.FilName,
            filialGroup.Key.FilNameFull,

            // Список претензий для таблицы претензий
            PretenseList = filialGroup
                .Select(pretense => new
                {
                    PretId = pretense.PretId,
                    OrgId = pretense.OrgId,
                    OrgName = pretense.OrgName,
                    UNP = pretense.UNP,
                    Address = pretense.Address,
                    NumberPret = pretense.NumberPret,
                    DatePret = pretense.DatePret,
                    Inout = pretense.Inout,
                    PredmetName = pretense.PredmetName,
                    UserMod = pretense.UserMod,
                    DateMod = pretense.DateMod,

                    // TablePretense для каждой претензии
                    TablePretenseList = pretense.TablePretenseList
                        .Select(tp => new
                        {
                            tp.TabPretId,
                            tp.DateTabPret,
                            tp.Result,
                            tp.Primechanie,
                            tp.StatusId,
                            tp.UserMod,
                            tp.DateMod,
                            ResultCurrencyGroups = tp.ResultCurrencyGroups
                        })
                        .OrderBy(x => x.DateTabPret)
                        .ToList(),

                    // Суммы по валютам для претензии
                    PretenseCurrencyGroups = pretense.CurrencyGroups
                        .Select(cg => new CurrencyGroup
                        {
                            CurrencyId = cg.CurrencyId,
                            CurrencyName = cg.CurrencyName,
                            SummaDolg = cg.SummaDolg,
                            SummaPeny = cg.SummaPeny,
                            SummaProc = cg.SummaProc,
                            SummaPoshlina = cg.SummaPoshlina
                        })
                        .ToList()
                })
                .OrderBy(x => x.DatePret)
                .ToList(),

            // Список организаций для таблицы организаций
            Organizations = filialGroup
                .GroupBy(x => new { x.OrgId, x.OrgName, x.UNP, x.Address })
                .Select(orgGroup => new
                {
                    orgGroup.Key.OrgId,
                    orgGroup.Key.OrgName,
                    orgGroup.Key.UNP,
                    orgGroup.Key.Address,
                    MinDatePret = orgGroup.Min(x => x.DatePret),
                    MaxDatePret = orgGroup.Max(x => x.DatePret),

                    // Суммы по валютам для организации
                    CurrencyGroups = orgGroup
                        .SelectMany(x => x.CurrencyGroups)
                        .GroupBy(cg => new { cg.CurrencyId, cg.CurrencyName })
                        .Select(g => new CurrencyGroup
                        {
                            CurrencyId = g.Key.CurrencyId,
                            CurrencyName = g.Key.CurrencyName,
                            SummaDolg = g.Sum(x => x.SummaDolg),
                            SummaPeny = g.Sum(x => x.SummaPeny),
                            SummaProc = g.Sum(x => x.SummaProc),
                            SummaPoshlina = g.Sum(x => x.SummaPoshlina)
                        })
                        .ToList(),

                    PredmetNames = string.Join(", ", orgGroup.Select(x => x.PredmetName).Distinct()),

                    // Добавляем ResultDetails
                    ResultDetails = orgGroup
                        .SelectMany(x => x.TablePretenseList)
                        .Select(tp => new
                        {
                            tp.TabPretId,
                            tp.DateTabPret,
                            tp.Result,
                            tp.Primechanie,
                            tp.StatusId,
                            tp.UserMod,
                            tp.DateMod,
                            CurrencyGroups = tp.ResultCurrencyGroups
                        })
                        .OrderBy(x => x.DateTabPret)
                        .ToList()
                })
                .OrderBy(x => x.OrgName)
                .ToList(),

            // Итоговые суммы по валютам для всего филиала
            FilialCurrencyGroups = filialGroup
                .SelectMany(x => x.CurrencyGroups)
                .GroupBy(cg => new { cg.CurrencyId, cg.CurrencyName })
                .Select(g => new CurrencyGroup
                {
                    CurrencyId = g.Key.CurrencyId,
                    CurrencyName = g.Key.CurrencyName,
                    SummaDolg = g.Sum(x => x.SummaDolg),
                    SummaPeny = g.Sum(x => x.SummaPeny),
                    SummaProc = g.Sum(x => x.SummaProc),
                    SummaPoshlina = g.Sum(x => x.SummaPoshlina)
                })
                .ToList()
        })
        .OrderBy(x => x.FilName)
        .ToList();

                    //-----------------------------------------------------------
                    int rowIndex = 1;

                    foreach (var pret in listpretense)
                    {
                        var groupedByCurrency = db.PretenseSummas
                            .Where(ps => ps.PretId == pret.PretId)
                            .Join(db.Valuta,
                                  ps => ps.ValId,
                                  v => v.ValId,
                                  (ps, v) => new { PretenseSumma = ps, Valuta = v })
                            .GroupBy(x => new { x.Valuta.ValId, x.Valuta.Name })
                            .Select(g => new
                            {
                                CurrencyId = g.Key.ValId,
                                CurrencyName = g.Key.Name,
                                SummaDolg = g.Where(x => x.PretenseSumma.SummaId == 1).Sum(x => x.PretenseSumma.Value),
                                SummaPeny = g.Where(x => x.PretenseSumma.SummaId == 2).Sum(x => x.PretenseSumma.Value),
                                SummaProc = g.Where(x => x.PretenseSumma.SummaId == 3).Sum(x => x.PretenseSumma.Value),
                                SummaPoshlina = g.Where(x => x.PretenseSumma.SummaId == 4).Sum(x => x.PretenseSumma.Value),
                                SummaItog = g.Sum(x => x.PretenseSumma.Value)
                            })
                            .ToList();
                    }

                    //-----------------------------------------------------------------------------
                    var total1 = new CategoryTotal(); // Претензий на стадии рассмотрения
                    var total2 = new CategoryTotal(); // Удовлетворенные претензии
                    var total3 = new CategoryTotal(); // Претензии в адрес ОАО
                    var total4 = new CategoryTotal(); // Исковые заявления в отношении ОАО
                    var total5 = new CategoryTotal(); // Заявления в судебном порядке на стадии рассмотрения
                    var total6 = new CategoryTotal(); // Удовлетворенные исковые требования
                    var total7 = new CategoryTotal(); // Предъявлены исполнительные документы
                    var total8 = new CategoryTotal(); // Дела на стадии исполнительного производства
                    var total9 = new CategoryTotal(); // Предъявлено требований кредитора
                    var total10 = new CategoryTotal(); // Заявлений о совершении исполнительной надписи
                    var total11 = new CategoryTotal(); // Удовлетворенные заявления о совершении исполнительной надписи

                    var allTotals = new List<CategoryTotal> { total1, total2, total3, total4, total5, total6, total7, total8, total9, total10, total11 };

                    foreach (var filial in groupedByFilial)
                    {
                        // Пустая строка перед филиалом
                        body.Append(new Paragraph());

                        // Название филиала по центру
                        var filialParagraph = new Paragraph(
                            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                            new Run(filialRunProps.CloneNode(true), new Text(filial.FilNameFull))
                        );
                        body.Append(filialParagraph);

                        // Создаем таблицу для текущего филиала
                        Table table = new Table();

                        // Настройки таблицы
                        TableProperties tblProps = new TableProperties(
                            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }, // 100%
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        );
                        table.AppendChild(tblProps);

                        // Заголовки и ширины
                        string[] headers = {
                "№", "Наименование организации (УНП)", "Предмет задолженности",
                "Дата образования", "Сумма", "Проделанная работа",
                "Взыскано", "Взыскано", "Остаток задолженности"
            };

                        string[] columnWidths = {
                "800", "2000", "800", "800", "800", "2000", "800", "800", "800"
            }; // в процентах от 10000

                        TableRow headerRow = new TableRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            TableCell cell = new TableCell(
                                new Paragraph(
                                    new Run(tableRunProps.CloneNode(true), new Text(headers[i]))
                                ),
                                new TableCellProperties(
                                    new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[i] }
                                )
                            );
                            headerRow.Append(cell);
                        }
                        table.Append(headerRow);

                        // Переменные для итоговых сумм по филиалу
                        var filialStatus9Sums = new Dictionary<string, decimal?>();
                        var filialStatus10Sums = new Dictionary<string, decimal?>();

                        foreach (var org in filial.Organizations)
                        {

                            TableRow dataRow = new TableRow();

                            dataRow.Append(new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[0] }),
                                new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(rowIndex.ToString())))
                            ));

                            dataRow.Append(new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[1] }),
                                new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{org.OrgName} (УНП: {org.UNP})")))
                            ));

                            dataRow.Append(new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[2] }),
                                new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(org.PredmetNames)))
                            ));
                            if (org.MinDatePret == org.MaxDatePret)
                            {
                                dataRow.Append(new TableCell(
                                    new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[3] }),
                                    new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("-")))
                                ));
                            }
                            else
                            {
                                dataRow.Append(new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[3] }),
                                new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(org.MinDatePret + " - " + org.MaxDatePret)))
                                ));
                            }

                            // Формируем текст для взысканных сумм по суду с группировкой по валютам
                            var strVziskanoSud = "";
                            var summaDolga = "";
                            //Сумма долга во всех валютах
                            foreach (var i in org.CurrencyGroups)
                            {
                                if (org.CurrencyGroups.Count == 1)
                                {
                                    summaDolga = summaDolga + i.SummaItog.ToString() + " " + i.CurrencyName;
                                }
                                else
                                {
                                    summaDolga = summaDolga + i.SummaItog.ToString() + " " + i.CurrencyName + ", ";
                                }
                            }

                            //---------------------------------------------------------------------------------------------------------------
                            dataRow.Append(new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[4] }),
                                new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(summaDolga)))
                            ));

                            var ResultText = org.ResultDetails?.FirstOrDefault()?.Result ?? "";
                            //Переберем таблицу результатов
                            var status9SumsByCurrency = new Dictionary<string, decimal>(); // Статус 9
                            var status10SumsByCurrency = new Dictionary<string, decimal>(); // Статус 10

                            var groupedResults = org.ResultDetails
                        .Where(r => r.StatusId == 9 || r.StatusId == 10)
                        .GroupBy(r => r.StatusId)
                        .Select(statusGroup => new
                        {
                            StatusId = statusGroup.Key,
                            CurrencyGroups = statusGroup
                                .SelectMany(r => r.CurrencyGroups.Select(cg => new
                                {
                                    CurrencyId = cg.CurrencyId,
                                    CurrencyName = cg.CurrencyName,
                                    SummaDolg = cg.SummaDolg,
                                    SummaPeny = cg.SummaPeny,
                                    SummaProc = cg.SummaProc,
                                    SummaPoshlina = cg.SummaPoshlina,
                                    SummaItog = cg.SummaItog
                                }))
                                .GroupBy(cg => new { cg.CurrencyId, cg.CurrencyName })
                                .Select(g => new
                                {
                                    CurrencyId = g.Key.CurrencyId,
                                    CurrencyName = g.Key.CurrencyName,
                                    SummaDolg = g.Sum(x => x.SummaDolg),
                                    SummaPeny = g.Sum(x => x.SummaPeny),
                                    SummaProc = g.Sum(x => x.SummaProc),
                                    SummaPoshlina = g.Sum(x => x.SummaPoshlina),
                                    SummaItog = g.Sum(x => x.SummaItog),
                                    ResultDetails = g.ToList()
                                })
                                .ToList()
                        })
                        .ToList();

                            string sumStatus9 = "-";
                            string sumStatus10 = "-";

                            //---Теперь выведем взысканные суммы с валюмами и статусами---
                            foreach (var res in groupedResults)
                            {
                                if (res.StatusId == 9)
                                {
                                    sumStatus9 = string.Join("; ", res.CurrencyGroups.Select(cg => $"{cg.SummaItog} {cg.CurrencyName}"));

                                    // Собираем итоги по филиалу для статуса 9
                                    foreach (var currency in res.CurrencyGroups)
                                    {
                                        if (filialStatus9Sums.ContainsKey(currency.CurrencyName))
                                        {
                                            filialStatus9Sums[currency.CurrencyName] += currency.SummaItog;
                                        }
                                        else
                                        {
                                            filialStatus9Sums[currency.CurrencyName] = currency.SummaItog;
                                        }
                                    }
                                }
                                if (res.StatusId == 10)
                                {
                                    sumStatus10 = string.Join("; ", res.CurrencyGroups.Select(cg => $"{cg.SummaItog} {cg.CurrencyName}"));

                                    // Собираем итоги по филиалу для статуса 10
                                    foreach (var currency in res.CurrencyGroups)
                                    {
                                        if (filialStatus10Sums.ContainsKey(currency.CurrencyName))
                                        {
                                            filialStatus10Sums[currency.CurrencyName] += currency.SummaItog;
                                        }
                                        else
                                        {
                                            filialStatus10Sums[currency.CurrencyName] = currency.SummaItog;
                                        }
                                    }
                                }
                            }
                            //--в этом списке получим сумму сразу по валютам без статусов----------------------
                            var groupedResVal = org.ResultDetails
        .Where(r => r.StatusId == 9 || r.StatusId == 10)
        .SelectMany(r => r.CurrencyGroups.Select(cg => new
        {
            TabPretId = r.TabPretId,
            DateTabPret = r.DateTabPret,
            Result = r.Result,
            Primechanie = r.Primechanie,
            StatusId = r.StatusId,
            UserMod = r.UserMod,
            DateMod = r.DateMod,
            CurrencyId = cg.CurrencyId,
            CurrencyName = cg.CurrencyName,
            SummaDolg = cg.SummaDolg,
            SummaPeny = cg.SummaPeny,
            SummaProc = cg.SummaProc,
            SummaPoshlina = cg.SummaPoshlina,
            SummaItog = cg.SummaItog
        }))
        .GroupBy(x => new { x.CurrencyId, x.CurrencyName })
        .Select(g => new
        {
            CurrencyId = g.Key.CurrencyId,
            CurrencyName = g.Key.CurrencyName,
            SummaDolg = g.Sum(x => x.SummaDolg),
            SummaPeny = g.Sum(x => x.SummaPeny),
            SummaProc = g.Sum(x => x.SummaProc),
            SummaPoshlina = g.Sum(x => x.SummaPoshlina),
            SummaItog = g.Sum(x => x.SummaItog),
            ResultDetails = g.ToList()
        })
        .ToList();

                            //--------Теперь нужно найти статок долга------------------------------------------------------
                            string oststokDolga = "";

                            foreach (var val in org.CurrencyGroups)
                            {
                                var matched = groupedResVal.FirstOrDefault(item => item.CurrencyId == val.CurrencyId);

                                if (matched != null)
                                {
                                    decimal? ost = val.SummaItog - matched.SummaItog;
                                    oststokDolga += $"{ost} {matched.CurrencyName}; ";
                                }
                                else
                                {
                                    oststokDolga += $"{val.SummaItog} {val.CurrencyName}; ";
                                }
                            }
                            //---------------------------------------------------------------------------------------------

                            //----------Теперь заполним маленькую таблицу с итогами взыскания------------------------------

                            dataRow.Append(new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[5] }),
                                new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(ResultText))) // Проделанная работа
                            ));

                            dataRow.Append(new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[6] }),
                            new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumStatus9))) // Сумма по статусу 9
                            ));

                            dataRow.Append(new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[7] }),
                                new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumStatus10))) // Сумма по статусу 10
                            ));

                            dataRow.Append(new TableCell(
                                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[8] }),
                                new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(oststokDolga)))
                            ));

                            table.Append(dataRow);
                            rowIndex++;
                        }

                        // РАСЧЕТ ОСТАТКА ЗАДОЛЖЕННОСТИ ПО ФИЛИАЛУ
                        string filialOststokDolga = string.Join("; ", filial.FilialCurrencyGroups.Select(val =>
                        {
                            decimal? status9Sum = filialStatus9Sums.GetValueOrDefault(val.CurrencyName, 0);
                            decimal? status10Sum = filialStatus10Sums.GetValueOrDefault(val.CurrencyName, 0);
                            decimal? ost = val.SummaItog - (status9Sum + status10Sum);
                            return $"{ost} {val.CurrencyName}";
                        }));

                        if (string.IsNullOrEmpty(filialOststokDolga))
                        {
                            filialOststokDolga = "-";
                        }

                        // ДОБАВЛЯЕМ СТРОКУ С ИТОГАМИ ПО ФИЛИАЛУ
                        TableRow totalRow = new TableRow();

                        // Объединяем первые 6 ячеек для надписи "Итого по филиалу:"
                        totalRow.Append(new TableCell(
                            new TableCellProperties(
                                new GridSpan { Val = 6 }, // Объединяем 6 ячеек
                                new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "5200" }
                            ),
                            new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Всего взыскано (нарастающий итог по году на отчетную дату)")))
                        ));

                        // Итоги по статусу 9
                        string totalStatus9 = filialStatus9Sums.Any() ?
                            string.Join("; ", filialStatus9Sums.Select(x => $"{x.Value} {x.Key}")) : "-";

                        totalRow.Append(new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[6] }),
                            new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(totalStatus9)))
                        ));

                        // Итоги по статусу 10
                        string totalStatus10 = filialStatus10Sums.Any() ?
                            string.Join("; ", filialStatus10Sums.Select(x => $"{x.Value} {x.Key}")) : "-";

                        totalRow.Append(new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[7] }),
                            new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(totalStatus10)))
                        ));

                        // Ячейка "Остаток задолженности" - пустая
                        totalRow.Append(new TableCell(
                            new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths[8] }),
                            new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(filialOststokDolga)))
                        ));

                        table.Append(totalRow);

                        // Добавляем таблицу текущего филиала в документ
                        body.Append(table);

                        body.Append(new Paragraph()); // пустая строка между таблицами

                        //--------Выведем вторую таблицу (Информация о количестве дел)-------------------------------------

                        int count1 = 0;
                        int count2 = 0;
                        int count3 = 0;
                        int count4 = 0;
                        int count5 = 0;
                        int count6 = 0;
                        int count7 = 0;
                        int count8 = 0;
                        int count9 = 0;
                        int count10 = 0;
                        int count11 = 0;
                        string str01 = "";
                        string str02 = "";
                        string str03 = "";
                        string str04 = "";
                        string str05 = "";
                        string str06 = "";
                        string str07 = "";
                        string str08 = "";
                        string str09 = "";
                        string str010 = "";
                        string str011 = "";
                        List<string> str1 = new List<string>();
                        List<string> str2 = new List<string>();
                        List<string> str3 = new List<string>();
                        List<string> str4 = new List<string>();
                        List<string> str5 = new List<string>();
                        List<string> str6 = new List<string>();
                        List<string> str7 = new List<string>();
                        List<string> str8 = new List<string>();
                        List<string> str9 = new List<string>();
                        List<string> str10 = new List<string>();
                        List<string> str11 = new List<string>();

                        //----------Название таблицы----------------

                        Paragraph tabinfoParagraph = new Paragraph(
                            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                            new Run(tableRunProps.CloneNode(true), new Text("Информация о количестве дел, находящихся в производстве")),
                            new Run(new Break()),
                            new Run(tableRunProps.CloneNode(true), new Text("по состоянию на " + date.ToString("g")))
                            );
                        body.Append(tabinfoParagraph);

                        //-------------------------------------------------------------------                   

                        foreach (var item in filial.PretenseList)
                        {
                            // Проверяем, что есть записи в TablePretenseList
                            if (item.TablePretenseList != null && item.TablePretenseList.Any())
                            {
                                // Берем последнюю запись TablePretense (самую свежую по дате)
                                var lastTablePretense = item.TablePretenseList
        .Where(tp => tp.DateTabPret != null && tp.DateTabPret <= date)
        .OrderByDescending(tp => tp.DateTabPret)
        .FirstOrDefault();

                                // Если нашли запись с датой
                                if (lastTablePretense != null)
                                {
                                    int? statusId = lastTablePretense.StatusId;
                                    int? inout = item.Inout;

                                    //System.IO.File.AppendAllText("debug.txt", $" filial: {filial.FilName},Org: {item.OrgName}, StatusId: {statusId}, Inout: {inout}\n");

                                    if ((statusId == 1 || statusId == 9) && inout == 1)
                                    {
                                        count1++;
                                        str1.Add(item.OrgName);
                                    }
                                    else if (statusId == 6 && inout == 1)
                                    {
                                        count2++;
                                        str2.Add(item.OrgName);
                                    }
                                    else if ((statusId == 1 || statusId == 9) && inout == 0)
                                    {
                                        count3++;
                                        str3.Add(item.OrgName);
                                    }
                                    else if ((statusId == 2 || statusId == 10) && inout == 0)
                                    {
                                        count4++;
                                        str4.Add(item.OrgName);
                                    }
                                    else if ((statusId == 2 || statusId == 10) && inout == 1)
                                    {
                                        count5++;
                                        str5.Add(item.OrgName);
                                    }
                                    else if (statusId == 12 && inout == 1)
                                    {
                                        count6++;
                                        str6.Add(item.OrgName);
                                    }
                                    else if (statusId == 13)
                                    {
                                        count7++;
                                        str7.Add(item.OrgName);
                                    }
                                    else if (statusId == 3)
                                    {
                                        count8++;
                                        str8.Add(item.OrgName);
                                    }
                                    else if (statusId == 4 || statusId == 5)
                                    {
                                        count9++;
                                        str9.Add(item.OrgName);
                                    }
                                    else if (statusId == 8)
                                    {
                                        count10++;
                                        str10.Add(item.OrgName);
                                    }
                                    else if (statusId == 14)
                                    {
                                        count11++;
                                        str11.Add(item.OrgName);
                                    }
                                }
                            }
                        }
                        Dictionary<string, int> orgsCount1 = str1.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount2 = str2.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount3 = str3.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount4 = str4.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount5 = str5.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount6 = str6.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount7 = str7.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount8 = str8.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount9 = str9.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount10 = str10.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        Dictionary<string, int> orgsCount11 = str11.GroupBy(org => org).ToDictionary(g => g.Key, g => g.Count());
                        //Подготовим строки для записи в таблицу
                        foreach (var it in orgsCount1)
                        {
                            str01 = str01 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount2)
                        {
                            str02 = str02 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount3)
                        {
                            str03 = str03 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount4)
                        {
                            str04 = str04 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount5)
                        {
                            str05 = str05 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount6)
                        {
                            str06 = str06 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount7)
                        {
                            str07 = str07 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount8)
                        {
                            str08 = str08 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount9)
                        {
                            str09 = str09 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount10)
                        {
                            str010 = str010 + " " + it.Key + "(" + it.Value + ") ";
                        }
                        foreach (var it in orgsCount11)
                        {
                            str011 = str011 + " " + it.Key + "(" + it.Value + ") ";
                        }

                        //-------------------------------------------------------------------

                        var pretenseListAsObjects = filial.PretenseList.Cast<object>().ToList();
                        //string sumInfo1 = GetCurrencySumInfo(str1, pretenseListAsObjects, date);
                        //string sumInfo2 = GetCurrencySumInfo(str2, pretenseListAsObjects, date);
                        //string sumInfo3 = GetCurrencySumInfo(str3, pretenseListAsObjects, date);
                        //string sumInfo4 = GetCurrencySumInfo(str4, pretenseListAsObjects, date);
                        //string sumInfo5 = GetCurrencySumInfo(str5, pretenseListAsObjects, date);
                        //string sumInfo6 = GetCurrencySumInfo(str6, pretenseListAsObjects, date);
                        //string sumInfo7 = GetCurrencySumInfo(str7, pretenseListAsObjects, date);
                        //string sumInfo8 = GetCurrencySumInfo(str8, pretenseListAsObjects, date);
                        //string sumInfo9 = GetCurrencySumInfo(str9, pretenseListAsObjects, date);
                        //string sumInfo10 = GetCurrencySumInfo(str10, pretenseListAsObjects, date);
                        //string sumInfo11 = GetCurrencySumInfo(str11, pretenseListAsObjects, date);
                        var sumInfoObj1 = GetCurrencySumInfoAsObject(str1, pretenseListAsObjects, date);
                        var sumInfoObj2 = GetCurrencySumInfoAsObject(str2, pretenseListAsObjects, date);
                        var sumInfoObj3 = GetCurrencySumInfoAsObject(str3, pretenseListAsObjects, date);
                        var sumInfoObj4 = GetCurrencySumInfoAsObject(str4, pretenseListAsObjects, date);
                        var sumInfoObj5 = GetCurrencySumInfoAsObject(str5, pretenseListAsObjects, date);
                        var sumInfoObj6 = GetCurrencySumInfoAsObject(str6, pretenseListAsObjects, date);
                        var sumInfoObj7 = GetCurrencySumInfoAsObject(str7, pretenseListAsObjects, date);
                        var sumInfoObj8 = GetCurrencySumInfoAsObject(str8, pretenseListAsObjects, date);
                        var sumInfoObj9 = GetCurrencySumInfoAsObject(str9, pretenseListAsObjects, date);
                        var sumInfoObj10 = GetCurrencySumInfoAsObject(str10, pretenseListAsObjects, date);
                        var sumInfoObj11 = GetCurrencySumInfoAsObject(str11, pretenseListAsObjects, date);

                        string sumInfo1 = FormatCurrencySums(sumInfoObj1);
                        string sumInfo2 = FormatCurrencySums(sumInfoObj2);
                        string sumInfo3 = FormatCurrencySums(sumInfoObj3);
                        string sumInfo4 = FormatCurrencySums(sumInfoObj4);
                        string sumInfo5 = FormatCurrencySums(sumInfoObj5);
                        string sumInfo6 = FormatCurrencySums(sumInfoObj6);
                        string sumInfo7 = FormatCurrencySums(sumInfoObj7);
                        string sumInfo8 = FormatCurrencySums(sumInfoObj8);
                        string sumInfo9 = FormatCurrencySums(sumInfoObj9);
                        string sumInfo10 = FormatCurrencySums(sumInfoObj10);
                        string sumInfo11 = FormatCurrencySums(sumInfoObj11);

                        //----------Попытаемся расчитать ИТОГИ для ИТОГОВОЙ таблицы-----------

                        if (filial.FilId != 10) // Исключаем филиал УСП Дружбинец
                        {
                            // Суммируем count'ы
                            total1.Count += count1;
                            total2.Count += count2;
                            total3.Count += count3;
                            total4.Count += count4;
                            total5.Count += count5;
                            total6.Count += count6;
                            total7.Count += count7;
                            total8.Count += count8;
                            total9.Count += count9;
                            total10.Count += count10;
                            total11.Count += count11;

                            // Добавляем организации
                            total1.Organizations.AddRange(str1);
                            total2.Organizations.AddRange(str2);
                            total3.Organizations.AddRange(str3);
                            total4.Organizations.AddRange(str4);
                            total5.Organizations.AddRange(str5);
                            total6.Organizations.AddRange(str6);
                            total7.Organizations.AddRange(str7);
                            total8.Organizations.AddRange(str8);
                            total9.Organizations.AddRange(str9);
                            total10.Organizations.AddRange(str10);
                            total11.Organizations.AddRange(str11);

                            // Суммируем валютные суммы (нужен метод для парсинга sumInfo1, sumInfo2 и т.д.)
                            //AddCurrencySums(total1.CurrencySums, sumInfo1);
                            //AddCurrencySums(total2.CurrencySums, sumInfo2);
                            //AddCurrencySums(total3.CurrencySums, sumInfo3);
                            //AddCurrencySums(total4.CurrencySums, sumInfo4);
                            //AddCurrencySums(total5.CurrencySums, sumInfo5);
                            //AddCurrencySums(total6.CurrencySums, sumInfo6);
                            //AddCurrencySums(total7.CurrencySums, sumInfo7);
                            //AddCurrencySums(total8.CurrencySums, sumInfo8);
                            //AddCurrencySums(total9.CurrencySums, sumInfo9);
                            //AddCurrencySums(total10.CurrencySums, sumInfo10);
                            //AddCurrencySums(total11.CurrencySums, sumInfo11);
                            total1.AddCurrencySumsFromObject(sumInfoObj1);
                            total2.AddCurrencySumsFromObject(sumInfoObj2);
                            total3.AddCurrencySumsFromObject(sumInfoObj3);
                            total4.AddCurrencySumsFromObject(sumInfoObj4);
                            total5.AddCurrencySumsFromObject(sumInfoObj5);
                            total6.AddCurrencySumsFromObject(sumInfoObj6);
                            total7.AddCurrencySumsFromObject(sumInfoObj7);
                            total8.AddCurrencySumsFromObject(sumInfoObj8);
                            total9.AddCurrencySumsFromObject(sumInfoObj9);
                            total10.AddCurrencySumsFromObject(sumInfoObj10);
                            total11.AddCurrencySumsFromObject(sumInfoObj11);
                        }

                        //-----------Сама таблица--------------------                
                        Table table2 = new Table();

                        // Настройки таблицы
                        TableProperties tblProps2 = new TableProperties(
                            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }, // 100%
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        );
                        table2.AppendChild(tblProps2);

                        // Заголовки для второй таблицы
                        string[] headers2 = {
    "Стадия рассмотрения", "Количество дел", "Сумма"
};

                        string[] columnWidths2 = {
    "2000", "2000", "1000"
}; // в процентах от 5000

                        // Создаем строку заголовков для второй таблицы
                        TableRow headerRow2 = new TableRow();
                        for (int i = 0; i < headers2.Length; i++)
                        {
                            TableCell cell = new TableCell(
                                new Paragraph(
                                    new Run(tableRunProps.CloneNode(true), new Text(headers2[i]))
                                ),
                                new TableCellProperties(
                                    new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[i] }
                                )
                            );
                            headerRow2.Append(cell);
                        }
                        table2.Append(headerRow2);

                        // Строка 1 
                        TableRow row1 = new TableRow();
                        row1.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Претензий на стадии рассмотрения")))));
                        row1.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count1 + " " + str01)))));
                        row1.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo1)))));
                        table2.Append(row1);

                        // Строка 2 
                        TableRow row2 = new TableRow();
                        row2.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Удовлетворенные претензии")))));
                        row2.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count2 + " " + str02)))));
                        row2.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo2)))));
                        table2.Append(row2);

                        // Строка 3 
                        TableRow row3 = new TableRow();
                        row3.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Претензии в адрес ОАО")))));
                        row3.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count3 + " " + str03)))));
                        row3.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo3)))));
                        table2.Append(row3);

                        // Строка 4 
                        TableRow row4 = new TableRow();
                        row4.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Исковые заявления в отношении ОАО")))));
                        row4.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count4 + " " + str04)))));
                        row4.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo4)))));
                        table2.Append(row4);

                        // Строка 5 
                        TableRow row5 = new TableRow();
                        row5.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Заявления в судебном порядке на стадии рассмотрения")))));
                        row5.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count5 + " " + str05)))));
                        row5.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo5)))));
                        table2.Append(row5);

                        // Строка 6 
                        TableRow row6 = new TableRow();
                        row6.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Удовлетворенные исковые требования")))));
                        row6.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count6 + " " + str06)))));
                        row6.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo6)))));
                        table2.Append(row6);

                        // Строка 7 
                        TableRow row7 = new TableRow();
                        row7.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Предъявлены исполнительные документы к счетам должников через банк")))));
                        row7.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count7 + " " + str07)))));
                        row7.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo7)))));
                        table2.Append(row7);

                        // Строка 8 
                        TableRow row8 = new TableRow();
                        row8.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Дела на стадии исполнительного производства в ОПИ")))));
                        row8.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count8 + " " + str08)))));
                        row8.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo8)))));
                        table2.Append(row8);

                        // Строка 9 
                        TableRow row9 = new TableRow();
                        row9.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Предъявлено требований кредитора")))));
                        row9.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count9 + " " + str09)))));
                        row9.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo9)))));
                        table2.Append(row9);

                        // Строка 10
                        TableRow row10 = new TableRow();
                        row10.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Заявлений о совершении исполнительной надписи на стадии рассмотрения")))));
                        row10.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count10 + " " + str010)))));
                        row10.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo10)))));
                        table2.Append(row10);

                        // Строка 11 
                        TableRow row11 = new TableRow();
                        row11.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Удовлетворенные заявления о совершении исполнительной надписи")))));
                        row11.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(count11 + " " + str011)))));
                        row11.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(sumInfo11)))));
                        table2.Append(row11);

                        body.Append(table2); // вторая таблица

                        //-----------------------------Третья таблица--------------------------------------------------

                        // Пустая строка между таблицами
                        body.Append(new Paragraph());

                        //--------ВЫВЕДЕМ ТРЕТЬЮ ТАБЛИЦУ (Сведения о количестве) ДЛЯ ТЕКУЩЕГО ФИЛИАЛА---------                    

                        //-----------Создаем третью таблицу для текущего филиала--------------------                
                        Table table3 = new Table();

                        // Настройки таблицы
                        TableProperties tblProps3 = new TableProperties(
                            new TableWidth { Width = "2500", Type = TableWidthUnitValues.Pct },
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                new RightBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        );
                        table3.AppendChild(tblProps3);

                        // Заголовки для третьей таблицы
                        string[] headers3 = {
        "Сведения о количестве", "по состоянию на " + date.ToString("g")
    };
                        string[] columnWidths3 = {
        "1250", "1250"
    };
                        // Создаем строку заголовков для третьей таблицы
                        TableRow headerRow3 = new TableRow();
                        for (int i = 0; i < headers3.Length; i++)
                        {
                            TableCell cell = new TableCell(
                                new Paragraph(
                                    new Run(tableRunProps.CloneNode(true), new Text(headers3[i]))
                                ),
                                new TableCellProperties(
                                    new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[i] }
                                )
                            );
                            headerRow3.Append(cell);
                        }
                        table3.Append(headerRow3);

                        TableRow row31 = new TableRow();
                        row31.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Претензии")))));
                        row31.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                        table3.Append(row31);

                        TableRow row32 = new TableRow();
                        row32.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Претензии в адрес ОАО")))));
                        row32.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                        table3.Append(row32);

                        TableRow row33 = new TableRow();
                        row33.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Исковые заявления")))));
                        row33.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                        table3.Append(row33);

                        TableRow row34 = new TableRow();
                        row34.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Исковые заявления в отношении ОАО")))));
                        row34.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                        table3.Append(row34);

                        TableRow row35 = new TableRow();
                        row35.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Заявления о совершении исполнительной надписи")))));
                        row35.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                        table3.Append(row35);

                        // Добавляем третью таблицу текущего филиала в документ
                        body.Append(table3);
                        //-----------------------------------------------------------------------------------------------
                    }
                    //-----------------------------Формируем ИТОГОВУЮ ТАБЛИЦУ--------------------------------------------

                    body.Append(new Paragraph()); // пустая строка между таблицами

                    Paragraph tabinfoParagraph1 = new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                        new Run(titleRunProps.CloneNode(true), new Text("ОАО Гомельтранснефть Дружба (без учета УСП Дружбинец)")),
                        new Run(new Break()),
                        new Run(titleRunProps.CloneNode(true), new Text("итоговая таблица на " + date.ToString("g")))
                        );
                    body.Append(tabinfoParagraph1);


                    Table table2ITOG = new Table();

                    // Настройки таблицы
                    TableProperties tblProps2ITOG = new TableProperties(
                        new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }, // 100%
                        new TableBorders(
                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                        )
                    );
                    table2ITOG.AppendChild(tblProps2ITOG);

                    // Заголовки для второй таблицы
                    string[] headers2I = {
    "Стадия рассмотрения", "Количество дел", "Сумма"
};

                    string[] columnWidths2I = {
    "2000", "2000", "1000"
}; // в процентах от 5000

                    // Создаем строку заголовков для второй таблицы
                    TableRow headerRow2ITOG = new TableRow();
                    for (int i = 0; i < headers2I.Length; i++)
                    {
                        TableCell cell = new TableCell(
                            new Paragraph(
                                new Run(tableRunProps.CloneNode(true), new Text(headers2I[i]))
                            ),
                            new TableCellProperties(
                                new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[i] }
                            )
                        );
                        headerRow2ITOG.Append(cell);
                    }
                    table2ITOG.Append(headerRow2ITOG);

                    // Строка 1 
                    TableRow row1I = new TableRow();
                    row1I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Претензий на стадии рассмотрения")))));
                    row1I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total1.Count} {FormatOrganizations(total1.Organizations)}")))));
                    row1I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total1.CurrencySums))))));
                    table2ITOG.Append(row1I);

                    // Строка 2 
                    TableRow row2I = new TableRow();
                    row2I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Удовлетворенные претензии")))));
                    row2I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total2.Count} {FormatOrganizations(total2.Organizations)}")))));
                    row2I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total2.CurrencySums))))));
                    table2ITOG.Append(row2I);

                    // Строка 3 
                    TableRow row3I = new TableRow();
                    row3I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Претензии в адрес ОАО")))));
                    row3I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total3.Count} {FormatOrganizations(total3.Organizations)}")))));
                    row3I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total3.CurrencySums))))));
                    table2ITOG.Append(row3I);

                    // Строка 4 
                    TableRow row4I = new TableRow();
                    row4I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Исковые заявления в отношении ОАО")))));
                    row4I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total4.Count} {FormatOrganizations(total4.Organizations)}")))));
                    row4I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total4.CurrencySums))))));
                    table2ITOG.Append(row4I);

                    // Строка 5 
                    TableRow row5I = new TableRow();
                    row5I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Заявления в судебном порядке на стадии рассмотрения")))));
                    row5I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total5.Count} {FormatOrganizations(total5.Organizations)}")))));
                    row5I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total5.CurrencySums))))));
                    table2ITOG.Append(row5I);

                    // Строка 6 
                    TableRow row6I = new TableRow();
                    row6I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Удовлетворенные исковые требования")))));
                    row6I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total6.Count} {FormatOrganizations(total6.Organizations)}")))));
                    row6I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total6.CurrencySums))))));
                    table2ITOG.Append(row6I);

                    // Строка 7 
                    TableRow row7I = new TableRow();
                    row7I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Предъявлены исполнительные документы к счетам должников через банк")))));
                    row7I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total7.Count} {FormatOrganizations(total7.Organizations)}")))));
                    row7I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total7.CurrencySums))))));
                    table2ITOG.Append(row7I);

                    // Строка 8 
                    TableRow row8I = new TableRow();
                    row8I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Дела на стадии исполнительного производства в ОПИ")))));
                    row8I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total8.Count} {FormatOrganizations(total8.Organizations)}")))));
                    row8I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total8.CurrencySums))))));
                    table2ITOG.Append(row8I);

                    // Строка 9 
                    TableRow row9I = new TableRow();
                    row9I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Предъявлено требований кредитора")))));
                    row9I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total9.Count} {FormatOrganizations(total9.Organizations)}")))));
                    row9I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total9.CurrencySums))))));
                    table2ITOG.Append(row9I);

                    // Строка 10
                    TableRow row10I = new TableRow();
                    row10I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Заявлений о совершении исполнительной надписи на стадии рассмотрения")))));
                    row10I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total10.Count} {FormatOrganizations(total10.Organizations)}")))));
                    row10I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total10.CurrencySums))))));
                    table2ITOG.Append(row10I);

                    // Строка 11 
                    TableRow row11I = new TableRow();
                    row11I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Удовлетворенные заявления о совершении исполнительной надписи")))));
                    row11I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text($"{total11.Count} {FormatOrganizations(total11.Organizations)}")))));
                    row11I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths2I[2] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text(FormatCurrencySums(total11.CurrencySums))))));
                    table2ITOG.Append(row11I);

                    body.Append(table2ITOG);

                    //-----------------------------Третья ИТОГОВАЯ таблица--------------------------------------------------

                    // Пустая строка между таблицами
                    body.Append(new Paragraph());

                    //--------ВЫВЕДЕМ ТРЕТЬЮ ИТОГОВУЮ ТАБЛИЦУ (Сведения о количестве) ДЛЯ ТЕКУЩЕГО ФИЛИАЛА---------                    

                    //-----------Создаем третью ИТОГОВУЮ таблицу--------------------                
                    Table table3ITOG = new Table();

                    // Настройки таблицы
                    TableProperties tblProps3ITOG = new TableProperties(
                        new TableWidth { Width = "2500", Type = TableWidthUnitValues.Pct },
                        new TableBorders(
                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                        )
                    );
                    table3ITOG.AppendChild(tblProps3ITOG);

                    // Заголовки для третьей таблицы
                    string[] headers3I = {
        "Сведения о количестве", "по состоянию на " + date.ToString("g")
    };
                    string[] columnWidths3I = {
        "1250", "1250"
    };
                    // Создаем строку заголовков для третьей таблицы
                    TableRow headerRow3ITOG = new TableRow();
                    for (int i = 0; i < headers3I.Length; i++)
                    {
                        TableCell cell = new TableCell(
                            new Paragraph(
                                new Run(tableRunProps.CloneNode(true), new Text(headers3I[i]))
                            ),
                            new TableCellProperties(
                                new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[i] }
                            )
                        );
                        headerRow3ITOG.Append(cell);
                    }
                    table3ITOG.Append(headerRow3ITOG);

                    TableRow row31I = new TableRow();
                    row31I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Претензии")))));
                    row31I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                    table3ITOG.Append(row31I);

                    TableRow row32I = new TableRow();
                    row32I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Претензии в адрес ОАО")))));
                    row32I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                    table3ITOG.Append(row32I);

                    TableRow row33I = new TableRow();
                    row33I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Исковые заявления")))));
                    row33I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                    table3ITOG.Append(row33I);

                    TableRow row34I = new TableRow();
                    row34I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Исковые заявления в отношении ОАО")))));
                    row34I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                    table3ITOG.Append(row34I);

                    TableRow row35I = new TableRow();
                    row35I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[0] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("Заявления о совершении исполнительной надписи")))));
                    row35I.Append(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = columnWidths3I[1] }), new Paragraph(new Run(tableRunProps.CloneNode(true), new Text("")))));
                    table3ITOG.Append(row35I);

                    // Добавляем третью таблицу текущего филиала в документ
                    body.Append(table3ITOG);

                    //---------------------------------------------------------------------------------------------------

                    body.Append(sectionProps); // применяем ориентацию

                    mainPart.Document.Append(body);
                    mainPart.Document.Save();
                }

                stream.Seek(0, SeekOrigin.Begin);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Отчет.docx");
            }
            catch (Exception ex)
            {
                // Записываем ошибку в лог-файл
                var logPath = @"C:/Temp/ReportSvod_debug.log";
                var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Ошибка: {ex}\n";
                System.IO.File.AppendAllText(logPath, logMessage);

                return StatusCode(500, "Ошибка при формировании отчёта. Подробности см. в ReportSvod_debug.log");
            }
}
        //--Вспомогательные методы для групировки данных претензий------------------------------------------
        private string FormatOrganizations(List<string> organizations)
        {
            if (organizations == null || !organizations.Any())
                return "";

            return string.Join(" ", organizations.GroupBy(x => x)
                                                .Select(g => $"{g.Key}({g.Count()})"));
        }
        //--------------------------------------------------------------------------------------------------
        //public class CategoryTotal
        //{
        //    public int Count { get; set; }
        //    public List<string> Organizations { get; set; } = new List<string>();
        //    public Dictionary<string, decimal> CurrencySums { get; set; } = new Dictionary<string, decimal>();
        //}
        public class CategoryTotal
        {
            public int Count { get; set; }
            public List<string> Organizations { get; set; } = new List<string>();
            public Dictionary<string, decimal> CurrencySums { get; set; } = new Dictionary<string, decimal>();

            // Новый метод для добавления валютных сумм из объекта
            public void AddCurrencySumsFromObject(Dictionary<string, decimal> sourceSums)
            {
                if (sourceSums == null) return;

                foreach (var currencySum in sourceSums)
                {
                    if (CurrencySums.ContainsKey(currencySum.Key))
                        CurrencySums[currencySum.Key] += currencySum.Value;
                    else
                        CurrencySums[currencySum.Key] = currencySum.Value;
                }
            }
        }
        //--------------------------------------------------------------------------------------------------
        private void AddCurrencySums(Dictionary<string, decimal> targetSums, string sumInfo)
        {
            if (string.IsNullOrEmpty(sumInfo) || sumInfo == "-") return;

            // Парсим строку вида "100 USD, 200 EUR, 300 RUB"
            var currencyParts = sumInfo.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in currencyParts)
            {
                var trimmed = part.Trim();
                var lastSpaceIndex = trimmed.LastIndexOf(' ');

                if (lastSpaceIndex > 0)
                {
                    var amountStr = trimmed.Substring(0, lastSpaceIndex).Trim();
                    var currency = trimmed.Substring(lastSpaceIndex + 1).Trim();

                    if (decimal.TryParse(amountStr, out decimal amount))
                    {
                        if (targetSums.ContainsKey(currency))
                            targetSums[currency] += amount;
                        else
                            targetSums[currency] = amount;
                    }
                }
            }
        }

        //--------------------------------------------------------------------------------------------------
        //private string GetCurrencySumInfo(List<string> organizationNames, List<object> pretenseList, DateTime reportDate)
        //{
        //    if (organizationNames == null || !organizationNames.Any() || pretenseList == null)
        //        return "-";

        //    var currencySums = new Dictionary<string, decimal>();

        //    foreach (var orgName in organizationNames.Distinct())
        //    {
        //        foreach (var pretenseObj in pretenseList)
        //        {
        //            try
        //            {
        //                var orgNameProp = pretenseObj.GetType().GetProperty("OrgName");
        //                var pretenseOrgName = orgNameProp?.GetValue(pretenseObj)?.ToString();

        //                if (pretenseOrgName == orgName)
        //                {
        //                    // Получаем исходные суммы долга из PretenseSumma
        //                    var currencyGroupsProp = pretenseObj.GetType().GetProperty("PretenseCurrencyGroups");
        //                    var currencyGroups = currencyGroupsProp?.GetValue(pretenseObj) as List<CurrencyGroup>;

        //                    // Получаем все результаты погашений до отчетной даты
        //                    var tablePretenseListProp = pretenseObj.GetType().GetProperty("TablePretenseList");
        //                    var tablePretenseList = tablePretenseListProp?.GetValue(pretenseObj) as IEnumerable<dynamic>;

        //                    if (currencyGroups != null)
        //                    {
        //                        foreach (var currencyGroup in currencyGroups)
        //                        {
        //                            if (currencyGroup != null && !string.IsNullOrEmpty(currencyGroup.CurrencyName))
        //                            {
        //                                string currencyName = currencyGroup.CurrencyName;

        //                                // Начальная сумма долга
        //                                decimal initialAmount = (currencyGroup.SummaDolg ?? 0) +
        //                                                       (currencyGroup.SummaPeny ?? 0) +
        //                                                       (currencyGroup.SummaProc ?? 0) +
        //                                                       (currencyGroup.SummaPoshlina ?? 0);

        //                                // Вычитаем все погашения до отчетной даты
        //                                decimal paidAmount = 0;

        //                                if (tablePretenseList != null)
        //                                {
        //                                    foreach (var tablePretense in tablePretenseList)
        //                                    {
        //                                        // Проверяем, что дата погашения не превышает отчетную дату
        //                                        var dateTabPretProp = tablePretense.GetType().GetProperty("DateTabPret");
        //                                        var dateTabPret = (DateTime?)dateTabPretProp?.GetValue(tablePretense);

        //                                        if (dateTabPret != null && dateTabPret <= reportDate)
        //                                        {
        //                                            // Получаем суммы погашения для этой валюты
        //                                            var resultCurrencyGroupsProp = tablePretense.GetType().GetProperty("ResultCurrencyGroups");
        //                                            var resultCurrencyGroups = resultCurrencyGroupsProp?.GetValue(tablePretense) as IEnumerable<dynamic>;

        //                                            if (resultCurrencyGroups != null)
        //                                            {
        //                                                var matchingCurrency = resultCurrencyGroups
        //                                                    .FirstOrDefault(rcg =>
        //                                                        rcg != null &&
        //                                                        rcg.CurrencyName?.ToString() == currencyName);

        //                                                if (matchingCurrency != null)
        //                                                {
        //                                                    var summaItogProp = matchingCurrency.GetType().GetProperty("SummaItog");
        //                                                    if (summaItogProp != null)
        //                                                    {
        //                                                        decimal? paidSum = summaItogProp.GetValue(matchingCurrency) as decimal?;
        //                                                        paidAmount += paidSum ?? 0;
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                }

        //                                // Остаток долга
        //                                decimal remainingAmount = initialAmount - paidAmount;

        //                                if (remainingAmount > 0)
        //                                {
        //                                    if (currencySums.ContainsKey(currencyName))
        //                                        currencySums[currencyName] += remainingAmount;
        //                                    else
        //                                        currencySums[currencyName] = remainingAmount;
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine($"Ошибка в GetCurrencySumInfo: {ex.Message}");
        //            }
        //        }
        //    }

        //    return currencySums.Count == 0 ? "-" :
        //        string.Join(", ", currencySums.OrderBy(x => x.Key).Select(x => $"{x.Value:N2} {x.Key}"));
        //}
        private Dictionary<string, decimal> GetCurrencySumInfoAsObject(List<string> organizationNames, List<object> pretenseList, DateTime reportDate)
        {
            if (organizationNames == null || !organizationNames.Any() || pretenseList == null)
                return new Dictionary<string, decimal>();

            var currencySums = new Dictionary<string, decimal>();

            foreach (var orgName in organizationNames.Distinct())
            {
                foreach (var pretenseObj in pretenseList)
                {
                    try
                    {
                        var orgNameProp = pretenseObj.GetType().GetProperty("OrgName");
                        var pretenseOrgName = orgNameProp?.GetValue(pretenseObj)?.ToString();

                        if (pretenseOrgName == orgName)
                        {
                            // Получаем исходные суммы долга из PretenseSumma
                            var currencyGroupsProp = pretenseObj.GetType().GetProperty("PretenseCurrencyGroups");
                            var currencyGroups = currencyGroupsProp?.GetValue(pretenseObj) as List<CurrencyGroup>;

                            // Получаем все результаты погашений до отчетной даты
                            var tablePretenseListProp = pretenseObj.GetType().GetProperty("TablePretenseList");
                            var tablePretenseList = tablePretenseListProp?.GetValue(pretenseObj) as IEnumerable<dynamic>;

                            if (currencyGroups != null)
                            {
                                foreach (var currencyGroup in currencyGroups)
                                {
                                    if (currencyGroup != null && !string.IsNullOrEmpty(currencyGroup.CurrencyName))
                                    {
                                        string currencyName = currencyGroup.CurrencyName;

                                        // Начальная сумма долга
                                        decimal initialAmount = (currencyGroup.SummaDolg ?? 0) +
                                                               (currencyGroup.SummaPeny ?? 0) +
                                                               (currencyGroup.SummaProc ?? 0) +
                                                               (currencyGroup.SummaPoshlina ?? 0);

                                        // Вычитаем все погашения до отчетной даты
                                        decimal paidAmount = 0;
                                        if (tablePretenseList != null)
                                        {
                                            foreach (var tablePretense in tablePretenseList)
                                            {
                                                var dateTabPretProp = tablePretense.GetType().GetProperty("DateTabPret");
                                                var dateTabPret = (DateTime?)dateTabPretProp?.GetValue(tablePretense);

                                                if (dateTabPret != null && dateTabPret <= reportDate)
                                                {
                                                    var resultCurrencyGroupsProp = tablePretense.GetType().GetProperty("ResultCurrencyGroups");
                                                    var resultCurrencyGroups = resultCurrencyGroupsProp?.GetValue(tablePretense) as IEnumerable<dynamic>;

                                                    if (resultCurrencyGroups != null)
                                                    {
                                                        var matchingCurrency = resultCurrencyGroups
                                                            .FirstOrDefault(rcg =>
                                                                rcg != null &&
                                                                rcg.CurrencyName?.ToString() == currencyName);

                                                        if (matchingCurrency != null)
                                                        {
                                                            var summaItogProp = matchingCurrency.GetType().GetProperty("SummaItog");
                                                            if (summaItogProp != null)
                                                            {
                                                                decimal? paidSum = summaItogProp.GetValue(matchingCurrency) as decimal?;
                                                                paidAmount += paidSum ?? 0;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        // Остаток долга
                                        decimal remainingAmount = initialAmount - paidAmount;
                                        if (remainingAmount > 0)
                                        {
                                            if (currencySums.ContainsKey(currencyName))
                                                currencySums[currencyName] += remainingAmount;
                                            else
                                                currencySums[currencyName] = remainingAmount;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка в GetCurrencySumInfoAsObject: {ex.Message}");
                    }
                }
            }

            return currencySums;
        }

        //--------------------------------------------------------------------------------------------------
        private Dictionary<string, decimal> GetRemainingSumByCurrency(dynamic pretense, DateTime reportDate)
        {
            var result = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

            // Сумма долга
            foreach (var debt in pretense.PretenseSummaList ?? Enumerable.Empty<dynamic>())
            {
                if (debt.Currency != null)
                {
                    if (!result.ContainsKey(debt.Currency))
                        result[debt.Currency] = 0;

                    result[debt.Currency] += debt.Summa ?? 0;
                }
            }

            // Погашенные суммы
            foreach (var tp in pretense.TablePretenseList ?? Enumerable.Empty<dynamic>())
            {
                if (tp.DateTabPret != null && tp.DateTabPret <= reportDate)
                {
                    foreach (var paid in tp.ResultSummaList ?? Enumerable.Empty<dynamic>())
                    {
                        if (paid.Currency != null)
                        {
                            if (!result.ContainsKey(paid.Currency))
                                result[paid.Currency] = 0;

                            result[paid.Currency] -= paid.Summa ?? 0;
                        }
                    }
                }
            }

            return result;
        }
        //--------------------------------------------------------------------------------------------------
        public class CurrencyGroup
        {
            public int CurrencyId { get; set; }
            public string CurrencyName { get; set; }
            public decimal? SummaDolg { get; set; }
            public decimal? SummaPeny { get; set; }
            public decimal? SummaProc { get; set; }
            public decimal? SummaPoshlina { get; set; }
            public decimal? SummaItog => SummaDolg + SummaPeny + SummaProc + SummaPoshlina;
        }
        private List<CurrencyGroup> GroupByCurrency(IEnumerable<dynamic> items)
        {
            return items
                .Where(x => x.Summa != null && x.Valuta != null)
                .GroupBy(x => new { x.Valuta.ValId, x.Valuta.Name })
                .Select(g => new CurrencyGroup
                {
                    CurrencyId = g.Key.ValId,
                    CurrencyName = g.Key.Name,
                    SummaDolg = g.Where(x => x.Type == "Dolg").Sum(x => x.Summa.Value),
                    SummaPeny = g.Where(x => x.Type == "Peny").Sum(x => x.Summa.Value),
                    SummaProc = g.Where(x => x.Type == "Proc").Sum(x => x.Summa.Value),
                    SummaPoshlina = g.Where(x => x.Type == "Poshlina").Sum(x => x.Summa.Value)
                })
                .ToList();
        }
        private List<CurrencyGroup> AggregateCurrencyGroups(IEnumerable<CurrencyGroup> groups)
        {
            return groups
                .GroupBy(g => new { g.CurrencyId, g.CurrencyName })
                .Select(g => new CurrencyGroup
                {
                    CurrencyId = g.Key.CurrencyId,
                    CurrencyName = g.Key.CurrencyName,
                    SummaDolg = g.Sum(x => x.SummaDolg),
                    SummaPeny = g.Sum(x => x.SummaPeny),
                    SummaProc = g.Sum(x => x.SummaProc),
                    SummaPoshlina = g.Sum(x => x.SummaPoshlina)
                })
                .ToList();
        }
        //--------------------------------------------------------------------------------------------------
        // Класс для хранения суммы в валютах
        public class CategorySum
        {
            public decimal TotalAmount { get; set; }
            public string CurrencyInfo { get; set; }
        }

        // Функция для расчета суммы остатков по списку претензий
        CategorySum CalculateCategorySum(List<string> organizationNames, List<dynamic> pretenseList)
        {
            var totalAmount = 0m;
            var currencyAmounts = new Dictionary<string, decimal>();

            foreach (var orgName in organizationNames)
            {
                // Находим претензии этой организации
                var orgPretenses = pretenseList.Where(p => p.OrgName == orgName);

                foreach (var pretense in orgPretenses)
                {
                    // Расчет остатка для каждой претензии
                    if (pretense.RemainingAmounts != null)
                    {
                        foreach (var amount in pretense.RemainingAmounts)
                        {
                            totalAmount += amount.Amount;

                            // Суммируем по валютам
                            if (currencyAmounts.ContainsKey(amount.CurrencyName))
                            {
                                currencyAmounts[amount.CurrencyName] += amount.Amount;
                            }
                            else
                            {
                                currencyAmounts[amount.CurrencyName] = amount.Amount;
                            }
                        }
                    }
                }
            }
            // Форматируем информацию о валютах
            string currencyInfo = string.Join(", ",
                currencyAmounts.Select(x => $"{x.Key}: {x.Value:N2}"));

            return new CategorySum
            {
                TotalAmount = totalAmount,
                CurrencyInfo = currencyInfo
            };
        }
        //--------------------------------------------------------------------------------------------------
        private string FormatCurrencySums(Dictionary<string, decimal> currencySums)
        {
            if (currencySums.Count == 0)
                return "-";

            var result = new StringBuilder();
            foreach (var currencySum in currencySums)
            {
                result.Append($"{currencySum.Value} {currencySum.Key}, ");
            }
            return result.ToString().TrimEnd(',', ' ');
        }
        //----------------Вспомогательный метод для объединения ячеек---------------------
        private TableCell CreateMergedCell(string text, RunProperties runProps, string width, int horizontalSpan, int verticalSpan)
        {
            var cell = new TableCell(
                new Paragraph(new Run(runProps.CloneNode(true), new Text(text))),
                new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = width })
            );

            if (horizontalSpan > 1)
            {
                cell.TableCellProperties.Append(new GridSpan { Val = horizontalSpan });
            }

            if (verticalSpan > 1)
            {
                cell.TableCellProperties.Append(new VerticalMerge { Val = MergedCellValues.Restart });
            }

            return cell;
        }
        //--------------------------------------------------------------------------------
        //----------------Список претензий------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult PretenseHome()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            List<Pretense> listpretense = new List<Pretense>();
            listpretense = db.Pretenses.ToList();

            List<Organization> listorganization = new List<Organization>();
            listorganization = db.Organizations.ToList();

            List<Valutum> listvaluta = new List<Valutum>();
            listvaluta = db.Valuta.OrderBy(l => l.Name).ToList();

            List<Filial> listfilial = new List<Filial>();
            listfilial = db.Filials.ToList();

            List<Predmet> listpredmet = new List<Predmet>();
            listpredmet = db.Predmets.ToList();

            List<Status> liststatus = new List<Status>();
            liststatus = db.Statuses.ToList();

            
            List<TablePretense> listTablePretense = new List<TablePretense>();
            listTablePretense = db.TablePretenses.Where(o => o.Delet != 1).ToList();

            var listpretenseJoin = (
    from pretense in listpretense
    join organization in listorganization on pretense.OrgId equals organization.OrgId
    join valuta in listvaluta on pretense.ValId equals valuta.ValId
    join filial in listfilial on pretense.FilId equals filial.FilId
    join predmet in listpredmet on pretense.PredmetId equals predmet.PredmetId
    select new
    {
        PretId = pretense.PretId,
        OrgId = pretense.OrgId,
        OrgName = organization.Name,
        UNP = organization.Unp,
        Address = organization.Address,
        NumberPret = pretense.NumberPret,
        DatePret = pretense.DatePret,
        SummaDolg = pretense.SummaDolg,
        SummaPeny = pretense.SummaPeny,
        SummaProc = pretense.SummaProc,
        SummaPoshlina = pretense.SummaPoshlina,
        SummaItog = pretense.SummaDolg + pretense.SummaPeny + pretense.SummaProc + pretense.SummaPoshlina,
        ValId = pretense.ValId,
        Inout = pretense.Inout,
        Visible = pretense.Visible,
        Arhiv = pretense.Arhiv,
        ValName = valuta.Name,
        ValFullName = valuta.NameFull,
        FilId = pretense.FilId,
        FilName = filial.Name,
        FilNameFull = filial.NameFull,
        PredmetId = pretense.PredmetId,
        PredmetName = predmet.Predmet1,
        UserMod = pretense.UserMod,
        DateMod = pretense.DateMod,
        TablePretenses = (
            from tp in listTablePretense
            join status in liststatus on tp.StatusId equals status.StatusId
            where tp.PretId == pretense.PretId
            select new
            {
                tp.TabPretId,
                tp.DateTabPret,
                tp.SummaDolg,
                tp.SummaPeny,
                tp.SummaProc,
                tp.SummaPoshlina,
                summaItog = tp.SummaDolg + tp.SummaPeny + tp.SummaProc + tp.SummaPoshlina,
                valName = valuta.Name,
                tp.ValId,
                tp.Result,
                tp.Primechanie,
                tp.UserMod,
                tp.DateMod,
                tp.StatusId,
                StatusName = status.Name
            }
        ).ToList()
    })
    .Where(i => i.Visible != 0 && i.Arhiv != 1)
    .OrderBy(x => x.FilName)
    .ThenBy(u => u.OrgName)
    .ToList();

            var groupedByFilial = listpretenseJoin
    .GroupBy(p => p.FilNameFull)
    .Select(g => new
    {
        Filial = g.Key,
        FilId = g.First().FilId,
        Count = g.Count(),
        TotalSum = g.Sum(x => x.SummaItog),
        OrgCount = g.Select(x => x.OrgId).Distinct().Count() 
    })
    .OrderBy(x => x.Filial)
    .ToList();
            if (admin == 1)
            {
                return Ok(groupedByFilial);
            }
            else
            {
              return Ok(groupedByFilial.Where(h=>h.FilId == filialId));
            }                
        }
        //--------------Отчёт для концерна EXCEL-----------------------------------------------------------------------

        //------------------------------Вспомогательные классы---------------------------------------------------------
        public class PretenseReport
        {
            public int PretId { get; set; }
            public string Status { get; set; }
            public string OrgName { get; set; }
            public string City { get; set; }
            public string Country { get; set; }
            public string Address { get; set; }
            public DateTime? DatePret { get; set; }

            // Сырые суммы (тип суммы + валюта + значение)
            public List<RawSum> RawSums { get; set; } = new();

            // Группировка по типам суммы: в каждой — список валют и значений
            public List<SumByType> SumsByType { get; set; } = new();
        }

        public class RawSum
        {
            public string SumType { get; set; }        // из таблицы Summa.Name
            public string Currency { get; set; }       // Valutum.Name
            public string CurrencyCode { get; set; }   // Valutum.CodeVal
            public decimal Value { get; set; }
        }

        public class SumByType
        {
            public string Type { get; set; }           // Summa.Name
            public List<SumItem> Items { get; set; } = new();
        }

        public class SumItem
        {
            public string Currency { get; set; }
            public string CurrencyCode { get; set; }
            public decimal Value { get; set; }
        }
        //-------------------------------------------------------------------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult ReportConcern([FromBody] DateTime date)
        {
            byte[] fileBytes = GenerateExcelReport(date);
            string fileName = $"Таблица о результатах ПИР_{date:yyyy-MM-dd}.xlsx";

            return File(fileBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }
        //-----------------------------------------------------------------------------
        private byte[] GenerateExcelReport(DateTime reportDate)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Отчёт");

            // Общий стиль
            sheet.Style.Font.FontName = "Times New Roman";
            sheet.Style.Font.FontSize = 12;

            sheet.Range("B4:Q4").Merge().Value = "Форма № 1 Претензии, предъявленные к контрагентам";
            sheet.Range("B4:Q4").Style.Font.FontSize = 14;
            sheet.Range("B4:Q4").Style.Font.Bold = true; // жирный
            sheet.Range("B4:Q4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range("B4:Q4").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range("B4:Q4").Style.Alignment.WrapText = true;
            //sheet.Row(4).Height = 40;

            var startDate = new DateTime(reportDate.Year, 1, 1);
            var endDate = reportDate;

            sheet.Range("B5:Q5").Merge().Value = $"в период с {startDate:dd.MM.yyyy} по {endDate:dd.MM.yyyy}";
            sheet.Range("B5:Q5").Style.Font.FontSize = 14;
            sheet.Range("B5:Q5").Style.Font.Bold = true; // жирный
            sheet.Range("B5:Q5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range("B5:Q5").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range("B5:Q5").Style.Alignment.WrapText = true;
            //sheet.Row(5).Height = 35;

            // Шапка таблицы
            sheet.Range("B7:B9").Merge().Value = "№";
            sheet.Range("C7:C9").Merge().Value = "Наименование должника";
            sheet.Range("D7:D9").Merge().Value = "Город (Страна)";
            sheet.Range("E7:E9").Merge().Value = "Дата предъявления претензии";
            sheet.Range("K7:K9").Merge().Value = "Срок рассмотрения претензии";

            sheet.Range("F7:J7").Merge().Value = "Заявлены требования";
            sheet.Range("F8:F9").Merge().Value = "Содержание требований";
            sheet.Range("G8:G9").Merge().Value = "Сумма основного долга";
            sheet.Range("H8:H9").Merge().Value = "Сумма неустойки";
            sheet.Range("I8:I9").Merge().Value = "Сумма процентов";
            sheet.Range("J8:J9").Merge().Value = "Валюта требований";

            sheet.Range("L7:Q7").Merge().Value = "Результат рассмотрения претензии должником";
            sheet.Range("L8:O8").Merge().Value = "Удовлетворена на сумму";
            sheet.Range("P8:Q8").Merge().Value = "Не удовлетворены";

            sheet.Cell("L9").Value = "Сумма основного долга";
            sheet.Cell("M9").Value = "Сумма неустойки";
            sheet.Cell("N9").Value = "Сумма процентов";
            sheet.Cell("O9").Value = "Валюта требований";
            sheet.Cell("P9").Value = "Предъявлен иск";
            sheet.Cell("Q9").Value = "Иск не предъявлен";

            // Стилизация шапки (без жирного)
            var headerRange = sheet.Range("B7:Q9");
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Alignment.WrapText = true;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // увеличил высоту строк шапки
            sheet.Row(7).Height = 25;
            sheet.Row(8).Height = 45;
            sheet.Row(9).Height = 45;

            //---------------------------------------------------------------
            List<Organization> listorganization = new List<Organization>();
            listorganization = db.Organizations.ToList();

            List<Valutum> listvaluta = new List<Valutum>();
            listvaluta = db.Valuta.OrderBy(l => l.Name).ToList();

            List<Filial> listfilial = new List<Filial>();
            listfilial = db.Filials.ToList();

            List<Predmet> listpredmet = new List<Predmet>();
            listpredmet = db.Predmets.ToList();

            List<Status> liststatus = new List<Status>();
            liststatus = db.Statuses.ToList();

            List<Pretense> listpretense = new List<Pretense>();
            listpretense = db.Pretenses.Where(j => j.DatePret <= reportDate && j.Visible == 1 && j.Arhiv == 0).ToList();
            List<Summa> listsumma = new List<Summa>(); listsumma = db.Summas.ToList();

            List<PretenseSumma> listpretensesumma = new List<PretenseSumma>();
            listpretensesumma = db.PretenseSummas.ToList();

            List<TablePretense> listtablepretense = new List<TablePretense>();
            listtablepretense = db.TablePretenses.Where(g => g.DateTabPret <= reportDate && g.Delet != 1).ToList();

            List<ResultSumma> listresultsumma = new List<ResultSumma>();
            listresultsumma = db.ResultSummas.ToList();

            List<TablePretense> listtablepretenseOrderBy = new List<TablePretense>();
            listtablepretenseOrderBy = db.TablePretenses.Where(g => g.DateTabPret <= reportDate && g.Delet != 1).OrderByDescending(o => o.DateTabPret).ToList();

            List<Models.City> listcity = new List<Models.City>();
            listcity = db.Cities.ToList();

            List<Country> listcountry = new List<Country>();
            listcountry = db.Countries.ToList();

            var listpretenseJoin =
    (from pretense in listpretense
     join organization in listorganization on pretense.OrgId equals organization.OrgId
     join city in listcity on organization.CityId equals city.CityId into cityJoin
     from city in cityJoin.DefaultIfEmpty()
     join country in listcountry on city.CountryId equals country.CountryId into countryJoin
     from country in countryJoin.DefaultIfEmpty()
     join predmet in listpredmet on pretense.PredmetId equals predmet.PredmetId

     // Суммы по претензии
     let currencyGroups =
         (from ps in listpretensesumma
          where ps.PretId == pretense.PretId
          join v in listvaluta on ps.ValId equals v.ValId
          join s in listsumma on ps.SummaId equals s.SummaId
          group new { ps, v, s } by new { v.ValId, v.Name } into g
          select new
          {
              CurrencyId = g.Key.ValId,
              CurrencyName = g.Key.Name,
              SummaDolg = g.Where(x => x.s.SummaId == 1).Sum(x => x.ps.Value) ?? 0,
              SummaPeny = g.Where(x => x.s.SummaId == 2).Sum(x => x.ps.Value) ?? 0,
              SummaProc = g.Where(x => x.s.SummaId == 3).Sum(x => x.ps.Value) ?? 0,
              SummaPoshlina = g.Where(x => x.s.SummaId == 4).Sum(x => x.ps.Value) ?? 0,
              SummaItog = g.Sum(x => x.ps.Value) ?? 0
          }).ToList()

     // История TablePretense
     let tablePretenseList =
         (from tp in listtablepretense
          where tp.PretId == pretense.PretId
          join st in liststatus on tp.StatusId equals st.StatusId
          let resultGroups =
              (from rs in listresultsumma
               where rs.ResultId == tp.TabPretId
               join v in listvaluta on rs.ValId equals v.ValId
               join s in listsumma on rs.SummaId equals s.SummaId
               group new { rs, v, s } by new { v.ValId, v.Name } into g
               select new
               {
                   CurrencyId = g.Key.ValId,
                   CurrencyName = g.Key.Name,
                   SummaDolg = g.Where(x => x.s.SummaId == 1).Sum(x => x.rs.Value) ?? 0,
                   SummaPeny = g.Where(x => x.s.SummaId == 2).Sum(x => x.rs.Value) ?? 0,
                   SummaProc = g.Where(x => x.s.SummaId == 3).Sum(x => x.rs.Value) ?? 0,
                   SummaPoshlina = g.Where(x => x.s.SummaId == 4).Sum(x => x.rs.Value) ?? 0,
                   SummaItog = g.Sum(x => x.rs.Value) ?? 0
               }).ToList()
          select new
          {
              tp.TabPretId,
              tp.DateTabPret,
              tp.Result,
              tp.Primechanie,
              tp.StatusId,
              StatusName = st.Name,
              tp.UserMod,
              tp.DateMod,
              ResultCurrencyGroups = resultGroups
          }).OrderBy(x => x.DateTabPret).ToList()

     select new
     {
         pretense.PretId,
         pretense.NumberPret,
         pretense.DatePret,
         pretense.Inout,
         pretense.Visible,
         pretense.Arhiv,
         OrgId = organization.OrgId,
         OrgName = organization.Name,
         UNP = organization.Unp,
         Address = organization.Address,
         CityName = city != null ? city.Name : string.Empty,
         CountryName = country != null ? country.Name : string.Empty,
         PredmetName = predmet.Predmet1,
         UserMod = pretense.UserMod,
         DateMod = pretense.DateMod,
         CurrencyGroups = currencyGroups,
         TablePretenseList = tablePretenseList
     })
    .Where(p => p.Visible == 1 && p.Arhiv == 0)
    .OrderBy(p => p.DatePret)
    .ToList();
            //--------------------------------------------------------------

            var dataRow = 10;
            int counter = 1;

            var totalRequested = new Dictionary<string, decimal>(); // Валюта -> Сумма
            var totalSatisfied = new Dictionary<string, decimal>(); // Валюта -> Сумма

            foreach (var pret in listpretenseJoin)
            {
                var lastTp1 = pret.TablePretenseList?
                    .Where(t => t.DateTabPret != null)
                    .OrderByDescending(t => t.DateTabPret)
                    .FirstOrDefault();

                if (lastTp1 != null && (lastTp1.StatusId == 1 || lastTp1.StatusId == 3 || lastTp1.StatusId == 6 || lastTp1.StatusId == 9))
                {
                    sheet.Cell(dataRow, 2).Value = counter;
                    sheet.Cell(dataRow, 3).Value = $"{pret.OrgName}";
                    sheet.Cell(dataRow, 4).Value = $"{pret.CityName} ({pret.CountryName})";
                    sheet.Cell(dataRow, 5).Value = pret.DatePret?.ToString("dd.MM.yyyy");
                    sheet.Cell(dataRow, 6).Value = pret.PredmetName;

                    //----Работаем с суммами группированными по валютам--------------------------------
                    if (pret.CurrencyGroups != null && pret.CurrencyGroups.Count > 0)
                    {
                        if (pret.CurrencyGroups.Count == 1)
                        {
                            var cg = pret.CurrencyGroups[0];

                            sheet.Cell(dataRow, 7).Value = cg.SummaDolg;
                            sheet.Cell(dataRow, 8).Value = cg.SummaPeny;
                            sheet.Cell(dataRow, 9).Value = cg.SummaProc;
                            sheet.Cell(dataRow, 10).Value = cg.CurrencyName;

                            // ДОБАВЛЕНО: Собираем итоги по заявленным требованиям
                            AddToTotals(totalRequested, cg.CurrencyName, cg.SummaDolg);
                            AddToTotals(totalRequested, cg.CurrencyName, cg.SummaPeny);
                            AddToTotals(totalRequested, cg.CurrencyName, cg.SummaProc);
                        }
                        else
                        {
                            // Формируем текстовые представления
                            string dolgText = string.Join("; ",
                                pret.CurrencyGroups
                                    .Where(x => x.SummaDolg != 0)
                                    .Select(x => $"{x.SummaDolg} {x.CurrencyName}"));

                            string penyText = string.Join("; ",
                                pret.CurrencyGroups
                                    .Where(x => x.SummaPeny != 0)
                                    .Select(x => $"{x.SummaPeny} {x.CurrencyName}"));

                            string procText = string.Join("; ",
                                pret.CurrencyGroups
                                    .Where(x => x.SummaProc != 0)
                                    .Select(x => $"{x.SummaProc} {x.CurrencyName}"));

                            sheet.Cell(dataRow, 7).Value = string.IsNullOrWhiteSpace(dolgText) ? "-" : dolgText;
                            sheet.Cell(dataRow, 8).Value = string.IsNullOrWhiteSpace(penyText) ? "-" : penyText;
                            sheet.Cell(dataRow, 9).Value = string.IsNullOrWhiteSpace(procText) ? "-" : procText;

                            var currenciesText = string.Join("; ",
                                pret.CurrencyGroups
                                    .Select(x => x.CurrencyName)
                                    .Distinct());
                            sheet.Cell(dataRow, 10).Value = currenciesText;

                            // ДОБАВЛЕНО: Собираем итоги по заявленным требованиям для всех валют
                            foreach (var cg in pret.CurrencyGroups)
                            {
                                AddToTotals(totalRequested, cg.CurrencyName, cg.SummaDolg);
                                AddToTotals(totalRequested, cg.CurrencyName, cg.SummaPeny);
                                AddToTotals(totalRequested, cg.CurrencyName, cg.SummaProc);
                            }
                        }
                    }
                    else
                    {
                        sheet.Cell(dataRow, 7).Value = "-";
                        sheet.Cell(dataRow, 8).Value = "-";
                        sheet.Cell(dataRow, 9).Value = "-";
                        sheet.Cell(dataRow, 10).Value = "-";
                    }

                    sheet.Cell(dataRow, 11).Value = pret.TablePretenseList.FirstOrDefault()?.DateTabPret?.ToString("dd.MM.yyyy");

                    //--------Результаты рассмотрения--------------------------------------------------------
                    if (pret.TablePretenseList != null && pret.TablePretenseList.Count > 0)
                    {
                        var allResults = pret.TablePretenseList
                            .SelectMany(tp => tp.ResultCurrencyGroups
                                .Select(rcg => new
                                {
                                    rcg.CurrencyName,
                                    rcg.SummaDolg,
                                    rcg.SummaPeny,
                                    rcg.SummaProc
                                }))
                            .GroupBy(x => x.CurrencyName)
                            .Select(g => new
                            {
                                CurrencyName = g.Key,
                                SummaDolg = g.Sum(x => x.SummaDolg),
                                SummaPeny = g.Sum(x => x.SummaPeny),
                                SummaProc = g.Sum(x => x.SummaProc)
                            })
                            .ToList();

                        if (allResults.Count == 1)
                        {
                            var r = allResults[0];
                            sheet.Cell(dataRow, 12).Value = r.SummaDolg;
                            sheet.Cell(dataRow, 13).Value = r.SummaPeny;
                            sheet.Cell(dataRow, 14).Value = r.SummaProc;
                            sheet.Cell(dataRow, 15).Value = r.CurrencyName;

                            // ДОБАВЛЕНО: Собираем итоги по удовлетворенным требованиям
                            AddToTotals(totalSatisfied, r.CurrencyName, r.SummaDolg);
                            AddToTotals(totalSatisfied, r.CurrencyName, r.SummaPeny);
                            AddToTotals(totalSatisfied, r.CurrencyName, r.SummaProc);
                        }
                        else
                        {
                            string dolgText = string.Join("; ",
                                allResults.Where(x => x.SummaDolg != 0)
                                          .Select(x => $"{x.SummaDolg} {x.CurrencyName}"));

                            string penyText = string.Join("; ",
                                allResults.Where(x => x.SummaPeny != 0)
                                          .Select(x => $"{x.SummaPeny} {x.CurrencyName}"));

                            string procText = string.Join("; ",
                                allResults.Where(x => x.SummaProc != 0)
                                          .Select(x => $"{x.SummaProc} {x.CurrencyName}"));

                            sheet.Cell(dataRow, 12).Value = string.IsNullOrWhiteSpace(dolgText) ? "-" : dolgText;
                            sheet.Cell(dataRow, 13).Value = string.IsNullOrWhiteSpace(penyText) ? "-" : penyText;
                            sheet.Cell(dataRow, 14).Value = string.IsNullOrWhiteSpace(procText) ? "-" : procText;
                            sheet.Cell(dataRow, 15).Value = string.Join("; ", allResults.Select(x => x.CurrencyName).Distinct());

                            // ДОБАВЛЕНО: Собираем итоги по удовлетворенным требованиям для всех валют
                            foreach (var r in allResults)
                            {
                                AddToTotals(totalSatisfied, r.CurrencyName, r.SummaDolg);
                                AddToTotals(totalSatisfied, r.CurrencyName, r.SummaPeny);
                                AddToTotals(totalSatisfied, r.CurrencyName, r.SummaProc);
                            }
                        }

                        var lastTp = pret.TablePretenseList.OrderByDescending(t => t.DateTabPret).FirstOrDefault();
                        if (lastTp != null)
                        {
                            sheet.Cell(dataRow, 16).Value = "";
                            sheet.Cell(dataRow, 17).Value = "";
                        }
                    }

                    var dataRange1 = sheet.Range($"B{dataRow}:Q{dataRow}");
                    dataRange1.Style.Alignment.WrapText = true;
                    dataRange1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    dataRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    dataRange1.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    OptimizeRowHeight(sheet, dataRow);

                    dataRow++;
                    counter++;
                }
            }
            //-----------------------------------
            var allCurrencies = totalRequested.Keys.Union(totalSatisfied.Keys).Distinct().ToList();
            int totalRowsCount = allCurrencies.Count;

            if (totalRowsCount > 0)
            {
                int totalStartRow = dataRow;

                // Заголовок "ИТОГО" объединяем на все строки валют
                sheet.Range(totalStartRow, 2, totalStartRow + totalRowsCount - 1, 6).Merge();
                sheet.Cell(totalStartRow, 2).Value = "ИТОГО";
                sheet.Cell(totalStartRow, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Cell(totalStartRow, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                sheet.Cell(totalStartRow, 2).Style.Font.Bold = true;

                // Заполняем данные по каждой валюте
                for (int i = 0; i < totalRowsCount; i++)
                {
                    int currentRow = totalStartRow + i;
                    string currency = allCurrencies[i];

                    // Название валюты в колонке J
                    sheet.Cell(currentRow, 10).Value = currency;
                    sheet.Cell(currentRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Cell(currentRow, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    // Суммы по заявленным требованиям
                    decimal requestedDolg = 0;
                    decimal requestedPeny = 0;
                    decimal requestedProc = 0;

                    // Пересчитываем суммы для этой валюты из исходных данных
                    foreach (var pret in listpretenseJoin)
                    {
                        var lastTp1 = pret.TablePretenseList?
                            .Where(t => t.DateTabPret != null)
                            .OrderByDescending(t => t.DateTabPret)
                            .FirstOrDefault();

                        if (lastTp1 != null && pret.Inout == 1 && (lastTp1.StatusId == 1 || lastTp1.StatusId == 3 || lastTp1.StatusId == 6 || lastTp1.StatusId == 9))
                        {
                            if (pret.CurrencyGroups != null)
                            {
                                foreach (var cg in pret.CurrencyGroups)
                                {
                                    if (cg.CurrencyName == currency)
                                    {
                                        requestedDolg += cg.SummaDolg;
                                        requestedPeny += cg.SummaPeny;
                                        requestedProc += cg.SummaProc;
                                    }
                                }
                            }
                        }
                    }

                    // Записываем суммы заявленных требований
                    sheet.Cell(currentRow, 7).Value = requestedDolg != 0 ? requestedDolg : "-";
                    sheet.Cell(currentRow, 8).Value = requestedPeny != 0 ? requestedPeny : "-";
                    sheet.Cell(currentRow, 9).Value = requestedProc != 0 ? requestedProc : "-";

                    // Суммы по удовлетворенным требованиям
                    decimal satisfiedDolg = 0;
                    decimal satisfiedPeny = 0;
                    decimal satisfiedProc = 0;

                    // Пересчитываем суммы для этой валюты из результатов
                    foreach (var pret in listpretenseJoin)
                    {
                        var lastTp1 = pret.TablePretenseList?
                            .Where(t => t.DateTabPret != null)
                            .OrderByDescending(t => t.DateTabPret)
                            .FirstOrDefault();

                        if (lastTp1 != null && (lastTp1.StatusId == 1 || lastTp1.StatusId == 3 || lastTp1.StatusId == 6 || lastTp1.StatusId == 9))
                        {
                            if (pret.TablePretenseList != null && pret.TablePretenseList.Count > 0)
                            {
                                var allResults = pret.TablePretenseList
                                    .SelectMany(tp => tp.ResultCurrencyGroups
                                        .Select(rcg => new
                                        {
                                            rcg.CurrencyName,
                                            rcg.SummaDolg,
                                            rcg.SummaPeny,
                                            rcg.SummaProc
                                        }))
                                    .GroupBy(x => x.CurrencyName)
                                    .Select(g => new
                                    {
                                        CurrencyName = g.Key,
                                        SummaDolg = g.Sum(x => x.SummaDolg),
                                        SummaPeny = g.Sum(x => x.SummaPeny),
                                        SummaProc = g.Sum(x => x.SummaProc)
                                    })
                                    .ToList();

                                foreach (var result in allResults)
                                {
                                    if (result.CurrencyName == currency)
                                    {
                                        satisfiedDolg += result.SummaDolg;
                                        satisfiedPeny += result.SummaPeny;
                                        satisfiedProc += result.SummaProc;
                                    }
                                }
                            }
                        }
                    }

                    // Записываем суммы удовлетворенных требований
                    sheet.Cell(currentRow, 12).Value = satisfiedDolg != 0 ? satisfiedDolg : "-";
                    sheet.Cell(currentRow, 13).Value = satisfiedPeny != 0 ? satisfiedPeny : "-";
                    sheet.Cell(currentRow, 14).Value = satisfiedProc != 0 ? satisfiedProc : "-";

                    // Валюта требований для удовлетворенных (дублируем название валюты)
                    sheet.Cell(currentRow, 15).Value = currency;
                    sheet.Cell(currentRow, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Cell(currentRow, 15).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    // Пустые значения для колонок "Предъявлен иск" и "Иск не предъявлен"
                    sheet.Cell(currentRow, 16).Value = "";
                    sheet.Cell(currentRow, 17).Value = "";

                    // Стили для строки итогов
                    var rowRange = sheet.Range(currentRow, 2, currentRow, 17);
                    rowRange.Style.Alignment.WrapText = true;
                    rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                                        
                    //rowRange.Style.Fill.BackgroundColor = currentRow % 2 == 0 ? XLColor.LightGray : XLColor.White;
                    rowRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                    // Автоматическая высота строки
                    OptimizeRowHeight(sheet, currentRow);
                }

                // Общие стили для блока итогов
                var totalRange = sheet.Range(totalStartRow, 2, totalStartRow + totalRowsCount - 1, 17);
                totalRange.Style.Font.Bold = true;

                // Толстая граница сверху всего блока итогов
                sheet.Range(totalStartRow, 2, totalStartRow, 17).Style.Border.TopBorder = XLBorderStyleValues.Medium;

                dataRow = totalStartRow + totalRowsCount;
            }
            else
            {
                // Если нет валют, создаем одну строку с прочерками
                int totalRow = dataRow;
                sheet.Range(totalRow, 2, totalRow, 6).Merge().Value = "ИТОГО";
                sheet.Range(totalRow, 7, totalRow, 10).Merge().Value = "-";
                sheet.Range(totalRow, 12, totalRow, 15).Merge().Value = "-";

                var totalRange = sheet.Range(totalRow, 2, totalRow, 17);
                totalRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                totalRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                totalRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                totalRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                totalRange.Style.Font.Bold = true;

                dataRow = totalRow + 1;
            }

            //-----------------------------------
            // Примечание после таблицы (без границ)
            int noteRow = dataRow;
            sheet.Cell(noteRow, 2).Value = "* - Комиссией по противодействию коррупции ОАО \"Гомельтранснефть Дружба\" принято решение дальнейшую претензионно-исковую работу не проводить.";
            sheet.Range(noteRow, 2, noteRow, 17).Merge();
            sheet.Range(noteRow, 2, noteRow, 17).Style.Alignment.WrapText = true;
            sheet.Range(noteRow, 2, noteRow, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            // Убираем границы для примечания
            sheet.Range(noteRow, 2, noteRow, 17).Style.Border.TopBorder = XLBorderStyleValues.None;
            sheet.Range(noteRow, 2, noteRow, 17).Style.Border.BottomBorder = XLBorderStyleValues.None;
            sheet.Range(noteRow, 2, noteRow, 17).Style.Border.LeftBorder = XLBorderStyleValues.None;
            sheet.Range(noteRow, 2, noteRow, 17).Style.Border.RightBorder = XLBorderStyleValues.None;

            // Подпись на следующей строке (без границ)
            int signatureRow = dataRow + 1;
            sheet.Cell(signatureRow, 2).Value = "Начальник юридического отдела ОАО \"Гомельтранснефть Дружба\"                                                    Ю.А.Лащенко";
            sheet.Range(signatureRow, 2, signatureRow, 17).Merge();
            sheet.Range(signatureRow, 2, signatureRow, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            sheet.Row(signatureRow).Height = 25;
            // Убираем границы для подписи
            sheet.Range(signatureRow, 2, signatureRow, 17).Style.Border.TopBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow, 2, signatureRow, 17).Style.Border.BottomBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow, 2, signatureRow, 17).Style.Border.LeftBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow, 2, signatureRow, 17).Style.Border.RightBorder = XLBorderStyleValues.None;

            //---------------Здесь формируем вторую таблицу Форма №2----------------------
            //*************************************************************
            dataRow = signatureRow + 2;

            // Создаем вторую таблицу "Форма №2 Претензии, предъявленные к организации"
            sheet.Range($"B{dataRow}:Q{dataRow}").Merge().Value = "Форма № 2 Претензии, предъявленные к организации";
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Font.FontSize = 14;
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Font.Bold = true;
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Alignment.WrapText = true;

            dataRow++;

            var startDate2 = new DateTime(reportDate.Year, 1, 1);
            var endDate2 = reportDate;

            sheet.Range($"B{dataRow}:Q{dataRow}").Merge().Value = $"в период с {startDate2:dd.MM.yyyy} по {endDate2:dd.MM.yyyy}";
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Font.FontSize = 14;
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Font.Bold = true;
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range($"B{dataRow}:Q{dataRow}").Style.Alignment.WrapText = true;

            dataRow += 2; // Пропускаем 2 строки перед шапкой таблицы

            // Шапка таблицы (такая же как у первой таблицы) - ВСЕГДА ОТОБРАЖАЕМ
            sheet.Range($"B{dataRow}:B{dataRow + 2}").Merge().Value = "№";
            sheet.Range($"C{dataRow}:C{dataRow + 2}").Merge().Value = "Наименование должника";
            sheet.Range($"D{dataRow}:D{dataRow + 2}").Merge().Value = "Город (Страна)";
            sheet.Range($"E{dataRow}:E{dataRow + 2}").Merge().Value = "Дата предъявления претензии";
            sheet.Range($"K{dataRow}:K{dataRow + 2}").Merge().Value = "Срок рассмотрения претензии";

            sheet.Range($"F{dataRow}:J{dataRow}").Merge().Value = "Заявлены требования";
            sheet.Range($"F{dataRow + 1}:F{dataRow + 2}").Merge().Value = "Содержание требований";
            sheet.Range($"G{dataRow + 1}:G{dataRow + 2}").Merge().Value = "Сумма основного долга";
            sheet.Range($"H{dataRow + 1}:H{dataRow + 2}").Merge().Value = "Сумма неустойки";
            sheet.Range($"I{dataRow + 1}:I{dataRow + 2}").Merge().Value = "Сумма процентов";
            sheet.Range($"J{dataRow + 1}:J{dataRow + 2}").Merge().Value = "Валюта требований";

            sheet.Range($"L{dataRow}:Q{dataRow}").Merge().Value = "Результат рассмотрения претензии должником";
            sheet.Range($"L{dataRow + 1}:O{dataRow + 1}").Merge().Value = "Удовлетворена на сумму";
            sheet.Range($"P{dataRow + 1}:Q{dataRow + 1}").Merge().Value = "Не удовлетворены";

            sheet.Cell($"L{dataRow + 2}").Value = "Сумма основного долга";
            sheet.Cell($"M{dataRow + 2}").Value = "Сумма неустойки";
            sheet.Cell($"N{dataRow + 2}").Value = "Сумма процентов";
            sheet.Cell($"O{dataRow + 2}").Value = "Валюта требований";
            sheet.Cell($"P{dataRow + 2}").Value = "Предъявлен иск";
            sheet.Cell($"Q{dataRow + 2}").Value = "Иск не предъявлен";

            // Стилизация шапки
            var headerRange2 = sheet.Range($"B{dataRow}:Q{dataRow + 2}");
            headerRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange2.Style.Alignment.WrapText = true;
            headerRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            headerRange2.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Высота строк шапки
            sheet.Row(dataRow).Height = 25;
            sheet.Row(dataRow + 1).Height = 45;
            sheet.Row(dataRow + 2).Height = 45;

            dataRow += 3; // Переходим к данным таблицы

            // Проверяем есть ли данные для второй таблицы (pret.Inout == 0)
            bool hasDataForSecondTable = listpretenseJoin.Any(pret =>
            {
                var lastTp1 = pret.TablePretenseList?
                    .Where(t => t.DateTabPret != null)
                    .OrderByDescending(t => t.DateTabPret)
                    .FirstOrDefault();

                return lastTp1 != null && pret.Inout == 0 && (lastTp1.StatusId == 1 || lastTp1.StatusId == 3 || lastTp1.StatusId == 6 || lastTp1.StatusId == 9);
            });

            if (hasDataForSecondTable)
            {
                int counter2 = 1;
                var totalRequested2 = new Dictionary<string, decimal>();
                var totalSatisfied2 = new Dictionary<string, decimal>();

                // Заполняем данные для второй таблицы (pret.Inout == 0)
                foreach (var pret in listpretenseJoin)
                {
                    var lastTp1 = pret.TablePretenseList?
                        .Where(t => t.DateTabPret != null)
                        .OrderByDescending(t => t.DateTabPret)
                        .FirstOrDefault();

                    // ОТЛИЧИЕ: pret.Inout == 0 вместо pret.Inout == 1
                    if (lastTp1 != null && pret.Inout == 0 && (lastTp1.StatusId == 1 || lastTp1.StatusId == 3 || lastTp1.StatusId == 6 || lastTp1.StatusId == 9))
                    {
                        // ... весь код заполнения данных таблицы (такой же как у вас в первой таблице) ...
                        sheet.Cell(dataRow, 2).Value = counter2;
                        sheet.Cell(dataRow, 3).Value = $"{pret.OrgName}";
                        sheet.Cell(dataRow, 4).Value = $"{pret.CityName} ({pret.CountryName})";
                        sheet.Cell(dataRow, 5).Value = pret.DatePret?.ToString("dd.MM.yyyy");
                        sheet.Cell(dataRow, 6).Value = pret.PredmetName;

                        //----Работаем с суммами группированными по валютам--------------------------------
                        if (pret.CurrencyGroups != null && pret.CurrencyGroups.Count > 0)
                        {
                            if (pret.CurrencyGroups.Count == 1)
                            {
                                var cg = pret.CurrencyGroups[0];

                                sheet.Cell(dataRow, 7).Value = cg.SummaDolg;
                                sheet.Cell(dataRow, 8).Value = cg.SummaPeny;
                                sheet.Cell(dataRow, 9).Value = cg.SummaProc;
                                sheet.Cell(dataRow, 10).Value = cg.CurrencyName;

                                // Собираем итоги по заявленным требованиям
                                AddToTotals(totalRequested2, cg.CurrencyName, cg.SummaDolg);
                                AddToTotals(totalRequested2, cg.CurrencyName, cg.SummaPeny);
                                AddToTotals(totalRequested2, cg.CurrencyName, cg.SummaProc);
                            }
                            else
                            {
                                // Формируем текстовые представления
                                string dolgText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaDolg != 0)
                                        .Select(x => $"{x.SummaDolg} {x.CurrencyName}"));

                                string penyText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaPeny != 0)
                                        .Select(x => $"{x.SummaPeny} {x.CurrencyName}"));

                                string procText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaProc != 0)
                                        .Select(x => $"{x.SummaProc} {x.CurrencyName}"));

                                sheet.Cell(dataRow, 7).Value = string.IsNullOrWhiteSpace(dolgText) ? "-" : dolgText;
                                sheet.Cell(dataRow, 8).Value = string.IsNullOrWhiteSpace(penyText) ? "-" : penyText;
                                sheet.Cell(dataRow, 9).Value = string.IsNullOrWhiteSpace(procText) ? "-" : procText;

                                var currenciesText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Select(x => x.CurrencyName)
                                        .Distinct());
                                sheet.Cell(dataRow, 10).Value = currenciesText;

                                // Собираем итоги по заявленным требованиям для всех валют
                                foreach (var cg in pret.CurrencyGroups)
                                {
                                    AddToTotals(totalRequested2, cg.CurrencyName, cg.SummaDolg);
                                    AddToTotals(totalRequested2, cg.CurrencyName, cg.SummaPeny);
                                    AddToTotals(totalRequested2, cg.CurrencyName, cg.SummaProc);
                                }
                            }
                        }
                        else
                        {
                            sheet.Cell(dataRow, 7).Value = "-";
                            sheet.Cell(dataRow, 8).Value = "-";
                            sheet.Cell(dataRow, 9).Value = "-";
                            sheet.Cell(dataRow, 10).Value = "-";
                        }

                        sheet.Cell(dataRow, 11).Value = pret.TablePretenseList.FirstOrDefault()?.DateTabPret?.ToString("dd.MM.yyyy");

                        //--------Результаты рассмотрения--------------------------------------------------------
                        if (pret.TablePretenseList != null && pret.TablePretenseList.Count > 0)
                        {
                            var allResults = pret.TablePretenseList
                                .SelectMany(tp => tp.ResultCurrencyGroups
                                    .Select(rcg => new
                                    {
                                        rcg.CurrencyName,
                                        rcg.SummaDolg,
                                        rcg.SummaPeny,
                                        rcg.SummaProc
                                    }))
                                .GroupBy(x => x.CurrencyName)
                                .Select(g => new
                                {
                                    CurrencyName = g.Key,
                                    SummaDolg = g.Sum(x => x.SummaDolg),
                                    SummaPeny = g.Sum(x => x.SummaPeny),
                                    SummaProc = g.Sum(x => x.SummaProc)
                                })
                                .ToList();

                            if (allResults.Count == 1)
                            {
                                var r = allResults[0];
                                sheet.Cell(dataRow, 12).Value = r.SummaDolg;
                                sheet.Cell(dataRow, 13).Value = r.SummaPeny;
                                sheet.Cell(dataRow, 14).Value = r.SummaProc;
                                sheet.Cell(dataRow, 15).Value = r.CurrencyName;

                                // Собираем итоги по удовлетворенным требованиям
                                AddToTotals(totalSatisfied2, r.CurrencyName, r.SummaDolg);
                                AddToTotals(totalSatisfied2, r.CurrencyName, r.SummaPeny);
                                AddToTotals(totalSatisfied2, r.CurrencyName, r.SummaProc);
                            }
                            else
                            {
                                string dolgText = string.Join("; ",
                                    allResults.Where(x => x.SummaDolg != 0)
                                              .Select(x => $"{x.SummaDolg} {x.CurrencyName}"));

                                string penyText = string.Join("; ",
                                    allResults.Where(x => x.SummaPeny != 0)
                                              .Select(x => $"{x.SummaPeny} {x.CurrencyName}"));

                                string procText = string.Join("; ",
                                    allResults.Where(x => x.SummaProc != 0)
                                              .Select(x => $"{x.SummaProc} {x.CurrencyName}"));

                                sheet.Cell(dataRow, 12).Value = string.IsNullOrWhiteSpace(dolgText) ? "-" : dolgText;
                                sheet.Cell(dataRow, 13).Value = string.IsNullOrWhiteSpace(penyText) ? "-" : penyText;
                                sheet.Cell(dataRow, 14).Value = string.IsNullOrWhiteSpace(procText) ? "-" : procText;
                                sheet.Cell(dataRow, 15).Value = string.Join("; ", allResults.Select(x => x.CurrencyName).Distinct());

                                // Собираем итоги по удовлетворенным требованиям для всех валют
                                foreach (var r in allResults)
                                {
                                    AddToTotals(totalSatisfied2, r.CurrencyName, r.SummaDolg);
                                    AddToTotals(totalSatisfied2, r.CurrencyName, r.SummaPeny);
                                    AddToTotals(totalSatisfied2, r.CurrencyName, r.SummaProc);
                                }
                            }

                            var lastTp = pret.TablePretenseList.OrderByDescending(t => t.DateTabPret).FirstOrDefault();
                            if (lastTp != null)
                            {
                                sheet.Cell(dataRow, 16).Value = "";
                                sheet.Cell(dataRow, 17).Value = "";
                            }
                        }

                        var dataRange2 = sheet.Range($"B{dataRow}:Q{dataRow}");
                        dataRange2.Style.Alignment.WrapText = true;
                        dataRange2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        dataRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        dataRange2.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        OptimizeRowHeight(sheet, dataRow);

                        dataRow++;
                        counter2++;
                    }
                }

                // Итоги для второй таблицы
                var allCurrencies2 = totalRequested2.Keys.Union(totalSatisfied2.Keys).Distinct().ToList();
                int totalRowsCount2 = allCurrencies2.Count;

                if (totalRowsCount2 > 0)
                {
                    int totalStartRow2 = dataRow;

                    // Заголовок "ИТОГО" объединяем на все строки валют
                    sheet.Range(totalStartRow2, 2, totalStartRow2 + totalRowsCount2 - 1, 6).Merge();
                    sheet.Cell(totalStartRow2, 2).Value = "ИТОГО";
                    sheet.Cell(totalStartRow2, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Cell(totalStartRow2, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    sheet.Cell(totalStartRow2, 2).Style.Font.Bold = true;

                    // Заполняем данные по каждой валюте
                    for (int i = 0; i < totalRowsCount2; i++)
                    {
                        int currentRow = totalStartRow2 + i;
                        string currency = allCurrencies2[i];

                        // Название валюты в колонке J
                        sheet.Cell(currentRow, 10).Value = currency;
                        sheet.Cell(currentRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        sheet.Cell(currentRow, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Суммы по заявленным требованиям
                        decimal requestedDolg = 0;
                        decimal requestedPeny = 0;
                        decimal requestedProc = 0;

                        // Пересчитываем суммы для этой валюты из исходных данных (pret.Inout == 0)
                        foreach (var pret in listpretenseJoin)
                        {
                            var lastTp1 = pret.TablePretenseList?
                                .Where(t => t.DateTabPret != null)
                                .OrderByDescending(t => t.DateTabPret)
                                .FirstOrDefault();

                            if (lastTp1 != null && pret.Inout == 0 && (lastTp1.StatusId == 1 || lastTp1.StatusId == 3 || lastTp1.StatusId == 6 || lastTp1.StatusId == 9))
                            {
                                if (pret.CurrencyGroups != null)
                                {
                                    foreach (var cg in pret.CurrencyGroups)
                                    {
                                        if (cg.CurrencyName == currency)
                                        {
                                            requestedDolg += cg.SummaDolg;
                                            requestedPeny += cg.SummaPeny;
                                            requestedProc += cg.SummaProc;
                                        }
                                    }
                                }
                            }
                        }

                        // Записываем суммы заявленных требований
                        sheet.Cell(currentRow, 7).Value = requestedDolg != 0 ? requestedDolg : "-";
                        sheet.Cell(currentRow, 8).Value = requestedPeny != 0 ? requestedPeny : "-";
                        sheet.Cell(currentRow, 9).Value = requestedProc != 0 ? requestedProc : "-";

                        // Суммы по удовлетворенным требованиям
                        decimal satisfiedDolg = 0;
                        decimal satisfiedPeny = 0;
                        decimal satisfiedProc = 0;

                        // Пересчитываем суммы для этой валюты из результатов (pret.Inout == 0)
                        foreach (var pret in listpretenseJoin)
                        {
                            var lastTp1 = pret.TablePretenseList?
                                .Where(t => t.DateTabPret != null)
                                .OrderByDescending(t => t.DateTabPret)
                                .FirstOrDefault();

                            if (lastTp1 != null && pret.Inout == 0 && (lastTp1.StatusId == 1 || lastTp1.StatusId == 3 || lastTp1.StatusId == 6 || lastTp1.StatusId == 9))
                            {
                                if (pret.TablePretenseList != null && pret.TablePretenseList.Count > 0)
                                {
                                    var allResults = pret.TablePretenseList
                                        .SelectMany(tp => tp.ResultCurrencyGroups
                                            .Select(rcg => new
                                            {
                                                rcg.CurrencyName,
                                                rcg.SummaDolg,
                                                rcg.SummaPeny,
                                                rcg.SummaProc
                                            }))
                                        .GroupBy(x => x.CurrencyName)
                                        .Select(g => new
                                        {
                                            CurrencyName = g.Key,
                                            SummaDolg = g.Sum(x => x.SummaDolg),
                                            SummaPeny = g.Sum(x => x.SummaPeny),
                                            SummaProc = g.Sum(x => x.SummaProc)
                                        })
                                        .ToList();

                                    foreach (var result in allResults)
                                    {
                                        if (result.CurrencyName == currency)
                                        {
                                            satisfiedDolg += result.SummaDolg;
                                            satisfiedPeny += result.SummaPeny;
                                            satisfiedProc += result.SummaProc;
                                        }
                                    }
                                }
                            }
                        }

                        // Записываем суммы удовлетворенных требований
                        sheet.Cell(currentRow, 12).Value = satisfiedDolg != 0 ? satisfiedDolg : "-";
                        sheet.Cell(currentRow, 13).Value = satisfiedPeny != 0 ? satisfiedPeny : "-";
                        sheet.Cell(currentRow, 14).Value = satisfiedProc != 0 ? satisfiedProc : "-";

                        // Валюта требований для удовлетворенных
                        sheet.Cell(currentRow, 15).Value = currency;
                        sheet.Cell(currentRow, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        sheet.Cell(currentRow, 15).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Пустые значения для колонок
                        sheet.Cell(currentRow, 16).Value = "";
                        sheet.Cell(currentRow, 17).Value = "";

                        // Стили для строки итогов
                        var rowRange = sheet.Range(currentRow, 2, currentRow, 17);
                        rowRange.Style.Alignment.WrapText = true;
                        rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        rowRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                        OptimizeRowHeight(sheet, currentRow);
                    }

                    // Общие стили для блока итогов
                    var totalRange2 = sheet.Range(totalStartRow2, 2, totalStartRow2 + totalRowsCount2 - 1, 17);
                    totalRange2.Style.Font.Bold = true;

                    // Толстая граница сверху всего блока итогов
                    sheet.Range(totalStartRow2, 2, totalStartRow2, 17).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                }
            }
            else
            {
                // Если данных нет - выводим сообщение "Данные отсутствуют" под шапкой
                sheet.Range($"B{dataRow}:Q{dataRow}").Merge().Value = "Данные отсутствуют";
                sheet.Range($"B{dataRow}:Q{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Range($"B{dataRow}:Q{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                sheet.Range($"B{dataRow}:Q{dataRow}").Style.Font.Bold = true;
                sheet.Range($"B{dataRow}:Q{dataRow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sheet.Range($"B{dataRow}:Q{dataRow}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                dataRow++;
            }
            //*************************************************************
            dataRow += 1;

            // Подпись для второй таблицы
            int signatureRow2 = dataRow;
            sheet.Cell(signatureRow2, 2).Value = "Начальник юридического отдела ОАО \"Гомельтранснефть Дружба\"                                                    Ю.А.Лащенко";
            sheet.Range(signatureRow2, 2, signatureRow2, 17).Merge();
            sheet.Range(signatureRow2, 2, signatureRow2, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            sheet.Row(signatureRow2).Height = 25;
            // Убираем границы для подписи
            sheet.Range(signatureRow2, 2, signatureRow2, 17).Style.Border.TopBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow2, 2, signatureRow2, 17).Style.Border.BottomBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow2, 2, signatureRow2, 17).Style.Border.LeftBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow2, 2, signatureRow2, 17).Style.Border.RightBorder = XLBorderStyleValues.None;
            //-------------------------------------------------------------------------------------------
            //------------------------Заполняем 3 Таблицу------------------------------------------------
            //********************************************************************************************
            dataRow = signatureRow2 + 2;

            // Создаем третью таблицу "Форма № 3 Иски"
            sheet.Range($"B{dataRow}:AB{dataRow}").Merge().Value = "Форма № 3 Иски (исковое, приказное производство, и/надписи), предъявленные к контрагентам";
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.FontSize = 14;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.Bold = true;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.WrapText = true;

            dataRow++;

            var startDate3 = new DateTime(reportDate.Year, 1, 1);
            var endDate3 = reportDate;

            sheet.Range($"B{dataRow}:AB{dataRow}").Merge().Value = $"в период с {startDate3:dd.MM.yyyy} по {endDate3:dd.MM.yyyy}";
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.FontSize = 14;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.Bold = true;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.WrapText = true;

            dataRow += 2; // Пропускаем 2 строки перед шапкой таблицы

            // Шапка таблицы
            sheet.Range($"B{dataRow}:B{dataRow + 2}").Merge().Value = "№";
            sheet.Range($"C{dataRow}:C{dataRow + 2}").Merge().Value = "Наименование должника";
            sheet.Range($"D{dataRow}:D{dataRow + 2}").Merge().Value = "Город (Страна)";
            sheet.Range($"E{dataRow}:E{dataRow + 2}").Merge().Value = "Дата предъявления иска";

            // Заявлены требования
            sheet.Range($"F{dataRow}:L{dataRow + 1}").Merge().Value = "Заявлены требования (в валюте требований)";
            sheet.Cell($"F{dataRow + 2}").Value = "Содержание требований";
            sheet.Cell($"G{dataRow + 2}").Value = "Сумма основного долга";
            sheet.Cell($"H{dataRow + 2}").Value = "Сумма неустойки";
            sheet.Cell($"I{dataRow + 2}").Value = "Сумма процентов";
            sheet.Cell($"J{dataRow + 2}").Value = "Валюта требований";
            sheet.Cell($"K{dataRow + 2}").Value = "Размер госпошлины";
            sheet.Cell($"L{dataRow + 2}").Value = "Валюта госпошлины";

            // Результат рассмотрения в суде первой апелляционной инстанции
            sheet.Range($"M{dataRow}:S{dataRow}").Merge().Value = "Результат рассмотрения заявленного иска в суде первой апелляционной инстанции";
            sheet.Range($"M{dataRow + 1}:S{dataRow + 1}").Merge().Value = "взыскано по вступившему в силу решению (определению) суда";
            sheet.Cell($"M{dataRow + 2}").Value = "Сумма основного долга";
            sheet.Cell($"N{dataRow + 2}").Value = "Сумма неустойки";
            sheet.Cell($"O{dataRow + 2}").Value = "Сумма процентов";
            sheet.Cell($"P{dataRow + 2}").Value = "Валюта требований";
            sheet.Cell($"Q{dataRow + 2}").Value = "Размер госпошлины";
            sheet.Cell($"R{dataRow + 2}").Value = "Валюта госпошлины";
            sheet.Cell($"S{dataRow + 2}").Value = "На решение инстанции подавалась апелляционная жалоба";

            // Результат рассмотрения в суде кассационной и надзорной инстанции
            sheet.Range($"T{dataRow}:Y{dataRow}").Merge().Value = "Результат рассмотрения иска в суде кассационной и надзорной инстанции";
            sheet.Range($"T{dataRow + 1}:Y{dataRow + 1}").Merge().Value = "Взыскано по итогу рассмотрения кассационной (надзорной) жалобы";
            sheet.Cell($"T{dataRow + 2}").Value = "Сумма основного долга";
            sheet.Cell($"U{dataRow + 2}").Value = "Сумма неустойки";
            sheet.Cell($"V{dataRow + 2}").Value = "Сумма процентов";
            sheet.Cell($"W{dataRow + 2}").Value = "Валюта требований";
            sheet.Cell($"X{dataRow + 2}").Value = "Размер госпошлины";
            sheet.Cell($"Y{dataRow + 2}").Value = "Валюта госпошлины";

            // Последние колонки
            sheet.Range($"Z{dataRow}:Z{dataRow + 2}").Merge().Value = "Дата вступления решения в законную силу";
            sheet.Range($"AA{dataRow}:AA{dataRow + 2}").Merge().Value = "Отозвано, оставлено без рассмотрения, возвращено без рассмотрения";
            sheet.Range($"AB{dataRow}:AB{dataRow + 2}").Merge().Value = "Предъявлено к исполнению";

            // Стилизация всей шапки
            var fullHeaderRange3 = sheet.Range($"B{dataRow}:AB{dataRow + 2}");
            fullHeaderRange3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            fullHeaderRange3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            fullHeaderRange3.Style.Alignment.WrapText = true;
            fullHeaderRange3.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            fullHeaderRange3.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Высота строк шапки
            sheet.Row(dataRow).Height = 25;
            sheet.Row(dataRow + 1).Height = 45;
            sheet.Row(dataRow + 2).Height = 45;

            dataRow += 3; // Переходим к данным таблицы

            // Проверяем есть ли данные для третьей таблицы (StatusId = 2, 10, 12)
            bool hasDataForThirdTable = listpretenseJoin.Any(pret =>
            {
                var lastTp1 = pret.TablePretenseList?
                    .Where(t => t.DateTabPret != null)
                    .OrderByDescending(t => t.DateTabPret)
                    .FirstOrDefault();

                return lastTp1 != null && pret.Inout == 1 && (lastTp1.StatusId == 2 || lastTp1.StatusId == 10 || lastTp1.StatusId == 12);
            });

            if (hasDataForThirdTable)
            {
                int counter3 = 1;
                var totalRequested3 = new Dictionary<string, decimal>();
                var totalSatisfied3 = new Dictionary<string, decimal>();
                var totalRequestedPoshlina3 = new Dictionary<string, decimal>();
                var totalSatisfiedPoshlina3 = new Dictionary<string, decimal>();

                // Заполняем данные для третьей таблицы (StatusId = 2, 10, 12)
                foreach (var pret in listpretenseJoin)
                {
                    var lastTp1 = pret.TablePretenseList?
                        .Where(t => t.DateTabPret != null)
                        .OrderByDescending(t => t.DateTabPret)
                        .FirstOrDefault();

                    if (lastTp1 != null && pret.Inout == 1 && (lastTp1.StatusId == 2 || lastTp1.StatusId == 10 || lastTp1.StatusId == 12))
                    {
                        sheet.Cell(dataRow, 2).Value = counter3;
                        sheet.Cell(dataRow, 3).Value = $"{pret.OrgName}";
                        sheet.Cell(dataRow, 4).Value = $"{pret.CityName} ({pret.CountryName})";
                        sheet.Cell(dataRow, 5).Value = pret.DatePret?.ToString("dd.MM.yyyy");
                        sheet.Cell(dataRow, 6).Value = pret.PredmetName;

                        //----Работаем с суммами группированными по валютам--------------------------------
                        if (pret.CurrencyGroups != null && pret.CurrencyGroups.Count > 0)
                        {
                            if (pret.CurrencyGroups.Count == 1)
                            {
                                var cg = pret.CurrencyGroups[0];

                                sheet.Cell(dataRow, 7).Value = cg.SummaDolg;
                                sheet.Cell(dataRow, 8).Value = cg.SummaPeny;
                                sheet.Cell(dataRow, 9).Value = cg.SummaProc;
                                sheet.Cell(dataRow, 10).Value = cg.CurrencyName;
                                // Госпошлина - используем SummaPoshlina
                                sheet.Cell(dataRow, 11).Value = cg.SummaPoshlina != 0 ? cg.SummaPoshlina : "-";
                                sheet.Cell(dataRow, 12).Value = cg.SummaPoshlina != 0 ? cg.CurrencyName : "-";

                                // Собираем итоги по заявленным требованиям
                                AddToTotals(totalRequested3, cg.CurrencyName, cg.SummaDolg);
                                AddToTotals(totalRequested3, cg.CurrencyName, cg.SummaPeny);
                                AddToTotals(totalRequested3, cg.CurrencyName, cg.SummaProc);
                                if (cg.SummaPoshlina != 0)
                                {
                                    AddToTotals(totalRequestedPoshlina3, cg.CurrencyName, cg.SummaPoshlina);
                                }
                            }
                            else
                            {
                                // Формируем текстовые представления
                                string dolgText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaDolg != 0)
                                        .Select(x => $"{x.SummaDolg} {x.CurrencyName}"));

                                string penyText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaPeny != 0)
                                        .Select(x => $"{x.SummaPeny} {x.CurrencyName}"));

                                string procText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaProc != 0)
                                        .Select(x => $"{x.SummaProc} {x.CurrencyName}"));

                                string poshlinaText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaPoshlina != 0)
                                        .Select(x => $"{x.SummaPoshlina} {x.CurrencyName}"));

                                sheet.Cell(dataRow, 7).Value = string.IsNullOrWhiteSpace(dolgText) ? "-" : dolgText;
                                sheet.Cell(dataRow, 8).Value = string.IsNullOrWhiteSpace(penyText) ? "-" : penyText;
                                sheet.Cell(dataRow, 9).Value = string.IsNullOrWhiteSpace(procText) ? "-" : procText;
                                sheet.Cell(dataRow, 10).Value = string.Join("; ", pret.CurrencyGroups.Select(x => x.CurrencyName).Distinct());
                                sheet.Cell(dataRow, 11).Value = string.IsNullOrWhiteSpace(poshlinaText) ? "-" : poshlinaText;
                                sheet.Cell(dataRow, 12).Value = string.IsNullOrWhiteSpace(poshlinaText) ? "-" : string.Join("; ", pret.CurrencyGroups.Where(x => x.SummaPoshlina != 0).Select(x => x.CurrencyName).Distinct());

                                // Собираем итоги по заявленным требованиям для всех валют
                                foreach (var cg in pret.CurrencyGroups)
                                {
                                    AddToTotals(totalRequested3, cg.CurrencyName, cg.SummaDolg);
                                    AddToTotals(totalRequested3, cg.CurrencyName, cg.SummaPeny);
                                    AddToTotals(totalRequested3, cg.CurrencyName, cg.SummaProc);
                                    if (cg.SummaPoshlina != 0)
                                    {
                                        AddToTotals(totalRequestedPoshlina3, cg.CurrencyName, cg.SummaPoshlina);
                                    }
                                }
                            }
                        }
                        else
                        {
                            sheet.Cell(dataRow, 7).Value = "-";
                            sheet.Cell(dataRow, 8).Value = "-";
                            sheet.Cell(dataRow, 9).Value = "-";
                            sheet.Cell(dataRow, 10).Value = "-";
                            sheet.Cell(dataRow, 11).Value = "-";
                            sheet.Cell(dataRow, 12).Value = "-";
                        }

                        //--------Результаты рассмотрения--------------------------------------------------------
                        if (pret.TablePretenseList != null && pret.TablePretenseList.Count > 0)
                        {
                            var allResults = pret.TablePretenseList
                                .SelectMany(tp => tp.ResultCurrencyGroups
                                    .Select(rcg => new
                                    {
                                        rcg.CurrencyName,
                                        rcg.SummaDolg,
                                        rcg.SummaPeny,
                                        rcg.SummaProc,
                                        rcg.SummaPoshlina
                                    }))
                                .GroupBy(x => x.CurrencyName)
                                .Select(g => new
                                {
                                    CurrencyName = g.Key,
                                    SummaDolg = g.Sum(x => x.SummaDolg),
                                    SummaPeny = g.Sum(x => x.SummaPeny),
                                    SummaProc = g.Sum(x => x.SummaProc),
                                    SummaPoshlina = g.Sum(x => x.SummaPoshlina),
                                })
                                .ToList();

                            if (allResults.Count == 1)
                            {
                                var r = allResults[0];
                                sheet.Cell(dataRow, 13).Value = r.SummaDolg;
                                sheet.Cell(dataRow, 14).Value = r.SummaPeny;
                                sheet.Cell(dataRow, 15).Value = r.SummaProc;
                                sheet.Cell(dataRow, 16).Value = r.CurrencyName;
                                sheet.Cell(dataRow, 17).Value = r.SummaPoshlina != 0 ? r.SummaPoshlina : "-"; // Госпошлина
                                sheet.Cell(dataRow, 18).Value = r.SummaPoshlina != 0 ? r.CurrencyName : "-"; // Валюта госпошлины
                                sheet.Cell(dataRow, 19).Value = ""; // Апелляционная жалоба

                                // Собираем итоги по удовлетворенным требованиям
                                AddToTotals(totalSatisfied3, r.CurrencyName, r.SummaDolg);
                                AddToTotals(totalSatisfied3, r.CurrencyName, r.SummaPeny);
                                AddToTotals(totalSatisfied3, r.CurrencyName, r.SummaProc);
                                if (r.SummaPoshlina != 0)
                                {
                                    AddToTotals(totalSatisfiedPoshlina3, r.CurrencyName, r.SummaPoshlina);
                                }
                            }
                            else
                            {
                                string dolgText = string.Join("; ",
                                    allResults.Where(x => x.SummaDolg != 0)
                                              .Select(x => $"{x.SummaDolg} {x.CurrencyName}"));

                                string penyText = string.Join("; ",
                                    allResults.Where(x => x.SummaPeny != 0)
                                              .Select(x => $"{x.SummaPeny} {x.CurrencyName}"));

                                string procText = string.Join("; ",
                                    allResults.Where(x => x.SummaProc != 0)
                                              .Select(x => $"{x.SummaProc} {x.CurrencyName}"));

                                string poshlinaText = string.Join("; ",
                                    allResults.Where(x => x.SummaPoshlina != 0)
                                              .Select(x => $"{x.SummaPoshlina} {x.CurrencyName}"));

                                sheet.Cell(dataRow, 13).Value = string.IsNullOrWhiteSpace(dolgText) ? "-" : dolgText;
                                sheet.Cell(dataRow, 14).Value = string.IsNullOrWhiteSpace(penyText) ? "-" : penyText;
                                sheet.Cell(dataRow, 15).Value = string.IsNullOrWhiteSpace(procText) ? "-" : procText;
                                sheet.Cell(dataRow, 16).Value = string.Join("; ", allResults.Select(x => x.CurrencyName).Distinct());
                                sheet.Cell(dataRow, 17).Value = string.IsNullOrWhiteSpace(poshlinaText) ? "-" : poshlinaText; // Госпошлина
                                sheet.Cell(dataRow, 18).Value = string.IsNullOrWhiteSpace(poshlinaText) ? "-" : string.Join("; ", allResults.Where(x => x.SummaPoshlina != 0).Select(x => x.CurrencyName).Distinct()); // Валюта госпошлины
                                sheet.Cell(dataRow, 19).Value = ""; // Апелляционная жалоба

                                // Собираем итоги по удовлетворенным требованиям для всех валют
                                foreach (var r in allResults)
                                {
                                    AddToTotals(totalSatisfied3, r.CurrencyName, r.SummaDolg);
                                    AddToTotals(totalSatisfied3, r.CurrencyName, r.SummaPeny);
                                    AddToTotals(totalSatisfied3, r.CurrencyName, r.SummaProc);
                                    if (r.SummaPoshlina != 0)
                                    {
                                        AddToTotals(totalSatisfiedPoshlina3, r.CurrencyName, r.SummaPoshlina);
                                    }
                                }
                            }

                            // Кассационная инстанция - аналогично (пока оставляем прочерки)
                            sheet.Cell(dataRow, 20).Value = "-";
                            sheet.Cell(dataRow, 21).Value = "-";
                            sheet.Cell(dataRow, 22).Value = "-";
                            sheet.Cell(dataRow, 23).Value = "-";
                            sheet.Cell(dataRow, 24).Value = "-";
                            sheet.Cell(dataRow, 25).Value = "-";

                            // Последние колонки
                            sheet.Cell(dataRow, 26).Value = ""; // Дата вступления решения
                            sheet.Cell(dataRow, 27).Value = ""; // Отозвано и т.д.
                            sheet.Cell(dataRow, 28).Value = ""; // Предъявлено к исполнению
                        }

                        var dataRange3 = sheet.Range($"B{dataRow}:AB{dataRow}");
                        dataRange3.Style.Alignment.WrapText = true;
                        dataRange3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        dataRange3.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        dataRange3.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        OptimizeRowHeight(sheet, dataRow);

                        dataRow++;
                        counter3++;
                    }
                }
                // Итоги для третьей таблицы
                var allCurrencies3 = totalRequested3.Keys.Union(totalSatisfied3.Keys).Union(totalRequestedPoshlina3.Keys).Union(totalSatisfiedPoshlina3.Keys).Distinct().ToList();
                int totalRowsCount3 = allCurrencies3.Count;

                if (totalRowsCount3 > 0)
                {
                    int totalStartRow3 = dataRow;

                    // Заголовок "ИТОГО" объединяем на все строки валют
                    sheet.Range(totalStartRow3, 2, totalStartRow3 + totalRowsCount3 - 1, 5).Merge();
                    sheet.Cell(totalStartRow3, 2).Value = "ИТОГО";
                    sheet.Cell(totalStartRow3, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Cell(totalStartRow3, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    sheet.Cell(totalStartRow3, 2).Style.Font.Bold = true;

                    // Заполняем данные по каждой валюте
                    for (int i = 0; i < totalRowsCount3; i++)
                    {
                        int currentRow = totalStartRow3 + i;
                        string currency = allCurrencies3[i];

                        // Название валюты в колонке J
                        sheet.Cell(currentRow, 10).Value = currency;
                        sheet.Cell(currentRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        sheet.Cell(currentRow, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Суммы по заявленным требованиям
                        decimal requestedDolg = 0;
                        decimal requestedPeny = 0;
                        decimal requestedProc = 0;
                        decimal requestedPoshlina = 0;

                        // Пересчитываем суммы для этой валюты из исходных данных (StatusId = 2, 10, 12)
                        foreach (var pret in listpretenseJoin)
                        {
                            var lastTp1 = pret.TablePretenseList?
                                .Where(t => t.DateTabPret != null)
                                .OrderByDescending(t => t.DateTabPret)
                                .FirstOrDefault();

                            if (lastTp1 != null && pret.Inout == 1 && (lastTp1.StatusId == 2 || lastTp1.StatusId == 10 || lastTp1.StatusId == 12))
                            {
                                if (pret.CurrencyGroups != null)
                                {
                                    foreach (var cg in pret.CurrencyGroups)
                                    {
                                        if (cg.CurrencyName == currency)
                                        {
                                            requestedDolg += cg.SummaDolg;
                                            requestedPeny += cg.SummaPeny;
                                            requestedProc += cg.SummaProc;
                                            requestedPoshlina += cg.SummaPoshlina;
                                        }
                                    }
                                }
                            }
                        }

                        // Записываем суммы заявленных требований
                        sheet.Cell(currentRow, 7).Value = requestedDolg != 0 ? requestedDolg : "-";
                        sheet.Cell(currentRow, 8).Value = requestedPeny != 0 ? requestedPeny : "-";
                        sheet.Cell(currentRow, 9).Value = requestedProc != 0 ? requestedProc : "-";
                        sheet.Cell(currentRow, 11).Value = requestedPoshlina != 0 ? requestedPoshlina : "-"; // Госпошлина заявленная
                        sheet.Cell(currentRow, 12).Value = requestedPoshlina != 0 ? currency : "-"; // Валюта госпошлины заявленная

                        // Суммы по удовлетворенным требованиям
                        decimal satisfiedDolg = 0;
                        decimal satisfiedPeny = 0;
                        decimal satisfiedProc = 0;
                        decimal satisfiedPoshlina = 0;

                        // Пересчитываем суммы для этой валюты из результатов (StatusId = 2, 10, 12)
                        foreach (var pret in listpretenseJoin)
                        {
                            var lastTp1 = pret.TablePretenseList?
                                .Where(t => t.DateTabPret != null)
                                .OrderByDescending(t => t.DateTabPret)
                                .FirstOrDefault();

                            if (lastTp1 != null && (lastTp1.StatusId == 2 || lastTp1.StatusId == 10 || lastTp1.StatusId == 12))
                            {
                                if (pret.TablePretenseList != null && pret.TablePretenseList.Count > 0)
                                {
                                    var allResults = pret.TablePretenseList
                                        .SelectMany(tp => tp.ResultCurrencyGroups
                                            .Select(rcg => new
                                            {
                                                rcg.CurrencyName,
                                                rcg.SummaDolg,
                                                rcg.SummaPeny,
                                                rcg.SummaProc,
                                                rcg.SummaPoshlina
                                            }))
                                        .GroupBy(x => x.CurrencyName)
                                        .Select(g => new
                                        {
                                            CurrencyName = g.Key,
                                            SummaDolg = g.Sum(x => x.SummaDolg),
                                            SummaPeny = g.Sum(x => x.SummaPeny),
                                            SummaProc = g.Sum(x => x.SummaProc),
                                            SummaPoshlina = g.Sum(x => x.SummaPoshlina)
                                        })
                                        .ToList();

                                    foreach (var result in allResults)
                                    {
                                        if (result.CurrencyName == currency)
                                        {
                                            satisfiedDolg += result.SummaDolg;
                                            satisfiedPeny += result.SummaPeny;
                                            satisfiedProc += result.SummaProc;
                                            satisfiedPoshlina += result.SummaPoshlina;
                                        }
                                    }
                                }
                            }
                        }
                        // Записываем суммы удовлетворенных требований (первая инстанция)
                        sheet.Cell(currentRow, 13).Value = satisfiedDolg != 0 ? satisfiedDolg : "-";
                        sheet.Cell(currentRow, 14).Value = satisfiedPeny != 0 ? satisfiedPeny : "-";
                        sheet.Cell(currentRow, 15).Value = satisfiedProc != 0 ? satisfiedProc : "-";
                        sheet.Cell(currentRow, 17).Value = satisfiedPoshlina != 0 ? satisfiedPoshlina : "-"; // Госпошлина удовлетворенная
                        sheet.Cell(currentRow, 18).Value = satisfiedPoshlina != 0 ? currency : "-"; // Валюта госпошлины удовлетворенная

                        // Валюта требований для удовлетворенных
                        sheet.Cell(currentRow, 16).Value = currency;
                        sheet.Cell(currentRow, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        sheet.Cell(currentRow, 16).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Пустые значения для остальных колонок
                        sheet.Cell(currentRow, 19).Value = ""; // Апелляционная жалоба
                        sheet.Cell(currentRow, 20).Value = "-"; // Кассация - долг
                        sheet.Cell(currentRow, 21).Value = "-"; // Кассация - неустойка
                        sheet.Cell(currentRow, 22).Value = "-"; // Кассация - проценты
                        sheet.Cell(currentRow, 23).Value = "-"; // Кассация - валюта
                        sheet.Cell(currentRow, 24).Value = "-"; // Кассация - госпошлина
                        sheet.Cell(currentRow, 25).Value = "-"; // Кассация - валюта госпошлины
                        sheet.Cell(currentRow, 26).Value = ""; // Дата вступления решения
                        sheet.Cell(currentRow, 27).Value = ""; // Отозвано и т.д.
                        sheet.Cell(currentRow, 28).Value = ""; // Предъявлено к исполнению

                        // Стили для строки итогов
                        var rowRange = sheet.Range(currentRow, 2, currentRow, 28);
                        rowRange.Style.Alignment.WrapText = true;
                        rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        rowRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                        OptimizeRowHeight(sheet, currentRow);
                    }

                    // Общие стили для блока итогов
                    var totalRange3 = sheet.Range(totalStartRow3, 2, totalStartRow3 + totalRowsCount3 - 1, 28);
                    totalRange3.Style.Font.Bold = true;

                    // Толстая граница сверху всего блока итогов
                    sheet.Range(totalStartRow3, 2, totalStartRow3, 28).Style.Border.TopBorder = XLBorderStyleValues.Medium;

                    dataRow = totalStartRow3 + totalRowsCount3;
                }
            }
            else
            {
                // Если данных нет - выводим сообщение "Данные отсутствуют" под шапкой
                sheet.Range($"B{dataRow}:AB{dataRow}").Merge().Value = "Данные отсутствуют";
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.Bold = true;
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                dataRow++;
            }

            // Пропускаем строку после третьей таблицы
            dataRow += 1;

            // Подпись для третьей таблицы
            int signatureRow3 = dataRow;
            sheet.Cell(signatureRow3, 2).Value = "Начальник юридического отдела ОАО \"Гомельтранснефть Дружба\"                                                    Ю.А.Лащенко";
            sheet.Range(signatureRow3, 2, signatureRow3, 17).Merge();
            sheet.Range(signatureRow3, 2, signatureRow3, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            sheet.Row(signatureRow3).Height = 25;
            // Убираем границы для подписи
            sheet.Range(signatureRow3, 2, signatureRow3, 17).Style.Border.TopBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow3, 2, signatureRow3, 17).Style.Border.BottomBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow3, 2, signatureRow3, 17).Style.Border.LeftBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow3, 2, signatureRow3, 17).Style.Border.RightBorder = XLBorderStyleValues.None;
            //********************************************************************************************
            //--------Создаем и заполняем 4 таблицу Форму №4----------------------------------------------

            dataRow = signatureRow3 + 2;

            // Создаем четвертую таблицу "Форма №4 Иски, предъявленные к организации"
            sheet.Range($"B{dataRow}:AB{dataRow}").Merge().Value = "Форма № 4 Иски (исковое, приказное производство, и/или надписи), предъявленные к организации";
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.FontSize = 14;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.Bold = true;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.WrapText = true;

            dataRow++;

            var startDate4 = new DateTime(reportDate.Year, 1, 1);
            var endDate4 = reportDate;

            sheet.Range($"B{dataRow}:AB{dataRow}").Merge().Value = $"в период с {startDate4:dd.MM.yyyy} по {endDate4:dd.MM.yyyy}";
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.FontSize = 14;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.Bold = true;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.WrapText = true;

            dataRow += 2; // Пропускаем 2 строки перед шапкой таблицы

            // Шапка таблицы (такая же как у третьей таблицы)
            sheet.Range($"B{dataRow}:B{dataRow + 2}").Merge().Value = "№";
            sheet.Range($"C{dataRow}:C{dataRow + 2}").Merge().Value = "Наименование должника";
            sheet.Range($"D{dataRow}:D{dataRow + 2}").Merge().Value = "Город (Страна)";
            sheet.Range($"E{dataRow}:E{dataRow + 2}").Merge().Value = "Дата предъявления иска";

            // Заявлены требования
            sheet.Range($"F{dataRow}:L{dataRow + 1}").Merge().Value = "Заявлены требования (в валюте требований)";
            sheet.Cell($"F{dataRow + 2}").Value = "Содержание требований";
            sheet.Cell($"G{dataRow + 2}").Value = "Сумма основного долга";
            sheet.Cell($"H{dataRow + 2}").Value = "Сумма неустойки";
            sheet.Cell($"I{dataRow + 2}").Value = "Сумма процентов";
            sheet.Cell($"J{dataRow + 2}").Value = "Валюта требований";
            sheet.Cell($"K{dataRow + 2}").Value = "Размер госпошлины";
            sheet.Cell($"L{dataRow + 2}").Value = "Валюта госпошлины";

            // Результат рассмотрения в суде первой апелляционной инстанции
            sheet.Range($"M{dataRow}:S{dataRow}").Merge().Value = "Результат рассмотрения заявленного иска в суде первой апелляционной инстанции";
            sheet.Range($"M{dataRow + 1}:S{dataRow + 1}").Merge().Value = "взыскано по вступившему в силу решению (определению) суда";
            sheet.Cell($"M{dataRow + 2}").Value = "Сумма основного долга";
            sheet.Cell($"N{dataRow + 2}").Value = "Сумма неустойки";
            sheet.Cell($"O{dataRow + 2}").Value = "Сумма процентов";
            sheet.Cell($"P{dataRow + 2}").Value = "Валюта требований";
            sheet.Cell($"Q{dataRow + 2}").Value = "Размер госпошлины";
            sheet.Cell($"R{dataRow + 2}").Value = "Валюта госпошлины";
            sheet.Cell($"S{dataRow + 2}").Value = "На решение инстанции подавалась апелляционная жалоба";

            // Результат рассмотрения в суде кассационной и надзорной инстанции
            sheet.Range($"T{dataRow}:Y{dataRow}").Merge().Value = "Результат рассмотрения иска в суде кассационной и надзорной инстанции";
            sheet.Range($"T{dataRow + 1}:Y{dataRow + 1}").Merge().Value = "Взыскано по итогу рассмотрения кассационной (надзорной) жалобы";
            sheet.Cell($"T{dataRow + 2}").Value = "Сумма основного долга";
            sheet.Cell($"U{dataRow + 2}").Value = "Сумма неустойки";
            sheet.Cell($"V{dataRow + 2}").Value = "Сумма процентов";
            sheet.Cell($"W{dataRow + 2}").Value = "Валюта требований";
            sheet.Cell($"X{dataRow + 2}").Value = "Размер госпошлины";
            sheet.Cell($"Y{dataRow + 2}").Value = "Валюта госпошлины";

            // Последние колонки
            sheet.Range($"Z{dataRow}:Z{dataRow + 2}").Merge().Value = "Дата вступления решения в законную силу";
            sheet.Range($"AA{dataRow}:AA{dataRow + 2}").Merge().Value = "Отозвано, оставлено без рассмотрения, возвращено без рассмотрения";
            sheet.Range($"AB{dataRow}:AB{dataRow + 2}").Merge().Value = "Предъявлено к исполнению";

            // Стилизация всей шапки
            var fullHeaderRange4 = sheet.Range($"B{dataRow}:AB{dataRow + 2}");
            fullHeaderRange4.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            fullHeaderRange4.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            fullHeaderRange4.Style.Alignment.WrapText = true;
            fullHeaderRange4.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            fullHeaderRange4.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Высота строк шапки
            sheet.Row(dataRow).Height = 25;
            sheet.Row(dataRow + 1).Height = 45;
            sheet.Row(dataRow + 2).Height = 45;

            dataRow += 3; // Переходим к данным таблицы

            // Проверяем есть ли данные для четвертой таблицы (StatusId = 2, 10, 12 и pret.Inout == 0)
            bool hasDataForFourthTable = listpretenseJoin.Any(pret =>
            {
                var lastTp1 = pret.TablePretenseList?
                    .Where(t => t.DateTabPret != null)
                    .OrderByDescending(t => t.DateTabPret)
                    .FirstOrDefault();

                return lastTp1 != null && pret.Inout == 0 && (lastTp1.StatusId == 2 || lastTp1.StatusId == 10 || lastTp1.StatusId == 12);
            });

            if (hasDataForFourthTable)
            {
                int counter4 = 1;
                var totalRequested4 = new Dictionary<string, decimal>();
                var totalSatisfied4 = new Dictionary<string, decimal>();
                var totalRequestedPoshlina4 = new Dictionary<string, decimal>();
                var totalSatisfiedPoshlina4 = new Dictionary<string, decimal>();

                // Заполняем данные для четвертой таблицы (StatusId = 2, 10, 12 и pret.Inout == 0)
                foreach (var pret in listpretenseJoin)
                {
                    var lastTp1 = pret.TablePretenseList?
                        .Where(t => t.DateTabPret != null)
                        .OrderByDescending(t => t.DateTabPret)
                        .FirstOrDefault();

                    // ОТЛИЧИЕ: pret.Inout == 0
                    if (lastTp1 != null && pret.Inout == 0 && (lastTp1.StatusId == 2 || lastTp1.StatusId == 10 || lastTp1.StatusId == 12))
                    {
                        sheet.Cell(dataRow, 2).Value = counter4;
                        sheet.Cell(dataRow, 3).Value = $"{pret.OrgName}";
                        sheet.Cell(dataRow, 4).Value = $"{pret.CityName} ({pret.CountryName})";
                        sheet.Cell(dataRow, 5).Value = pret.DatePret?.ToString("dd.MM.yyyy");
                        sheet.Cell(dataRow, 6).Value = pret.PredmetName;

                        //----Работаем с суммами группированными по валютам--------------------------------
                        if (pret.CurrencyGroups != null && pret.CurrencyGroups.Count > 0)
                        {
                            if (pret.CurrencyGroups.Count == 1)
                            {
                                var cg = pret.CurrencyGroups[0];

                                sheet.Cell(dataRow, 7).Value = cg.SummaDolg;
                                sheet.Cell(dataRow, 8).Value = cg.SummaPeny;
                                sheet.Cell(dataRow, 9).Value = cg.SummaProc;
                                sheet.Cell(dataRow, 10).Value = cg.CurrencyName;
                                // Госпошлина - используем SummaPoshlina
                                sheet.Cell(dataRow, 11).Value = cg.SummaPoshlina != 0 ? cg.SummaPoshlina : "-";
                                sheet.Cell(dataRow, 12).Value = cg.SummaPoshlina != 0 ? cg.CurrencyName : "-";

                                // Собираем итоги по заявленным требованиям
                                AddToTotals(totalRequested4, cg.CurrencyName, cg.SummaDolg);
                                AddToTotals(totalRequested4, cg.CurrencyName, cg.SummaPeny);
                                AddToTotals(totalRequested4, cg.CurrencyName, cg.SummaProc);
                                if (cg.SummaPoshlina != 0)
                                {
                                    AddToTotals(totalRequestedPoshlina4, cg.CurrencyName, cg.SummaPoshlina);
                                }
                            }
                            else
                            {
                                // Формируем текстовые представления
                                string dolgText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaDolg != 0)
                                        .Select(x => $"{x.SummaDolg} {x.CurrencyName}"));

                                string penyText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaPeny != 0)
                                        .Select(x => $"{x.SummaPeny} {x.CurrencyName}"));

                                string procText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaProc != 0)
                                        .Select(x => $"{x.SummaProc} {x.CurrencyName}"));

                                string poshlinaText = string.Join("; ",
                                    pret.CurrencyGroups
                                        .Where(x => x.SummaPoshlina != 0)
                                        .Select(x => $"{x.SummaPoshlina} {x.CurrencyName}"));

                                sheet.Cell(dataRow, 7).Value = string.IsNullOrWhiteSpace(dolgText) ? "-" : dolgText;
                                sheet.Cell(dataRow, 8).Value = string.IsNullOrWhiteSpace(penyText) ? "-" : penyText;
                                sheet.Cell(dataRow, 9).Value = string.IsNullOrWhiteSpace(procText) ? "-" : procText;
                                sheet.Cell(dataRow, 10).Value = string.Join("; ", pret.CurrencyGroups.Select(x => x.CurrencyName).Distinct());
                                sheet.Cell(dataRow, 11).Value = string.IsNullOrWhiteSpace(poshlinaText) ? "-" : poshlinaText;
                                sheet.Cell(dataRow, 12).Value = string.IsNullOrWhiteSpace(poshlinaText) ? "-" : string.Join("; ", pret.CurrencyGroups.Where(x => x.SummaPoshlina != 0).Select(x => x.CurrencyName).Distinct());

                                // Собираем итоги по заявленным требованиям для всех валют
                                foreach (var cg in pret.CurrencyGroups)
                                {
                                    AddToTotals(totalRequested4, cg.CurrencyName, cg.SummaDolg);
                                    AddToTotals(totalRequested4, cg.CurrencyName, cg.SummaPeny);
                                    AddToTotals(totalRequested4, cg.CurrencyName, cg.SummaProc);
                                    if (cg.SummaPoshlina != 0)
                                    {
                                        AddToTotals(totalRequestedPoshlina4, cg.CurrencyName, cg.SummaPoshlina);
                                    }
                                }
                            }
                        }
                        else
                        {
                            sheet.Cell(dataRow, 7).Value = "-";
                            sheet.Cell(dataRow, 8).Value = "-";
                            sheet.Cell(dataRow, 9).Value = "-";
                            sheet.Cell(dataRow, 10).Value = "-";
                            sheet.Cell(dataRow, 11).Value = "-";
                            sheet.Cell(dataRow, 12).Value = "-";
                        }

                        //--------Результаты рассмотрения--------------------------------------------------------
                        if (pret.TablePretenseList != null && pret.TablePretenseList.Count > 0)
                        {
                            var allResults = pret.TablePretenseList
                                .SelectMany(tp => tp.ResultCurrencyGroups
                                    .Select(rcg => new
                                    {
                                        rcg.CurrencyName,
                                        rcg.SummaDolg,
                                        rcg.SummaPeny,
                                        rcg.SummaProc,
                                        rcg.SummaPoshlina
                                    }))
                                .GroupBy(x => x.CurrencyName)
                                .Select(g => new
                                {
                                    CurrencyName = g.Key,
                                    SummaDolg = g.Sum(x => x.SummaDolg),
                                    SummaPeny = g.Sum(x => x.SummaPeny),
                                    SummaProc = g.Sum(x => x.SummaProc),
                                    SummaPoshlina = g.Sum(x => x.SummaPoshlina),
                                })
                                .ToList();

                            if (allResults.Count == 1)
                            {
                                var r = allResults[0];
                                sheet.Cell(dataRow, 13).Value = r.SummaDolg;
                                sheet.Cell(dataRow, 14).Value = r.SummaPeny;
                                sheet.Cell(dataRow, 15).Value = r.SummaProc;
                                sheet.Cell(dataRow, 16).Value = r.CurrencyName;
                                sheet.Cell(dataRow, 17).Value = r.SummaPoshlina != 0 ? r.SummaPoshlina : "-"; // Госпошлина
                                sheet.Cell(dataRow, 18).Value = r.SummaPoshlina != 0 ? r.CurrencyName : "-"; // Валюта госпошлины
                                sheet.Cell(dataRow, 19).Value = ""; // Апелляционная жалоба

                                // Собираем итоги по удовлетворенным требованиям
                                AddToTotals(totalSatisfied4, r.CurrencyName, r.SummaDolg);
                                AddToTotals(totalSatisfied4, r.CurrencyName, r.SummaPeny);
                                AddToTotals(totalSatisfied4, r.CurrencyName, r.SummaProc);
                                if (r.SummaPoshlina != 0)
                                {
                                    AddToTotals(totalSatisfiedPoshlina4, r.CurrencyName, r.SummaPoshlina);
                                }
                            }
                            else
                            {
                                string dolgText = string.Join("; ",
                                    allResults.Where(x => x.SummaDolg != 0)
                                              .Select(x => $"{x.SummaDolg} {x.CurrencyName}"));

                                string penyText = string.Join("; ",
                                    allResults.Where(x => x.SummaPeny != 0)
                                              .Select(x => $"{x.SummaPeny} {x.CurrencyName}"));

                                string procText = string.Join("; ",
                                    allResults.Where(x => x.SummaProc != 0)
                                              .Select(x => $"{x.SummaProc} {x.CurrencyName}"));

                                string poshlinaText = string.Join("; ",
                                    allResults.Where(x => x.SummaPoshlina != 0)
                                              .Select(x => $"{x.SummaPoshlina} {x.CurrencyName}"));

                                sheet.Cell(dataRow, 13).Value = string.IsNullOrWhiteSpace(dolgText) ? "-" : dolgText;
                                sheet.Cell(dataRow, 14).Value = string.IsNullOrWhiteSpace(penyText) ? "-" : penyText;
                                sheet.Cell(dataRow, 15).Value = string.IsNullOrWhiteSpace(procText) ? "-" : procText;
                                sheet.Cell(dataRow, 16).Value = string.Join("; ", allResults.Select(x => x.CurrencyName).Distinct());
                                sheet.Cell(dataRow, 17).Value = string.IsNullOrWhiteSpace(poshlinaText) ? "-" : poshlinaText; // Госпошлина
                                sheet.Cell(dataRow, 18).Value = string.IsNullOrWhiteSpace(poshlinaText) ? "-" : string.Join("; ", allResults.Where(x => x.SummaPoshlina != 0).Select(x => x.CurrencyName).Distinct()); // Валюта госпошлины
                                sheet.Cell(dataRow, 19).Value = ""; // Апелляционная жалоба

                                // Собираем итоги по удовлетворенным требованиям для всех валют
                                foreach (var r in allResults)
                                {
                                    AddToTotals(totalSatisfied4, r.CurrencyName, r.SummaDolg);
                                    AddToTotals(totalSatisfied4, r.CurrencyName, r.SummaPeny);
                                    AddToTotals(totalSatisfied4, r.CurrencyName, r.SummaProc);
                                    if (r.SummaPoshlina != 0)
                                    {
                                        AddToTotals(totalSatisfiedPoshlina4, r.CurrencyName, r.SummaPoshlina);
                                    }
                                }
                            }

                            // Кассационная инстанция - аналогично (пока оставляем прочерки)
                            sheet.Cell(dataRow, 20).Value = "-";
                            sheet.Cell(dataRow, 21).Value = "-";
                            sheet.Cell(dataRow, 22).Value = "-";
                            sheet.Cell(dataRow, 23).Value = "-";
                            sheet.Cell(dataRow, 24).Value = "-";
                            sheet.Cell(dataRow, 25).Value = "-";

                            // Последние колонки
                            sheet.Cell(dataRow, 26).Value = ""; // Дата вступления решения
                            sheet.Cell(dataRow, 27).Value = ""; // Отозвано и т.д.
                            sheet.Cell(dataRow, 28).Value = ""; // Предъявлено к исполнению
                        }

                        var dataRange4 = sheet.Range($"B{dataRow}:AB{dataRow}");
                        dataRange4.Style.Alignment.WrapText = true;
                        dataRange4.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        dataRange4.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        dataRange4.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        OptimizeRowHeight(sheet, dataRow);

                        dataRow++;
                        counter4++;
                    }
                }

                // Итоги для четвертой таблицы
                var allCurrencies4 = totalRequested4.Keys.Union(totalSatisfied4.Keys).Union(totalRequestedPoshlina4.Keys).Union(totalSatisfiedPoshlina4.Keys).Distinct().ToList();
                int totalRowsCount4 = allCurrencies4.Count;

                if (totalRowsCount4 > 0)
                {
                    int totalStartRow4 = dataRow;

                    // Заголовок "ИТОГО" объединяем на все строки валют
                    sheet.Range(totalStartRow4, 2, totalStartRow4 + totalRowsCount4 - 1, 5).Merge();
                    sheet.Cell(totalStartRow4, 2).Value = "ИТОГО";
                    sheet.Cell(totalStartRow4, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    sheet.Cell(totalStartRow4, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    sheet.Cell(totalStartRow4, 2).Style.Font.Bold = true;

                    // Заполняем данные по каждой валюте
                    for (int i = 0; i < totalRowsCount4; i++)
                    {
                        int currentRow = totalStartRow4 + i;
                        string currency = allCurrencies4[i];

                        // Название валюты в колонке J
                        sheet.Cell(currentRow, 10).Value = currency;
                        sheet.Cell(currentRow, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        sheet.Cell(currentRow, 10).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Суммы по заявленным требованиям
                        decimal requestedDolg = 0;
                        decimal requestedPeny = 0;
                        decimal requestedProc = 0;
                        decimal requestedPoshlina = 0;

                        // Пересчитываем суммы для этой валюты из исходных данных (StatusId = 2, 10, 12 и pret.Inout == 0)
                        foreach (var pret in listpretenseJoin)
                        {
                            var lastTp1 = pret.TablePretenseList?
                                .Where(t => t.DateTabPret != null)
                                .OrderByDescending(t => t.DateTabPret)
                                .FirstOrDefault();

                            if (lastTp1 != null && pret.Inout == 0 && (lastTp1.StatusId == 2 || lastTp1.StatusId == 10 || lastTp1.StatusId == 12))
                            {
                                if (pret.CurrencyGroups != null)
                                {
                                    foreach (var cg in pret.CurrencyGroups)
                                    {
                                        if (cg.CurrencyName == currency)
                                        {
                                            requestedDolg += cg.SummaDolg;
                                            requestedPeny += cg.SummaPeny;
                                            requestedProc += cg.SummaProc;
                                            requestedPoshlina += cg.SummaPoshlina;
                                        }
                                    }
                                }
                            }
                        }

                        // Записываем суммы заявленных требований
                        sheet.Cell(currentRow, 7).Value = requestedDolg != 0 ? requestedDolg : "-";
                        sheet.Cell(currentRow, 8).Value = requestedPeny != 0 ? requestedPeny : "-";
                        sheet.Cell(currentRow, 9).Value = requestedProc != 0 ? requestedProc : "-";
                        sheet.Cell(currentRow, 11).Value = requestedPoshlina != 0 ? requestedPoshlina : "-"; // Госпошлина заявленная
                        sheet.Cell(currentRow, 12).Value = requestedPoshlina != 0 ? currency : "-"; // Валюта госпошлины заявленная

                        // Суммы по удовлетворенным требованиям
                        decimal satisfiedDolg = 0;
                        decimal satisfiedPeny = 0;
                        decimal satisfiedProc = 0;
                        decimal satisfiedPoshlina = 0;

                        // Пересчитываем суммы для этой валюты из результатов (StatusId = 2, 10, 12 и pret.Inout == 0)
                        foreach (var pret in listpretenseJoin)
                        {
                            var lastTp1 = pret.TablePretenseList?
                                .Where(t => t.DateTabPret != null)
                                .OrderByDescending(t => t.DateTabPret)
                                .FirstOrDefault();

                            if (lastTp1 != null && pret.Inout == 0 && (lastTp1.StatusId == 2 || lastTp1.StatusId == 10 || lastTp1.StatusId == 12))
                            {
                                if (pret.TablePretenseList != null && pret.TablePretenseList.Count > 0)
                                {
                                    var allResults = pret.TablePretenseList
                                        .SelectMany(tp => tp.ResultCurrencyGroups
                                            .Select(rcg => new
                                            {
                                                rcg.CurrencyName,
                                                rcg.SummaDolg,
                                                rcg.SummaPeny,
                                                rcg.SummaProc,
                                                rcg.SummaPoshlina
                                            }))
                                        .GroupBy(x => x.CurrencyName)
                                        .Select(g => new
                                        {
                                            CurrencyName = g.Key,
                                            SummaDolg = g.Sum(x => x.SummaDolg),
                                            SummaPeny = g.Sum(x => x.SummaPeny),
                                            SummaProc = g.Sum(x => x.SummaProc),
                                            SummaPoshlina = g.Sum(x => x.SummaPoshlina)
                                        })
                                        .ToList();

                                    foreach (var result in allResults)
                                    {
                                        if (result.CurrencyName == currency)
                                        {
                                            satisfiedDolg += result.SummaDolg;
                                            satisfiedPeny += result.SummaPeny;
                                            satisfiedProc += result.SummaProc;
                                            satisfiedPoshlina += result.SummaPoshlina;
                                        }
                                    }
                                }
                            }
                        }

                        // Записываем суммы удовлетворенных требований (первая инстанция)
                        sheet.Cell(currentRow, 13).Value = satisfiedDolg != 0 ? satisfiedDolg : "-";
                        sheet.Cell(currentRow, 14).Value = satisfiedPeny != 0 ? satisfiedPeny : "-";
                        sheet.Cell(currentRow, 15).Value = satisfiedProc != 0 ? satisfiedProc : "-";
                        sheet.Cell(currentRow, 17).Value = satisfiedPoshlina != 0 ? satisfiedPoshlina : "-"; // Госпошлина удовлетворенная
                        sheet.Cell(currentRow, 18).Value = satisfiedPoshlina != 0 ? currency : "-"; // Валюта госпошлины удовлетворенная

                        // Валюта требований для удовлетворенных
                        sheet.Cell(currentRow, 16).Value = currency;
                        sheet.Cell(currentRow, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        sheet.Cell(currentRow, 16).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Пустые значения для остальных колонок
                        sheet.Cell(currentRow, 19).Value = ""; // Апелляционная жалоба
                        sheet.Cell(currentRow, 20).Value = "-"; // Кассация - долг
                        sheet.Cell(currentRow, 21).Value = "-"; // Кассация - неустойка
                        sheet.Cell(currentRow, 22).Value = "-"; // Кассация - проценты
                        sheet.Cell(currentRow, 23).Value = "-"; // Кассация - валюта
                        sheet.Cell(currentRow, 24).Value = "-"; // Кассация - госпошлина
                        sheet.Cell(currentRow, 25).Value = "-"; // Кассация - валюта госпошлины
                        sheet.Cell(currentRow, 26).Value = ""; // Дата вступления решения
                        sheet.Cell(currentRow, 27).Value = ""; // Отозвано и т.д.
                        sheet.Cell(currentRow, 28).Value = ""; // Предъявлено к исполнению

                        // Стили для строки итогов
                        var rowRange = sheet.Range(currentRow, 2, currentRow, 28);
                        rowRange.Style.Alignment.WrapText = true;
                        rowRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        rowRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                        OptimizeRowHeight(sheet, currentRow);
                    }

                    // Общие стили для блока итогов
                    var totalRange4 = sheet.Range(totalStartRow4, 2, totalStartRow4 + totalRowsCount4 - 1, 28);
                    totalRange4.Style.Font.Bold = true;

                    // Толстая граница сверху всего блока итогов
                    sheet.Range(totalStartRow4, 2, totalStartRow4, 28).Style.Border.TopBorder = XLBorderStyleValues.Medium;

                    dataRow = totalStartRow4 + totalRowsCount4;
                }
            }
            else
            {
                // Если данных нет - выводим сообщение "Данные отсутствуют" под шапкой
                sheet.Range($"B{dataRow}:AB{dataRow}").Merge().Value = "Данные отсутствуют";
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Font.Bold = true;
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                sheet.Range($"B{dataRow}:AB{dataRow}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                dataRow++;
            }

            // Пропускаем строку после четвертой таблицы
            dataRow += 1;

            // Подпись для четвертой таблицы
            int signatureRow4 = dataRow;
            sheet.Cell(signatureRow4, 2).Value = "Начальник юридического отдела ОАО \"Гомельтранснефть Дружба\"                                                    Ю.А.Лащенко";
            sheet.Range(signatureRow4, 2, signatureRow4, 17).Merge();
            sheet.Range(signatureRow4, 2, signatureRow4, 17).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            sheet.Row(signatureRow4).Height = 25;
            // Убираем границы для подписи
            sheet.Range(signatureRow4, 2, signatureRow4, 17).Style.Border.TopBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow4, 2, signatureRow4, 17).Style.Border.BottomBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow4, 2, signatureRow4, 17).Style.Border.LeftBorder = XLBorderStyleValues.None;
            sheet.Range(signatureRow4, 2, signatureRow4, 17).Style.Border.RightBorder = XLBorderStyleValues.None;

            //--------------------------------------------------------------------------------------------

            // Индивидуальная ширина колонок
            sheet.Column("B").Width = 4;
            sheet.Column("C").Width = 28;
            sheet.Column("D").Width = 20;
            sheet.Column("E").Width = 14;
            sheet.Column("F").Width = 28;
            sheet.Column("G").Width = 12;
            sheet.Column("H").Width = 12;
            sheet.Column("I").Width = 12;
            sheet.Column("J").Width = 12;
            sheet.Column("K").Width = 14;
            sheet.Column("L").Width = 12;
            sheet.Column("M").Width = 12;
            sheet.Column("N").Width = 12;
            sheet.Column("O").Width = 12;
            sheet.Column("P").Width = 12;
            sheet.Column("Q").Width = 12;
            sheet.Column("R").Width = 12;
            sheet.Column("S").Width = 15;
            sheet.Column("T").Width = 12;
            sheet.Column("U").Width = 12;
            sheet.Column("V").Width = 12;
            sheet.Column("W").Width = 12;
            sheet.Column("X").Width = 12;
            sheet.Column("Y").Width = 12;
            sheet.Column("Z").Width = 14;
            sheet.Column("AA").Width = 20;
            sheet.Column("AB").Width = 15;

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        //**************************************************************************************************************************************************
        private static void AddToTotals(Dictionary<string, decimal> totals, string currency, decimal amount)
        {
            if (totals.ContainsKey(currency))
            {
                totals[currency] += amount;
            }
            else
            {
                totals[currency] = amount;
            }
        }
        //--------------------------------------------------------------------------------------------------------------------------------------------------
        private static void OptimizeRowHeight(IXLWorksheet sheet, int rowNumber)
        {
            var row = sheet.Row(rowNumber);
            int maxLines = 1;

            foreach (var cell in row.CellsUsed())
            {
                var text = cell.GetString();
                if (string.IsNullOrEmpty(text)) continue;

                // Более точный расчет количества строк
                double columnWidth = sheet.Column(cell.Address.ColumnNumber).Width;

                // Приблизительное количество символов, которые помещаются в ширину колонки
                // Учитываем, что разные символы имеют разную ширину
                int approxCharsPerLine = Math.Max((int)(columnWidth * 1.8), 10); // Более консервативный коэффициент

                // Разделяем текст по переносам строк
                var lines = text.Split('\n');
                int totalLines = 0;

                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line)) continue;

                    // Для каждой линии считаем, сколько строк займет текст
                    int linesForThisText = Math.Max(1, (int)Math.Ceiling((double)line.Length / approxCharsPerLine));
                    totalLines += linesForThisText;
                }

                maxLines = Math.Max(maxLines, totalLines);
            }

            // Ограничиваем максимальную высоту
            double maxHeight = 100; // Максимальная высота строки
            double lineHeight = 15.0; // Высота одной линии
            double calculatedHeight = Math.Min(maxLines * lineHeight, maxHeight);

            row.Height = calculatedHeight;
        }
        //------Новый алгоритм вывода претензий ( с таблицами PretenseSumma, ResultSumma, Summa)-------------------------------------------------------------
        //----------------Список претензий------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult PretenseT()
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            List<Pretense> listpretense = new List<Pretense>();
            listpretense = db.Pretenses.ToList();

            List<Organization> listorganization = new List<Organization>();
            listorganization = db.Organizations.ToList();

            List<Valutum> listvaluta = new List<Valutum>();
            listvaluta = db.Valuta.OrderBy(l => l.Name).ToList();

            List<Filial> listfilial = new List<Filial>();
            listfilial = db.Filials.ToList();

            List<Predmet> listpredmet = new List<Predmet>();
            listpredmet = db.Predmets.ToList();

            List<Status> liststatus = new List<Status>();
            liststatus = db.Statuses.ToList();

            List<Summa> listsumma = new List<Summa>();
            listsumma = db.Summas.ToList();

            List<PretenseSumma> listpretensesumma = new List<PretenseSumma>();
            listpretensesumma = db.PretenseSummas.ToList();

            List<ResultSumma> listresultsumma = new List<ResultSumma>();
            listresultsumma = db.ResultSummas.ToList();

            List<TablePretense> listTablePretense = new List<TablePretense>();
            listTablePretense = db.TablePretenses.Where(o => o.Delet != 1).ToList();

            var listpretenseJoin = (
from pretense in listpretense
join organization in listorganization on pretense.OrgId equals organization.OrgId
join filial in listfilial on pretense.FilId equals filial.FilId
join predmet in listpredmet on pretense.PredmetId equals predmet.PredmetId

// Получаем суммы претензии с типами и валютами
let pretenseSummas = (
    from ps in listpretensesumma
    join summaType in listsumma on ps.SummaId equals summaType.SummaId
    join valuta in listvaluta on ps.ValId equals valuta.ValId
    where ps.PretId == pretense.PretId
    select new
    {
        SummaId = summaType.SummaId,
        SummaType = summaType.Name,
        Value = ps.Value,
        ValId = ps.ValId,
        ValName = valuta.Name,
        ValFullName = valuta.NameFull
    }
).ToList()

select new
{
    PretId = pretense.PretId,
    OrgId = pretense.OrgId,
    OrgName = organization.Name,
    UNP = organization.Unp,
    Address = organization.Address,
    NumberPret = pretense.NumberPret,
    DatePret = pretense.DatePret,
    Inout = pretense.Inout,
    Visible = pretense.Visible,
    Arhiv = pretense.Arhiv,
    FilId = pretense.FilId,
    FilName = filial.Name,
    PredmetId = pretense.PredmetId,
    PredmetName = predmet.Predmet1,
    UserMod = pretense.UserMod,
    DateMod = pretense.DateMod,

    // Извлекаем конкретные суммы по типам из PretenseSumma
    SummaDolg = pretenseSummas.FirstOrDefault(s => s.SummaId == 1)?.Value ?? 0,
    SummaDolgValId = pretenseSummas.FirstOrDefault(s => s.SummaId == 1)?.ValId,
    SummaDolgValName = pretenseSummas.FirstOrDefault(s => s.SummaId == 1)?.ValName ?? string.Empty,

    SummaPeny = pretenseSummas.FirstOrDefault(s => s.SummaId == 2)?.Value ?? 0,
    SummaPenyValId = pretenseSummas.FirstOrDefault(s => s.SummaId == 2)?.ValId,
    SummaPenyValName = pretenseSummas.FirstOrDefault(s => s.SummaId == 2)?.ValName ?? string.Empty,

    SummaProc = pretenseSummas.FirstOrDefault(s => s.SummaId == 3)?.Value ?? 0,
    SummaProcValId = pretenseSummas.FirstOrDefault(s => s.SummaId == 3)?.ValId,
    SummaProcValName = pretenseSummas.FirstOrDefault(s => s.SummaId == 3)?.ValName ?? string.Empty,

    SummaPoshlina = pretenseSummas.FirstOrDefault(s => s.SummaId == 4)?.Value ?? 0,
    SummaPoshlinaValId = pretenseSummas.FirstOrDefault(s => s.SummaId == 4)?.ValId,
    SummaPoshlinaValName = pretenseSummas.FirstOrDefault(s => s.SummaId == 4)?.ValName ?? string.Empty,

    TablePretenses = (
        from tp in listTablePretense
        join status in liststatus on tp.StatusId equals status.StatusId

        // Получаем суммы результатов с типами и валютами из ResultSumma
        let resultSummas = (
            from rs in listresultsumma
            join summaType in listsumma on rs.SummaId equals summaType.SummaId
            join valuta in listvaluta on rs.ValId equals valuta.ValId
            where rs.ResultId == tp.TabPretId
            select new
            {
                SummaId = summaType.SummaId,
                SummaType = summaType.Name,
                Value = rs.Value,
                ValId = rs.ValId,
                ValName = valuta.Name
            }
        ).ToList()

        where tp.PretId == pretense.PretId
        select new
        {
            tp.TabPretId,
            tp.DateTabPret,

            // Извлекаем конкретные суммы результатов по типам
            SummaDolg = resultSummas.FirstOrDefault(s => s.SummaId == 1)?.Value ?? 0,
            SummaDolgValId = resultSummas.FirstOrDefault(s => s.SummaId == 1)?.ValId,
            SummaDolgValName = resultSummas.FirstOrDefault(s => s.SummaId == 1)?.ValName ?? string.Empty,

            SummaPeny = resultSummas.FirstOrDefault(s => s.SummaId == 2)?.Value ?? 0,
            SummaPenyValId = resultSummas.FirstOrDefault(s => s.SummaId == 2)?.ValId,
            SummaPenyValName = resultSummas.FirstOrDefault(s => s.SummaId == 2)?.ValName ?? string.Empty,

            SummaProc = resultSummas.FirstOrDefault(s => s.SummaId == 3)?.Value ?? 0,
            SummaProcValId = resultSummas.FirstOrDefault(s => s.SummaId == 3)?.ValId,
            SummaProcValName = resultSummas.FirstOrDefault(s => s.SummaId == 3)?.ValName ?? string.Empty,

            SummaPoshlina = resultSummas.FirstOrDefault(s => s.SummaId == 4)?.Value ?? 0,
            SummaPoshlinaValId = resultSummas.FirstOrDefault(s => s.SummaId == 4)?.ValId,
            SummaPoshlinaValName = resultSummas.FirstOrDefault(s => s.SummaId == 4)?.ValName ?? string.Empty,

            summaItog = resultSummas.Sum(s => s.Value ?? 0),

            tp.Result,
            tp.Primechanie,
            tp.UserMod,
            tp.DateMod,
            tp.StatusId,
            StatusName = status.Name
        }
    ).ToList()
})
.Where(i => i.Visible != 0 && i.Arhiv != 1)
.OrderBy(x => x.FilName)
.ThenBy(u => u.OrgName)
.ToList();

            return Ok(listpretenseJoin);
        }
        //--Добавление претензии------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddPretenseT([FromBody] PretenseValId pretense)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                Pretense pret = new Pretense();
                pret.NumberPret = pretense.NumberPret;
                pret.DatePret = pretense.DatePret;
                pret.OrgId = pretense.OrgId;
                pret.DateRassmPret = pretense.DateRassmPret;
                pret.FilId = pretense.FilId;
                pret.Inout = pretense.Inout;
                pret.PredmetId = pretense.PredmetId;
                pret.Visible = 1;
                pret.Arhiv = 0;
                pret.UserMod = username;
                pret.DateMod = DateTime.Now;
                db.Pretenses.Add(pret);
                db.SaveChanges();

                // Получаем ID только что созданной претензии
                int newPretId = pret.PretId;

                //Заполняем таблицу с суммами и валютами для этих сумм
                PretenseSumma summaDolg = new PretenseSumma();
                summaDolg.PretId = newPretId;
                summaDolg.SummaId = 1;
                summaDolg.ValId = pretense.SummaDolgValId;
                summaDolg.Value = pretense.SummaDolg ?? 0;
                db.PretenseSummas.Add(summaDolg);

                PretenseSumma summaPeny = new PretenseSumma();
                summaPeny.PretId = newPretId;
                summaPeny.SummaId = 2;
                summaPeny.ValId = pretense.SummaPenyValId;
                summaPeny.Value = pretense.SummaPeny ?? 0;
                db.PretenseSummas.Add(summaPeny);

                PretenseSumma summaProc = new PretenseSumma();
                summaProc.PretId = newPretId;
                summaProc.SummaId = 3;
                summaProc.ValId = pretense.SummaProcValId;
                summaProc.Value = pretense.SummaProc ?? 0;
                db.PretenseSummas.Add(summaProc);

                PretenseSumma summaPoshlina = new PretenseSumma();
                summaPoshlina.PretId = newPretId;
                summaPoshlina.SummaId = 4;
                summaPoshlina.ValId = pretense.SummaPoshlinaValId;
                summaPoshlina.Value = pretense.SummaPoshlina ?? 0;
                db.PretenseSummas.Add(summaPoshlina);
                db.SaveChanges();

                return Ok(pret);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }
        //--Редактирование претензии------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditPretenseT([FromBody] PretenseValId pretense)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                Pretense pret = new Pretense();
                pret = db.Pretenses.FirstOrDefault(s => s.PretId == pretense.PretId);

                pret.NumberPret = pretense.NumberPret;
                pret.DatePret = pretense.DatePret;
                pret.OrgId = pretense.OrgId;
                pret.DateRassmPret = pretense.DateRassmPret;
                pret.FilId = pretense.FilId;
                pret.Inout = pretense.Inout;
                pret.PredmetId = pretense.PredmetId;
                pret.UserMod = username;
                pret.DateMod = DateTime.Now;
                
                List<PretenseSumma> listpretsumma = new List<PretenseSumma>();
                listpretsumma = db.PretenseSummas.Where(pr => pr.PretId == pretense.PretId).ToList();

                foreach(var item in listpretsumma)
                {
                    if (item.SummaId == 1)
                    {                        
                        item.ValId = pretense.SummaDolgValId;
                        item.Value = pretense.SummaDolg ?? 0;
                    }
                    else if(item.SummaId == 2)
                    {
                        item.ValId = pretense.SummaPenyValId;
                        item.Value = pretense.SummaPeny ?? 0;
                    }
                    else if(item.SummaId == 3)
                    {
                        item.ValId = pretense.SummaProcValId;
                        item.Value = pretense.SummaProc ?? 0;
                    }
                    else if(item.SummaId == 4)
                    {
                        item.ValId = pretense.SummaPoshlinaValId;
                        item.Value = pretense.SummaPoshlina ?? 0;
                    }
                }
                db.SaveChanges();

                return Ok(pret);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //--------------------------------------------------------------------------------------
        //--Добавление результата------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult AddResultT([FromBody] TablePretenseValId result)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------

            try
            {
                TablePretense tab = new TablePretense();
                tab.PretId = result.PretId;
                tab.DateTabPret = result.DateTabPret;
                tab.Result = result.Result;
                tab.StatusId = result.StatusId;
                tab.UserMod = username;
                tab.DateMod = DateTime.Now;
                db.TablePretenses.Add(tab);
                db.SaveChanges();

                // Получаем ID только что созданной претензии
                int newTabPretId = tab.TabPretId;

                //Заполняем таблицу с суммами и валютами для этих сумм
                ResultSumma summaDolg = new ResultSumma();
                summaDolg.ResultId = newTabPretId;
                summaDolg.SummaId = 1;
                summaDolg.ValId = result.SummaDolgValId;
                summaDolg.Value = result.SummaDolg ?? 0;
                db.ResultSummas.Add(summaDolg);

                ResultSumma summaPeny = new ResultSumma();
                summaPeny.ResultId = newTabPretId;
                summaPeny.SummaId = 2;
                summaPeny.ValId = result.SummaPenyValId;
                summaPeny.Value = result.SummaPeny ?? 0;
                db.ResultSummas.Add(summaPeny);

                ResultSumma summaProc = new ResultSumma();
                summaProc.ResultId = newTabPretId;
                summaProc.SummaId = 3;
                summaProc.ValId = result.SummaProcValId;
                summaProc.Value = result.SummaProc ?? 0;
                db.ResultSummas.Add(summaProc);

                ResultSumma summaPoshlina = new ResultSumma();
                summaPoshlina.ResultId = newTabPretId;
                summaPoshlina.SummaId = 4;
                summaPoshlina.ValId = result.SummaPoshlinaValId;
                summaPoshlina.Value = result.SummaPoshlina ?? 0;
                db.ResultSummas.Add(summaPoshlina);
                db.SaveChanges();

                return Ok(tab);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при добавлении записи");
            }
        }
        //--Редактирование результата------------------------------------------------

        [HttpPost]
        [Route("[action]")]
        public IActionResult EditResultT([FromBody] TablePretenseValId result)
        {
            //-------------------------------------------------------------------------------------------------
            // Извлекаем токен из заголовков запроса
            var authHeader = Request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ") ? authHeader.Substring("Bearer ".Length) : authHeader;

            // Разбираем токен
            var handler = new JwtSecurityTokenHandler();
            var jwtToken = handler.ReadToken(token) as JwtSecurityToken;

            // Получаем утверждения
            var username = jwtToken.Claims.First(claim => claim.Type == JwtRegisteredClaimNames.Sub).Value;
            var filialId = int.Parse(jwtToken.Claims.First(claim => claim.Type == "FilialId").Value);
            var fio = jwtToken.Claims.First(claim => claim.Type == "FIO").Value;
            var admin = int.Parse(jwtToken.Claims.First(claim => claim.Type == "Admin").Value);
            //--------------------------------------------------------------
            try
            {
                TablePretense tab = new TablePretense();
                tab = db.TablePretenses.FirstOrDefault(s => s.TabPretId == result.TabPretId);
                //tab.PretId = result.PretId;
                tab.DateTabPret = result.DateTabPret;
                tab.Result = result.Result;
                tab.StatusId = result.StatusId;
                tab.UserMod = username;
                tab.DateMod = DateTime.Now;

                List<ResultSumma> listressumma = new List<ResultSumma>();
                listressumma = db.ResultSummas.Where(pr => pr.ResultId == result.TabPretId).ToList();

                foreach (var item in listressumma)
                {
                    if (item.SummaId == 1)
                    {
                        item.ValId = result.SummaDolgValId;
                        item.Value = result.SummaDolg ?? 0;
                    }
                    else if (item.SummaId == 2)
                    {
                        item.ValId = result.SummaPenyValId;
                        item.Value = result.SummaPeny ?? 0;
                    }
                    else if (item.SummaId == 3)
                    {
                        item.ValId = result.SummaProcValId;
                        item.Value = result.SummaProc ?? 0;
                    }
                    else if (item.SummaId == 4)
                    {
                        item.ValId = result.SummaPoshlinaValId;
                        item.Value = result.SummaPoshlina ?? 0;
                    }
                }

                db.SaveChanges();
                return Ok(tab);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка: {ex.Message}");
                return StatusCode(500, "Произошла ошибка при редактировании записи");
            }
        }
        //---------------------------------------------------------------------------------------------------------------------------------------------------
    }
}
