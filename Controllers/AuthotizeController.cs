using Jurist.Models;
using Jurist.Services;
using DocumentFormat.OpenXml.Bibliography;
using Jurist.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Cors.Infrastructure;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using Microsoft.VisualStudio.Web.CodeGenerators.Mvc.Templates.BlazorIdentity.Pages;
using System.Security.Cryptography;
using System.Text;

namespace ApiMetrolog.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    public class AuthorizeController : Controller
    {
        public JuristContext db;
        private readonly TokenService _tokenService;

        // Объединенный конструктор
        public AuthorizeController(JuristContext context, TokenService tokenService)
        {
            db = context;
            _tokenService = tokenService;
        }


        //[HttpPost("login")]
        [HttpPost]
        [Route("[action]")]
        public IActionResult Login([FromBody] LoginJ loginModel)
        {
            //Найдем пользователя в базе данных
            LoginPassword login = new LoginPassword();
            
            login = db.LoginPasswords.FirstOrDefault(u => u.Login == loginModel.Login1);
            string pass = login.Password;
            string pass1 = BCrypt.Net.BCrypt.HashPassword(loginModel.Password);
            var test = BCrypt.Net.BCrypt.Verify(loginModel.Password, login.Password);

            if (login != null && BCrypt.Net.BCrypt.Verify(loginModel.Password, login.Password))
            {
                // Найдем отдельно всю инфо по пользователю
                Employee responsible = new Employee();
                responsible = db.Employees.FirstOrDefault(u => u.EmplId == login.EmplId);

                string FIO = "";
                //FIO = responsible.LastName + " " + responsible.FirstName + " " + responsible.MiddleName;
                FIO = responsible.LastName + " " +
                (string.IsNullOrEmpty(responsible.FirstName) ? "" : responsible.FirstName[0] + ".") +
                (string.IsNullOrEmpty(responsible.MiddleName) ? "" : responsible.MiddleName[0] + ".");

                List<Filial> fillist = new List<Filial>();
                fillist = db.Filials.ToList();

                Filial fil = new Filial();
                fil = fillist.FirstOrDefault(h => h.FilId == responsible.FilId);

                LoginModel loginmodel = new LoginModel();
                var token = _tokenService.GenerateToken(loginModel.Login1, responsible.LastName, responsible.FirstName, responsible.MiddleName, FIO, fil.FilId, fil.Name, login.Admin, login.ThemeColor);
                loginmodel.token = token;
                loginmodel.Login1 = login.Login;
                loginmodel.LastName = responsible.LastName;
                loginmodel.FirstName = responsible.FirstName;
                loginmodel.MiddleName = responsible.MiddleName;
                loginmodel.fio = FIO;
                loginmodel.Admin = login.Admin;
                loginmodel.FilialId = fil.FilId;
                loginmodel.FilialName = fil.Name;
                loginmodel.ThemeColor = login.ThemeColor;

                return Ok(loginmodel);
            }
            else
            {
                return Unauthorized();
            }
            //login = db.LoginPasswords.FirstOrDefault(u => u.Login == loginModel.Login1 && u.Password == loginModel.Password);
            //login = db.Logins.FirstOrDefault(g => g.Login1 == loginModel.Login1 & g.Password == md5.hashPassword(loginModel.Password));
       }

    }
}
