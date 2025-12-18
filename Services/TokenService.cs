using System;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using Jurist.Models;

namespace Jurist.Services
{
    public class TokenService
    {

        private readonly IConfiguration _configuration;

        public TokenService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public string GenerateToken(string username, string LastName, string FirstName, string MiddleName, string FIO, int filialId, string filName, int? admin, int? adds, int? upds, int? dels, string ThemeColor)
        {
            var claims = new[]
    {
        new Claim(JwtRegisteredClaimNames.Sub, username),
        new Claim("LastName", LastName),
        new Claim("FirstName", FirstName),
        new Claim("MiddleName", MiddleName),
        new Claim("FIO", FIO),
        new Claim("FilialId", filialId.ToString()),
        new Claim("filName", filName),
        new Claim("Admin", admin.ToString()),
        new Claim("Adds", adds.ToString()),
        new Claim("Upds", upds.ToString()),
        new Claim("Dels", dels.ToString()),
        new Claim("ThemeColor", ThemeColor),
        new Claim(JwtRegisteredClaimNames.Jti, Guid.NewGuid().ToString())
    };

            var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_configuration["Jwt:Key"]));
            var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);

            var token = new JwtSecurityToken(
                issuer: _configuration["Jwt:Issuer"],
                audience: _configuration["Jwt:Audience"],
                claims: claims,
                expires: DateTime.Now.AddMinutes(30),
                signingCredentials: creds);

            return new JwtSecurityTokenHandler().WriteToken(token);
        }

    }
}
