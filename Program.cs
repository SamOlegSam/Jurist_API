using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Jurist.Models;
using Microsoft.AspNetCore.Cors.Infrastructure;
using Microsoft.AspNetCore.Builder;
using Jurist.Services;
using System.Text.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

// Регистрация TokenService
builder.Services.AddSingleton<TokenService>();

var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");

builder.Services.AddDbContext<JuristContext>(options =>
    options.UseSqlServer(connectionString));
builder.Services.AddDatabaseDeveloperPageExceptionFilter();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Добавление CORS
//builder.Services.AddCors(options =>
//{
//    options.AddPolicy(
//    "AllowEverything",
//    builder => builder
//        .AllowAnyOrigin()
//        .AllowAnyHeader()
//        .AllowAnyMethod());

//});

builder.Services.AddCors(options =>
{    
    options.AddPolicy(
        "AllowLocalhost",
        builder => builder
            .WithOrigins("http://localhost:3000", "https://localhost:3000", "http://s38vm:88", "https://s38vm:88", "http://s38vm:89", "https://s38vm:89")
            .AllowAnyHeader()
            .AllowAnyMethod()
            .AllowCredentials()
            .WithExposedHeaders("*"));
});


// Добавьте сервисы для JWT аутентификации
builder.Services.AddAuthentication(options =>
{
    options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
    options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
})
.AddJwtBearer(options =>
{
    options.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuer = true,
        ValidateAudience = true,
        ValidateLifetime = true,
        ValidateIssuerSigningKey = true,
        ValidIssuer = builder.Configuration["Jwt:Issuer"],
        ValidAudience = builder.Configuration["Jwt:Audience"],
        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(builder.Configuration["Jwt:Key"]))
    };
});

// Добавление контроллеров
builder.Services.AddControllers();
//Добавление контроллеров с настройкой JSON сериализации
//builder.Services.AddControllers().AddJsonOptions(options =>
//    options.JsonSerializerOptions.ReferenceHandler = ReferenceHandler.Preserve);

var app = builder.Build();

app.UseRouting();

// Использование CORS
app.UseCors("AllowLocalhost");

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

// Использование аутентификации и авторизации
app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

app.Run();



