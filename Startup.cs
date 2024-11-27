using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using reportesApi.Extensions;
using reportesApi.Models;
using reportesApi.Services;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;

namespace reportesApi
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
      
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            var key = "BusinessCaseMarcatelApiV2Tibs!*";

            // requires using Microsoft.Extensions.Options
            services.Configure<MarcatelDatabaseSetting>(
            Configuration.GetSection(nameof(MarcatelDatabaseSetting)));

            services.Configure<UrlUpload>(
            Configuration.GetSection(nameof(UrlUpload)));

            //services.Configure<EmailstoreSettings>(
                //Configuration.GetSection(nameof(EmailstoreSettings)));

            // Get KudotibsDatabaseSettings
            services.AddSingleton<IMarcatelDatabaseSetting>(sp =>
            sp.GetRequiredService<IOptions<MarcatelDatabaseSetting>>().Value);

            services.AddSingleton<UrlUpload>(sp =>
            sp.GetRequiredService<IOptions<UrlUpload>>().Value);



            // Get EmailstoreSettings
            //services.AddSingleton<IEmailstoreSettings>(sp =>
            //    sp.GetRequiredService<IOptions<EmailstoreSettings>>().Value);

            // Get Services LB
            services.AddSingleton<LoginService>();
            // services.AddSingleton<FuncionesExtraService>();
            services.AddSingleton<ExistenciasService>();
            services.AddSingleton<ExcelService>();
            services.AddSingleton<TipoMovimientoService>();
            services.AddSingleton<MovimientosService>();
            services.AddSingleton<RenglonMovimientoService>();
            services.AddSingleton<TRSPService>();
            services.AddSingleton<RecetasService>();
            services.AddSingleton<DetalleRecetaService>();
            services.AddSingleton<OrdenCompraService>();
            services.AddSingleton<DetalleOrdenCompraService>();
            services.AddSingleton<EntradaService>();
            services.AddSingleton<DetalleEntradaService>();
            services.AddSingleton<ProveedorService>();
            services.AddSingleton<InsumoService>();
            services.AddSingleton<AlmacenService>();
            services.AddSingleton<PersonasService>();

            // services.AddSingleton<GrupoService>();
            // services.AddSingleton<GrupoAlumnoService>();
            // services.AddSingleton<MateriaAlumnoService>();
            // services.AddSingleton<GrupoMateriaService>();



            services.AddCors();

            services.AddControllers()
                .AddNewtonsoftJson(options => options.UseMemberCasing())
                .AddJsonOptions(options =>
                {
                    options.JsonSerializerOptions.IgnoreNullValues = false;
                    options.JsonSerializerOptions.WriteIndented = true;
                });

            services
                .AddAuthentication(x =>
                {
                    x.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
                    x.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
                }).AddJwtBearer(x =>
                {
                    x.RequireHttpsMetadata = false;
                    x.SaveToken = true;
                    x.TokenValidationParameters = new TokenValidationParameters
                    {
                        ValidateIssuerSigningKey = true,
                        ValidateIssuer = false,
                        ValidateAudience = false,
                        ValidateLifetime = true,
                        ValidIssuer = "https://marcatelapi.herokuapp.com/",
                        ValidAudience = "https://marcatelapi.herokuapp.com/",
                        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(key))
                    };
                });

            services.AddAuthorization();
            services.AddSingleton<IJwtAuthenticationService>(new JwtAuthenticationService(key));
            services.AddHttpClient();
            services.AddSingleton<IHttpContextAccessor, HttpContextAccessor>();


            // Context DB
            //services.AddDbContext<TodoContext>(opt => opt.UseInMemoryDatabase("TodoList"));

            // Register the Swagger generator, defening 1 or more Swgager documents
            services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
                {
                    Version = "v1",
                    Title = "Backend",
                    Description = "ASP.NET Core Web API for DataBrain reports",
                    TermsOfService = new Uri("https://example.com/terms"),
                    Contact = new Microsoft.OpenApi.Models.OpenApiContact
                    {
                        Name = "Next Generation Software",
                        Email = string.Empty,
                        Url = new Uri("https://ainextgen.mx/"),
                    },
                    License = new Microsoft.OpenApi.Models.OpenApiLicense
                    {
                        Name = "License type",
                        Url = new Uri("https://example.com/license"),
                    }
                });
            });
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            
            // Enable middleware to serve generated swagger as a JSON endpoint.
            app.UseSwagger();

            // Enable middleware to serve swagger-ui (HTML, JS, CSS, etc.),
            // specifying the Swagger JSON endpoint.
            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "Marcatel API V1");
                c.RoutePrefix = string.Empty;
            });

            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            //  app.ConfigureExceptionHandler(logger);
            app.ConfigureCustomExceptionMiddleware();
            app.UseHttpsRedirection();

            app.UseRouting();

            // global cors policy
            app.UseCors(x => x
                .AllowAnyMethod()
                .AllowAnyHeader()
                .SetIsOriginAllowed(origin => true) // allow any origin
                .AllowCredentials()); // allow credentials

            app.UseAuthorization();

            app.UseAuthentication();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
