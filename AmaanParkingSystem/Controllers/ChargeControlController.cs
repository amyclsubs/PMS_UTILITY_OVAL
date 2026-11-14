using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using AmaanParkingSystem.Services;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace AmaanParkingSystem.Controllers
{
    public class ChargeControlController : BaseController
    {
        private readonly IConnectionStringService _connStrService;
        private readonly IConfiguration _configuration;
        private readonly ILogger<ChargeControlController> _logger;

        private const string PENALTY = "isPenaltyActive";
        private const string SPECIAL = "IsSpecialOccasionActive";
        private const string SPECIAL_AT = "SpecialOccasionActivatedAt";

        public ChargeControlController(
            IConnectionStringService connStrService,
            IConfiguration configuration,
            ILogger<ChargeControlController> logger)
        {
            _connStrService = connStrService;
            _configuration = configuration;
            _logger = logger;
        }

        private string GetConnectionString()
        {
            return _connStrService.GetConnectionString()
                   ?? _configuration.GetConnectionString("DefaultConnection")
                   ?? throw new InvalidOperationException("No connection string found.");
        }

        // GET: /ChargeControl
        public IActionResult Index()
        {
            return View();
        }

        // GET: /ChargeControl/GetSettings
        [HttpGet]
        public async Task<IActionResult> GetSettings()
        {
            try
            {
                using var conn = new SqlConnection(GetConnectionString());
                var rows = (await conn.QueryAsync(
                    "SELECT [SettingName], [SettingValue] FROM [AMAAN_PMS].[dbo].[SystemSettings]")).ToList();

                string Get(string name)
                {
                    var r = rows.FirstOrDefault(x =>
                        string.Equals((string)x.SettingName, name, StringComparison.OrdinalIgnoreCase));
                    return r == null ? null : (string)r.SettingValue;
                }

                bool IsTrue(string v) => string.Equals(v?.Trim(), "true", StringComparison.OrdinalIgnoreCase);

                return Json(new
                {
                    success = true,
                    isPenaltyActive = IsTrue(Get(PENALTY)),
                    isSpecialOccasionActive = IsTrue(Get(SPECIAL)),
                    specialOccasionActivatedAt = Get(SPECIAL_AT)
                });
            }
            catch (Exception ex)
            {
                _logger.LogError($"ChargeControl GetSettings Error: {ex.Message}");
                return Json(new { success = false, message = ex.Message });
            }
        }

        // POST: /ChargeControl/SetSetting
        [HttpPost]
        public async Task<IActionResult> SetSetting([FromBody] SetSettingRequest request)
        {
            try
            {
                if (request == null ||
                    (request.SettingName != PENALTY && request.SettingName != SPECIAL))
                    return Json(new { success = false, message = "Invalid setting." });

                using var conn = new SqlConnection(GetConnectionString());

                await UpsertAsync(conn, request.SettingName, request.Enable ? "true" : "false");

                if (request.SettingName == SPECIAL)
                {
                    string activatedAt = request.Enable ? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") : null;
                    await UpsertAsync(conn, SPECIAL_AT, activatedAt);
                }

                string msg = request.SettingName == PENALTY
                    ? (request.Enable ? "✅ Penalty ACTIVATED" : "⏸️ Penalty DEACTIVATED")
                    : (request.Enable ? "🎉 Special Occasion ACTIVATED — temporary cars are free" : "⏸️ Special Occasion DEACTIVATED — charges active");

                _logger.LogInformation($"[ChargeControl] {msg}");
                return Json(new { success = true, message = msg });
            }
            catch (Exception ex)
            {
                _logger.LogError($"ChargeControl SetSetting Error: {ex.Message}");
                return Json(new { success = false, message = ex.Message });
            }
        }

        private static async Task UpsertAsync(SqlConnection conn, string name, string value)
        {
            const string sql = @"
                UPDATE [AMAAN_PMS].[dbo].[SystemSettings]
                   SET [SettingValue] = @Value
                 WHERE [SettingName] = @Name;
                IF @@ROWCOUNT = 0
                    INSERT INTO [AMAAN_PMS].[dbo].[SystemSettings] ([SettingName], [SettingValue])
                    VALUES (@Name, @Value);";

            await conn.ExecuteAsync(sql, new { Name = name, Value = value });
        }

        public class SetSettingRequest
        {
            public string SettingName { get; set; }
            public bool Enable { get; set; }
        }
    }
}