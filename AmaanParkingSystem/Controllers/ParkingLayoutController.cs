using Dapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace AmaanParkingSystem.Controllers
{
    // Row shown on the Parking Layout page (one per zone)
    public class ParkingLayoutZoneVm
    {
        public string Id { get; set; } = "";
        public string SiteId { get; set; } = "";
        public string ZoneId { get; set; } = "";
        public string ZoneName { get; set; } = "";
        public int Capacity { get; set; }
        public string Status { get; set; } = "";
        public int SlotCount { get; set; }
        public string CreatedBy { get; set; } = "";
        public string CreatedOn { get; set; } = "";
        public string UpdatedBy { get; set; } = "";
        public string UpdatedOn { get; set; } = "";

        public bool IsActive
        {
            get
            {
                var s = (Status ?? "").Trim().ToLowerInvariant();
                return s == "1" || s == "true" || s == "active" || s == "y";
            }
        }
    }

    public class ParkingLayoutController : BaseController
    {
        private const string PlZonesTable = "[AMAAN_PMS].[dbo].[db_tbl_04_zones]";
        private const string PlLayoutTable = "[AMAAN_PMS].[dbo].[db_tbl_31_Parking_layout]";

        // OPERATOR role can view only; MANAGER / ADMIN can add, edit and delete
        private bool PlCanWrite()
        {
            var op = HttpContext.Session.GetString("OperatorType") ?? string.Empty;
            return !string.Equals(op, "OPERATOR", StringComparison.OrdinalIgnoreCase);
        }

        [HttpGet]
        public IActionResult Index()
        {
            using var conn = new SqlConnection(GetDynamicConnectionString());

            var zones = conn.Query<ParkingLayoutZoneVm>($@"
                SELECT
                    CAST(z.id AS NVARCHAR(50)) AS Id,
                    ISNULL(CAST(z.site_id AS NVARCHAR(50)), '') AS SiteId,
                    ISNULL(CAST(z.zone_id AS NVARCHAR(50)), '') AS ZoneId,
                    ISNULL(CAST(z.zone_name AS NVARCHAR(200)), '') AS ZoneName,
                    ISNULL(TRY_CAST(z.capacity AS INT), 0) AS Capacity,
                    ISNULL(CAST(z.status AS NVARCHAR(20)), '0') AS Status,
                    (SELECT COUNT(*) FROM {PlLayoutTable} p
                      WHERE CAST(p.zone_id AS NVARCHAR(50)) = CAST(z.zone_id AS NVARCHAR(50))) AS SlotCount,
                    ISNULL(CAST(z.created_by AS NVARCHAR(100)), '') AS CreatedBy,
                    ISNULL(CONVERT(NVARCHAR(30), z.created_on, 120), '') AS CreatedOn,
                    ISNULL(CAST(z.updated_by AS NVARCHAR(100)), '') AS UpdatedBy,
                    ISNULL(CONVERT(NVARCHAR(30), z.updated_on, 120), '') AS UpdatedOn
                FROM {PlZonesTable} z
                ORDER BY z.zone_name ASC").ToList();

            ViewBag.DefaultSiteId = zones.Select(z => z.SiteId).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? "";

            return View("~/Views/ParkingLayout/Index.cshtml", zones);
        }

        [HttpPost]
        public IActionResult SaveZone(string id, string site_id, string zone_name, int capacity, string status)
        {
            if (!PlCanWrite())
            {
                TempData["ErrorMessage"] = "You do not have permission to change zones.";
                return RedirectToAction("Index");
            }

            var zoneName = (zone_name ?? "").Trim();
            var siteId = (site_id ?? "").Trim();
            var zoneKey = (id ?? "").Trim();
            var statusValue = status == "0" ? "0" : "1";
            if (capacity < 0) capacity = 0;

            if (zoneName.Length == 0)
            {
                TempData["ErrorMessage"] = "Zone name is required.";
                return RedirectToAction("Index");
            }
            if (siteId.Length == 0)
            {
                TempData["ErrorMessage"] = "Site ID is required.";
                return RedirectToAction("Index");
            }

            using var conn = new SqlConnection(GetDynamicConnectionString());
            var currentUser = GetLoggedInUser();

            // Duplicate zone name within the same site
            var duplicateExists = conn.ExecuteScalar<int>($@"
                SELECT COUNT(*) FROM {PlZonesTable}
                WHERE LTRIM(RTRIM(LOWER(CAST(zone_name AS NVARCHAR(200))))) = LOWER(@zoneName)
                  AND CAST(site_id AS NVARCHAR(50)) = @siteId
                  AND CAST(id AS NVARCHAR(50)) <> @zoneKey",
                new { zoneName, siteId, zoneKey }) > 0;

            if (duplicateExists)
            {
                TempData["ErrorMessage"] = $"Zone '{zoneName}' already exists for this site.";
                return RedirectToAction("Index");
            }

            if (zoneKey.Length > 0)
            {
                // UPDATE
                var rows = conn.Execute($@"
                    UPDATE {PlZonesTable} SET
                        site_id    = @siteId,
                        zone_name  = @zoneName,
                        capacity   = @capacity,
                        status     = @statusValue,
                        updated_by = @currentUser,
                        updated_on = GETDATE()
                    WHERE CAST(id AS NVARCHAR(50)) = @zoneKey",
                    new { siteId, zoneName, capacity, statusValue, currentUser, zoneKey });

                TempData[rows > 0 ? "SuccessMessage" : "ErrorMessage"] =
                    rows > 0 ? $"Zone '{zoneName}' updated successfully!" : "Zone not found.";
            }
            else
            {
                // INSERT - next zone_id = highest numeric zone_id + 1
                var nextZoneNo = (conn.ExecuteScalar<int?>($"SELECT MAX(TRY_CAST(zone_id AS INT)) FROM {PlZonesTable}") ?? 0) + 1;
                while (conn.ExecuteScalar<int>(
                    $"SELECT COUNT(*) FROM {PlZonesTable} WHERE CAST(zone_id AS NVARCHAR(50)) = @zn",
                    new { zn = nextZoneNo.ToString() }) > 0)
                {
                    nextZoneNo++;
                }
                var newZoneId = nextZoneNo.ToString();

                var cols = "site_id, zone_id, zone_name, capacity, status, created_by, created_on";
                var vals = "@siteId, @newZoneId, @zoneName, @capacity, @statusValue, @currentUser, GETDATE()";
                string newRowId = "";

                // If [id] is not an IDENTITY column, generate it the same way
                var idIsIdentity = conn.ExecuteScalar<int?>(
                    $"SELECT COLUMNPROPERTY(OBJECT_ID('{PlZonesTable}'), 'id', 'IsIdentity')") ?? 0;
                if (idIsIdentity == 0)
                {
                    var nextId = (conn.ExecuteScalar<int?>($"SELECT MAX(TRY_CAST(id AS INT)) FROM {PlZonesTable}") ?? 0) + 1;
                    while (conn.ExecuteScalar<int>(
                        $"SELECT COUNT(*) FROM {PlZonesTable} WHERE CAST(id AS NVARCHAR(50)) = @nid",
                        new { nid = nextId.ToString() }) > 0)
                    {
                        nextId++;
                    }
                    newRowId = nextId.ToString();
                    cols = "id, " + cols;
                    vals = "@newRowId, " + vals;
                }

                conn.Execute($"INSERT INTO {PlZonesTable} ({cols}) VALUES ({vals})",
                    new { newRowId, siteId, newZoneId, zoneName, capacity, statusValue, currentUser });

                TempData["SuccessMessage"] = $"Zone '{zoneName}' added successfully! (Zone ID: {newZoneId})";
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult DeleteZone(string id)
        {
            if (!PlCanWrite())
            {
                TempData["ErrorMessage"] = "You do not have permission to change zones.";
                return RedirectToAction("Index");
            }

            var zoneKey = (id ?? "").Trim();
            using var conn = new SqlConnection(GetDynamicConnectionString());

            var zoneId = conn.QueryFirstOrDefault<string>(
                $"SELECT CAST(zone_id AS NVARCHAR(50)) FROM {PlZonesTable} WHERE CAST(id AS NVARCHAR(50)) = @zoneKey",
                new { zoneKey });

            if (zoneId == null)
            {
                TempData["ErrorMessage"] = "Zone not found.";
                return RedirectToAction("Index");
            }

            var hasParking = conn.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {PlLayoutTable} WHERE CAST(zone_id AS NVARCHAR(50)) = @zoneId",
                new { zoneId }) > 0;

            if (hasParking)
            {
                TempData["ErrorMessage"] = "This zone has parking slots. Remove them before deleting the zone.";
                return RedirectToAction("Index");
            }

            conn.Execute($"DELETE FROM {PlZonesTable} WHERE CAST(id AS NVARCHAR(50)) = @zoneKey", new { zoneKey });
            TempData["SuccessMessage"] = "Zone deleted successfully!";
            return RedirectToAction("Index");
        }
    }
}
