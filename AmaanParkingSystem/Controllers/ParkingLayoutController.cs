using Dapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace AmaanParkingSystem.Controllers
{
    // One parking slot shown under a zone
    public class ParkingLayoutSlotVm
    {
        public string ZoneId { get; set; } = "";
        public string ParkingNo { get; set; } = "";
        public bool IsOccupied { get; set; }
    }

    // One zone row shown on the Parking Layout page
    public class ParkingLayoutZoneVm
    {
        public string Id { get; set; } = "";
        public string SiteId { get; set; } = "";
        public string ZoneId { get; set; } = "";
        public string ZoneName { get; set; } = "";
        public int Capacity { get; set; }
        public string Status { get; set; } = "";
        public string CreatedBy { get; set; } = "";
        public string CreatedOn { get; set; } = "";
        public string UpdatedBy { get; set; } = "";
        public string UpdatedOn { get; set; } = "";
        public List<ParkingLayoutSlotVm> Slots { get; set; } = new List<ParkingLayoutSlotVm>();

        public int SlotCount
        {
            get { return Slots == null ? 0 : Slots.Count; }
        }

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
        private const string PlCardsTable = "[AMAAN_PMS].[dbo].[db_tbl_11_cards_allocated]";
        private const string PlCompanyMasterTable = "[AMAAN_PMS].[dbo].[db_tbl_08A_parkers_company_master]";

        // OPERATOR role can view only; MANAGER / ADMIN can add, edit and delete
        private bool PlCanWrite()
        {
            var op = HttpContext.Session.GetString("OperatorType") ?? string.Empty;
            return !string.Equals(op, "OPERATOR", StringComparison.OrdinalIgnoreCase);
        }

        // customer_id for new parking rows: copied from existing data
        private string PlGetCustomerId(SqlConnection conn)
        {
            var sources = new[] { PlLayoutTable, PlCardsTable, PlCompanyMasterTable };
            foreach (var table in sources)
            {
                try
                {
                    var value = conn.QueryFirstOrDefault<string>($@"
                        SELECT TOP 1 CAST(customer_id AS NVARCHAR(50))
                        FROM {table}
                        WHERE customer_id IS NOT NULL AND CAST(customer_id AS NVARCHAR(50)) <> ''");
                    if (!string.IsNullOrWhiteSpace(value)) return value.Trim();
                }
                catch
                {
                    // try the next table
                }
            }
            return "";
        }

        private bool PlIsAllocated(SqlConnection conn, string parkingNo)
        {
            return conn.ExecuteScalar<int>($@"
                SELECT COUNT(*) FROM {PlCardsTable}
                WHERE LTRIM(RTRIM(CAST(parking_no AS NVARCHAR(100)))) = @parkingNo",
                new { parkingNo }) > 0;
        }

        private bool PlParkingExists(SqlConnection conn, string parkingNo)
        {
            return conn.ExecuteScalar<int>($@"
                SELECT COUNT(*) FROM {PlLayoutTable}
                WHERE LOWER(LTRIM(RTRIM(CAST(parking_no AS NVARCHAR(100))))) = LOWER(@parkingNo)",
                new { parkingNo }) > 0;
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
                    ISNULL(CAST(z.created_by AS NVARCHAR(100)), '') AS CreatedBy,
                    ISNULL(CONVERT(NVARCHAR(30), z.created_on, 120), '') AS CreatedOn,
                    ISNULL(CAST(z.updated_by AS NVARCHAR(100)), '') AS UpdatedBy,
                    ISNULL(CONVERT(NVARCHAR(30), z.updated_on, 120), '') AS UpdatedOn
                FROM {PlZonesTable} z
                ORDER BY z.zone_name ASC").ToList();

            var slots = conn.Query<ParkingLayoutSlotVm>($@"
                SELECT
                    ISNULL(CAST(p.zone_id AS NVARCHAR(50)), '') AS ZoneId,
                    LTRIM(RTRIM(CAST(p.parking_no AS NVARCHAR(100)))) AS ParkingNo,
                    CAST(CASE WHEN EXISTS (
                            SELECT 1 FROM {PlCardsTable} c
                            WHERE LTRIM(RTRIM(CAST(c.parking_no AS NVARCHAR(100)))) = LTRIM(RTRIM(CAST(p.parking_no AS NVARCHAR(100)))))
                         THEN 1 ELSE 0 END AS BIT) AS IsOccupied
                FROM {PlLayoutTable} p
                WHERE ISNULL(LTRIM(RTRIM(CAST(p.parking_no AS NVARCHAR(100)))), '') <> ''
                ORDER BY LEN(LTRIM(RTRIM(CAST(p.parking_no AS NVARCHAR(100))))), p.parking_no").ToList();

            foreach (var z in zones)
            {
                z.Slots = slots.Where(s => s.ZoneId == z.ZoneId).ToList();
            }

            ViewBag.DefaultSiteId = zones.Select(z => z.SiteId).FirstOrDefault(s => !string.IsNullOrWhiteSpace(s)) ?? "";

            return View("~/Views/ParkingLayout/Index.cshtml", zones);
        }

        // ---------------------------------------------------------------- ZONES

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

        // -------------------------------------------------------------- PARKING

        // Add (original_no empty) or edit (original_no = current parking number)
        [HttpPost]
        public IActionResult SaveParking(string zone_id, string parking_nos, string original_no)
        {
            if (!PlCanWrite())
            {
                TempData["ErrorMessage"] = "You do not have permission to change parking slots.";
                return RedirectToAction("Index");
            }

            var zoneId = (zone_id ?? "").Trim();
            var originalNo = (original_no ?? "").Trim();
            var numbers = (parking_nos ?? "")
                .Split(new[] { ',', ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(n => n.Trim())
                .Where(n => n.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (zoneId.Length == 0)
            {
                TempData["ErrorMessage"] = "Please select a zone.";
                return RedirectToAction("Index");
            }
            if (numbers.Count == 0)
            {
                TempData["ErrorMessage"] = "Parking number is required.";
                return RedirectToAction("Index");
            }
            if (numbers.Any(n => n.Length > 50))
            {
                TempData["ErrorMessage"] = "Each parking number can be up to 50 characters.";
                return RedirectToAction("Index");
            }
            if (numbers.Count > 300)
            {
                TempData["ErrorMessage"] = "You can add up to 300 parking numbers at a time.";
                return RedirectToAction("Index");
            }

            using var conn = new SqlConnection(GetDynamicConnectionString());

            var zoneExists = conn.ExecuteScalar<int>(
                $"SELECT COUNT(*) FROM {PlZonesTable} WHERE CAST(zone_id AS NVARCHAR(50)) = @zoneId",
                new { zoneId }) > 0;
            if (!zoneExists)
            {
                TempData["ErrorMessage"] = "Selected zone was not found.";
                return RedirectToAction("Index");
            }

            if (originalNo.Length > 0)
            {
                // EDIT one parking slot
                if (numbers.Count != 1)
                {
                    TempData["ErrorMessage"] = "Enter a single parking number when editing.";
                    return RedirectToAction("Index");
                }

                var newNo = numbers[0];
                var renamed = !string.Equals(newNo, originalNo, StringComparison.Ordinal);

                if (renamed)
                {
                    if (PlIsAllocated(conn, originalNo))
                    {
                        TempData["ErrorMessage"] = "This parking is allocated to a card, so it cannot be renamed.";
                        return RedirectToAction("Index");
                    }
                    var sameNameOnly = string.Equals(newNo, originalNo, StringComparison.OrdinalIgnoreCase);
                    if (!sameNameOnly && PlParkingExists(conn, newNo))
                    {
                        TempData["ErrorMessage"] = $"Parking '{newNo}' already exists.";
                        return RedirectToAction("Index");
                    }
                }

                var rows = conn.Execute($@"
                    UPDATE {PlLayoutTable} SET
                        parking_no = @newNo,
                        zone_id    = @zoneId
                    WHERE LTRIM(RTRIM(CAST(parking_no AS NVARCHAR(100)))) = @originalNo",
                    new { newNo, zoneId, originalNo });

                TempData[rows > 0 ? "SuccessMessage" : "ErrorMessage"] =
                    rows > 0 ? $"Parking '{newNo}' updated successfully!" : "Parking not found.";
                return RedirectToAction("Index");
            }

            // ADD one or more parking slots
            var customerId = PlGetCustomerId(conn);
            if (customerId.Length == 0)
            {
                TempData["ErrorMessage"] = "Customer ID could not be found from existing data, so parking cannot be added.";
                return RedirectToAction("Index");
            }

            var added = new List<string>();
            var skipped = new List<string>();
            foreach (var no in numbers)
            {
                if (PlParkingExists(conn, no))
                {
                    skipped.Add(no);
                    continue;
                }

                conn.Execute($@"
                    INSERT INTO {PlLayoutTable} (customer_id, parking_no, zone_id)
                    VALUES (@customerId, @no, @zoneId)",
                    new { customerId, no, zoneId });
                added.Add(no);
            }

            if (added.Count > 0)
            {
                var msg = added.Count == 1
                    ? $"Parking '{added[0]}' added successfully!"
                    : $"{added.Count} parking slots added successfully!";
                if (skipped.Count > 0)
                    msg += " Skipped (already exist): " + string.Join(", ", skipped) + ".";
                TempData["SuccessMessage"] = msg;
            }
            else
            {
                TempData["ErrorMessage"] = "Already exists: " + string.Join(", ", skipped) + ".";
            }

            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult DeleteParking(string parking_no)
        {
            if (!PlCanWrite())
            {
                TempData["ErrorMessage"] = "You do not have permission to change parking slots.";
                return RedirectToAction("Index");
            }

            var parkingNo = (parking_no ?? "").Trim();
            if (parkingNo.Length == 0)
            {
                TempData["ErrorMessage"] = "Parking not found.";
                return RedirectToAction("Index");
            }

            using var conn = new SqlConnection(GetDynamicConnectionString());

            if (PlIsAllocated(conn, parkingNo))
            {
                TempData["ErrorMessage"] = "This parking is allocated to a card, so it cannot be deleted.";
                return RedirectToAction("Index");
            }

            conn.Execute($@"
                DELETE FROM {PlLayoutTable}
                WHERE LTRIM(RTRIM(CAST(parking_no AS NVARCHAR(100)))) = @parkingNo",
                new { parkingNo });

            TempData["SuccessMessage"] = $"Parking '{parkingNo}' deleted successfully!";
            return RedirectToAction("Index");
        }
    }
}
