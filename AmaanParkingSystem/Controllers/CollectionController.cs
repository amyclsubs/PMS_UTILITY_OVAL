using AmaanParkingSystem.Models.Collection;
using AmaanParkingSystem.Services;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace AmaanParkingSystem.Controllers
{
    public class CollectionController : BaseController
    {
        private readonly CollectionService _collectionService;
        private readonly ILogger<CollectionController> _logger;

        public CollectionController(CollectionService collectionService, ILogger<CollectionController> logger)
        {
            _collectionService = collectionService;
            _logger = logger;
        }

        public IActionResult Daily()
        {
            return View();
        }

        public IActionResult Monthly()
        {
            return View();
        }

        // ============================================
        // EXITED CARS METHODS - ✅ UPDATED FOR DATE RANGE
        // ============================================
        [HttpPost]
        public async Task<IActionResult> GetDailyData([FromBody] CollectionFilterRequest request)
        {
            try
            {
                _logger.LogInformation(
                    $"GetDailyData called with date range: {request.StartDate} to {request.EndDate}");

                var data = await _collectionService.GetDailyCollectionData(request.StartDate, request.EndDate);
                var summary = await _collectionService.GetPaymentMethodSummary(request.StartDate, request.EndDate);
                var statistics = await _collectionService.GetDailySummaryStatistics(request.StartDate, request.EndDate);
                var cardTypeSummary = await _collectionService.GetCardTypeSummary(request.StartDate, request.EndDate);

                _logger.LogInformation(
                    $"GetDailyData: {data.Count} records, {summary.Count} summary rows");

                return Json(new
                {
                    success = true,
                    data,
                    summary,
                    statistics,
                    cardTypeSummary,
                    recordCount = data.Count
                });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetDailyData: {ex.Message}\n{ex.StackTrace}");
                return Json(new { success = false, message = ex.Message });
            }
        }
        // ============================================
        // ✅ NEW: API to get card type summary - UPDATED FOR DATE RANGE
        // ============================================
        [HttpPost]
        public async Task<IActionResult> GetCardTypeSummary([FromBody] DateRequest request)
        {
            try
            {
                var cardTypeSummary = await _collectionService.GetCardTypeSummary(
                    request.StartDate, request.EndDate);

                return Ok(new { success = true, data = cardTypeSummary });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetCardTypeSummary: {ex.Message}");
                return BadRequest(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public async Task<IActionResult> GetMonthlyData([FromBody] MonthlyFilterRequest request)
        {
            try
            {
                var data = await _collectionService.GetMonthlyCollectionData(request.Year, request.Month);
                return Json(new { success = true, data });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetMonthlyData: {ex.Message}");
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public async Task<IActionResult> ExportDailyExcel(string startDate, string endDate)
        {
            try
            {
                if (!DateTime.TryParse(startDate, out var start))
                    return BadRequest("Invalid start date format");

                if (!DateTime.TryParse(endDate, out var end))
                    return BadRequest("Invalid end date format");

                var data = await _collectionService.GetDailyCollectionData(start, end);
                var stats = await _collectionService.GetDailySummaryStatistics(start, end);

                using (var workbook = new XLWorkbook())
                {
                    workbook.Properties.Author = "APMS System";
                    workbook.Properties.Title = $"Daily Collection Report - {start:dd-MMM-yyyy} to {end:dd-MMM-yyyy}";

                    var reportSheet = workbook.Worksheets.Add("Report");

                    // ── HEADERS (26 columns) ──────────────────────────────────────
                    reportSheet.Cell(1, 1).Value = "#";
                    reportSheet.Cell(1, 2).Value = "ID";
                    reportSheet.Cell(1, 3).Value = "Transaction ID";
                    reportSheet.Cell(1, 4).Value = "Card Number";
                    reportSheet.Cell(1, 5).Value = "Card Name";
                    reportSheet.Cell(1, 6).Value = "Card Type";
                    reportSheet.Cell(1, 7).Value = "Entry Date";
                    reportSheet.Cell(1, 8).Value = "Entry Time";
                    reportSheet.Cell(1, 9).Value = "Exit Date";
                    reportSheet.Cell(1, 10).Value = "Exit Time";
                    reportSheet.Cell(1, 11).Value = "Park Time";
                    reportSheet.Cell(1, 12).Value = "Ad. Charges";
                    reportSheet.Cell(1, 13).Value = "Validation Date";
                    reportSheet.Cell(1, 14).Value = "Validation Time";
                    reportSheet.Cell(1, 15).Value = "Revalidation Date";   // ✅ NEW
                    reportSheet.Cell(1, 16).Value = "Revalidation Time";   // ✅ NEW
                    reportSheet.Cell(1, 17).Value = "Charge (KES)";
                    reportSheet.Cell(1, 18).Value = "Discount (KES)";
                    reportSheet.Cell(1, 19).Value = "Remark";
                    reportSheet.Cell(1, 20).Value = "Reason";
                    reportSheet.Cell(1, 21).Value = "Source";
                    reportSheet.Cell(1, 22).Value = "Pay Mode";
                    reportSheet.Cell(1, 23).Value = "Paid (KES)";
                    reportSheet.Cell(1, 24).Value = "FOC Amount (KES)";    // ✅ from charges
                    reportSheet.Cell(1, 25).Value = "Revalidation Charge (KES)"; // ✅ NEW
                    reportSheet.Cell(1, 26).Value = "Terminal";

                    var headerRange = reportSheet.Range(1, 1, 1, 26);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    reportSheet.SheetView.FreezeRows(1);

                    // ── DATA ROWS ────────────────────────────────────────────────
                    int row = 2;
                    int counter = 1;

                    foreach (var record in data)
                    {
                        try
                        {
                            reportSheet.Cell(row, 1).Value = counter;
                            reportSheet.Cell(row, 2).Value = record.id;
                            reportSheet.Cell(row, 3).Value = record.transaction_id ?? "";
                            reportSheet.Cell(row, 4).Value = record.cardno ?? "";
                            reportSheet.Cell(row, 5).Value = record.card_name ?? "";
                            reportSheet.Cell(row, 6).Value = record.card_type ?? "temporary";

                            // Entry date/time split
                            reportSheet.Cell(row, 7).Value = record.entry_time?.ToString("dd-MMM-yyyy").ToUpper() ?? "--";
                            reportSheet.Cell(row, 8).Value = record.entry_time?.ToString("HH:mm:ss") ?? "--";

                            // Exit date/time split
                            reportSheet.Cell(row, 9).Value = record.exit_time?.ToString("dd-MMM-yyyy").ToUpper() ?? "--";
                            reportSheet.Cell(row, 10).Value = record.exit_time?.ToString("HH:mm:ss") ?? "--";

                            // Park time
                            string parkTime = "--";
                            if (record.park_time_total.HasValue && record.park_time_total.Value > 0)
                            {
                                int hours = record.park_time_total.Value / 60;
                                int minutes = record.park_time_total.Value % 60;
                                parkTime = $"{hours:D2}hr, {minutes:D2}min, 00sec";
                            }
                            reportSheet.Cell(row, 11).Value = parkTime;

                            // Additional charges
                            reportSheet.Cell(row, 12).Value = (decimal)(record.addn_chgs ?? 0m);
                            reportSheet.Cell(row, 12).Style.NumberFormat.Format = "#,##0.00";

                            // Validation date/time split
                            reportSheet.Cell(row, 13).Value = record.validation_time?.ToString("dd-MMM-yyyy").ToUpper() ?? "--";
                            reportSheet.Cell(row, 14).Value = record.validation_time?.ToString("HH:mm:ss") ?? "--";

                            // ✅ Revalidation date/time split (from joined transactions table)
                            reportSheet.Cell(row, 15).Value = record.revalidation_time?.ToString("dd-MMM-yyyy").ToUpper() ?? "--";
                            reportSheet.Cell(row, 16).Value = record.revalidation_time?.ToString("HH:mm:ss") ?? "--";

                            // Charge (KES) from charges column
                            decimal charge = record.charges ?? 0m;
                            reportSheet.Cell(row, 17).Value = charge;
                            reportSheet.Cell(row, 17).Style.NumberFormat.Format = "#,##0.00";

                            // Discount
                            reportSheet.Cell(row, 18).Value = (decimal)(record.discount_amount ?? 0m);
                            reportSheet.Cell(row, 18).Style.NumberFormat.Format = "#,##0.00";

                            reportSheet.Cell(row, 19).Value = record.remark ?? "Hourly Rate Applied";
                            reportSheet.Cell(row, 20).Value = record.reason ?? "Normal Exit";
                            reportSheet.Cell(row, 21).Value = record.pay_source ?? "Server";
                            reportSheet.Cell(row, 22).Value = record.pay_mode ?? "Auto";

                            // Paid
                            decimal paid = record.paid_amount ?? 0m;
                            reportSheet.Cell(row, 23).Value = paid;
                            reportSheet.Cell(row, 23).Style.NumberFormat.Format = "#,##0.00";

                            // ✅ FOC Amount = charges column (only when pay_mode is FOC)
                            decimal focAmount = 0m;
                            if (!string.IsNullOrWhiteSpace(record.pay_mode))
                            {
                                string normalizedPayMode = record.pay_mode.Trim().ToUpperInvariant()
                                    .Replace(" ", "").Replace("-", "").Replace("_", "");
                                if (normalizedPayMode == "FREEOFCHARGE" || normalizedPayMode == "FOC")
                                    focAmount = record.charges ?? 0m;   // ✅ charges, not paid_amount
                            }
                            reportSheet.Cell(row, 24).Value = focAmount;
                            reportSheet.Cell(row, 24).Style.NumberFormat.Format = "#,##0.00";

                            // ✅ Revalidation charge
                            decimal revalCharge = record.revalidation_charges ?? 0m;
                            reportSheet.Cell(row, 25).Value = revalCharge;
                            reportSheet.Cell(row, 25).Style.NumberFormat.Format = "#,##0.00";

                            reportSheet.Cell(row, 26).Value = "Server_Exit1";

                            row++;
                            counter++;
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"Error processing row {row}: {ex.Message}");
                            row++;
                        }
                    }

                    // Column widths
                    reportSheet.Columns().AdjustToContents();
                    reportSheet.Column(3).Width = 20;
                    reportSheet.Column(11).Width = 18;
                    reportSheet.Column(13).Width = 18;
                    reportSheet.Column(14).Width = 15;
                    reportSheet.Column(15).Width = 18;  // Revalidation Date
                    reportSheet.Column(16).Width = 15;  // Revalidation Time
                    reportSheet.Column(19).Width = 25;

                    // ── SUMMARY STATISTICS (2 blank rows after data) ─────────────
                    int statsRow = row + 2;

                    try
                    {
                        dynamic statisticsData = stats;
                        var devices = new[] { "Server", "Paystation1", "Paystation2", "Paystation3", "Paystation4", "HHD1", "HHD2", "HHD3", "HHD4", "Web" };
                        var deviceStats = (IDictionary<string, object>)statisticsData.devices;

                        reportSheet.Cell(statsRow, 1).Value = "SUMMARY STATISTICS";
                        reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                        reportSheet.Cell(statsRow, 1).Style.Font.FontSize = 12;
                        reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                        statsRow += 2;

                        foreach (var device in devices)
                        {
                            try
                            {
                                if (!deviceStats.ContainsKey(device)) continue;
                                dynamic devData = deviceStats[device];

                                reportSheet.Cell(statsRow, 1).Value = $"{device} Cash Count:";
                                reportSheet.Cell(statsRow, 2).Value = devData.cashCount;
                                statsRow++;
                                reportSheet.Cell(statsRow, 1).Value = $"{device} Cash Total:";
                                reportSheet.Cell(statsRow, 2).Value = (decimal)devData.cashTotal;
                                reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                                statsRow++;
                                reportSheet.Cell(statsRow, 1).Value = $"{device} M-Pesa Count:";
                                reportSheet.Cell(statsRow, 2).Value = devData.mPesaCount;
                                statsRow++;
                                reportSheet.Cell(statsRow, 1).Value = $"{device} M-Pesa Total:";
                                reportSheet.Cell(statsRow, 2).Value = (decimal)devData.mPesaTotal;
                                reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                                statsRow++;
                                reportSheet.Cell(statsRow, 1).Value = $"{device} M-Pesa Via Paybill Count:";
                                reportSheet.Cell(statsRow, 2).Value = devData.mPesaViaPaybillCount;
                                statsRow++;
                                reportSheet.Cell(statsRow, 1).Value = $"{device} M-Pesa Via Paybill Total:";
                                reportSheet.Cell(statsRow, 2).Value = (decimal)devData.mPesaViaPaybillTotal;
                                reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                                statsRow++;
                                reportSheet.Cell(statsRow, 1).Value = $"{device} FOC Count:";
                                reportSheet.Cell(statsRow, 2).Value = devData.focCount;
                                reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                statsRow++;
                                reportSheet.Cell(statsRow, 1).Value = $"{device} FOC Amount:";
                                reportSheet.Cell(statsRow, 2).Value = (decimal)devData.focTotal;
                                reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                                reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                statsRow++;
                            }
                            catch (Exception devEx)
                            {
                                _logger.LogError($"Error processing device {device}: {devEx.Message}");
                                continue;
                            }
                        }

                        statsRow++;

                        // ── Grand totals (yellow) ──────────────────────────────────
                        void AddYellowRowInt(string label, int value)
                        {
                            reportSheet.Cell(statsRow, 1).Value = label;
                            reportSheet.Cell(statsRow, 2).Value = value;
                            reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                            reportSheet.Cell(statsRow, 2).Style.Font.Bold = true;
                            reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                            reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
                            statsRow++;
                        }

                        void AddYellowRowDecimal(string label, decimal value)
                        {
                            reportSheet.Cell(statsRow, 1).Value = label;
                            reportSheet.Cell(statsRow, 2).Value = value;
                            reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                            reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                            reportSheet.Cell(statsRow, 2).Style.Font.Bold = true;
                            reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                            reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
                            statsRow++;
                        }

                        AddYellowRowInt("- Free of Charge Count:", (int)statisticsData.totalFocCount);
                        AddYellowRowInt("- Equity Count:", (int)statisticsData.equityCount);
                        AddYellowRowInt("- Credit Taken Count:", (int)statisticsData.creditTakenCount);
                        AddYellowRowInt("- Advance Paid Count:", (int)statisticsData.advancePaidCount);
                        AddYellowRowInt("- Repark Count:", 0);
                        statsRow++;

                        AddYellowRowInt("GRAND TOTAL Count:", (int)statisticsData.grandTotal);
                        AddYellowRowDecimal("TOTAL CASH Collected:", (decimal)statisticsData.totalCashCollected);
                        AddYellowRowDecimal("TOTAL M-Pesa Collected:",
                            (decimal)statisticsData.totalMPesaCollected + (decimal)statisticsData.totalMPesaViaPaybillCollected);
                        AddYellowRowDecimal("TOTAL FOC Amount (from Charges):", (decimal)statisticsData.totalFocAmount);
                        AddYellowRowDecimal("TOTAL Amount Collected:", (decimal)statisticsData.grandTotalAmount);
                        AddYellowRowDecimal("Total Amount Charged:", (decimal)statisticsData.totalCharges);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error processing statistics: {ex.Message}");
                        reportSheet.Cell(statsRow, 1).Value = $"Error: {ex.Message}";
                    }

                    reportSheet.Column(1).Width = 40;
                    reportSheet.Column(2).Width = 20;

                    using var stream = new MemoryStream();
                    workbook.SaveAs(stream);
                    stream.Position = 0;
                    var fileName = $"DAY_REPORT_{start:dd-MMM-yyyy}_to_{end:dd-MMM-yyyy}_{DateTime.Now:HH-mm-ss}.xlsx";
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ExportDailyExcel: {ex.Message}\n{ex.StackTrace}");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        private static string NormalizePayModeLocal(string payMode)
        {
            if (string.IsNullOrWhiteSpace(payMode)) return "Cash";
            string n = payMode.Trim().ToUpperInvariant().Replace(" ", "").Replace("-", "").Replace("_", "");
            if (n == "CASH") return "Cash";
            if (n == "MPESA") return "M-Pesa";
            if (n == "MPESAKIAPAYBILL" || n == "MPESAVIAPAYBILL") return "M-Pesa Via Paybill";
            if (n == "FOC" || n == "FREEOFCHARGE") return "Free of Charge";
            if (n == "AUTO" || n == "AUTOMATIC") return "Auto";
            return payMode.Trim();
        }

        [HttpGet]
        public async Task<IActionResult> ExportMonthlyExcel(int month, int year)
        {
            try
            {
                // ============================================================
                // LOAD MONTHLY DATA
                // ============================================================
                var monthlyData = await _collectionService.GetMonthlyCollectionData(year, month);
                if (monthlyData == null || monthlyData.Count == 0)
                    return NotFound(new { success = false, message = "No data found for the selected month" });

                using var workbook = new XLWorkbook();

                var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
                var worksheet = workbook.Worksheets.Add($"{monthName} {year}");
                int currentRow = 1;

                // ============================================================
                // STATIC CONFIG FOR SUMMARY SHEET (UNCHANGED)
                // ============================================================
                var devices = new[]
                {
            "Server", "Paystation1", "Paystation2", "Paystation3", "Paystation4",
            "HHD1", "HHD2", "HHD3", "HHD4", "Web"
        };

                var deviceColors = new Dictionary<string, XLColor>
        {
            {"Server", XLColor.FromArgb(255, 242, 204)},
            {"Paystation1", XLColor.FromArgb(252, 213, 180)},
            {"Paystation2", XLColor.FromArgb(198, 224, 180)},
            {"Paystation3", XLColor.FromArgb(180, 199, 231)},
            {"Paystation4", XLColor.FromArgb(217, 210, 233)},
            {"HHD1", XLColor.FromArgb(234, 209, 220)},
            {"HHD2", XLColor.FromArgb(207, 226, 243)},
            {"HHD3", XLColor.FromArgb(255, 230, 153)},
            {"HHD4", XLColor.FromArgb(182, 215, 168)},
            {"Web", XLColor.FromArgb(213, 166, 189)}
        };

                // ============================================================
                // SORT MONTHLY DATA
                // ============================================================
                var sortedItems = monthlyData
                    .Cast<dynamic>()
                    .Select(item => new
                    {
                        Item = item,
                        Date = DateTime.Parse(item.date.ToString())
                    })
                    .OrderBy(x => x.Date)
                    .ToList();

                var sortedDates = sortedItems.Select(x => x.Date).ToList();
                var rowsByDateKey = sortedItems.ToDictionary(x => x.Date.ToString("yyyy-MM-dd"), x => (dynamic)x.Item);

                // ============================================================
                // TAB 1 - SPLIT DATA WRITING FOR MAIN REPORT
                // Left table  : columns 1 - 19
                // Gap         : column 20
                // Right table : columns 21 - 33
                // FIXED: single title row to avoid double-header visual issue
                // ============================================================

                worksheet.Cell(currentRow, 1).Value = $"MONTHLY EXITED CARS SUMMARY - {monthName.ToUpper()} {year}";
                worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
                worksheet.Cell(currentRow, 1).Style.Font.FontSize = 16;
                worksheet.Cell(currentRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Cell(currentRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(currentRow, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range(currentRow, 1, currentRow, 33).Merge();

                worksheet.Column(20).Width = 3;
                currentRow++;

                int topHeaderRow = currentRow;
                int subHeaderRowTab1 = currentRow + 1;
                int dataStartRow = currentRow + 2;

                // ------------------------------------------------------------
                // LEFT TABLE HEADERS (1 - 19)
                // ------------------------------------------------------------
                var leftHeaders = new Dictionary<int, string>
        {
            { 1, "#" },
            { 2, "Date" },
            { 3, "Day" },
            { 4, "Total Vehicles" },
            { 5, "Cash Count" },
            { 6, "Cash Total (KES)" },
            { 7, "M-Pesa Count" },
            { 8, "M-Pesa Total (KES)" },
            { 9, "FOC Count" },
            { 10, "FOC Amount (KES)" },
            { 11, "HHD-4 FOC Count" },
            { 12, "HHD-4 FOC Amount (KES)" },
            { 13, "Total Charges (KES)" },
            { 14, "Total Paid (KES)" },
            { 15, "Discount (KES)" },
            { 16, "Zero Charge Count" },
            { 17, "Permanent Parkers" },
            { 18, "Revalidation Count" },
            { 19, "Revalidation Amt (KES)" }
        };

                foreach (var h in leftHeaders)
                {
                    worksheet.Cell(topHeaderRow, h.Key).Value = h.Value;
                    worksheet.Range(topHeaderRow, h.Key, subHeaderRowTab1, h.Key).Merge();

                    var cell = worksheet.Cell(topHeaderRow, h.Key);
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.Gray;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Alignment.WrapText = true;
                    worksheet.Column(h.Key).Width = 12;
                }

                // ------------------------------------------------------------
                // RIGHT TABLE GROUP HEADERS (21 - 33)
                // ------------------------------------------------------------
                worksheet.Cell(topHeaderRow, 21).Value = "TEMPORARY";
                worksheet.Range(topHeaderRow, 21, topHeaderRow, 23).Merge();

                worksheet.Cell(topHeaderRow, 24).Value = "FOC_PER_UNIT";
                worksheet.Range(topHeaderRow, 24, topHeaderRow, 26).Merge();

                worksheet.Cell(topHeaderRow, 27).Value = "DAILY ACCESS";
                worksheet.Range(topHeaderRow, 27, topHeaderRow, 29).Merge();

                worksheet.Cell(topHeaderRow, 30).Value = "CASH CARD";
                worksheet.Range(topHeaderRow, 30, topHeaderRow, 32).Merge();

                worksheet.Cell(topHeaderRow, 33).Value = "VIP";
                worksheet.Range(topHeaderRow, 33, subHeaderRowTab1, 33).Merge();

                foreach (var col in new[] { 21, 24, 27, 30, 33 })
                {
                    var cell = worksheet.Cell(topHeaderRow, col);
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.Gray;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Alignment.WrapText = true;
                }

                var rightSubHeaders = new Dictionary<int, string>
        {
            { 21, "Zero Charge" },
            { 22, "Charged" },
            { 23, "Paid Amount (KES)" },
            { 24, "Zero Charge" },
            { 25, "Charged" },
            { 26, "Paid Amount (KES)" },
            { 27, "Zero Charge" },
            { 28, "Charged" },
            { 29, "Paid Amount (KES)" },
            { 30, "Zero Charge" },
            { 31, "Charged" },
            { 32, "Paid Amount (KES)" }
        };

                foreach (var h in rightSubHeaders)
                {
                    var cell = worksheet.Cell(subHeaderRowTab1, h.Key);
                    cell.Value = h.Value;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Alignment.WrapText = true;
                    worksheet.Column(h.Key).Width = 11;
                }

                worksheet.Row(topHeaderRow).Height = 24;
                worksheet.Row(subHeaderRowTab1).Height = 32;
                worksheet.SheetView.Freeze(0, 0);
                // ------------------------------------------------------------
                // WRITE DATA ROWS
                // ------------------------------------------------------------
                int rowNumber = 1;
                currentRow = dataStartRow;

                foreach (var entry in sortedItems)
                {
                    dynamic item = entry.Item;
                    DateTime date = entry.Date;

                    // LEFT TABLE
                    worksheet.Cell(currentRow, 1).Value = rowNumber++;
                    worksheet.Cell(currentRow, 2).Value = date.ToString("dd-MMM-yyyy").ToUpper();
                    worksheet.Cell(currentRow, 3).Value = date.ToString("dddd");
                    worksheet.Cell(currentRow, 4).Value = (int)item.totalVehicles;
                    worksheet.Cell(currentRow, 5).Value = (int)item.cashCount;
                    worksheet.Cell(currentRow, 6).Value = (decimal)item.cashTotal;
                    worksheet.Cell(currentRow, 7).Value = (int)item.mPesaCount;
                    worksheet.Cell(currentRow, 8).Value = (decimal)item.mPesaTotal;
                    worksheet.Cell(currentRow, 9).Value = (int)item.focCount;
                    worksheet.Cell(currentRow, 10).Value = (decimal)item.focAmount;
                    worksheet.Cell(currentRow, 11).Value = (int)item.ctmFocCount;
                    worksheet.Cell(currentRow, 12).Value = (decimal)item.ctmFocAmount;
                    worksheet.Cell(currentRow, 13).Value = (decimal)item.totalCharges;
                    worksheet.Cell(currentRow, 14).Value = (decimal)item.totalPaid;
                    worksheet.Cell(currentRow, 15).Value = (decimal)item.discount;
                    worksheet.Cell(currentRow, 16).Value = (int)item.zeroChargeCount;
                    worksheet.Cell(currentRow, 17).Value = (int)item.permanentParkersCount;
                    worksheet.Cell(currentRow, 18).Value = (int)item.revalidationCount;
                    worksheet.Cell(currentRow, 19).Value = (decimal)item.revalidationAmount;

                    // RIGHT TABLE
                    worksheet.Cell(currentRow, 21).Value = (int)item.tempZeroChargeCount;
                    worksheet.Cell(currentRow, 22).Value = (int)item.tempChargedCount;
                    worksheet.Cell(currentRow, 23).Value = (decimal)item.tempPaidAmount;

                    worksheet.Cell(currentRow, 24).Value = (int)item.focPerUnitZeroCount;
                    worksheet.Cell(currentRow, 25).Value = (int)item.focPerUnitChargedCount;
                    worksheet.Cell(currentRow, 26).Value = (decimal)item.focPerUnitPaidAmount;

                    worksheet.Cell(currentRow, 27).Value = (int)item.dailyAccessZeroCount;
                    worksheet.Cell(currentRow, 28).Value = (int)item.dailyAccessChargedCount;
                    worksheet.Cell(currentRow, 29).Value = (decimal)item.dailyAccessPaidAmount;

                    worksheet.Cell(currentRow, 30).Value = (int)item.cashCardZeroCount;
                    worksheet.Cell(currentRow, 31).Value = (int)item.cashCardChargedCount;
                    worksheet.Cell(currentRow, 32).Value = (decimal)item.cashCardPaidAmount;

                    worksheet.Cell(currentRow, 33).Value = (int)item.vipCount;

                    currentRow++;
                }

                int dataEndRow = currentRow - 1;

                // ------------------------------------------------------------
                // NUMBER FORMAT + COLUMN COLORS
                // ------------------------------------------------------------
                var amountColumns = new[] { 6, 8, 10, 12, 13, 15, 19, 23, 26, 29, 32 };
                foreach (var col in amountColumns)
                {
                    if (dataEndRow >= dataStartRow)
                    {
                        var range = worksheet.Range(dataStartRow, col, dataEndRow, col);
                        range.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFEBD6");
                        range.Style.NumberFormat.Format = "#,##0.00";
                        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    }
                }

                if (dataEndRow >= dataStartRow)
                {
                    var totalPaidRange = worksheet.Range(dataStartRow, 14, dataEndRow, 14);
                    totalPaidRange.Style.Fill.BackgroundColor = XLColor.FromArgb(146, 208, 80);
                    totalPaidRange.Style.NumberFormat.Format = "#,##0.00";
                    totalPaidRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }

                // ------------------------------------------------------------
                // TOTAL ROW
                // ------------------------------------------------------------
                worksheet.Cell(currentRow, 1).Value = "TOTAL";

                for (int c = 4; c <= 19; c++)
                {
                    string colLetter = worksheet.Cell(dataStartRow, c).Address.ColumnLetter;
                    worksheet.Cell(currentRow, c).FormulaA1 = $"=SUM({colLetter}{dataStartRow}:{colLetter}{dataEndRow})";

                    if (amountColumns.Contains(c) || c == 14)
                        worksheet.Cell(currentRow, c).Style.NumberFormat.Format = "#,##0.00";
                }

                for (int c = 21; c <= 33; c++)
                {
                    string colLetter = worksheet.Cell(dataStartRow, c).Address.ColumnLetter;
                    worksheet.Cell(currentRow, c).FormulaA1 = $"=SUM({colLetter}{dataStartRow}:{colLetter}{dataEndRow})";

                    if (amountColumns.Contains(c))
                        worksheet.Cell(currentRow, c).Style.NumberFormat.Format = "#,##0.00";
                }

                var totalRowLeft = worksheet.Range(currentRow, 1, currentRow, 19);
                totalRowLeft.Style.Font.Bold = true;
                totalRowLeft.Style.Fill.BackgroundColor = XLColor.Gold;
                totalRowLeft.Style.Border.TopBorder = XLBorderStyleValues.Thick;

                var totalRowRight = worksheet.Range(currentRow, 21, currentRow, 33);
                totalRowRight.Style.Font.Bold = true;
                totalRowRight.Style.Fill.BackgroundColor = XLColor.Gold;
                totalRowRight.Style.Border.TopBorder = XLBorderStyleValues.Thick;

                worksheet.Cell(currentRow, 14).Style.Fill.BackgroundColor = XLColor.FromArgb(146, 208, 80);

                // ------------------------------------------------------------
                // BORDERS + GAP
                // ------------------------------------------------------------
                int lastDataRow = currentRow;

                var leftTableRange = worksheet.Range(1, 1, lastDataRow, 19);
                leftTableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                leftTableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                var gapRange = worksheet.Range(1, 20, lastDataRow, 20);
                gapRange.Style.Fill.BackgroundColor = XLColor.NoColor;
                gapRange.Style.Border.LeftBorder = XLBorderStyleValues.None;
                gapRange.Style.Border.RightBorder = XLBorderStyleValues.None;
                gapRange.Style.Border.TopBorder = XLBorderStyleValues.None;
                gapRange.Style.Border.BottomBorder = XLBorderStyleValues.None;
                gapRange.Style.Border.InsideBorder = XLBorderStyleValues.None;

                var rightTableRange = worksheet.Range(1, 21, lastDataRow, 33);
                rightTableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                rightTableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                currentRow += 2;

                // ============================================================
                // BUILD DAILY STATS FOR SUMMARY SHEET
                // UNCHANGED
                // ============================================================
                var dailyStats = new Dictionary<string, dynamic>();
                foreach (var date in sortedDates)
                {
                    string dateKey = date.ToString("yyyy-MM-dd");
                    dailyStats[dateKey] = await _collectionService.GetDailySummaryStatistics(date, date);
                }

                // ============================================================
                // SUMMARY SHEET DEVICE METRIC DEFINITIONS
                // UNCHANGED
                // ============================================================
                var allMetricsByDevice = new Dictionary<string, List<(string Label, string Key, bool IsCurrency)>>();
                foreach (var device in devices)
                {
                    if (device.StartsWith("Paystation"))
                    {
                        allMetricsByDevice[device] = new List<(string, string, bool)>
                {
                    ("Cash Count", "cashCount", false),
                    ("Cash Total", "cashTotal", true),
                    ("M-Pesa Count", "mPesaCount", false),
                    ("M-Pesa Total", "mPesaTotal", true)
                };
                    }
                    else if (device == "Web")
                    {
                        allMetricsByDevice[device] = new List<(string, string, bool)>
                {
                    ("M-Pesa Count", "mPesaCount", false),
                    ("M-Pesa Total", "mPesaTotal", true)
                };
                    }
                    else
                    {
                        allMetricsByDevice[device] = new List<(string, string, bool)>
                {
                    ("Cash Count", "cashCount", false),
                    ("Cash Total", "cashTotal", true),
                    ("M-Pesa Count", "mPesaCount", false),
                    ("M-Pesa Total", "mPesaTotal", true),
                    ("M-Pesa Via Paybill Count", "mPesaViaPaybillCount", false),
                    ("M-Pesa Via Paybill Total", "mPesaViaPaybillTotal", true),
                    ("FOC Count", "focCount", false),
                    ("FOC Amount", "focTotal", true)
                };
                    }
                }

                // ============================================================
                // SUMMARY SHEET DEVICE VALUES
                // UNCHANGED
                // ============================================================
                var values = new Dictionary<string, Dictionary<string, List<decimal>>>();
                foreach (var device in devices)
                {
                    values[device] = new Dictionary<string, List<decimal>>();
                    foreach (var m in allMetricsByDevice[device])
                        values[device][m.Key] = new List<decimal>();

                    foreach (var date in sortedDates)
                    {
                        string dateKey = date.ToString("yyyy-MM-dd");
                        dynamic stats = dailyStats.ContainsKey(dateKey) ? dailyStats[dateKey] : null;
                        object devicesData = stats != null ? GetPropertyValue(stats, "devices") : null;
                        object deviceData = devicesData != null ? GetPropertyValue(devicesData, device) : null;

                        foreach (var m in allMetricsByDevice[device])
                        {
                            decimal v = 0;
                            if (deviceData != null)
                                v = Convert.ToDecimal(GetPropertyValue(deviceData, m.Key));
                            values[device][m.Key].Add(v);
                        }
                    }
                }

                // ============================================================
                // HIDE ZERO-ONLY DEVICE METRICS / DEVICES ON SUMMARY SHEET
                // UNCHANGED
                // ============================================================
                var keptMetricsByDevice = new Dictionary<string, List<(string Label, string Key, bool IsCurrency)>>();
                foreach (var device in devices)
                {
                    var kept = allMetricsByDevice[device]
                        .Where(m => values[device][m.Key].Any(v => v != 0))
                        .ToList();

                    if (kept.Count > 0)
                        keptMetricsByDevice[device] = kept;
                }

                var keptDevices = devices.Where(d => keptMetricsByDevice.ContainsKey(d)).ToList();

                // ============================================================
                // SUMMARY SHEET
                // UNCHANGED
                // ============================================================
                var summarySheet = workbook.Worksheets.Add("Summary Statistics");
                int sRow = 1;

                summarySheet.Cell(sRow, 1).Value = "SUMMARY STATISTICS (EXITED CARS)";
                summarySheet.Cell(sRow, 1).Style.Font.Bold = true;
                summarySheet.Cell(sRow, 1).Style.Font.FontSize = 14;
                summarySheet.Cell(sRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                sRow += 2;

                int deviceHeaderRow = sRow;
                int subHeaderRow = sRow + 1;
                int dStart = sRow + 2;

                summarySheet.Cell(subHeaderRow, 1).Value = "DATE";
                summarySheet.Cell(subHeaderRow, 1).Style.Font.Bold = true;
                summarySheet.Cell(subHeaderRow, 1).Style.Alignment.WrapText = true;
                summarySheet.Cell(subHeaderRow, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                summarySheet.Range(deviceHeaderRow, 1, subHeaderRow, 1).Merge();
                summarySheet.Row(subHeaderRow).Height = 32;
                summarySheet.Column(1).Width = 13;

                var deviceColumnMap = new Dictionary<string, (int start, int end)>();
                int currentCol = 2;

                if (keptDevices.Count == 0)
                {
                    summarySheet.Cell(dStart, 1).Value = "No non-zero collection data found for this period.";
                }
                else
                {
                    foreach (var device in keptDevices)
                    {
                        var metrics = keptMetricsByDevice[device];
                        int startCol = currentCol;

                        foreach (var m in metrics)
                        {
                            var cell = summarySheet.Cell(subHeaderRow, currentCol);
                            cell.Value = m.Label.ToUpper();
                            cell.Style.Font.Bold = true;
                            cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            cell.Style.Alignment.WrapText = true;
                            summarySheet.Column(currentCol).Width = 11;
                            currentCol++;
                        }

                        int endCol = currentCol - 1;

                        summarySheet.Cell(deviceHeaderRow, startCol).Value = device.ToUpper();
                        summarySheet.Range(deviceHeaderRow, startCol, deviceHeaderRow, endCol).Merge();
                        summarySheet.Cell(deviceHeaderRow, startCol).Style.Font.Bold = true;
                        summarySheet.Cell(deviceHeaderRow, startCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        summarySheet.Range(deviceHeaderRow, startCol, deviceHeaderRow, endCol).Style.Fill.BackgroundColor =
                            deviceColors.ContainsKey(device) ? deviceColors[device] : XLColor.LightBlue;

                        deviceColumnMap[device] = (startCol, endCol);
                    }

                    int lastCol = currentCol - 1;
                    int rIdx = dStart;

                    for (int di = 0; di < sortedDates.Count; di++)
                    {
                        var date = sortedDates[di];
                        summarySheet.Cell(rIdx, 1).Value = date.ToString("dd-MMM-yyyy").ToUpper();
                        summarySheet.Cell(rIdx, 1).Style.Font.Bold = true;

                        foreach (var device in keptDevices)
                        {
                            var (startCol, _) = deviceColumnMap[device];
                            int col = startCol;

                            foreach (var m in keptMetricsByDevice[device])
                            {
                                decimal value = values[device][m.Key][di];
                                summarySheet.Cell(rIdx, col).Value = value;
                                if (m.IsCurrency)
                                    summarySheet.Cell(rIdx, col).Style.NumberFormat.Format = "#,##0.00";
                                col++;
                            }
                        }

                        rIdx++;
                    }

                    summarySheet.Cell(rIdx, 1).Value = "MONTHLY TOTAL";
                    summarySheet.Cell(rIdx, 1).Style.Font.Bold = true;
                    summarySheet.Cell(rIdx, 1).Style.Fill.BackgroundColor = XLColor.Gold;

                    for (int c = 2; c <= lastCol; c++)
                    {
                        string colLetter = summarySheet.Cell(dStart, c).Address.ColumnLetter;
                        summarySheet.Cell(rIdx, c).FormulaA1 = $"=SUM({colLetter}{dStart}:{colLetter}{rIdx - 1})";
                        summarySheet.Cell(rIdx, c).Style.Font.Bold = true;
                        summarySheet.Cell(rIdx, c).Style.Fill.BackgroundColor = XLColor.Gold;
                    }

                    var deviceTableRange = summarySheet.Range(deviceHeaderRow, 1, rIdx, lastCol);
                    deviceTableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                    deviceTableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    foreach (var device in keptDevices)
                    {
                        var (startCol, endCol) = deviceColumnMap[device];
                        var deviceRange = summarySheet.Range(deviceHeaderRow, startCol, rIdx, endCol);
                        deviceRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                    }

                    // ============================================================
                    // GRAND TOTALS SECTION ON SUMMARY SHEET
                    // UNCHANGED
                    // ============================================================
                    int grandRow = rIdx + 3;

                    summarySheet.Cell(grandRow, 1).Value = "GRAND TOTALS";
                    summarySheet.Cell(grandRow, 1).Style.Font.Bold = true;
                    summarySheet.Cell(grandRow, 1).Style.Font.FontSize = 14;
                    summarySheet.Cell(grandRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                    grandRow += 2;

                    var gtCols = new List<(string Header, bool AlwaysKeep, bool IsCurrency, Func<dynamic, object> ValueSelector)>
            {
                ("DATE", true, false, item => DateTime.Parse(item.date.ToString()).ToString("dd-MMM-yyyy").ToUpper()),
                ("TOTAL VEHICLES", false, false, item => (int)item.totalVehicles),
                ("CASH COUNT", false, false, item => (int)item.cashCount),
                ("CASH COLLECTED (KES)", false, true, item => (decimal)item.cashTotal),
                ("M-PESA COUNT", false, false, item => (int)item.mPesaCount),
                ("M-PESA COLLECTED (KES)", false, true, item => (decimal)item.mPesaTotal),
                ("FOC COUNT", false, false, item => (int)item.focCount),
                ("FOC AMOUNT (KES)", false, true, item => (decimal)item.focAmount),
                ("HHD-4 FOC COUNT", false, false, item => (int)item.ctmFocCount),
                ("HHD-4 FOC AMOUNT (KES)", false, true, item => (decimal)item.ctmFocAmount),
                ("ZERO CHARGE COUNT", false, false, item => (int)item.zeroChargeCount),
                ("PERMANENT PARKERS", false, false, item => (int)item.permanentParkersCount),
                ("REVALIDATION COUNT", false, false, item => (int)item.revalidationCount),
                ("REVALIDATION AMT (KES)", false, true, item => (decimal)item.revalidationAmount),
                ("DISCOUNT (KES)", false, true, item => (decimal)item.discount),
                ("TOTAL PAID (KES)", false, true, item => (decimal)item.totalPaid),
                ("TOTAL CHARGES (KES)", false, true, item => (decimal)item.totalCharges),
                ("VIP", false, false, item => (int)item.vipCount)
            };

                    var gtRows = new List<object[]>();
                    foreach (var entry in sortedItems)
                    {
                        dynamic item = entry.Item;
                        var row = new object[gtCols.Count];
                        for (int i = 0; i < gtCols.Count; i++)
                            row[i] = gtCols[i].ValueSelector(item);

                        gtRows.Add(row);
                    }

                    var keepGt = new List<int>();
                    for (int c = 0; c < gtCols.Count; c++)
                    {
                        if (gtCols[c].AlwaysKeep)
                        {
                            keepGt.Add(c);
                            continue;
                        }

                        bool hasNonZero = false;
                        foreach (var row in gtRows)
                        {
                            var value = row[c];
                            if (value == null) continue;
                            if (value is string) continue;

                            if (Convert.ToDecimal(value) != 0)
                            {
                                hasNonZero = true;
                                break;
                            }
                        }

                        if (hasNonZero)
                            keepGt.Add(c);
                    }

                    int gtHeaderRow = grandRow;
                    for (int pos = 0; pos < keepGt.Count; pos++)
                    {
                        var col = gtCols[keepGt[pos]];
                        var cell = summarySheet.Cell(gtHeaderRow, pos + 1);
                        cell.Value = col.Header;
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.Gray;
                        cell.Style.Font.FontColor = XLColor.White;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        cell.Style.Alignment.WrapText = true;
                        summarySheet.Column(pos + 1).Width = 13;
                    }

                    summarySheet.Row(gtHeaderRow).Height = 32;

                    int gtDataStart = gtHeaderRow + 1;
                    for (int r = 0; r < gtRows.Count; r++)
                    {
                        int rr = gtDataStart + r;
                        for (int pos = 0; pos < keepGt.Count; pos++)
                        {
                            int origCol = keepGt[pos];
                            var cell = summarySheet.Cell(rr, pos + 1);
                            cell.Value = XLCellValue.FromObject(gtRows[r][origCol]);

                            if (gtCols[origCol].IsCurrency)
                                cell.Style.NumberFormat.Format = "#,##0.00";
                        }
                    }

                    int gtDataEnd = gtDataStart + gtRows.Count - 1;

                    int totalAmtPos = keepGt.FindIndex(c => gtCols[c].Header == "TOTAL PAID (KES)");
                    if (totalAmtPos >= 0)
                    {
                        summarySheet.Range(gtDataStart, totalAmtPos + 1, gtDataEnd, totalAmtPos + 1)
                            .Style.Fill.BackgroundColor = XLColor.FromArgb(146, 208, 80);
                    }

                    int gtTotalRow = gtDataEnd + 1;
                    for (int pos = 0; pos < keepGt.Count; pos++)
                    {
                        int origCol = keepGt[pos];
                        var cell = summarySheet.Cell(gtTotalRow, pos + 1);

                        if (gtCols[origCol].Header == "DATE")
                        {
                            cell.Value = "TOTAL";
                        }
                        else
                        {
                            string colLetter = summarySheet.Cell(gtDataStart, pos + 1).Address.ColumnLetter;
                            cell.FormulaA1 = $"=SUM({colLetter}{gtDataStart}:{colLetter}{gtDataEnd})";
                            if (gtCols[origCol].IsCurrency)
                                cell.Style.NumberFormat.Format = "#,##0.00";
                        }

                        cell.Style.Font.Bold = true;
                    }

                    if (totalAmtPos >= 0)
                        summarySheet.Cell(gtTotalRow, totalAmtPos + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(146, 208, 80);

                    var gtRange = summarySheet.Range(gtHeaderRow, 1, gtTotalRow, keepGt.Count);
                    gtRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                    gtRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    summarySheet.Range(gtTotalRow, 1, gtTotalRow, keepGt.Count).Style.Fill.BackgroundColor = XLColor.Gold;

                    summarySheet.SheetView.FreezeRows(subHeaderRow);
                    summarySheet.SheetView.FreezeColumns(1);
                }

                // ============================================================
                // ADDITIONAL SUMMARY SECTION ON SUMMARY SHEET
                // UNCHANGED
                // ============================================================
                int extraRow = summarySheet.LastRowUsed()?.RowNumber() + 3 ?? 1;

                var summaryRows = new List<(string Label, bool IsCurrency, XLColor Fill, Func<dynamic, decimal> Selector)>
        {
            ("FOC Count:", false, XLColor.LightCyan, item => Convert.ToDecimal(item.focCount)),
            ("FOC Amount:", true, XLColor.LightCyan, item => Convert.ToDecimal(item.focAmount)),
            ("HHD-4 FOC Count:", false, XLColor.FromArgb(255, 230, 153), item => Convert.ToDecimal(item.ctmFocCount)),
            ("HHD-4 FOC Amount:", true, XLColor.FromArgb(255, 230, 153), item => Convert.ToDecimal(item.ctmFocAmount)),
            ("Zero Charge Count:", false, XLColor.LightYellow, item => Convert.ToDecimal(item.zeroChargeCount)),
            ("Permanent Parkers:", false, XLColor.LightYellow, item => Convert.ToDecimal(item.permanentParkersCount)),
            ("Total Number of Vehicles:", false, XLColor.LightSteelBlue, item => Convert.ToDecimal(item.totalVehicles)),
            ("TOTAL CASH Collected:", true, XLColor.LightGreen, item => Convert.ToDecimal(item.cashTotal)),
            ("TOTAL M-Pesa Collected:", true, XLColor.LightCoral, item => Convert.ToDecimal(item.mPesaTotal)),
            ("TOTAL Amount Collected:", true, XLColor.Gold, item => Convert.ToDecimal(item.totalPaid))
        };

                summarySheet.Cell(extraRow, 1).Value = "ADDITIONAL SUMMARY";
                summarySheet.Cell(extraRow, 1).Style.Font.Bold = true;
                summarySheet.Cell(extraRow, 1).Style.Font.FontSize = 14;
                summarySheet.Cell(extraRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                extraRow += 2;

                int addHeaderRow = extraRow;
                summarySheet.Cell(addHeaderRow, 1).Value = "METRIC";
                summarySheet.Cell(addHeaderRow, 2).Value = "MONTHLY TOTAL";
                summarySheet.Cell(addHeaderRow, 1).Style.Font.Bold = true;
                summarySheet.Cell(addHeaderRow, 2).Style.Font.Bold = true;
                summarySheet.Cell(addHeaderRow, 1).Style.Fill.BackgroundColor = XLColor.Gold;
                summarySheet.Cell(addHeaderRow, 2).Style.Fill.BackgroundColor = XLColor.Gold;

                int addCol = 3;
                foreach (var date in sortedDates)
                {
                    summarySheet.Cell(addHeaderRow, addCol).Value = date.ToString("dd-MMM-yyyy").ToUpper();
                    summarySheet.Cell(addHeaderRow, addCol).Style.Font.Bold = true;
                    summarySheet.Cell(addHeaderRow, addCol).Style.Fill.BackgroundColor = XLColor.LightGray;
                    summarySheet.Cell(addHeaderRow, addCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    summarySheet.Column(addCol).Width = 13;
                    addCol++;
                }

                extraRow++;

                foreach (var metric in summaryRows)
                {
                    summarySheet.Cell(extraRow, 1).Value = metric.Label;
                    summarySheet.Cell(extraRow, 1).Style.Font.Bold = true;
                    summarySheet.Cell(extraRow, 1).Style.Fill.BackgroundColor = metric.Fill;

                    decimal monthlyTotal = 0;
                    foreach (var entry in sortedItems)
                        monthlyTotal += metric.Selector(entry.Item);

                    summarySheet.Cell(extraRow, 2).Value = monthlyTotal;
                    summarySheet.Cell(extraRow, 2).Style.Font.Bold = true;
                    summarySheet.Cell(extraRow, 2).Style.Fill.BackgroundColor = metric.Fill;
                    if (metric.IsCurrency)
                        summarySheet.Cell(extraRow, 2).Style.NumberFormat.Format = "#,##0.00";

                    int dc = 3;
                    foreach (var date in sortedDates)
                    {
                        string dateKey = date.ToString("yyyy-MM-dd");
                        dynamic item = rowsByDateKey[dateKey];
                        decimal value = metric.Selector(item);

                        summarySheet.Cell(extraRow, dc).Value = value;
                        summarySheet.Cell(extraRow, dc).Style.Fill.BackgroundColor = metric.Fill;
                        if (metric.IsCurrency)
                            summarySheet.Cell(extraRow, dc).Style.NumberFormat.Format = "#,##0.00";
                        dc++;
                    }

                    extraRow++;
                }

                var addRange = summarySheet.Range(addHeaderRow, 1, extraRow - 1, 2 + sortedDates.Count);
                addRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                addRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                summarySheet.Column(1).Width = 28;
                summarySheet.Column(2).Width = 16;

                // ============================================================
                // SAVE FILE
                // ============================================================
                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;

                var fileName = $"MONTHLY_COLLECTION_{monthName}_{year}_{DateTime.Now:HH-mm-ss}.xlsx";
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ExportMonthlyExcel: {ex.Message}\n{ex.StackTrace}");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        // ============================================
        // MONTHLY EXITED CARS DATA (TAB 1)
        // ============================================

        [HttpGet]
        public async Task<IActionResult> GetMonthlyExitedCarsData(int year, int month)
        {
            try
            {
                var data = await _collectionService.GetMonthlyCollectionData(year, month);

                if (data == null || data.Count == 0)
                    return Json(new { success = false, message = "No data found for the selected month." });

                return Json(new { success = true, data });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetMonthlyExitedCarsData: {ex.Message}");
                return Json(new { success = false, message = ex.Message });
            }
        }


        // ============================================
        // MONTHLY TRANSACTIONS DATA (TAB 2)
        // ============================================

        [HttpGet]
        public async Task<IActionResult> GetMonthlyTransactionsData(int year, int month)
        {
            try
            {
                var data = await _collectionService.GetMonthlyTransactionsDataWithVehicles(year, month);

                if (data == null || !data.Any())
                {
                    return Ok(new { success = false, message = "No data found for the selected month" });
                }

                return Ok(new { success = true, data = data });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetMonthlyTransactionsDataWithVehicles: {ex.Message}");
                return BadRequest(new { success = false, message = ex.Message });
            }
        }

        // ============================================
        // EXPORT MONTHLY TRANSACTIONS EXCEL (TAB 2)
        // UPDATED:
        //   1) Headers wrap onto 2 lines with narrow fixed column widths (matches screenshot style) - ALL 3 tables.
        //   2) ALL 3 tables (Main report, Device Summary Stats, Grand Totals) hide any column
        //      that is ZERO for every row. Device groups with ALL metrics zero are hidden entirely.
        // Underlying data/business logic is UNCHANGED - only presentation + column visibility.
        // ============================================
        [HttpGet]
        public async Task<IActionResult> ExportMonthlyTransactionsExcel(int month, int year)
        {
            var monthlyData = await _collectionService.GetMonthlyTransactionsDataWithVehicles(year, month);
            if (monthlyData == null || !monthlyData.Any())
                return NotFound(new { success = false, message = "No data found for the selected month" });

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month) + " " + year + " - Transactions");
            int currentRow = 1;

            // TITLE
            worksheet.Cell(currentRow, 1).Value = "MONTHLY TRANSACTIONS SUMMARY - " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month).ToUpper() + " " + year;
            worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
            worksheet.Cell(currentRow, 1).Style.Font.FontSize = 16;
            worksheet.Cell(currentRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
            currentRow++;

            // ============================================================
            // BUILD RAW DATA FOR MAIN TABLE (21 logical columns) BEFORE WRITING
            // ============================================================
            var mainCols = new List<(string Header, bool AlwaysKeep, bool IsCurrency, string ColorHex)>
    {
        ("#", true, false, null),
        ("Date", true, false, null),
        ("Day", true, false, null),
        ("Total Vehicles", false, false, null),
        ("Total Transactions", false, false, null),
        ("Cash Count", false, false, null),
        ("Cash Total (KES)", false, true, "FFEBD6"),
        ("M-Pesa Count", false, false, null),
        ("M-Pesa Total (KES)", false, true, "FFEBD6"),
        ("Cash Count (TOP-UP)", false, false, null),
        ("Cash Total (KES)(TOP-UP)", false, true, "FFEBD6"),
        ("M-Pesa Count (TOP-UP)", false, false, null),
        ("M-Pesa Total (KES)(TOP-UP)", false, true, "FFEBD6"),
        ("Total Collected Cash (KES)", false, true, "E2EFDA"),
        ("Total Collected M-Pesa (KES)", false, true, "E2EFDA"),
        ("Total Collected Amount (KES)", false, true, "92D050"),
        ("FOC Count", false, false, null),
        ("FOC Amount (KES)", false, true, "FFEBD6"),
        ("HHD-4 FOC Count", false, false, null),
        ("HHD-4 FOC Amount (KES)", false, true, "FFEBD6"),
        ("Total Charged Amount (KES)", false, true, "FFEBD6"),
    };

            var mainRows = new List<object[]>();
            var dateList = new List<string>();
            int rowNumber = 1;

            foreach (dynamic item in monthlyData)
            {
                DateTime date = DateTime.Parse(item.date.ToString());
                decimal totalCollCash = (decimal)item.cashTotal + (decimal)item.cashTopUpTotal;
                decimal totalCollMpesa = (decimal)item.mPesaTotal + (decimal)item.mPesaTopUpTotal;
                decimal totalCollAmount = totalCollCash + totalCollMpesa;

                var row = new object[]
                {
            rowNumber++,
            date.ToString("dd-MMM-yyyy").ToUpper(),
            date.ToString("dddd"),
            (int)item.totalVehicles,
            (int)item.totalTransactions,
            (int)item.cashCount,
            (decimal)item.cashTotal,
            (int)item.mPesaCount,
            (decimal)item.mPesaTotal,
            (int)item.cashTopUpCount,
            (decimal)item.cashTopUpTotal,
            (int)item.mPesaTopUpCount,
            (decimal)item.mPesaTopUpTotal,
            totalCollCash,
            totalCollMpesa,
            totalCollAmount,
            (int)item.focCount,
            (decimal)item.focAmount,
            (int)item.ctmFocCount,
            (decimal)item.ctmFocAmount,
            (decimal)item.totalChargedAmount
                };
                mainRows.Add(row);
                dateList.Add(date.ToString("yyyy-MM-dd"));
            }

            // Determine which main columns to keep (always-keep OR has a non-zero value somewhere)
            var keepMain = new List<int>();
            for (int c = 0; c < mainCols.Count; c++)
            {
                if (mainCols[c].AlwaysKeep || mainRows.Any(r => Convert.ToDecimal(r[c]) != 0))
                    keepMain.Add(c);
            }

            // Title merge now that we know final column count
            worksheet.Range(1, 1, 1, keepMain.Count).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // HEADER ROW (wrapped, narrow)
            int headerRow = currentRow;
            for (int pos = 0; pos < keepMain.Count; pos++)
            {
                var col = mainCols[keepMain[pos]];
                var cell = worksheet.Cell(headerRow, pos + 1);
                cell.Value = col.Header;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.Gray;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Alignment.WrapText = true;
                worksheet.Column(pos + 1).Width = 12;
            }
            worksheet.Row(headerRow).Height = 32;
            currentRow++;

            int dataStartRow = currentRow;
            for (int r = 0; r < mainRows.Count; r++)
            {
                for (int pos = 0; pos < keepMain.Count; pos++)
                {
                    int origCol = keepMain[pos];
                    var cell = worksheet.Cell(currentRow, pos + 1);
                    cell.Value = XLCellValue.FromObject(mainRows[r][origCol]);
                    if (mainCols[origCol].IsCurrency)
                        cell.Style.NumberFormat.Format = "#,##0.00";
                }
                currentRow++;
            }
            int dataEndRow = currentRow - 1;

            // Column background colors (only for kept columns)
            for (int pos = 0; pos < keepMain.Count; pos++)
            {
                var col = mainCols[keepMain[pos]];
                if (!string.IsNullOrEmpty(col.ColorHex))
                {
                    var range = worksheet.Range(dataStartRow, pos + 1, dataEndRow, pos + 1);
                    range.Style.Fill.BackgroundColor = XLColor.FromHtml("#" + col.ColorHex);
                    range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                }
            }

            // ===== TOTALS ROW =====
            for (int pos = 0; pos < keepMain.Count; pos++)
            {
                int origCol = keepMain[pos];
                var header = mainCols[origCol].Header;
                var cell = worksheet.Cell(currentRow, pos + 1);

                if (header == "#")
                {
                    cell.Value = "TOTAL";
                }
                else if (header == "Date" || header == "Day")
                {
                    // leave blank
                }
                else
                {
                    string colLetter = worksheet.Cell(dataStartRow, pos + 1).Address.ColumnLetter;
                    cell.FormulaA1 = $"=SUM({colLetter}{dataStartRow}:{colLetter}{dataEndRow})";
                    if (mainCols[origCol].IsCurrency)
                        cell.Style.NumberFormat.Format = "0.00";
                }

                cell.Style.Font.Bold = true;
                cell.Style.Font.FontSize = 12;
                cell.Style.Fill.BackgroundColor = XLColor.Gold;
                cell.Style.Border.TopBorder = XLBorderStyleValues.Thick;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                if (mainCols[origCol].IsCurrency)
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            }

            // Keep the strong-green highlight on Total Collected Amount, if that column survived
            int totalCollectedPos = keepMain.FindIndex(c => mainCols[c].Header == "Total Collected Amount (KES)");
            if (totalCollectedPos >= 0)
                worksheet.Cell(currentRow, totalCollectedPos + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(146, 208, 80);

            currentRow++;

            // ============================================================
            // SUMMARY STATISTICS + GRAND TOTALS -> SEPARATE TAB, VERTICAL, ZERO COLUMNS/DEVICES HIDDEN
            // ============================================================
            var dailyStats = new Dictionary<string, dynamic>();
            foreach (dynamic item in monthlyData)
            {
                DateTime date = DateTime.Parse(item.date.ToString());
                string dateKey = date.ToString("yyyy-MM-dd");
                var stats = await _collectionService.GetTransactionStatistics(date, date);
                dailyStats[dateKey] = stats;
            }

            var sortedDates = monthlyData.Select(x => DateTime.Parse((string)x.date)).OrderBy(d => d).ToList();
            var devices = new[] { "Server", "Paystation1", "Paystation2", "Paystation3", "Paystation4", "HHD1", "HHD2", "HHD3", "HHD4", "Web" };

            var deviceColors = new Dictionary<string, XLColor>
    {
        {"Server", XLColor.FromArgb(255, 242, 204)},
        {"Paystation1", XLColor.FromArgb(252, 213, 180)},
        {"Paystation2", XLColor.FromArgb(198, 224, 180)},
        {"Paystation3", XLColor.FromArgb(180, 199, 231)},
        {"Paystation4", XLColor.FromArgb(217, 210, 233)},
        {"HHD1", XLColor.FromArgb(234, 209, 220)},
        {"HHD2", XLColor.FromArgb(207, 226, 243)},
        {"HHD3", XLColor.FromArgb(255, 230, 153)},
        {"HHD4", XLColor.FromArgb(182, 215, 168)},
        {"Web", XLColor.FromArgb(213, 166, 189)}
    };

            // ---- PHASE 1: define all possible metrics per device, then compute values per date ----
            var allMetricsByDevice = new Dictionary<string, List<(string Label, string Key, bool IsCurrency)>>();
            foreach (var device in devices)
            {
                if (device.StartsWith("Paystation"))
                {
                    allMetricsByDevice[device] = new List<(string, string, bool)>
            {
                ("Cash Count", "cashCount", false),
                ("Cash Total", "cashTotal", true),
                ("M-Pesa Count", "mPesaCount", false),
                ("M-Pesa Total", "mPesaTotal", true)
            };
                }
                else if (device == "Web")
                {
                    allMetricsByDevice[device] = new List<(string, string, bool)>
            {
                ("M-Pesa Count", "mPesaCount", false),
                ("M-Pesa Total", "mPesaTotal", true)
            };
                }
                else // Server or HHD
                {
                    allMetricsByDevice[device] = new List<(string, string, bool)>
            {
                ("Cash Count", "cashCount", false),
                ("Cash Total", "cashTotal", true),
                ("M-Pesa Count", "mPesaCount", false),
                ("M-Pesa Total", "mPesaTotal", true),
                ("M-Pesa Via Paybill Count", "mPesaViaPaybillCount", false),
                ("M-Pesa Via Paybill Total", "mPesaViaPaybillTotal", true),
                ("FOC Count", "focCount", false),
                ("FOC Amount", "focTotal", true)
            };
                }
            }

            // values[device][metricKey] = list of decimals, one per sortedDates entry
            var values = new Dictionary<string, Dictionary<string, List<decimal>>>();
            foreach (var device in devices)
            {
                values[device] = new Dictionary<string, List<decimal>>();
                foreach (var m in allMetricsByDevice[device])
                    values[device][m.Key] = new List<decimal>();

                foreach (var date in sortedDates)
                {
                    string dateKey = date.ToString("yyyy-MM-dd");
                    dynamic deviceStats = dailyStats.ContainsKey(dateKey) ? GetPropertyValue(dailyStats[dateKey].devices, device) : null;
                    foreach (var m in allMetricsByDevice[device])
                    {
                        decimal v = 0;
                        if (deviceStats != null)
                            v = Convert.ToDecimal(GetPropertyValue(deviceStats, m.Key));
                        values[device][m.Key].Add(v);
                    }
                }
            }

            // ---- PHASE 2: filter out all-zero metric columns, then all-zero devices ----
            var keptMetricsByDevice = new Dictionary<string, List<(string Label, string Key, bool IsCurrency)>>();
            foreach (var device in devices)
            {
                var kept = allMetricsByDevice[device]
                    .Where(m => values[device][m.Key].Any(v => v != 0))
                    .ToList();
                if (kept.Count > 0)
                    keptMetricsByDevice[device] = kept;
            }
            var keptDevices = devices.Where(d => keptMetricsByDevice.ContainsKey(d)).ToList();

            var summarySheet = workbook.Worksheets.Add("Summary Statistics");
            int sRow = 1;

            summarySheet.Cell(sRow, 1).Value = "SUMMARY STATISTICS (TRANSACTIONS)";
            summarySheet.Cell(sRow, 1).Style.Font.Bold = true;
            summarySheet.Cell(sRow, 1).Style.Font.FontSize = 14;
            summarySheet.Cell(sRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
            sRow += 2;

            int deviceHeaderRow = sRow;
            int subHeaderRow = sRow + 1;
            int dStart = sRow + 2;

            summarySheet.Cell(subHeaderRow, 1).Value = "DATE";
            summarySheet.Cell(subHeaderRow, 1).Style.Font.Bold = true;
            summarySheet.Cell(subHeaderRow, 1).Style.Alignment.WrapText = true;
            summarySheet.Cell(subHeaderRow, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            summarySheet.Range(deviceHeaderRow, 1, subHeaderRow, 1).Merge();
            summarySheet.Row(subHeaderRow).Height = 32;
            summarySheet.Column(1).Width = 13;

            var deviceColumnMap = new Dictionary<string, (int start, int end)>();
            int currentCol = 2;

            if (keptDevices.Count == 0)
            {
                summarySheet.Cell(dStart, 1).Value = "No non-zero collection data found for this period.";
            }
            else
            {
                foreach (var device in keptDevices)
                {
                    var metrics = keptMetricsByDevice[device];
                    int startCol = currentCol;
                    foreach (var m in metrics)
                    {
                        var cell = summarySheet.Cell(subHeaderRow, currentCol);
                        cell.Value = m.Label.ToUpper();
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        cell.Style.Alignment.WrapText = true;
                        summarySheet.Column(currentCol).Width = 11;
                        currentCol++;
                    }
                    int endCol = currentCol - 1;

                    summarySheet.Cell(deviceHeaderRow, startCol).Value = device.ToUpper();
                    summarySheet.Range(deviceHeaderRow, startCol, deviceHeaderRow, endCol).Merge();
                    summarySheet.Cell(deviceHeaderRow, startCol).Style.Font.Bold = true;
                    summarySheet.Cell(deviceHeaderRow, startCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    summarySheet.Range(deviceHeaderRow, startCol, deviceHeaderRow, endCol).Style.Fill.BackgroundColor =
                        deviceColors.ContainsKey(device) ? deviceColors[device] : XLColor.LightBlue;

                    deviceColumnMap[device] = (startCol, endCol);
                }

                int lastCol = currentCol - 1;

                // Data rows: one row per date (vertical), reading precomputed values
                int rIdx = dStart;
                for (int di = 0; di < sortedDates.Count; di++)
                {
                    var date = sortedDates[di];
                    summarySheet.Cell(rIdx, 1).Value = date.ToString("dd-MMM-yyyy").ToUpper();
                    summarySheet.Cell(rIdx, 1).Style.Font.Bold = true;

                    foreach (var device in keptDevices)
                    {
                        var (startCol, _) = deviceColumnMap[device];
                        int col = startCol;
                        foreach (var m in keptMetricsByDevice[device])
                        {
                            decimal value = values[device][m.Key][di];
                            summarySheet.Cell(rIdx, col).Value = value;
                            if (m.IsCurrency)
                                summarySheet.Cell(rIdx, col).Style.NumberFormat.Format = "#,##0.00";
                            col++;
                        }
                    }
                    rIdx++;
                }

                // Monthly total row
                summarySheet.Cell(rIdx, 1).Value = "MONTHLY TOTAL";
                summarySheet.Cell(rIdx, 1).Style.Font.Bold = true;
                summarySheet.Cell(rIdx, 1).Style.Fill.BackgroundColor = XLColor.Gold;
                for (int c = 2; c <= lastCol; c++)
                {
                    string colLetter = summarySheet.Cell(dStart, c).Address.ColumnLetter;
                    summarySheet.Cell(rIdx, c).FormulaA1 = $"=SUM({colLetter}{dStart}:{colLetter}{rIdx - 1})";
                    summarySheet.Cell(rIdx, c).Style.Font.Bold = true;
                    summarySheet.Cell(rIdx, c).Style.Fill.BackgroundColor = XLColor.Gold;
                }

                var deviceTableRange = summarySheet.Range(deviceHeaderRow, 1, rIdx, lastCol);
                deviceTableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                deviceTableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Thick outside border per device group for visual separation
                foreach (var device in keptDevices)
                {
                    var (startCol, endCol) = deviceColumnMap[device];
                    var deviceRange = summarySheet.Range(deviceHeaderRow, startCol, rIdx, endCol);
                    deviceRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                }

                int grandRow = rIdx + 3;

                // ============================================================
                // GRAND TOTALS - vertical, ZERO COLUMNS HIDDEN
                // ============================================================
                summarySheet.Cell(grandRow, 1).Value = "GRAND TOTALS";
                summarySheet.Cell(grandRow, 1).Style.Font.Bold = true;
                summarySheet.Cell(grandRow, 1).Style.Font.FontSize = 14;
                summarySheet.Cell(grandRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                grandRow += 2;

                var gtCols = new List<(string Header, bool AlwaysKeep, bool IsCurrency)>
        {
            ("DATE", true, false),
            ("TOTAL VEHICLES", false, false),
            ("TOTAL TRANSACTIONS", false, false),
            ("CASH COLLECTED (KES)", false, true),
            ("M-PESA COLLECTED (KES)", false, true),
            ("CASH TOP-UP COLLECTED (KES)", false, true),
            ("M-PESA TOP-UP COLLECTED (KES)", false, true),
            ("TOTAL TOP-UP COLLECTED (KES)", false, true),
            ("FOC AMOUNT (KES)", false, true),
            ("HHD-4 FOC AMOUNT (KES)", false, true),
            ("TOTAL AMOUNT COLLECTED (KES)", false, true)
        };

                var gtRows = new List<object[]>();
                foreach (dynamic item in monthlyData)
                {
                    DateTime date = DateTime.Parse(item.date.ToString());
                    decimal cashTopUp = (decimal)item.cashTopUpTotal;
                    decimal mpesaTopUp = (decimal)item.mPesaTopUpTotal;
                    decimal totalTopUp = cashTopUp + mpesaTopUp;
                    decimal totalCollected = (decimal)item.cashTotal + (decimal)item.mPesaTotal + cashTopUp + mpesaTopUp;

                    gtRows.Add(new object[]
                    {
                date.ToString("dd-MMM-yyyy").ToUpper(),
                (int)item.totalVehicles,
                (int)item.totalTransactions,
                (decimal)item.cashTotal,
                (decimal)item.mPesaTotal,
                cashTopUp,
                mpesaTopUp,
                totalTopUp,
                (decimal)item.focAmount,
                (decimal)item.ctmFocAmount,
                totalCollected
                    });
                }

                var keepGt = new List<int>();
                for (int c = 0; c < gtCols.Count; c++)
                {
                    if (gtCols[c].AlwaysKeep || gtRows.Any(r => Convert.ToDecimal(r[c]) != 0))
                        keepGt.Add(c);
                }

                int gtHeaderRow = grandRow;
                for (int pos = 0; pos < keepGt.Count; pos++)
                {
                    var col = gtCols[keepGt[pos]];
                    var cell = summarySheet.Cell(gtHeaderRow, pos + 1);
                    cell.Value = col.Header;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.Gray;
                    cell.Style.Font.FontColor = XLColor.White;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Alignment.WrapText = true;
                    summarySheet.Column(pos + 1).Width = 13;
                }
                summarySheet.Row(gtHeaderRow).Height = 32;

                int gtDataStart = gtHeaderRow + 1;
                for (int r = 0; r < gtRows.Count; r++)
                {
                    int rr = gtDataStart + r;
                    for (int pos = 0; pos < keepGt.Count; pos++)
                    {
                        int origCol = keepGt[pos];
                        var cell = summarySheet.Cell(rr, pos + 1);
                        cell.Value = XLCellValue.FromObject(gtRows[r][origCol]);
                        if (gtCols[origCol].IsCurrency)
                            cell.Style.NumberFormat.Format = "0.00";
                    }
                }
                int gtDataEnd = gtDataStart + gtRows.Count - 1;

                // Highlight Total Amount Collected column if it survived
                int totalAmtPos = keepGt.FindIndex(c => gtCols[c].Header == "TOTAL AMOUNT COLLECTED (KES)");
                if (totalAmtPos >= 0)
                    summarySheet.Range(gtDataStart, totalAmtPos + 1, gtDataEnd, totalAmtPos + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(146, 208, 80);

                int gtTotalRow = gtDataEnd + 1;
                for (int pos = 0; pos < keepGt.Count; pos++)
                {
                    int origCol = keepGt[pos];
                    var cell = summarySheet.Cell(gtTotalRow, pos + 1);
                    if (gtCols[origCol].Header == "DATE")
                    {
                        cell.Value = "TOTAL";
                    }
                    else
                    {
                        string colLetter = summarySheet.Cell(gtDataStart, pos + 1).Address.ColumnLetter;
                        cell.FormulaA1 = $"=SUM({colLetter}{gtDataStart}:{colLetter}{gtDataEnd})";
                        if (gtCols[origCol].IsCurrency)
                            cell.Style.NumberFormat.Format = "0.00";
                    }
                    cell.Style.Font.Bold = true;
                }
                if (totalAmtPos >= 0)
                    summarySheet.Cell(gtTotalRow, totalAmtPos + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(146, 208, 80);

                var gtRange = summarySheet.Range(gtHeaderRow, 1, gtTotalRow, keepGt.Count);
                gtRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                gtRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                summarySheet.Range(gtTotalRow, 1, gtTotalRow, keepGt.Count).Style.Fill.BackgroundColor = XLColor.Gold;

                summarySheet.SheetView.FreezeRows(subHeaderRow);
                summarySheet.SheetView.FreezeColumns(1);
            }

            // ============================================================
            // SAVE
            // ============================================================
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;
            var fileName = $"MonthlyTransactionsReport_{CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month)}_{year}.xlsx";
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        // ============================================
        // HOURLY COLLECTION REPORT (NEW)
        // ============================================
        public class HourlyFilterRequest
        {
            public DateTime Date { get; set; }
        }

        public class HourlyRow
        {
            public int hour { get; set; }
            public string label { get; set; }
            public int count { get; set; }
            public int cashCount { get; set; }
            public decimal cashAmount { get; set; }
            public int mpesaCount { get; set; }
            public decimal mpesaAmount { get; set; }
            public int paybillCount { get; set; }
            public decimal paybillAmount { get; set; }
            public int focCount { get; set; }
            public decimal focAmount { get; set; }
            public int otherCount { get; set; }
            public decimal otherAmount { get; set; }
            public decimal totalCharges { get; set; }
            public decimal discount { get; set; }
            public decimal totalPaid { get; set; }
            public Dictionary<string, int> deviceCount { get; set; } = new Dictionary<string, int>();
            public Dictionary<string, decimal> deviceAmount { get; set; } = new Dictionary<string, decimal>();
        }

        public class HourlyTxn
        {
            public string hourLabel { get; set; }
            public string transactionId { get; set; }
            public string cardNo { get; set; }
            public string cardName { get; set; }
            public string entryTime { get; set; }
            public string exitTime { get; set; }
            public string parkTime { get; set; }
            public decimal charge { get; set; }
            public decimal paid { get; set; }
            public decimal discount { get; set; }
            public string payMode { get; set; }
            public string paySource { get; set; }
        }

        private static readonly string[] HourlyDevices = new[]
        {
            "Server", "Paystation1", "Paystation2", "Paystation3", "Paystation4",
            "HHD1", "HHD2", "HHD3", "HHD4", "Web"
        };

        private static string HourlyDeviceKey(string src)
        {
            if (string.IsNullOrWhiteSpace(src)) return "Server";
            string n = src.Trim().ToUpperInvariant().Replace(" ", "").Replace("-", "").Replace("_", "");
            foreach (var d in HourlyDevices)
            {
                if (n.Contains(d.ToUpperInvariant())) return d;
            }
            return "Server";
        }

        private async Task<(List<HourlyRow> rows, List<HourlyTxn> txns)> BuildHourlyReport(DateTime date)
        {
            var data = await _collectionService.GetDailyCollectionData(date, date);

            var rows = new List<HourlyRow>();
            for (int h = 0; h < 24; h++)
            {
                var r = new HourlyRow { hour = h, label = $"{h:D2}:00 - {h:D2}:59" };
                foreach (var d in HourlyDevices)
                {
                    r.deviceCount[d] = 0;
                    r.deviceAmount[d] = 0m;
                }
                rows.Add(r);
            }

            var txns = new List<HourlyTxn>();

            foreach (var rec in data)
            {
                if (!rec.exit_time.HasValue) continue;

                int hour = rec.exit_time.Value.Hour;
                var row = rows[hour];

                decimal charge = rec.charges ?? 0m;
                decimal paid = rec.paid_amount ?? 0m;
                decimal disc = (decimal)(rec.discount_amount ?? 0m);
                string payMode = NormalizePayModeLocal(rec.pay_mode);
                string device = HourlyDeviceKey(rec.pay_source);

                row.count++;
                row.totalCharges += charge;
                row.discount += disc;
                row.totalPaid += paid;
                row.deviceCount[device]++;
                row.deviceAmount[device] += paid;

                if (payMode == "Cash") { row.cashCount++; row.cashAmount += paid; }
                else if (payMode == "M-Pesa") { row.mpesaCount++; row.mpesaAmount += paid; }
                else if (payMode == "M-Pesa Via Paybill") { row.paybillCount++; row.paybillAmount += paid; }
                else if (payMode == "Free of Charge") { row.focCount++; row.focAmount += charge; }
                else { row.otherCount++; row.otherAmount += paid; }

                string parkTime = "N/A";
                if (rec.park_time_total.HasValue && rec.park_time_total.Value > 0)
                {
                    int hrs = rec.park_time_total.Value / 60;
                    int mins = rec.park_time_total.Value % 60;
                    parkTime = $"{hrs}h {mins}m";
                }

                txns.Add(new HourlyTxn
                {
                    hourLabel = row.label,
                    transactionId = rec.transaction_id,
                    cardNo = rec.cardno,
                    cardName = rec.card_name,
                    entryTime = rec.entry_time.HasValue ? rec.entry_time.Value.ToString("dd-MMM-yyyy HH:mm:ss").ToUpper() : "--",
                    exitTime = rec.exit_time.Value.ToString("dd-MMM-yyyy HH:mm:ss").ToUpper(),
                    parkTime = parkTime,
                    charge = charge,
                    paid = paid,
                    discount = disc,
                    payMode = rec.pay_mode,
                    paySource = rec.pay_source
                });
            }

            txns = txns.OrderBy(t => t.exitTime == null ? DateTime.MinValue : DateTime.ParseExact(t.exitTime, "dd-MMM-yyyy HH:mm:ss", CultureInfo.InvariantCulture)).ToList();
            return (rows, txns);
        }

        public IActionResult Hourly()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> GetHourlyData([FromBody] HourlyFilterRequest request)
        {
            try
            {
                var (rows, txns) = await BuildHourlyReport(request.Date.Date);
                return Json(new
                {
                    success = true,
                    hours = rows,
                    transactions = txns,
                    recordCount = txns.Count
                });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetHourlyData: {ex.Message}\n{ex.StackTrace}");
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public async Task<IActionResult> ExportHourlyExcel(string date)
        {
            try
            {
                if (!DateTime.TryParse(date, out var d))
                    return BadRequest("Invalid date format");

                var (rows, txns) = await BuildHourlyReport(d.Date);

                using var workbook = new XLWorkbook();

                // ---------- Sheet 1: Hourly Summary ----------
                var ws = workbook.Worksheets.Add("Hourly Summary");
                string[] headers = {
                    "Hour", "Vehicles", "Cash Count", "Cash (KES)", "M-Pesa Count", "M-Pesa (KES)",
                    "Paybill Count", "Paybill (KES)", "FOC Count", "FOC Amount (KES)",
                    "Other Count", "Other (KES)", "Total Charges (KES)", "Discount (KES)", "Total Paid (KES)"
                };
                for (int c = 0; c < headers.Length; c++)
                    ws.Cell(1, c + 1).Value = headers[c];

                var hdr = ws.Range(1, 1, 1, headers.Length);
                hdr.Style.Font.Bold = true;
                hdr.Style.Fill.BackgroundColor = XLColor.LightGray;
                hdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                hdr.Style.Alignment.WrapText = true;
                ws.SheetView.FreezeRows(1);

                int r = 2;
                foreach (var h in rows)
                {
                    object[] vals = {
                        h.label, h.count, h.cashCount, h.cashAmount, h.mpesaCount, h.mpesaAmount,
                        h.paybillCount, h.paybillAmount, h.focCount, h.focAmount,
                        h.otherCount, h.otherAmount, h.totalCharges, h.discount, h.totalPaid
                    };
                    for (int c = 0; c < vals.Length; c++)
                        ws.Cell(r, c + 1).Value = XLCellValue.FromObject(vals[c]);
                    r++;
                }

                ws.Cell(r, 1).Value = "TOTAL";
                for (int c = 2; c <= headers.Length; c++)
                {
                    string col = ws.Cell(2, c).Address.ColumnLetter;
                    ws.Cell(r, c).FormulaA1 = $"=SUM({col}2:{col}{r - 1})";
                }
                var totalRange = ws.Range(r, 1, r, headers.Length);
                totalRange.Style.Font.Bold = true;
                totalRange.Style.Fill.BackgroundColor = XLColor.Gold;

                foreach (var c in new[] { 4, 6, 8, 10, 12, 13, 14, 15 })
                    ws.Range(2, c, r, c).Style.NumberFormat.Format = "#,##0.00";

                ws.Range(1, 1, r, headers.Length).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                ws.Range(1, 1, r, headers.Length).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ws.Column(1).Width = 16;
                for (int c = 2; c <= headers.Length; c++) ws.Column(c).Width = 13;
                ws.Row(1).Height = 32;

                // ---------- Sheet 2: Payment Source Summary ----------
                var ds = workbook.Worksheets.Add("Payment Source Summary");
                ds.Cell(1, 1).Value = "Hour";
                ds.Range(1, 1, 2, 1).Merge();
                int col2 = 2;
                foreach (var dev in HourlyDevices)
                {
                    ds.Cell(1, col2).Value = dev;
                    ds.Range(1, col2, 1, col2 + 1).Merge();
                    ds.Cell(2, col2).Value = "Count";
                    ds.Cell(2, col2 + 1).Value = "Amount (KES)";
                    col2 += 2;
                }
                ds.Cell(1, col2).Value = "Total";
                ds.Range(1, col2, 1, col2 + 1).Merge();
                ds.Cell(2, col2).Value = "Count";
                ds.Cell(2, col2 + 1).Value = "Amount (KES)";
                int lastCol = col2 + 1;

                var dsHdr = ds.Range(1, 1, 2, lastCol);
                dsHdr.Style.Font.Bold = true;
                dsHdr.Style.Fill.BackgroundColor = XLColor.LightGray;
                dsHdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                dsHdr.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ds.SheetView.FreezeRows(2);

                int dr = 3;
                foreach (var h in rows)
                {
                    ds.Cell(dr, 1).Value = h.label;
                    int cc = 2;
                    foreach (var dev in HourlyDevices)
                    {
                        ds.Cell(dr, cc).Value = h.deviceCount[dev];
                        ds.Cell(dr, cc + 1).Value = h.deviceAmount[dev];
                        ds.Cell(dr, cc + 1).Style.NumberFormat.Format = "#,##0.00";
                        cc += 2;
                    }
                    ds.Cell(dr, cc).Value = h.count;
                    ds.Cell(dr, cc + 1).Value = h.totalPaid;
                    ds.Cell(dr, cc + 1).Style.NumberFormat.Format = "#,##0.00";
                    dr++;
                }

                ds.Cell(dr, 1).Value = "TOTAL";
                for (int c = 2; c <= lastCol; c++)
                {
                    string col = ds.Cell(3, c).Address.ColumnLetter;
                    ds.Cell(dr, c).FormulaA1 = $"=SUM({col}3:{col}{dr - 1})";
                    if (c % 2 == 1) ds.Cell(dr, c).Style.NumberFormat.Format = "#,##0.00";
                }
                var dsTotal = ds.Range(dr, 1, dr, lastCol);
                dsTotal.Style.Font.Bold = true;
                dsTotal.Style.Fill.BackgroundColor = XLColor.Gold;
                ds.Range(1, 1, dr, lastCol).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                ds.Range(1, 1, dr, lastCol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ds.Column(1).Width = 16;
                for (int c = 2; c <= lastCol; c++) ds.Column(c).Width = 13;

                // ---------- Sheet 3: Transactions ----------
                var ts = workbook.Worksheets.Add("Transactions");
                string[] th = {
                    "#", "Hour", "Transaction ID", "Card Number", "Card Name", "Entry Time", "Exit Time",
                    "Park Time", "Charge (KES)", "Paid (KES)", "Discount (KES)", "Pay Mode", "Payment Source"
                };
                for (int c = 0; c < th.Length; c++)
                    ts.Cell(1, c + 1).Value = th[c];
                var tsHdr = ts.Range(1, 1, 1, th.Length);
                tsHdr.Style.Font.Bold = true;
                tsHdr.Style.Fill.BackgroundColor = XLColor.LightGray;
                tsHdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ts.SheetView.FreezeRows(1);

                int tr = 2;
                int no = 1;
                foreach (var t in txns)
                {
                    ts.Cell(tr, 1).Value = no++;
                    ts.Cell(tr, 2).Value = t.hourLabel ?? "";
                    ts.Cell(tr, 3).Value = t.transactionId ?? "";
                    ts.Cell(tr, 4).Value = t.cardNo ?? "";
                    ts.Cell(tr, 5).Value = t.cardName ?? "";
                    ts.Cell(tr, 6).Value = t.entryTime ?? "";
                    ts.Cell(tr, 7).Value = t.exitTime ?? "";
                    ts.Cell(tr, 8).Value = t.parkTime ?? "";
                    ts.Cell(tr, 9).Value = t.charge;
                    ts.Cell(tr, 10).Value = t.paid;
                    ts.Cell(tr, 11).Value = t.discount;
                    ts.Cell(tr, 12).Value = t.payMode ?? "";
                    ts.Cell(tr, 13).Value = t.paySource ?? "";
                    ts.Range(tr, 9, tr, 11).Style.NumberFormat.Format = "#,##0.00";
                    tr++;
                }
                ts.Columns().AdjustToContents();

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                var fileName = $"HOURLY_REPORT_{d:dd-MMM-yyyy}_{DateTime.Now:HH-mm-ss}.xlsx";
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ExportHourlyExcel: {ex.Message}\n{ex.StackTrace}");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        // ============================================
        // HOURLY COLLECTION: TRANSACTIONS TAB (NEW)
        // ============================================
        public class HourlyTransTxn
        {
            public string hourLabel { get; set; }
            public string transactionId { get; set; }
            public string cardNo { get; set; }
            public string cardName { get; set; }
            public string entryTime { get; set; }
            public string validationTime { get; set; }
            public decimal amount { get; set; }
            public string payMode { get; set; }
            public string paySource { get; set; }
            public string status { get; set; }
        }

        private async Task<(List<HourlyRow> rows, List<HourlyTransTxn> txns)> BuildHourlyTransactionReport(DateTime date)
        {
            var data = await _collectionService.GetTransactionSummary(date, date);

            var rows = new List<HourlyRow>();
            for (int h = 0; h < 24; h++)
            {
                var r = new HourlyRow { hour = h, label = $"{h:D2}:00 - {h:D2}:59" };
                foreach (var d in HourlyDevices)
                {
                    r.deviceCount[d] = 0;
                    r.deviceAmount[d] = 0m;
                }
                rows.Add(r);
            }

            var list = new List<(DateTime when, HourlyTransTxn txn)>();

            foreach (var rec in data)
            {
                DateTime when;
                if (rec.validation_time.HasValue) when = rec.validation_time.Value;
                else if (rec.revalidation_time.HasValue) when = rec.revalidation_time.Value;
                else continue;

                var row = rows[when.Hour];

                string rawMode = !string.IsNullOrWhiteSpace(rec.validation_paymode)
                    ? rec.validation_paymode
                    : rec.revalidation_paymode;

                decimal amount = rec.paidamount ?? 0m;
                string device = HourlyDeviceKey(rec.payment_source);

                row.count++;
                row.totalPaid += amount;
                row.deviceCount[device]++;
                row.deviceAmount[device] += amount;

                string payMode = string.IsNullOrWhiteSpace(rawMode) ? "Other" : NormalizePayModeLocal(rawMode);
                if (payMode == "Cash") { row.cashCount++; row.cashAmount += amount; }
                else if (payMode == "M-Pesa") { row.mpesaCount++; row.mpesaAmount += amount; }
                else if (payMode == "M-Pesa Via Paybill") { row.paybillCount++; row.paybillAmount += amount; }
                else if (payMode == "Free of Charge") { row.focCount++; row.focAmount += amount; }
                else { row.otherCount++; row.otherAmount += amount; }

                list.Add((when, new HourlyTransTxn
                {
                    hourLabel = row.label,
                    transactionId = rec.transaction_id,
                    cardNo = rec.cardno,
                    cardName = rec.cardname,
                    entryTime = rec.entry_time.HasValue ? rec.entry_time.Value.ToString("dd-MMM-yyyy HH:mm:ss").ToUpper() : "--",
                    validationTime = when.ToString("dd-MMM-yyyy HH:mm:ss").ToUpper(),
                    amount = amount,
                    payMode = string.IsNullOrWhiteSpace(rawMode) ? "N/A" : rawMode,
                    paySource = rec.payment_source,
                    status = rec.status?.ToString()
                }));
            }

            var txns = list.OrderBy(x => x.when).Select(x => x.txn).ToList();
            return (rows, txns);
        }

        [HttpPost]
        public async Task<IActionResult> GetHourlyTransactionData([FromBody] HourlyFilterRequest request)
        {
            try
            {
                var (rows, txns) = await BuildHourlyTransactionReport(request.Date.Date);
                return Json(new
                {
                    success = true,
                    hours = rows,
                    transactions = txns,
                    recordCount = txns.Count
                });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetHourlyTransactionData: {ex.Message}\n{ex.StackTrace}");
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public async Task<IActionResult> ExportHourlyTransactionExcel(string date)
        {
            try
            {
                if (!DateTime.TryParse(date, out var d))
                    return BadRequest("Invalid date format");

                var (rows, txns) = await BuildHourlyTransactionReport(d.Date);

                using var workbook = new XLWorkbook();

                // ---------- Sheet 1: Hourly Summary ----------
                var ws = workbook.Worksheets.Add("Hourly Summary");
                string[] headers = {
                    "Hour", "Transactions", "Cash Count", "Cash (KES)", "M-Pesa Count", "M-Pesa (KES)",
                    "Paybill Count", "Paybill (KES)", "FOC Count", "FOC Amount (KES)",
                    "Other Count", "Other (KES)", "Total Paid (KES)"
                };
                for (int c = 0; c < headers.Length; c++)
                    ws.Cell(1, c + 1).Value = headers[c];

                var hdr = ws.Range(1, 1, 1, headers.Length);
                hdr.Style.Font.Bold = true;
                hdr.Style.Fill.BackgroundColor = XLColor.LightGray;
                hdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                hdr.Style.Alignment.WrapText = true;
                ws.SheetView.FreezeRows(1);

                int r = 2;
                foreach (var h in rows)
                {
                    object[] vals = {
                        h.label, h.count, h.cashCount, h.cashAmount, h.mpesaCount, h.mpesaAmount,
                        h.paybillCount, h.paybillAmount, h.focCount, h.focAmount,
                        h.otherCount, h.otherAmount, h.totalPaid
                    };
                    for (int c = 0; c < vals.Length; c++)
                        ws.Cell(r, c + 1).Value = XLCellValue.FromObject(vals[c]);
                    r++;
                }

                ws.Cell(r, 1).Value = "TOTAL";
                for (int c = 2; c <= headers.Length; c++)
                {
                    string col = ws.Cell(2, c).Address.ColumnLetter;
                    ws.Cell(r, c).FormulaA1 = $"=SUM({col}2:{col}{r - 1})";
                }
                var totalRange = ws.Range(r, 1, r, headers.Length);
                totalRange.Style.Font.Bold = true;
                totalRange.Style.Fill.BackgroundColor = XLColor.Gold;

                foreach (var c in new[] { 4, 6, 8, 10, 12, 13 })
                    ws.Range(2, c, r, c).Style.NumberFormat.Format = "#,##0.00";

                ws.Range(1, 1, r, headers.Length).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                ws.Range(1, 1, r, headers.Length).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ws.Column(1).Width = 16;
                for (int c = 2; c <= headers.Length; c++) ws.Column(c).Width = 13;
                ws.Row(1).Height = 32;

                // ---------- Sheet 2: Payment Source Summary ----------
                var ds = workbook.Worksheets.Add("Payment Source Summary");
                ds.Cell(1, 1).Value = "Hour";
                ds.Range(1, 1, 2, 1).Merge();
                int col2 = 2;
                foreach (var dev in HourlyDevices)
                {
                    ds.Cell(1, col2).Value = dev;
                    ds.Range(1, col2, 1, col2 + 1).Merge();
                    ds.Cell(2, col2).Value = "Count";
                    ds.Cell(2, col2 + 1).Value = "Amount (KES)";
                    col2 += 2;
                }
                ds.Cell(1, col2).Value = "Total";
                ds.Range(1, col2, 1, col2 + 1).Merge();
                ds.Cell(2, col2).Value = "Count";
                ds.Cell(2, col2 + 1).Value = "Amount (KES)";
                int lastCol = col2 + 1;

                var dsHdr = ds.Range(1, 1, 2, lastCol);
                dsHdr.Style.Font.Bold = true;
                dsHdr.Style.Fill.BackgroundColor = XLColor.LightGray;
                dsHdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                dsHdr.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ds.SheetView.FreezeRows(2);

                int dr = 3;
                foreach (var h in rows)
                {
                    ds.Cell(dr, 1).Value = h.label;
                    int cc = 2;
                    foreach (var dev in HourlyDevices)
                    {
                        ds.Cell(dr, cc).Value = h.deviceCount[dev];
                        ds.Cell(dr, cc + 1).Value = h.deviceAmount[dev];
                        ds.Cell(dr, cc + 1).Style.NumberFormat.Format = "#,##0.00";
                        cc += 2;
                    }
                    ds.Cell(dr, cc).Value = h.count;
                    ds.Cell(dr, cc + 1).Value = h.totalPaid;
                    ds.Cell(dr, cc + 1).Style.NumberFormat.Format = "#,##0.00";
                    dr++;
                }

                ds.Cell(dr, 1).Value = "TOTAL";
                for (int c = 2; c <= lastCol; c++)
                {
                    string col = ds.Cell(3, c).Address.ColumnLetter;
                    ds.Cell(dr, c).FormulaA1 = $"=SUM({col}3:{col}{dr - 1})";
                    if (c % 2 == 1) ds.Cell(dr, c).Style.NumberFormat.Format = "#,##0.00";
                }
                var dsTotal = ds.Range(dr, 1, dr, lastCol);
                dsTotal.Style.Font.Bold = true;
                dsTotal.Style.Fill.BackgroundColor = XLColor.Gold;
                ds.Range(1, 1, dr, lastCol).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                ds.Range(1, 1, dr, lastCol).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ds.Column(1).Width = 16;
                for (int c = 2; c <= lastCol; c++) ds.Column(c).Width = 13;

                // ---------- Sheet 3: Transactions ----------
                var ts = workbook.Worksheets.Add("Transactions");
                string[] th = {
                    "#", "Hour", "Transaction ID", "Card Number", "Card Name", "Entry Time",
                    "Validation Time", "Amount (KES)", "Pay Mode", "Payment Source", "Status"
                };
                for (int c = 0; c < th.Length; c++)
                    ts.Cell(1, c + 1).Value = th[c];
                var tsHdr = ts.Range(1, 1, 1, th.Length);
                tsHdr.Style.Font.Bold = true;
                tsHdr.Style.Fill.BackgroundColor = XLColor.LightGray;
                tsHdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ts.SheetView.FreezeRows(1);

                int tr = 2;
                int no = 1;
                foreach (var t in txns)
                {
                    ts.Cell(tr, 1).Value = no++;
                    ts.Cell(tr, 2).Value = t.hourLabel ?? "";
                    ts.Cell(tr, 3).Value = t.transactionId ?? "";
                    ts.Cell(tr, 4).Value = t.cardNo ?? "";
                    ts.Cell(tr, 5).Value = t.cardName ?? "";
                    ts.Cell(tr, 6).Value = t.entryTime ?? "";
                    ts.Cell(tr, 7).Value = t.validationTime ?? "";
                    ts.Cell(tr, 8).Value = t.amount;
                    ts.Cell(tr, 8).Style.NumberFormat.Format = "#,##0.00";
                    ts.Cell(tr, 9).Value = t.payMode ?? "";
                    ts.Cell(tr, 10).Value = t.paySource ?? "";
                    ts.Cell(tr, 11).Value = t.status ?? "";
                    tr++;
                }
                ts.Columns().AdjustToContents();

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                var fileName = $"HOURLY_TRANSACTION_REPORT_{d:dd-MMM-yyyy}_{DateTime.Now:HH-mm-ss}.xlsx";
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ExportHourlyTransactionExcel: {ex.Message}\n{ex.StackTrace}");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        // ============================================
        // YEARLY COLLECTION REPORT (NEW)
        // Built from existing monthly service methods (month-wise totals)
        // ============================================
        public class YearlyExitedRow
        {
            public int month { get; set; }
            public string label { get; set; }
            public int totalVehicles { get; set; }
            public int cashCount { get; set; }
            public decimal cashTotal { get; set; }
            public int mPesaCount { get; set; }
            public decimal mPesaTotal { get; set; }
            public int focCount { get; set; }
            public decimal focAmount { get; set; }
            public int ctmFocCount { get; set; }
            public decimal ctmFocAmount { get; set; }
            public int zeroChargeCount { get; set; }
            public int permanentParkersCount { get; set; }
            public int revalidationCount { get; set; }
            public decimal revalidationAmount { get; set; }
            public decimal totalCharges { get; set; }
            public decimal discount { get; set; }
            public decimal totalPaid { get; set; }
            public int vipCount { get; set; }
        }

        public class YearlyTransRow
        {
            public int month { get; set; }
            public string label { get; set; }
            public int totalVehicles { get; set; }
            public int totalTransactions { get; set; }
            public int cashCount { get; set; }
            public decimal cashTotal { get; set; }
            public int mPesaCount { get; set; }
            public decimal mPesaTotal { get; set; }
            public int cashTopUpCount { get; set; }
            public decimal cashTopUpTotal { get; set; }
            public int mPesaTopUpCount { get; set; }
            public decimal mPesaTopUpTotal { get; set; }
            public int focCount { get; set; }
            public decimal focAmount { get; set; }
            public int ctmFocCount { get; set; }
            public decimal ctmFocAmount { get; set; }
            public decimal totalChargedAmount { get; set; }
            public decimal totalCollected { get; set; }
        }

        private static bool IsFutureMonth(int year, int month)
        {
            var now = DateTime.Now;
            return year > now.Year || (year == now.Year && month > now.Month);
        }

        private async Task<List<YearlyExitedRow>> BuildYearlyExited(int year)
        {
            var rows = new List<YearlyExitedRow>();
            for (int m = 1; m <= 12; m++)
            {
                var row = new YearlyExitedRow
                {
                    month = m,
                    label = new DateTime(year, m, 1).ToString("MMMM yyyy", CultureInfo.InvariantCulture)
                };
                rows.Add(row);

                if (IsFutureMonth(year, m)) continue;

                var data = await _collectionService.GetMonthlyCollectionData(year, m);
                if (data == null) continue;

                foreach (dynamic item in data)
                {
                    row.totalVehicles += Convert.ToInt32(item.totalVehicles);
                    row.cashCount += Convert.ToInt32(item.cashCount);
                    row.cashTotal += Convert.ToDecimal(item.cashTotal);
                    row.mPesaCount += Convert.ToInt32(item.mPesaCount);
                    row.mPesaTotal += Convert.ToDecimal(item.mPesaTotal);
                    row.focCount += Convert.ToInt32(item.focCount);
                    row.focAmount += Convert.ToDecimal(item.focAmount);
                    row.ctmFocCount += Convert.ToInt32(item.ctmFocCount);
                    row.ctmFocAmount += Convert.ToDecimal(item.ctmFocAmount);
                    row.zeroChargeCount += Convert.ToInt32(item.zeroChargeCount);
                    row.permanentParkersCount += Convert.ToInt32(item.permanentParkersCount);
                    row.revalidationCount += Convert.ToInt32(item.revalidationCount);
                    row.revalidationAmount += Convert.ToDecimal(item.revalidationAmount);
                    row.totalCharges += Convert.ToDecimal(item.totalCharges);
                    row.discount += Convert.ToDecimal(item.discount);
                    row.totalPaid += Convert.ToDecimal(item.totalPaid);
                    row.vipCount += Convert.ToInt32(item.vipCount);
                }
            }
            return rows;
        }

        private async Task<List<YearlyTransRow>> BuildYearlyTransactions(int year)
        {
            var rows = new List<YearlyTransRow>();
            for (int m = 1; m <= 12; m++)
            {
                var row = new YearlyTransRow
                {
                    month = m,
                    label = new DateTime(year, m, 1).ToString("MMMM yyyy", CultureInfo.InvariantCulture)
                };
                rows.Add(row);

                if (IsFutureMonth(year, m)) continue;

                var data = await _collectionService.GetMonthlyTransactionsDataWithVehicles(year, m);
                if (data == null) continue;

                foreach (dynamic item in data)
                {
                    row.totalVehicles += Convert.ToInt32(item.totalVehicles);
                    row.totalTransactions += Convert.ToInt32(item.totalTransactions);
                    row.cashCount += Convert.ToInt32(item.cashCount);
                    row.cashTotal += Convert.ToDecimal(item.cashTotal);
                    row.mPesaCount += Convert.ToInt32(item.mPesaCount);
                    row.mPesaTotal += Convert.ToDecimal(item.mPesaTotal);
                    row.cashTopUpCount += Convert.ToInt32(item.cashTopUpCount);
                    row.cashTopUpTotal += Convert.ToDecimal(item.cashTopUpTotal);
                    row.mPesaTopUpCount += Convert.ToInt32(item.mPesaTopUpCount);
                    row.mPesaTopUpTotal += Convert.ToDecimal(item.mPesaTopUpTotal);
                    row.focCount += Convert.ToInt32(item.focCount);
                    row.focAmount += Convert.ToDecimal(item.focAmount);
                    row.ctmFocCount += Convert.ToInt32(item.ctmFocCount);
                    row.ctmFocAmount += Convert.ToDecimal(item.ctmFocAmount);
                    row.totalChargedAmount += Convert.ToDecimal(item.totalChargedAmount);
                }

                row.totalCollected = row.cashTotal + row.mPesaTotal + row.cashTopUpTotal + row.mPesaTopUpTotal;
            }
            return rows;
        }

        private static void WriteYearlySheet(XLWorkbook workbook, string sheetName, string title,
            string[] headers, List<object[]> data, int[] amountCols)
        {
            var ws = workbook.Worksheets.Add(sheetName);

            ws.Cell(1, 1).Value = title;
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 14;
            ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
            ws.Range(1, 1, 1, headers.Length).Merge();

            for (int c = 0; c < headers.Length; c++)
            {
                var cell = ws.Cell(2, c + 1);
                cell.Value = headers[c];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.Gray;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Alignment.WrapText = true;
                ws.Column(c + 1).Width = 13;
            }
            ws.Column(1).Width = 18;
            ws.Row(2).Height = 32;

            int r = 3;
            foreach (var vals in data)
            {
                for (int c = 0; c < vals.Length; c++)
                    ws.Cell(r, c + 1).Value = XLCellValue.FromObject(vals[c]);
                r++;
            }

            ws.Cell(r, 1).Value = "TOTAL";
            for (int c = 2; c <= headers.Length; c++)
            {
                string col = ws.Cell(3, c).Address.ColumnLetter;
                ws.Cell(r, c).FormulaA1 = $"=SUM({col}3:{col}{r - 1})";
            }

            foreach (var c in amountCols)
                ws.Range(3, c, r, c).Style.NumberFormat.Format = "#,##0.00";

            var totalRange = ws.Range(r, 1, r, headers.Length);
            totalRange.Style.Font.Bold = true;
            totalRange.Style.Fill.BackgroundColor = XLColor.Gold;

            var all = ws.Range(2, 1, r, headers.Length);
            all.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            all.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.SheetView.FreezeRows(2);
            ws.SheetView.FreezeColumns(1);
        }

        public IActionResult Yearly()
        {
            return View();
        }

        [HttpGet]
        public async Task<IActionResult> GetYearlyExitedCarsData(int year)
        {
            try
            {
                var rows = await BuildYearlyExited(year);
                return Json(new { success = true, data = rows });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetYearlyExitedCarsData: {ex.Message}\n{ex.StackTrace}");
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public async Task<IActionResult> GetYearlyTransactionsData(int year)
        {
            try
            {
                var rows = await BuildYearlyTransactions(year);
                return Json(new { success = true, data = rows });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in GetYearlyTransactionsData: {ex.Message}\n{ex.StackTrace}");
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public async Task<IActionResult> ExportYearlyExcel(int year)
        {
            try
            {
                var rows = await BuildYearlyExited(year);

                string[] headers = {
                    "Month", "Vehicles", "Cash Count", "Cash (KES)", "M-Pesa Count", "M-Pesa (KES)",
                    "FOC Count", "FOC Amount (KES)", "HHD-4 FOC Count", "HHD-4 FOC Amount (KES)",
                    "Zero Charge Count", "Permanent Parkers", "Revalidation Count", "Revalidation Amt (KES)",
                    "Total Charges (KES)", "Discount (KES)", "Total Paid (KES)", "VIP"
                };

                var data = rows.Select(r => new object[]
                {
                    r.label, r.totalVehicles, r.cashCount, r.cashTotal, r.mPesaCount, r.mPesaTotal,
                    r.focCount, r.focAmount, r.ctmFocCount, r.ctmFocAmount,
                    r.zeroChargeCount, r.permanentParkersCount, r.revalidationCount, r.revalidationAmount,
                    r.totalCharges, r.discount, r.totalPaid, r.vipCount
                }).ToList();

                using var workbook = new XLWorkbook();
                WriteYearlySheet(workbook, $"Yearly {year}", $"YEARLY EXITED CARS SUMMARY - {year}",
                    headers, data, new[] { 4, 6, 8, 10, 14, 15, 16, 17 });

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"YEARLY_COLLECTION_{year}_{DateTime.Now:HH-mm-ss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ExportYearlyExcel: {ex.Message}\n{ex.StackTrace}");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        [HttpGet]
        public async Task<IActionResult> ExportYearlyTransactionsExcel(int year)
        {
            try
            {
                var rows = await BuildYearlyTransactions(year);

                string[] headers = {
                    "Month", "Vehicles", "Transactions", "Cash Count", "Cash (KES)", "M-Pesa Count", "M-Pesa (KES)",
                    "Cash Count (TOP-UP)", "Cash (KES) (TOP-UP)", "M-Pesa Count (TOP-UP)", "M-Pesa (KES) (TOP-UP)",
                    "FOC Count", "FOC Amount (KES)", "HHD-4 FOC Count", "HHD-4 FOC Amount (KES)",
                    "Total Charged (KES)", "Total Collected (KES)"
                };

                var data = rows.Select(r => new object[]
                {
                    r.label, r.totalVehicles, r.totalTransactions, r.cashCount, r.cashTotal, r.mPesaCount, r.mPesaTotal,
                    r.cashTopUpCount, r.cashTopUpTotal, r.mPesaTopUpCount, r.mPesaTopUpTotal,
                    r.focCount, r.focAmount, r.ctmFocCount, r.ctmFocAmount,
                    r.totalChargedAmount, r.totalCollected
                }).ToList();

                using var workbook = new XLWorkbook();
                WriteYearlySheet(workbook, $"Yearly {year}", $"YEARLY TRANSACTIONS SUMMARY - {year}",
                    headers, data, new[] { 5, 7, 9, 11, 13, 15, 16, 17 });

                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return File(stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"YEARLY_TRANSACTIONS_{year}_{DateTime.Now:HH-mm-ss}.xlsx");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ExportYearlyTransactionsExcel: {ex.Message}\n{ex.StackTrace}");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        // Helper method to get property value from anonymous/dynamic objects
        private object GetPropertyValue(object obj, string propertyName)
        {
            if (obj == null) return 0;

            var type = obj.GetType();
            var property = type.GetProperty(propertyName);

            if (property != null)
            {
                return property.GetValue(obj) ?? 0;
            }

            // Try to get from dictionary if it's a dictionary
            if (obj is IDictionary<string, object> dict && dict.ContainsKey(propertyName))
            {
                return dict[propertyName] ?? 0;
            }

            return 0;
        }

        // ============================================
        // TRANSACTION METHODS - ✅ UPDATED FOR DATE RANGE
        // ============================================
        [HttpPost]
        public async Task<IActionResult> GetTransactionData([FromBody] CollectionFilterRequest request)
        {
            try
            {
                var data = await _collectionService.GetTransactionSummary(request.StartDate, request.EndDate);
                var summary = await _collectionService.GetTransactionPaymentMethodSummary(request.StartDate, request.EndDate);

                // SIMPLIFIED: Don't include complex statistics in API response
                // Statistics are only used in Excel export which handles them internally
                return Json(new
                {
                    success = true,
                    data,
                    summary,
                    recordCount = data.Count
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error in GetTransactionData: {ex.Message}", ex);
            }
        }

        [HttpGet]
        public async Task<IActionResult> ExportTransactionExcel(string startDate, string endDate)
        {
            try
            {
                if (!DateTime.TryParse(startDate, out var start))
                    return BadRequest("Invalid start date format");

                if (!DateTime.TryParse(endDate, out var end))
                    return BadRequest("Invalid end date format");

                var data = await _collectionService.GetTransactionSummary(start, end);
                var statsRaw = await _collectionService.GetTransactionStatistics(start, end);

                using var workbook = new XLWorkbook();
                workbook.Properties.Author = "APMS System";
                workbook.Properties.Title = $"Transaction Report - {start:dd-MMM-yyyy} to {end:dd-MMM-yyyy}";

                var reportSheet = workbook.Worksheets.Add("Report");

                // ============================================
                // REPORT HEADERS (19 COLUMNS) - UPDATED: Added Paid Amount
                // ============================================
                reportSheet.Cell(1, 1).Value = "#";
                reportSheet.Cell(1, 2).Value = "Transaction ID";
                reportSheet.Cell(1, 3).Value = "Card Number";
                reportSheet.Cell(1, 4).Value = "Card Name";
                reportSheet.Cell(1, 5).Value = "Vehicle Plate";
                reportSheet.Cell(1, 6).Value = "Transaction Type";

                // Entry Timestamp (Split)
                reportSheet.Cell(1, 7).Value = "Entry Date";
                reportSheet.Cell(1, 8).Value = "Entry Time";

                // Validation Timestamp + PayMode + Amount (Split)
                reportSheet.Cell(1, 9).Value = "Validation Date";
                reportSheet.Cell(1, 10).Value = "Validation Time";
                reportSheet.Cell(1, 11).Value = "Validation Pay Mode";
                reportSheet.Cell(1, 12).Value = "Validation Amount";

                // Revalidation Timestamp + PayMode + Amount (Split)
                reportSheet.Cell(1, 13).Value = "Revalidation Date";
                reportSheet.Cell(1, 14).Value = "Revalidation Time";
                reportSheet.Cell(1, 15).Value = "Revalidation Pay Mode";
                reportSheet.Cell(1, 16).Value = "Revalidation Amount";

                // Final Columns - UPDATED: Added Paid Amount (Column 19)
                reportSheet.Cell(1, 17).Value = "Payment Source";
                reportSheet.Cell(1, 18).Value = "Paid Amount";      // ✅ NEW: Paid Amount
                reportSheet.Cell(1, 19).Value = "Status";

                // Style header
                var headerRange = reportSheet.Range(1, 1, 1, 19);  // ✅ Changed to 19 columns
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                reportSheet.SheetView.FreezeRows(1);

                // ============================================
                // ADD DATA ROWS
                // ============================================
                int row = 2;
                int counter = 1;

                foreach (var record in data)
                {
                    try
                    {
                        reportSheet.Cell(row, 1).Value = counter;
                        reportSheet.Cell(row, 2).Value = record.transaction_id ?? "";
                        reportSheet.Cell(row, 3).Value = record.cardno ?? "";
                        reportSheet.Cell(row, 4).Value = record.cardname ?? "";
                        reportSheet.Cell(row, 5).Value = record.car_plate_no ?? "";
                        reportSheet.Cell(row, 6).Value = record.transaction_type ?? "";

                        // Entry Date/Time (Split)
                        if (record.entry_time.HasValue)
                        {
                            reportSheet.Cell(row, 7).Value = record.entry_time.Value.ToString("dd-MMM-yyyy").ToUpper();
                            reportSheet.Cell(row, 8).Value = record.entry_time.Value.ToString("HH:mm:ss");
                        }
                        else
                        {
                            reportSheet.Cell(row, 7).Value = "--";
                            reportSheet.Cell(row, 8).Value = "--";
                        }

                        // Validation Date/Time (Split)
                        if (record.validation_time.HasValue)
                        {
                            reportSheet.Cell(row, 9).Value = record.validation_time.Value.ToString("dd-MMM-yyyy").ToUpper();
                            reportSheet.Cell(row, 10).Value = record.validation_time.Value.ToString("HH:mm:ss");
                        }
                        else
                        {
                            reportSheet.Cell(row, 9).Value = "--";
                            reportSheet.Cell(row, 10).Value = "--";
                        }

                        // Validation PayMode + Amount
                        reportSheet.Cell(row, 11).Value = record.validation_paymode ?? "";
                        decimal validationAmount = record.validation_charges ?? 0m;
                        reportSheet.Cell(row, 12).Value = validationAmount;
                        reportSheet.Cell(row, 12).Style.NumberFormat.Format = "#,##0.00";

                        // Revalidation Date/Time (Split)
                        if (record.revalidation_time.HasValue)
                        {
                            reportSheet.Cell(row, 13).Value = record.revalidation_time.Value.ToString("dd-MMM-yyyy").ToUpper();
                            reportSheet.Cell(row, 14).Value = record.revalidation_time.Value.ToString("HH:mm:ss");
                        }
                        else
                        {
                            reportSheet.Cell(row, 13).Value = "--";
                            reportSheet.Cell(row, 14).Value = "--";
                        }

                        // Revalidation PayMode + Amount
                        reportSheet.Cell(row, 15).Value = record.revalidation_paymode ?? "";
                        decimal revalidationAmount = record.revalidation_charges ?? 0m;
                        reportSheet.Cell(row, 16).Value = revalidationAmount;
                        reportSheet.Cell(row, 16).Style.NumberFormat.Format = "#,##0.00";

                        // Payment Source + ✅ Paid Amount + Status
                        reportSheet.Cell(row, 17).Value = record.payment_source ?? "";

                        // ✅ NEW: Add paid_amount from database
                        decimal paidAmount = record.paidamount ?? 0m;
                        reportSheet.Cell(row, 18).Value = paidAmount;
                        reportSheet.Cell(row, 18).Style.NumberFormat.Format = "#,##0.00";

                        // Status
                        reportSheet.Cell(row, 19).Value = record.status?.ToString() ?? "";

                        row++;
                        counter++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Error processing transaction row {row}: {ex.Message}");
                        row++;
                    }
                }

                // Adjust columns
                reportSheet.Columns().AdjustToContents();
                reportSheet.Column(7).Width = 15;  // Entry Date
                reportSheet.Column(8).Width = 12;  // Entry Time
                reportSheet.Column(9).Width = 15;  // Validation Date
                reportSheet.Column(10).Width = 12; // Validation Time
                reportSheet.Column(13).Width = 15; // Revalidation Date
                reportSheet.Column(14).Width = 12; // Revalidation Time
                reportSheet.Column(18).Width = 15; // ✅ Paid Amount

                // ============================================
                // SUMMARY STATISTICS - SAME SHEET
                // ============================================
                int statsRow = row + 2;  // 2 blank rows after data

                try
                {
                    dynamic statisticsData = statsRaw;
                    var devices = new[] { "Server", "Paystation1", "Paystation2", "Paystation3", "Paystation4", "HHD1", "HHD2", "HHD3", "HHD4", "Web" };

                    // SUMMARY STATISTICS HEADER
                    reportSheet.Cell(statsRow, 1).Value = "SUMMARY STATISTICS";
                    reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 1).Style.Font.FontSize = 12;
                    reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.LightBlue;
                    statsRow += 2;

                    // PER-DEVICE BREAKDOWN
                    foreach (var device in devices)
                    {
                        try
                        {
                            var devicesProperty = statisticsData.devices;
                            var deviceData = GetPropertyValue(devicesProperty, device);

                            if (deviceData == null)
                                continue;

                            // Cash Count
                            reportSheet.Cell(statsRow, 1).Value = $"{device} Cash Count:";
                            reportSheet.Cell(statsRow, 2).Value = (int)GetPropertyValue(deviceData, "cashCount");
                            statsRow++;

                            // Cash Total
                            reportSheet.Cell(statsRow, 1).Value = $"{device} Cash Total:";
                            reportSheet.Cell(statsRow, 2).Value = (decimal)GetPropertyValue(deviceData, "cashTotal");
                            reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                            statsRow++;

                            // M-Pesa Count
                            reportSheet.Cell(statsRow, 1).Value = $"{device} M-Pesa Count:";
                            reportSheet.Cell(statsRow, 2).Value = (int)GetPropertyValue(deviceData, "mPesaCount");
                            statsRow++;

                            // M-Pesa Total
                            reportSheet.Cell(statsRow, 1).Value = $"{device} M-Pesa Total:";
                            reportSheet.Cell(statsRow, 2).Value = (decimal)GetPropertyValue(deviceData, "mPesaTotal");
                            reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                            statsRow++;

                            // M-Pesa Via Paybill Count
                            reportSheet.Cell(statsRow, 1).Value = $"{device} M-Pesa Via Paybill Count:";
                            reportSheet.Cell(statsRow, 2).Value = (int)GetPropertyValue(deviceData, "mPesaViaPaybillCount");
                            statsRow++;

                            // M-Pesa Via Paybill Total
                            reportSheet.Cell(statsRow, 1).Value = $"{device} M-Pesa Via Paybill Total:";
                            reportSheet.Cell(statsRow, 2).Value = (decimal)GetPropertyValue(deviceData, "mPesaViaPaybillTotal");
                            reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                            statsRow++;

                            // FOC Count
                            reportSheet.Cell(statsRow, 1).Value = $"{device} FOC Count:";
                            reportSheet.Cell(statsRow, 2).Value = (int)GetPropertyValue(deviceData, "focCount");
                            reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            statsRow++;

                            // FOC Amount
                            reportSheet.Cell(statsRow, 1).Value = $"{device} FOC Amount:";
                            reportSheet.Cell(statsRow, 2).Value = (decimal)GetPropertyValue(deviceData, "focTotal");
                            reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                            reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            statsRow++;

                        }
                        catch (Exception devEx)
                        {
                            _logger.LogError($"Error processing device {device}: {devEx.Message}");
                            continue;
                        }
                    }

                    statsRow++;

                    // ✅ GRAND TOTALS (YELLOW BACKGROUND)
                    reportSheet.Cell(statsRow, 1).Value = "GRAND TOTAL Count:";
                    reportSheet.Cell(statsRow, 2).Value = (int)statisticsData.grandTotal;
                    reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 2).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
                    statsRow++;

                    // Total Cash Collected
                    reportSheet.Cell(statsRow, 1).Value = "TOTAL CASH Collected:";
                    reportSheet.Cell(statsRow, 2).Value = (decimal)statisticsData.totalCashCollected;
                    reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                    reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 2).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
                    statsRow++;

                    // Total M-Pesa Collected (DIRECT APP ONLY - NOT Paybill)
                    reportSheet.Cell(statsRow, 1).Value = "TOTAL M-Pesa Collected:";
                    reportSheet.Cell(statsRow, 2).Value = (decimal)statisticsData.totalMPesaCollected;
                    reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                    reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 2).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
                    statsRow++;

                    // Total M-Pesa Via Paybill Collected (SEPARATE from M-Pesa)
                    reportSheet.Cell(statsRow, 1).Value = "TOTAL M-Pesa Via Paybill Collected:";
                    reportSheet.Cell(statsRow, 2).Value = (decimal)statisticsData.totalMPesaViaPaybillCollected;
                    reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                    reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 2).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
                    statsRow++;

                    // Total FOC Amount
                    reportSheet.Cell(statsRow, 1).Value = "TOTAL FOC Amount:";
                    reportSheet.Cell(statsRow, 2).Value = (decimal)statisticsData.totalFocAmount;
                    reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                    reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 2).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
                    statsRow++;

                    // Grand Total Amount (All collections)
                    decimal grandTotalAmount = (decimal)statisticsData.totalCashCollected + (decimal)statisticsData.totalMPesaCollected + (decimal)statisticsData.totalMPesaViaPaybillCollected;
                    reportSheet.Cell(statsRow, 1).Value = "TOTAL Amount Collected:";
                    reportSheet.Cell(statsRow, 2).Value = grandTotalAmount;
                    reportSheet.Cell(statsRow, 2).Style.NumberFormat.Format = "#,##0.00";
                    reportSheet.Cell(statsRow, 1).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 2).Style.Font.Bold = true;
                    reportSheet.Cell(statsRow, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    reportSheet.Cell(statsRow, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
                    statsRow++;
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error processing statistics: {ex.Message}");
                    reportSheet.Cell(statsRow, 1).Value = $"Error: {ex.Message}";
                }

                // Adjust column widths
                reportSheet.Column(1).Width = 40;
                reportSheet.Column(2).Width = 20;

                // Save and return
                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                var fileName = $"TRANSACTION_REPORT_{start:dd-MMM-yyyy}_to_{end:dd-MMM-yyyy}_{DateTime.Now:HH-mm-ss}.xlsx";

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ExportTransactionExcel: {ex.Message}\n{ex.StackTrace}");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        // ============================================
        // EMAIL REPORT SERVICE METHODS - SIMPLE WRAPPERS (ADD AT END OF CLASS)
        // ============================================

        /// <summary>
        /// Generate Daily Exited Cars Excel - Returns byte[] for email attachment
        /// Used by: EmailReportService.SendDailyExitedCarsReportAsync()
        /// </summary>
        public async Task<byte[]> GenerateDailyExitedCarsExcelBytes(DateTime startDate, DateTime endDate)
        {
            // ✅ Call existing ExportDailyExcel method
            var result = await ExportDailyExcel(startDate.ToString("yyyy-MM-dd"), endDate.ToString("yyyy-MM-dd"));

            if (result is FileContentResult fileResult)
            {
                return fileResult.FileContents;
            }

            throw new Exception("Failed to generate daily exited cars Excel");
        }

        /// <summary>
        /// Generate Daily Transactions Excel - Returns byte[] for email attachment
        /// Used by: EmailReportService.SendDailyTransactionsReportAsync()
        /// </summary>
        public async Task<byte[]> GenerateDailyTransactionsExcelBytes(DateTime startDate, DateTime endDate)
        {
            // ✅ Call existing ExportTransactionExcel method
            var result = await ExportTransactionExcel(startDate.ToString("yyyy-MM-dd"), endDate.ToString("yyyy-MM-dd"));

            if (result is FileContentResult fileResult)
            {
                return fileResult.FileContents;
            }

            throw new Exception("Failed to generate daily transactions Excel");
        }

        /// <summary>
        /// Generate Monthly Exited Cars Excel - Returns byte[] for email attachment
        /// Used by: EmailReportService.SendMonthlyExitedCarsReportAsync()
        /// </summary>
        public async Task<byte[]> GenerateMonthlyExitedCarsExcelBytes(int year, int month)
        {
            // ✅ Call existing ExportMonthlyExcel method
            var result = await ExportMonthlyExcel(month, year);

            if (result is FileContentResult fileResult)
            {
                return fileResult.FileContents;
            }

            throw new Exception("Failed to generate monthly exited cars Excel");
        }

        /// <summary>
        /// Generate Monthly Transactions Excel - Returns byte[] for email attachment0
        /// Used by: EmailReportService.SendMonthlyTransactionsReportAsync()
        /// </summary>
        public async Task<byte[]> GenerateMonthlyTransactionsExcelBytes(int year, int month)
        {
            // ✅ Call existing ExportMonthlyTransactionsExcel method
            var result = await ExportMonthlyTransactionsExcel(month, year);

            if (result is FileContentResult fileResult)
            {
                return fileResult.FileContents;
            }

            throw new Exception("Failed to generate monthly transactions Excel");
        }
    }
}