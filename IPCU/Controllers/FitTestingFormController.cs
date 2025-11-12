using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using IPCU.Data;
using IPCU.Models;
using X.PagedList;
using X.PagedList.Extensions;
using ClosedXML.Excel;
using System.IO;
using System.Data.SqlClient;

namespace IPCU.Controllers
{
    public class FitTestingFormController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _configuration;

        public FitTestingFormController(ApplicationDbContext context, IConfiguration configuration)
        {
            _context = context;
            _configuration = configuration;
        }

        // GET: FitTestingForm
        public async Task<IActionResult> Index(int? page, bool? filterExpiring, string testResult, string searchTerm)
        {
            int pageSize = 20;
            int pageNumber = page ?? 1;
            var fitTestingForm = _context.FitTestingForm.AsQueryable();

            if (filterExpiring == true)
            {
                DateTime today = DateTime.Today;
                DateTime thresholdDate = today.AddDays(30);
                fitTestingForm = fitTestingForm
                    .Where(f => f.ExpiringAt >= today && f.ExpiringAt <= thresholdDate);
            }

            if (!string.IsNullOrEmpty(testResult))
            {
                fitTestingForm = fitTestingForm.Where(f => f.Test_Results == testResult);
            }

            if (!string.IsNullOrEmpty(searchTerm))
            {
                fitTestingForm = fitTestingForm.Where(f => f.HCW_Name.Contains(searchTerm) || f.DUO.Contains(searchTerm));
            }

            fitTestingForm = fitTestingForm.OrderByDescending(f => f.ExpiringAt);
            var pagedList = fitTestingForm.ToPagedList(pageNumber, pageSize);

            ViewData["FilterExpiring"] = filterExpiring;
            ViewData["SelectedTestResult"] = testResult;
            ViewBag.SearchTerm = searchTerm;

            return View(pagedList);
        }

        // GET: FitTestingForm/Details/5
        public IActionResult Details(int id)
        {
            var fitTestingForm = _context.FitTestingForm.FirstOrDefault(f => f.Id == id);
            if (fitTestingForm == null) return NotFound();

            var history = _context.FitTestingFormHistory
                .Where(h => h.FitTestingFormId == id)
                .OrderBy(h => h.SubmittedAt)
                .ToList();

            ViewData["FirstAttempt"] = history.ElementAtOrDefault(0);
            ViewData["SecondAttempt"] = history.ElementAtOrDefault(1);
            ViewData["LastAttempt"] = history.Count > 2 ? history.LastOrDefault() : null;

            return View(fitTestingForm);
        }

        // GET: FitTestingForm/Create
        public IActionResult Create()
        {
            ViewBag.DUO_Tester = new SelectList(new List<string>
            {
                "Unit 2A", "Unit 2B", "Unit 2C", "Unit 2D", "Unit 2E/Ext",
                "Unit 2F/2G", "Unit 2H", "Unit 3A", "Unit 3B", "Unit 3C", "Unit 3D/Celtran",
                "Unit 3E", "Unit 3F", "ICU", "ER", "PD", "HDU", "AEUC", "ORU", "CCRU",
                "iVASC", "OPS", "AITU", "PCU", "IPCU"
            });

            return View();
        }

        // POST: FitTestingForm/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(FitTestingForm fitTestingForm, string? OtherLimitation)
        {
            ModelState.Remove("OtherLimitation");

            try
            {
                var limitations = Request.Form["Limitation"].ToList();

                if (limitations != null && limitations.Any())
                {
                    if (limitations.Contains("Other"))
                    {
                        if (string.IsNullOrWhiteSpace(OtherLimitation))
                        {
                            ModelState.AddModelError("Limitation", "Please specify the other limitation");
                            ViewBag.DUO_Tester = new SelectList(new List<string> { /* same list */ });
                            return View(fitTestingForm);
                        }
                        limitations.Remove("Other");
                        limitations.Add(OtherLimitation.Trim());
                    }
                    fitTestingForm.Limitation = string.Join(", ", limitations.Where(x => !string.IsNullOrWhiteSpace(x)));
                }
                else
                {
                    fitTestingForm.Limitation = "None";
                }

                if (ModelState.IsValid)
                {
                    fitTestingForm.SubmittedAt = DateTime.Now;
                    fitTestingForm.ExpiringAt = fitTestingForm.SubmittedAt.AddYears(1);

                    _context.Add(fitTestingForm);
                    await _context.SaveChangesAsync();

                    var history = new FitTestingFormHistory
                    {
                        FitTestingFormId = fitTestingForm.Id,
                        Fit_Test_Solution = fitTestingForm.Fit_Test_Solution,
                        Sensitivity_Test = fitTestingForm.Sensitivity_Test,
                        Respiratory_Type = fitTestingForm.Respiratory_Type,
                        Model = fitTestingForm.Model,
                        Size = fitTestingForm.Size,
                        Normal_Breathing = fitTestingForm.Normal_Breathing,
                        Deep_Breathing = fitTestingForm.Deep_Breathing,
                        Turn_head_side_to_side = fitTestingForm.Turn_head_side_to_side,
                        Move_head_up_and_down = fitTestingForm.Move_head_up_and_down,
                        Reading = fitTestingForm.Reading,
                        Bending_Jogging = fitTestingForm.Bending_Jogging,
                        Normal_Breathing_2 = fitTestingForm.Normal_Breathing_2,
                        Test_Results = fitTestingForm.Test_Results,
                        SubmittedAt = fitTestingForm.SubmittedAt
                    };

                    _context.FitTestingFormHistory.Add(history);
                    await _context.SaveChangesAsync();

                    return RedirectToAction(nameof(Index));
                }
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("", "Error: " + ex.Message);
            }

            ViewBag.DUO_Tester = new SelectList(new List<string>
            {
                "Unit 2A", "Unit 2B", "Unit 2C", "Unit 2D", "Unit 2E/Ext",
                "Unit 2F/2G", "Unit 2H", "Unit 3A", "Unit 3B", "Unit 3C", "Unit 3D/Celtran",
                "Unit 3E", "Unit 3F", "ICU", "ER", "PD", "HDU", "AEUC", "ORU", "CCRU",
                "iVASC", "OPS", "AITU", "PCU", "IPCU"
            });

            return View(fitTestingForm);
        }

        // GET: FitTestingForm/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var fitTestingForm = await _context.FitTestingForm.FindAsync(id);
            if (fitTestingForm == null)
            {
                return NotFound();
            }
            return View(fitTestingForm);
        }

        // POST: FitTestingForm/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("Id,HCW_Name,DUO,Limitation,Fit_Test_Solution,Sensitivity_Test,Respiratory_Type,Model,Size,Normal_Breathing,Deep_Breathing,Turn_head_side_to_side,Move_head_up_and_down,Reading,Bending_Jogging,Normal_Breathing_2,Test_Results,Name_of_Fit_Tester,DUO_Tester,SubmittedAt")] FitTestingForm fitTestingForm)
        {
            if (id != fitTestingForm.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(fitTestingForm);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!FitTestingFormExists(fitTestingForm.Id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(fitTestingForm);
        }

        // GET: FitTestingForm/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var fitTestingForm = await _context.FitTestingForm
                .FirstOrDefaultAsync(m => m.Id == id);
            if (fitTestingForm == null)
            {
                return NotFound();
            }

            return View(fitTestingForm);
        }

        // POST: FitTestingForm/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var fitTestingForm = await _context.FitTestingForm.FindAsync(id);
            if (fitTestingForm != null)
            {
                _context.FitTestingForm.Remove(fitTestingForm);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool FitTestingFormExists(int id)
        {
            return _context.FitTestingForm.Any(e => e.Id == id);
        }

        [HttpGet("GeneratePdf/{id}")]
        public IActionResult GeneratePdf(int id)
        {
            var form = _context.FitTestingForm.FirstOrDefault(f => f.Id == id);
            if (form == null) return NotFound();

            var pdfService = new FitTestingFormPdfService(_context);
            var pdfBytes = pdfService.GeneratePdf(form);

            return File(pdfBytes, "application/pdf"); // This ensures the browser previews it properly
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult SubmitFitTest(int id, FitTestingForm updatedForm)
        {
            var fitTest = _context.FitTestingForm.FirstOrDefault(f => f.Id == id);
            if (fitTest != null && fitTest.SubmissionCount < fitTest.MaxRetakes)
            {
                // Update the main form with the new data FIRST
                fitTest.Fit_Test_Solution = updatedForm.Fit_Test_Solution;
                fitTest.Sensitivity_Test = updatedForm.Sensitivity_Test;
                fitTest.Respiratory_Type = updatedForm.Respiratory_Type;
                fitTest.Model = updatedForm.Model;
                fitTest.Size = updatedForm.Size;
                fitTest.Normal_Breathing = updatedForm.Normal_Breathing;
                fitTest.Deep_Breathing = updatedForm.Deep_Breathing;
                fitTest.Turn_head_side_to_side = updatedForm.Turn_head_side_to_side;
                fitTest.Move_head_up_and_down = updatedForm.Move_head_up_and_down;
                fitTest.Reading = updatedForm.Reading;
                fitTest.Bending_Jogging = updatedForm.Bending_Jogging;
                fitTest.Normal_Breathing_2 = updatedForm.Normal_Breathing_2;

                // Update the submission count and submission date
                fitTest.SubmissionCount++;
                fitTest.SubmittedAt = DateTime.Now; // Update the submission date for the main form

                // Save the updated FitTestingForm to the database
                _context.SaveChanges(); // Save the updated main form

                // NOW, save the current state to FitTestingFormHistory
                var history = new FitTestingFormHistory
                {
                    FitTestingFormId = fitTest.Id,
                    Fit_Test_Solution = fitTest.Fit_Test_Solution, // Use the updated data
                    Sensitivity_Test = fitTest.Sensitivity_Test,
                    Respiratory_Type = fitTest.Respiratory_Type,
                    Model = fitTest.Model,
                    Size = fitTest.Size,
                    Normal_Breathing = fitTest.Normal_Breathing,
                    Deep_Breathing = fitTest.Deep_Breathing,
                    Turn_head_side_to_side = fitTest.Turn_head_side_to_side,
                    Move_head_up_and_down = fitTest.Move_head_up_and_down,
                    Reading = fitTest.Reading,
                    Bending_Jogging = fitTest.Bending_Jogging,
                    Normal_Breathing_2 = fitTest.Normal_Breathing_2,
                    Test_Results = fitTest.Test_Results,
                    SubmittedAt = fitTest.SubmittedAt // Use the updated submission date
                };

                // Add the history entry to the database
                _context.FitTestingFormHistory.Add(history);
                _context.SaveChanges(); // Save history entry
            }

            return RedirectToAction("Details", new { id });
        }

        public IActionResult Reports()
        {
            var physicianCategories = new List<string>
    {
        "Consultant - Plantilla", "Physician - Plantilla", "Resident - Plantilla",
        "Fellows - Plantilla", "Consultant-Active Non-Plantilla", "Resident - Plantilla Second Year",
        "Resident - Plantilla First Year", "Resident - Third Year - Plantilla",
        "Resident - Second Year - Plantilla", "Resident - First Year - Plantilla",
        "Plantilla - MS II", "Plantilla - DM III", "Plantilla - MS I",
        "Plantilla - DED IV", "Plantilla - MS III", "Fellow Plantilla - 1st Year",
        "Fellow Plantilla - 2nd Year", "Active Non-Plantilla", "Fellow - 3rd Year",
        "Fellow - 2nd Year", "Fellow - 1st Year", "Medical Officer III", "Fellow",
        "Resident", "Consultants - Plantilla", "Visiting Consultant", "Consultant",
        "Consultant - Non-Plantilla", "MO III"
    };

            var duoList = new List<string>
    {
        "Unit 2A", "Unit 2B", "Unit 2C", "Unit 2D", "Unit 2E/Ext",
        "Unit 2F/2G", "Unit 2H", "Unit 3A", "Unit 3B", "Unit 3C", "Unit 3D/Celtran",
        "Unit 3E", "Unit 3F", "ICU", "ER", "PD", "HDU", "AEUC", "ORU", "CCRU",
        "iVASC", "OPS", "AITU", "PCU", "IPCU"
    };

            var duoMedical = new List<string>
    {
        "Adult Nephrology", "Surgery", "OTVS", "Pedia Nephrology", "Urology", "IM", "Anesthesiology"
    };

            var duoAllied = new List<string>
    {
        "Cardiology", "HOPE", "Nuclear", "Radiology/DMITRI", "PMRS", "PLMD", "Pulmonology"
    };

            var fitTests = _context.FitTestingForm.ToList();
            var currentDate = DateTime.Now;

            // ------------------------------
            // Attendance: Physicians
            // ------------------------------
            var attendanceForPhysicians = fitTests
                .Where(f => physicianCategories.Contains(f.Professional_Category) && f.Test_Results == "Passed")
                .ToList();

            // ------------------------------
            // Attendance: Nursing and Allied
            // ------------------------------
            var attendanceForNursingAndAllied = fitTests
                .Where(f => !physicianCategories.Contains(f.Professional_Category) && f.Test_Results == "Passed")
                .Select(f => new FitTestingReportViewModel
                {
                    HCW_Name = f.HCW_Name,
                    DUO = f.DUO,
                    Professional_Category = f.Professional_Category,
                    Fit_Test_Solution = f.Fit_Test_Solution,
                    Test_Results = f.ExpiringAt < currentDate ? "Expired" : "Passed",
                    Name_of_Fit_Tester = f.Name_of_Fit_Tester,
                    SubmittedAt = f.SubmittedAt,
                    ExpiringAt = f.ExpiringAt
                })
                .ToList();

            // ------------------------------
            // DUO (Department) Summary - UPDATED WITH NEW PROPERTIES
            // ------------------------------
            var tallyReport = duoList
                .Select(unit => new
                {
                    Unit = unit,
                    TotalStaff = fitTests.Count(f => f.DUO_Tester == unit), // Total staff in the unit
                    TotalFitTested = fitTests.Count(f => f.DUO_Tester == unit && f.Test_Results == "Passed"),
                    Passed = fitTests.Count(f => f.DUO_Tester == unit && f.Test_Results == "Passed"),
                    Failed = fitTests.Count(f => f.DUO_Tester == unit && f.Test_Results == "Failed"),
                    NotDone = fitTests.Count(f => f.DUO_Tester == unit && f.Test_Results != "Passed" && f.Test_Results != "Failed"), // Not done = neither passed nor failed
                    Expired = fitTests.Count(f => f.DUO_Tester == unit && f.ExpiringAt < currentDate && f.Test_Results == "Passed"),
                    ThreeMAura = fitTests.Count(f => f.DUO_Tester == unit && f.Fit_Test_Solution != null && f.Fit_Test_Solution.Contains("3M")), // Count 3M Aura tests
                    GrandTotal = fitTests.Count(f => f.DUO_Tester == unit) // Grand total = all tests in the unit
                })
                .ToList();

            // ------------------------------
            // Per DUO_Tester Summary (NEW)
            // ------------------------------
            var testerSummary = fitTests
                .GroupBy(f => f.DUO_Tester)
                .Select(g => new
                {
                    Tester = g.Key,
                    Total = g.Count(),
                    Passed = g.Count(f => f.Test_Results == "Passed"),
                    Failed = g.Count(f => f.Test_Results == "Failed"),
                    Expired = g.Count(f => f.ExpiringAt < currentDate && f.Test_Results == "Passed")
                })
                .OrderBy(t => t.Tester)
                .ToList();

            // ------------------------------
            // Totals - UPDATED WITH NEW TOTALS
            // ------------------------------
            int totalFitTestedReport = tallyReport.Sum(t => t.TotalFitTested);
            int totalExpiredReport = tallyReport.Sum(t => t.Expired);
            int totalPassedReport = tallyReport.Sum(t => t.Passed);
            int totalFailedReport = tallyReport.Sum(t => t.Failed);
            int totalStaffReport = tallyReport.Sum(t => t.TotalStaff);
            int totalNotDoneReport = tallyReport.Sum(t => t.NotDone);
            int totalThreeMAuraReport = tallyReport.Sum(t => t.ThreeMAura);
            int totalGrandTotalReport = tallyReport.Sum(t => t.GrandTotal);

            // Grand totals
            int grandTotalFitTested = fitTests.Count(f => f.Test_Results == "Passed");
            int grandTotalPassed = fitTests.Count(f => f.Test_Results == "Passed");
            int grandTotalFailed = fitTests.Count(f => f.Test_Results == "Failed");
            int grandTotalExpired = fitTests.Count(f => f.ExpiringAt < currentDate && f.Test_Results == "Passed");

            // ------------------------------
            // Pass to view - UPDATED WITH NEW VIEWBAG PROPERTIES
            // ------------------------------
            ViewBag.AttendanceForPhysicians = attendanceForPhysicians;
            ViewBag.AttendanceForNursingAndAllied = attendanceForNursingAndAllied;
            ViewBag.TallyReport = tallyReport;
            ViewBag.TesterSummary = testerSummary;

            ViewBag.TotalFitTestedReport = totalFitTestedReport;
            ViewBag.TotalExpiredReport = totalExpiredReport;
            ViewBag.TotalPassedReport = totalPassedReport;
            ViewBag.TotalFailedReport = totalFailedReport;
            ViewBag.TotalStaffReport = totalStaffReport;
            ViewBag.TotalNotDoneReport = totalNotDoneReport;
            ViewBag.TotalThreeMAuraReport = totalThreeMAuraReport;
            ViewBag.TotalGrandTotalReport = totalGrandTotalReport;

            ViewBag.GrandTotalFitTested = grandTotalFitTested;
            ViewBag.GrandTotalPassed = grandTotalPassed;
            ViewBag.GrandTotalFailed = grandTotalFailed;
            ViewBag.GrandTotalExpired = grandTotalExpired;

            return View();
        }

        public IActionResult ExportToExcel()
        {
            var physicianCategories = new List<string>
    {
        "Consultant - Plantilla", "Physician - Plantilla", "Resident - Plantilla",
        "Fellows - Plantilla", "Consultant-Active Non-Plantilla", "Resident - Plantilla Second Year",
        "Resident - Plantilla First Year", "Resident - Third Year - Plantilla",
        "Resident - Second Year - Plantilla", "Resident - First Year - Plantilla",
        "Plantilla - MS II", "Plantilla - DM III", "Plantilla - MS I",
        "Plantilla - DED IV", "Plantilla - MS III", "Fellow Plantilla - 1st Year",
        "Fellow Plantilla - 2nd Year", "Active Non-Plantilla", "Fellow - 3rd Year",
        "Fellow - 2nd Year", "Fellow - 1st Year", "Medical Officer III", "Fellow",
        "Resident", "Consultants - Plantilla", "Visiting Consultant", "Consultant",
        "Consultant - Non-Plantilla", "MO III"
    };

            var duoList = new List<string>
    {
        "Unit 2A", "Unit 2B", "Unit 2C", "Unit 2D", "Unit 2E/Ext",
        "Unit 2F/2G", "Unit 2H", "Unit 3A", "Unit 3B", "Unit 3C", "Unit 3D/Celtran",
        "Unit 3E", "Unit 3F", "ICU", "ER", "PD", "HDU", "AEUC", "ORU", "CCRU",
        "iVASC", "OPS", "AITU", "PCU", "IPCU"
    };

            var fitTests = _context.FitTestingForm.ToList();
            var currentDate = DateTime.Now;

            var attendanceForPhysicians = fitTests
                .Where(f => physicianCategories.Contains(f.Professional_Category))
                .Select(f => new FitTestingForm
                {
                    HCW_Name = f.HCW_Name,
                    DUO = f.DUO,
                    Professional_Category = f.Professional_Category,
                    Fit_Test_Solution = f.Fit_Test_Solution,
                    Test_Results = f.ExpiringAt < currentDate ? "Expired" : f.Test_Results,
                    Name_of_Fit_Tester = f.Name_of_Fit_Tester,
                    SubmittedAt = f.SubmittedAt
                })
                .Where(f => f.Test_Results == "Passed" || f.Test_Results == "Expired")
                .ToList();

            var attendanceForNursingAndAllied = fitTests
                .Where(f => !physicianCategories.Contains(f.Professional_Category))
                .Select(f => new FitTestingForm
                {
                    HCW_Name = f.HCW_Name,
                    DUO = f.DUO,
                    Professional_Category = f.Professional_Category,
                    Fit_Test_Solution = f.Fit_Test_Solution,
                    Test_Results = f.ExpiringAt < currentDate ? "Expired" : f.Test_Results,
                    Name_of_Fit_Tester = f.Name_of_Fit_Tester,
                    SubmittedAt = f.SubmittedAt
                })
                .Where(f => f.Test_Results == "Passed" || f.Test_Results == "Expired")
                .ToList();

            var tallyReport = duoList
                .Select(unit => new
                {
                    Unit = unit,
                    TotalStaff = fitTests.Count(f => f.DUO_Tester == unit),
                    TotalFitTested = fitTests.Count(f => f.DUO_Tester == unit && f.Test_Results == "Passed"),
                    Passed = fitTests.Count(f => f.DUO_Tester == unit && f.Test_Results == "Passed"),
                    Failed = fitTests.Count(f => f.DUO_Tester == unit && f.Test_Results == "Failed"),
                    NotDone = fitTests.Count(f => f.DUO_Tester == unit && f.Test_Results != "Passed" && f.Test_Results != "Failed"),
                    Expired = fitTests.Count(f => f.DUO_Tester == unit && f.ExpiringAt < currentDate && f.Test_Results == "Passed"),
                    ThreeMAura = fitTests.Count(f => f.DUO_Tester == unit && f.Fit_Test_Solution != null && f.Fit_Test_Solution.Contains("3M")),
                    GrandTotal = fitTests.Count(f => f.DUO_Tester == unit)
                })
                .ToList();

            using (var workbook = new XLWorkbook())
            {
                void FormatHeaders(IXLWorksheet ws, int cols)
                {
                    ws.Row(1).Style.Font.Bold = true;
                    ws.Row(1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    ws.Columns(1, cols).AdjustToContents();
                }

                // Physicians Sheet
                var ws1 = workbook.Worksheets.Add("Physicians");
                ws1.Cell(1, 1).Value = "Name"; ws1.Cell(1, 2).Value = "Department/Unit/Office";
                ws1.Cell(1, 3).Value = "Professional Category"; ws1.Cell(1, 4).Value = "Fit Test Solution";
                ws1.Cell(1, 5).Value = "Status"; ws1.Cell(1, 6).Value = "Fit Tester"; ws1.Cell(1, 7).Value = "Date Fit Tested";
                FormatHeaders(ws1, 7);

                int row = 2;
                foreach (var p in attendanceForPhysicians)
                {
                    ws1.Cell(row, 1).Value = p.HCW_Name;
                    ws1.Cell(row, 2).Value = p.DUO;
                    ws1.Cell(row, 3).Value = p.Professional_Category;
                    ws1.Cell(row, 4).Value = p.Fit_Test_Solution;
                    ws1.Cell(row, 5).Value = p.Test_Results;
                    ws1.Cell(row, 6).Value = p.Name_of_Fit_Tester;
                    ws1.Cell(row, 7).Value = p.SubmittedAt.ToString("yyyy-MM-dd");
                    row++;
                }

                // Nursing & Allied Sheet
                var ws2 = workbook.Worksheets.Add("Nursing & Allied");
                ws2.Cell(1, 1).Value = "Name"; ws2.Cell(1, 2).Value = "Department/Unit/Office";
                ws2.Cell(1, 3).Value = "Professional Category"; ws2.Cell(1, 4).Value = "Fit Test Solution";
                ws2.Cell(1, 5).Value = "Status"; ws2.Cell(1, 6).Value = "Fit Tester"; ws2.Cell(1, 7).Value = "Date Fit Tested";
                FormatHeaders(ws2, 7);

                row = 2;
                foreach (var n in attendanceForNursingAndAllied)
                {
                    ws2.Cell(row, 1).Value = n.HCW_Name;
                    ws2.Cell(row, 2).Value = n.DUO;
                    ws2.Cell(row, 3).Value = n.Professional_Category;
                    ws2.Cell(row, 4).Value = n.Fit_Test_Solution;
                    ws2.Cell(row, 5).Value = n.Test_Results;
                    ws2.Cell(row, 6).Value = n.Name_of_Fit_Tester;
                    ws2.Cell(row, 7).Value = n.SubmittedAt.ToString("yyyy-MM-dd");
                    row++;
                }

                // Tally Report Sheet
                var ws3 = workbook.Worksheets.Add("DUO Unit Tally Report");
                ws3.Cell(1, 1).Value = "Department / Unit"; ws3.Cell(1, 2).Value = "Total Staff";
                ws3.Cell(1, 3).Value = "Total Fit Tested"; ws3.Cell(1, 4).Value = "Passed";
                ws3.Cell(1, 5).Value = "Failed"; ws3.Cell(1, 6).Value = "Not Done";
                ws3.Cell(1, 7).Value = "Expired"; ws3.Cell(1, 8).Value = "Rate (%)";
                ws3.Cell(1, 9).Value = "3M Aura"; ws3.Cell(1, 10).Value = "Grand Total";
                FormatHeaders(ws3, 10);

                row = 2;
                foreach (var t in tallyReport)
                {
                    var rate = t.TotalStaff > 0 ? (t.TotalFitTested * 100.0 / t.TotalStaff) : 0;
                    ws3.Cell(row, 1).Value = t.Unit;
                    ws3.Cell(row, 2).Value = t.TotalStaff;
                    ws3.Cell(row, 3).Value = t.TotalFitTested;
                    ws3.Cell(row, 4).Value = t.Passed;
                    ws3.Cell(row, 5).Value = t.Failed;
                    ws3.Cell(row, 6).Value = t.NotDone;
                    ws3.Cell(row, 7).Value = t.Expired;
                    ws3.Cell(row, 8).Value = Math.Round(rate, 1);
                    ws3.Cell(row, 9).Value = t.ThreeMAura;
                    ws3.Cell(row, 10).Value = t.GrandTotal;
                    row++;
                }

                // Total Row
                ws3.Cell(row, 1).Value = "TOTAL";
                ws3.Cell(row, 2).Value = tallyReport.Sum(x => x.TotalStaff);
                ws3.Cell(row, 3).Value = tallyReport.Sum(x => x.TotalFitTested);
                ws3.Cell(row, 4).Value = tallyReport.Sum(x => x.Passed);
                ws3.Cell(row, 5).Value = tallyReport.Sum(x => x.Failed);

                ws3.Cell(row, 6).Value = tallyReport.Sum(x => x.NotDone);
                ws3.Cell(row, 7).Value = tallyReport.Sum(x => x.Expired);
                var totalRate = tallyReport.Sum(x => x.TotalStaff) > 0
                    ? (tallyReport.Sum(x => x.TotalFitTested) * 100.0 / tallyReport.Sum(x => x.TotalStaff)) : 0;
                ws3.Cell(row, 8).Value = Math.Round(totalRate, 1);
                ws3.Cell(row, 9).Value = tallyReport.Sum(x => x.ThreeMAura);
                ws3.Cell(row, 10).Value = tallyReport.Sum(x => x.GrandTotal);

                ws3.Range(row, 1, row, 10).Style.Font.Bold = true;
                ws3.Range(row, 1, row, 10).Style.Fill.BackgroundColor = XLColor.LightGray;

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        $"Fit_Testing_Reports_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
                }
            }
        }
        [HttpGet]
        public async Task<IActionResult> GetEmployeeDetails(string employeeId)
        {
            var connectionString = _configuration.GetConnectionString("EmployeeConnection");
            var result = new { fullName = "", position = "", department = "" };

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"
                    SELECT TOP 1 
                        LastName + ' ' + FirstName + ' ' + MiddleName AS [Name],
                        Position,
                        Department
                    FROM UNIFIEDSVR.payroll.dbo.vwSPMS_User 
                    WHERE EmpNum = @EmpNum";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@EmpNum", employeeId);
                    await conn.OpenAsync();
                    using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                    {
                        if (await reader.ReadAsync())
                        {
                            result = new
                            {
                                fullName = reader["Name"].ToString().Trim(),
                                position = reader["Position"].ToString().Trim(),
                                department = reader["Department"].ToString().Trim()
                            };
                            return Json(result);
                        }
                    }
                }
            }
            return NotFound();
        }
    }
}