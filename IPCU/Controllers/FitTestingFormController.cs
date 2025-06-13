﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using IPCU.Data;
using IPCU.Models;
using X.PagedList;
using X.PagedList.Mvc.Core;
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

            // Check if filterExpiring is true
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

            // Order the data in descending order (e.g., by ExpiringAt)
            fitTestingForm = fitTestingForm.OrderByDescending(f => f.ExpiringAt);

            var pagedList = fitTestingForm.ToPagedList(pageNumber, pageSize);

            // Store selected filters
            ViewData["FilterExpiring"] = filterExpiring;
            ViewData["SelectedTestResult"] = testResult;
            ViewBag.SearchTerm = searchTerm;

            return View(pagedList);
        }




        // GET: FitTestingForm/Details/5
        public IActionResult Details(int id)
        {
            // Fetch the main FitTestingForm record
            var fitTestingForm = _context.FitTestingForm.FirstOrDefault(f => f.Id == id);
            if (fitTestingForm == null)
            {
                return NotFound();
            }

            // Fetch the history of attempts for the given form
            var history = _context.FitTestingFormHistory
                .Where(h => h.FitTestingFormId == id)
                .OrderBy(h => h.SubmittedAt) // Ensure chronological order
                .ToList();

            // Assign attempts based on their position in the history
            ViewData["FirstAttempt"] = history.ElementAtOrDefault(0);
            ViewData["SecondAttempt"] = history.ElementAtOrDefault(1);
            ViewData["LastAttempt"] = history.Count > 2 ? history.LastOrDefault() : null;

            return View(fitTestingForm);
        }


        // GET: FitTestingForm/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: FitTestingForm/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(FitTestingForm fitTestingForm, string? OtherLimitation)
        {
            // Remove OtherLimitation from ModelState since it's not in our model
            ModelState.Remove("OtherLimitation");

            try
            {
                var limitations = Request.Form["Limitation"].ToList();

                if (limitations != null && limitations.Any())
                {
                    // If "Other" is selected and has a value
                    if (limitations.Contains("Other"))
                    {
                        if (string.IsNullOrWhiteSpace(OtherLimitation))
                        {
                            ModelState.AddModelError("Limitation", "Please specify the other limitation");
                            return View(fitTestingForm);
                        }
                        limitations.Remove("Other");
                        limitations.Add(OtherLimitation.Trim());
                    }

                    // Concatenate the limitations into a single string
                    fitTestingForm.Limitation = string.Join(", ", limitations.Where(x => !string.IsNullOrWhiteSpace(x)));
                }
                else
                {
                    fitTestingForm.Limitation = "None";
                }

                if (ModelState.IsValid)
                {
                    // Set the SubmittedAt and ExpiringAt fields
                    fitTestingForm.SubmittedAt = DateTime.Now;
                    fitTestingForm.ExpiringAt = fitTestingForm.SubmittedAt.AddYears(1); // Set ExpiringAt

                    // Add the new FitTestingForm record
                    _context.Add(fitTestingForm);
                    await _context.SaveChangesAsync();

                    // Save the initial state to FitTestingFormHistory
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
                ModelState.AddModelError("", "Error processing limitations: " + ex.Message);
            }

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


            // Get "Attendance for Physicians" (only Passed results)
            var attendanceForPhysicians = _context.FitTestingForm.FromSqlRaw(@"
    SELECT * FROM FitTestingForm
    WHERE Professional_Category IN ({0}) AND Test_Results = {1}",
                string.Join(",", physicianCategories.Select(c => $"'{c}'")), "Passed")
                .ToList();



            var attendanceForNursingAndAllied = _context.FitTestingForm
    .FromSqlRaw(@"
        SELECT 
            HCW_Name,
            DUO,
            Professional_Category,
            Fit_Test_Solution,
            CASE 
                WHEN ExpiringAt < {0} THEN 'Expired' 
                ELSE 'Passed' 
            END AS Test_Results,
            Name_of_Fit_Tester,
            SubmittedAt,
            ExpiringAt
        FROM FitTestingForm
        WHERE Professional_Category NOT IN ({1}) AND Test_Results = 'Passed'",
        currentDate,
        string.Join(",", physicianCategories.Select(c => $"'{c}'")))
    .Select(f => new FitTestingReportViewModel
    {
        HCW_Name = f.HCW_Name,
        DUO = f.DUO,
        Professional_Category = f.Professional_Category,
        Fit_Test_Solution = f.Fit_Test_Solution,
        Test_Results = f.Test_Results, // Now handled by SQL
        Name_of_Fit_Tester = f.Name_of_Fit_Tester,
        SubmittedAt = f.SubmittedAt,
        ExpiringAt = f.ExpiringAt
    })
    .ToList();

            var tallyReport = duoList
                .Select(unit => new
                {
                    Unit = unit,
                    TotalFitTested = fitTests.Count(f => f.DUO == unit && f.Test_Results == "Passed"),
                    Expired = fitTests.Count(f => f.DUO == unit && f.Test_Results == "Passed" && f.ExpiringAt < currentDate)
                })
                .ToList();

            var tallyMedical = duoMedical
                .Select(unit => new
                {
                    Unit = unit,
                    TotalFitTested = fitTests.Count(f => f.DUO == unit && f.Test_Results == "Passed"),
                    Expired = fitTests.Count(f => f.DUO == unit && f.Test_Results == "Passed" && f.ExpiringAt < currentDate)
                })
                .ToList();

            var tallyAllied = duoAllied
               .Select(unit => new
               {
                   Unit = unit,
                   TotalFitTested = fitTests.Count(f => f.DUO == unit && f.Test_Results == "Passed"),
                   Expired = fitTests.Count(f => f.DUO == unit && f.Test_Results == "Passed" && f.ExpiringAt < currentDate)
               })
               .ToList();

            // Totals for tallyReport
            int totalFitTestedReport = tallyReport.Sum(t => t.TotalFitTested);
            int totalExpiredReport = tallyReport.Sum(t => t.Expired);

            // Totals for tallyMedical
            int totalFitTestedMedical = tallyMedical.Sum(t => t.TotalFitTested);
            int totalExpiredMedical = tallyMedical.Sum(t => t.Expired);

            // Totals for tallyAllied
            int totalFitTestedAllied = tallyAllied.Sum(t => t.TotalFitTested);
            int totalExpiredAllied = tallyAllied.Sum(t => t.Expired);

            // Grand Totals (sum of all three)
            int grandTotalFitTested = totalFitTestedReport + totalFitTestedMedical + totalFitTestedAllied;
            int grandTotalExpired = totalExpiredReport + totalExpiredMedical + totalExpiredAllied;

            // Pass all lists to the view
            ViewBag.AttendanceForPhysicians = attendanceForPhysicians;
            ViewBag.AttendanceForNursingAndAllied = attendanceForNursingAndAllied;
            ViewBag.TallyReport = tallyReport;
            ViewBag.TallyMedical = tallyMedical;
            ViewBag.TallyAllied = tallyAllied;

            ViewBag.TotalFitTestedReport = totalFitTestedReport;
            ViewBag.TotalExpiredReport = totalExpiredReport;
            ViewBag.TotalFitTestedMedical = totalFitTestedMedical;
            ViewBag.TotalExpiredMedical = totalExpiredMedical;
            ViewBag.TotalFitTestedAllied = totalFitTestedAllied;
            ViewBag.TotalExpiredAllied = totalExpiredAllied;

            ViewBag.GrandTotalFitTested = grandTotalFitTested;
            ViewBag.GrandTotalExpired = grandTotalExpired;


            return View();
        }
        public IActionResult ExportToExcel()
        {
            var attendanceForPhysicians = _context.FitTestingForm
                .Where(f => f.Test_Results == "Passed")
                .ToList();

            var attendanceForNursingAndAllied = _context.FitTestingForm
                .Where(f => f.Test_Results == "Passed")
                .ToList();

            var tallyReport = _context.FitTestingForm
                .GroupBy(f => f.DUO)
                .Select(g => new
                {
                    Unit = g.Key,
                    TotalFitTested = g.Count(f => f.Test_Results == "Passed"),
                    Expired = g.Count(f => f.ExpiringAt < DateTime.Now && f.Test_Results == "Passed")
                }).ToList();

            using (var workbook = new XLWorkbook())
            {
                void FormatHeaders(IXLWorksheet sheet, int columnCount)
                {
                    var headerRow = sheet.Row(1);
                    headerRow.Style.Font.Bold = true;
                    headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;
                    sheet.Columns(1, columnCount).AdjustToContents();
                }

                var worksheet1 = workbook.Worksheets.Add("Physicians");
                worksheet1.Cell(1, 1).Value = "Name";
                worksheet1.Cell(1, 2).Value = "Department/Unit/Office";
                worksheet1.Cell(1, 3).Value = "Professional Category";
                worksheet1.Cell(1, 4).Value = "Fit Test Solution";
                worksheet1.Cell(1, 5).Value = "Status";
                worksheet1.Cell(1, 6).Value = "Fit Tester";
                worksheet1.Cell(1, 7).Value = "Date Fit Tested";
                FormatHeaders(worksheet1, 7);

                int row = 2;
                foreach (var item in attendanceForPhysicians)
                {
                    worksheet1.Cell(row, 1).Value = item.HCW_Name;
                    worksheet1.Cell(row, 2).Value = item.DUO;
                    worksheet1.Cell(row, 3).Value = item.Professional_Category;
                    worksheet1.Cell(row, 4).Value = item.Fit_Test_Solution;
                    worksheet1.Cell(row, 5).Value = item.Test_Results;
                    worksheet1.Cell(row, 6).Value = item.Name_of_Fit_Tester;
                    worksheet1.Cell(row, 7).Value = item.SubmittedAt.ToString("yyyy-MM-dd");
                    row++;
                }

                var worksheet2 = workbook.Worksheets.Add("Nursing & Allied");
                worksheet2.Cell(1, 1).Value = "Name";
                worksheet2.Cell(1, 2).Value = "Department/Unit/Office";
                worksheet2.Cell(1, 3).Value = "Professional Category";
                worksheet2.Cell(1, 4).Value = "Fit Test Solution";
                worksheet2.Cell(1, 5).Value = "Status";
                worksheet2.Cell(1, 6).Value = "Fit Tester";
                worksheet2.Cell(1, 7).Value = "Date Fit Tested";
                FormatHeaders(worksheet2, 7);

                row = 2;
                foreach (var item in attendanceForNursingAndAllied)
                {
                    worksheet2.Cell(row, 1).Value = item.HCW_Name;
                    worksheet2.Cell(row, 2).Value = item.DUO;
                    worksheet2.Cell(row, 3).Value = item.Professional_Category;
                    worksheet2.Cell(row, 4).Value = item.Fit_Test_Solution;
                    worksheet2.Cell(row, 5).Value = item.Test_Results;
                    worksheet2.Cell(row, 6).Value = item.Name_of_Fit_Tester;
                    worksheet2.Cell(row, 7).Value = item.SubmittedAt.ToString("yyyy-MM-dd");
                    row++;
                }

                var worksheet3 = workbook.Worksheets.Add("Tally Report");
                worksheet3.Cell(1, 1).Value = "Unit";
                worksheet3.Cell(1, 2).Value = "Total Fit Tested";
                worksheet3.Cell(1, 3).Value = "Expired";
                FormatHeaders(worksheet3, 3);

                row = 2;
                foreach (var item in tallyReport)
                {
                    worksheet3.Cell(row, 1).Value = item.Unit;
                    worksheet3.Cell(row, 2).Value = item.TotalFitTested;
                    worksheet3.Cell(row, 3).Value = item.Expired;
                    row++;
                }

                var summarySheet = workbook.Worksheets.Add("Summary");
                summarySheet.Cell(1, 1).Value = "Category";
                summarySheet.Cell(1, 2).Value = "Total Fit Tested";
                summarySheet.Cell(1, 3).Value = "Expired";
                FormatHeaders(summarySheet, 3);

                summarySheet.Cell(2, 1).Value = "Physicians";
                summarySheet.Cell(2, 2).Value = attendanceForPhysicians.Count;
                summarySheet.Cell(2, 3).Value = attendanceForPhysicians.Count(f => f.ExpiringAt < DateTime.Now);

                summarySheet.Cell(3, 1).Value = "Nursing & Allied";
                summarySheet.Cell(3, 2).Value = attendanceForNursingAndAllied.Count;
                summarySheet.Cell(3, 3).Value = attendanceForNursingAndAllied.Count(f => f.ExpiringAt < DateTime.Now);

                summarySheet.Cell(4, 1).Value = "Grand Total";
                summarySheet.Cell(4, 2).FormulaA1 = "=B2+B3";
                summarySheet.Cell(4, 3).FormulaA1 = "=C2+C3";

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Fit_Testing_Reports.xlsx");
                }
            }
        }


        [HttpGet]
        public async Task<IActionResult> GetEmployeeDetails(string employeeId)
        {
            var connectionString = _configuration.GetConnectionString("EmployeeConnection");

            var result = new
            {
                fullName = "",
                position = "",
                department = ""
            };

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                string sql = @"
            SELECT TOP 1 
                EmpNum, 
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