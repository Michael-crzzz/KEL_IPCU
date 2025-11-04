using IPCU.Data;
using IPCU.Models;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using Microsoft.EntityFrameworkCore;

public class FitTestingFormPdfService
{
    private readonly ApplicationDbContext _context;

    public FitTestingFormPdfService(ApplicationDbContext context)
    {
        _context = context;
    }

    private static readonly string DarkColor = Colors.Grey.Darken3;
    private static readonly string AccentColor = Colors.Indigo.Darken4;
    private static readonly float HeaderSize = 12;
    private static readonly float BodySize = 10;

    private byte[] LoadLogo()
    {
        return File.ReadAllBytes("wwwroot/images/NKTI.png");
    }

    public byte[] GeneratePdf(FitTestingForm fitTestingForm)
    {
        QuestPDF.Settings.License = LicenseType.Community;

        // Fetch the history records for the given FitTestingForm
        var history = _context.FitTestingFormHistory
            .Where(h => h.FitTestingFormId == fitTestingForm.Id)
            .OrderBy(h => h.SubmittedAt)
            .ToList();

        return Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(15);

                // Header Section
                page.Header().Row(header =>
                {
                    header.ConstantColumn(50).Image(LoadLogo()).FitArea();
                    header.RelativeColumn().Column(col =>
                    {
                        col.Spacing(5);
                        col.Item().AlignCenter().Text("NATIONAL KIDNEY AND TRANSPLANT INSTITUTE")
                            .FontSize(14).Bold().FontColor(AccentColor);
                        col.Item().AlignCenter().Text("East Avenue, Quezon City")
                            .FontSize(9).Bold().FontColor(DarkColor);
                        col.Item().PaddingVertical(5).LineHorizontal(1).LineColor(Colors.Grey.Darken3);
                        col.Item().AlignCenter().Text("INFECTION PREVENTION AND CONTROL UNIT")
                            .FontSize(9).Bold().FontColor(DarkColor);
                        col.Item().AlignCenter().Text("QUALITATIVE RESPIRATOR FIT TESTING FORM")
                            .FontSize(9).Bold().FontColor(DarkColor);
                    });
                });

                page.Content().Column(content =>
                {
                    content.Spacing(8);

                    // Healthcare Worker Information
                    content.Item().Border(1).BorderColor(Colors.Grey.Darken2).Column(col =>
                    {
                        col.Item().Padding(5).Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                            });
                            AddTableRow(table, "HCW Name:", fitTestingForm.HCW_Name);
                            AddTableRow(table, "Department/Unit/Office:", fitTestingForm.DUO);
                            AddTableRow(table, "Limitations:", fitTestingForm.Limitation);
                        });
                    });

                    // Fit Test Details
                    

                    // Fit Test Activities Table with History
                    content.Item().Border(1).BorderColor(Colors.Grey.Darken2).Column(col =>
                    {
                        col.Item().Padding(5).Background(Colors.Grey.Lighten3)
                            .Text("FIT TEST ACTIVITIES AND HISTORY")
                            .FontSize(HeaderSize).Bold().FontColor(DarkColor);

                        col.Item().Padding(5).Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.RelativeColumn(2); // Activity Column
                                columns.RelativeColumn();  // First Attempt
                                columns.RelativeColumn();  // Second Attempt
                                columns.RelativeColumn();  // Third Attempt
                            });

                            // Table Header
                            table.Header(header =>
                            {
                                header.Cell().BorderBottom(1).PaddingVertical(3)
                                    .Text("Activity").Bold().FontColor(AccentColor);
                                header.Cell().BorderBottom(1).PaddingVertical(3)
                                    .Text("First Attempt").Bold().FontColor(AccentColor);
                                header.Cell().BorderBottom(1).PaddingVertical(3)
                                    .Text("Second Attempt").Bold().FontColor(AccentColor);
                                header.Cell().BorderBottom(1).PaddingVertical(3)
                                    .Text("Third Attempt").Bold().FontColor(AccentColor);
                            });

                            // Helper function to display checkmarks and Xs
                            string DisplayBoolean(bool? value) => value.HasValue ? (value.Value ? "✔" : "✘") : "";

                            // Fit Test Details (Solution, Sensitivity Test, Respiratory Type, Model, Size)
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text("Fit Test Solution:").FontSize(BodySize).FontColor(DarkColor);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(0)?.Fit_Test_Solution ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(1)?.Fit_Test_Solution ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(2)?.Fit_Test_Solution ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);

                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text("Sensitivity Test:").FontSize(BodySize).FontColor(DarkColor);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(0)?.Sensitivity_Test ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(1)?.Sensitivity_Test ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(2)?.Sensitivity_Test ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);

                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text("Respiratory Type:").FontSize(BodySize).FontColor(DarkColor);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(0)?.Respiratory_Type ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(1)?.Respiratory_Type ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(2)?.Respiratory_Type ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);

                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text("Model:").FontSize(BodySize).FontColor(DarkColor);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(0)?.Model ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(1)?.Model ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(2)?.Model ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);

                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text("Size:").FontSize(BodySize).FontColor(DarkColor);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(0)?.Size ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(1)?.Size ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text(history.ElementAtOrDefault(2)?.Size ?? "").FontSize(BodySize).FontColor(Colors.Grey.Darken2);

                            // Activity Rows
                            AddActivityRow(table, "Normal Breathing", history, h => h.Normal_Breathing, DisplayBoolean);
                            AddActivityRow(table, "Deep Breathing", history, h => h.Deep_Breathing, DisplayBoolean);
                            AddActivityRow(table, "Turn Head Side to Side", history, h => h.Turn_head_side_to_side, DisplayBoolean);
                            AddActivityRow(table, "Move Head Up and Down", history, h => h.Move_head_up_and_down, DisplayBoolean);
                            AddActivityRow(table, "Reading", history, h => h.Reading, DisplayBoolean);
                            AddActivityRow(table, "Bending/Jogging", history, h => h.Bending_Jogging, DisplayBoolean);
                            AddActivityRow(table, "Normal Breathing (2)", history, h => h.Normal_Breathing_2, DisplayBoolean);

                            // Test Results
                            table.Cell().BorderBottom(1).PaddingVertical(3)
                                .Text("Result").FontSize(BodySize).Bold().FontColor(DarkColor);
                            for (int i = 0; i < 3; i++)
                            {
                                table.Cell().BorderBottom(1).PaddingVertical(3)
                                    .Text(history.ElementAtOrDefault(i)?.Test_Results)
                                    .FontSize(BodySize)
                                    .FontColor(history.ElementAtOrDefault(i)?.Test_Results == "Failed" ? Colors.Red.Darken3 : Colors.Green.Darken3);
                            }
                        });
                    });


                    // Certification Section
                    content.Item().Border(1).BorderColor(Colors.Grey.Darken2).Column(col =>
                    {
                        col.Item().Padding(5).Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.RelativeColumn();
                                columns.RelativeColumn();
                            });
                            AddTableRow(table, "Fit Tester:", fitTestingForm.Name_of_Fit_Tester);
                            AddTableRow(table, "Date:", fitTestingForm.SubmittedAt.ToString("MMM dd, yyyy"));
                            AddTableRow(table, "Expiration Date:", fitTestingForm.ExpiringAt.ToString("MMM dd, yyyy"));
                        });
                    });

                    // Disclaimer Section
                    content.Item().Border(1).BorderColor(Colors.Grey.Darken2).Column(col =>
                    {
                        col.Item().Padding(5).Background(Colors.Grey.Lighten3)
                            .Text("DISCLAIMER")
                            .FontSize(HeaderSize).Bold().FontColor(DarkColor);

                        col.Item().Padding(5).Text(text =>
                        {
                            text.Span("These procedures are in accordance with accepted standards of fit testing...")
                                .FontSize(BodySize - 1).FontColor(Colors.Grey.Darken2).LineHeight(1);
                        });
                    });
                });

                // Footer Section
                page.Footer().AlignCenter().Text(t =>
                {
                    t.Span("Page ").FontSize(BodySize).FontColor(Colors.Grey.Medium);
                    t.CurrentPageNumber().FontSize(BodySize).FontColor(AccentColor);
                    t.Span(" of ").FontSize(BodySize).FontColor(Colors.Grey.Medium);
                    t.TotalPages().FontSize(BodySize).FontColor(AccentColor);
                });
            });
        }).GeneratePdf();
    }

    private void AddTableRow(TableDescriptor table, string label, string value)
    {
        table.Cell().BorderBottom(1).PaddingVertical(3)
            .Text(label).FontSize(BodySize).Bold().FontColor(DarkColor);
        table.Cell().BorderBottom(1).PaddingVertical(3)
            .Text(value).FontSize(BodySize).FontColor(Colors.Grey.Darken2);
    }

    private void AddActivityRow(TableDescriptor table, string activity, List<FitTestingFormHistory> history, Func<FitTestingFormHistory, bool?> selector, Func<bool?, string> displayBoolean)
    {
        table.Cell().BorderBottom(1).PaddingVertical(3)
            .Text(activity).FontSize(BodySize).FontColor(DarkColor);
        for (int i = 0; i < 3; i++)
        {
            table.Cell().BorderBottom(1).PaddingVertical(3)
                .AlignCenter().Text(displayBoolean(history.ElementAtOrDefault(i) != null ? selector(history[i]) : null))
                .FontColor(history.ElementAtOrDefault(i) != null && selector(history[i]) == true ? Colors.Green.Darken3 : Colors.Red.Darken3);
        }
    }
}