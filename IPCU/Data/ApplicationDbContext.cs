using IPCU.Models;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using IPCU.Models.InfectionControl.Models;

namespace IPCU.Data
{
    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }
        public DbSet<PatientForm> PatientForms { get; set; }
        public DbSet<IPCU.Models.TrainingEvaluation> TrainingEvaluation { get; set; } = default!;
        public DbSet<IPCU.Models.FitTestingForm> FitTestingForm { get; set; } = default!;
        public DbSet<HandHygieneForm> HandHygieneForms { get; set; }
        public DbSet<HHActivity> HHActivities { get; set; }
        public DbSet<FitTestingFormHistory> FitTestingFormHistory { get; set; }
        public DbSet<Trainee> Trainees { get; set; }

        public DbSet<HandHygieneComplianceSummary> HandHygieneComplianceSummary { get; set; }

        public DbSet<PreTestNonClinical> PreTestNonClinicals { get; set; }
        public DbSet<PreTestClinical> PreTestClinicals { get; set; }
        public DbSet<PostTestNonClinical> PostTestNonCLinicals { get; set; }
        public DbSet<PostTestClinical> PostTestClinicals { get; set; }
        public DbSet<IPCU.Models.EvaluationViewModel> EvaluationViewModel { get; set; } = default!;

        public DbSet<TrainingSummary> TrainingSummaries { get; set; }
        public DbSet<EvaluationViewModel> Evaluations { get; set; }
        public DbSet<TrainingEvaluation> TrainingEvaluations { get; set; } // ✅ Ensure this is present
        public DbSet<Pneumonia> Pneumonias { get; set; }
        public DbSet<Usi> Usi { get; set; }

        public DbSet<Patient> Patients { get; set; }
        public DbSet<PatientMaster> PatientMasters { get; set; }
        public DbSet<VitalSigns> VitalSigns { get; set; }
        public DbSet<DeviceConnected> DeviceConnected { get; set; }

        public DbSet<SSTInfectionModel> SSTInfectionModels { get; set; }
        public DbSet<UTIFormModel> UTIModels { get; set; }

        public DbSet<CardiovascularSystemInfection> CardiovascularSystemInfection { get; set; }
        public DbSet<PediatricVAEChecklist> PediatricVAEChecklist { get; set; }
        public DbSet<LaboratoryConfirmedBSI> LaboratoryConfirmedBSI { get; set; }

        public DbSet<VentilatorEventChecklist> VentilatorEventChecklists { get; set; }
        public DbSet<Diagnostics> Diagnostics { get; set; }
        public DbSet<Antibiotic> Antibiotics { get; set; }
        public DbSet<DiagnosticsTreatment> DiagnosticsTreatments { get; set; }
        public DbSet<SurgicalSiteInfectionChecklist> SurgicalSiteInfectionChecklist { get; set; }
        public DbSet<GIInfectionChecklist> GIInfectionChecklists { get; set; }
        public DbSet<PatientMovement> PatientMovements { get; set; }
        public DbSet<PatientMovementDetail> PatientMovementDetails { get; set; }

        public DbSet<InfectionControlMonitoringForm> InfectionControlMonitoringForms { get; set; }
        public DbSet<SSIP> SSIP { get; set; }
        public DbSet<DailyCentralLineMaintenanceChecklist> DailyCentralLineMaintenanceChecklists { get; set; }
        public DbSet<IPCU.Models.EnvironmentalzControl> EnvironmentalzControl { get; set; } = default!;
        public DbSet<MDROrderSheet> MDROrderSheets { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Create a unique index on PatientMovement for date and area combination
            modelBuilder.Entity<PatientMovement>()
                .HasIndex(p => new { p.MovementDate, p.Area })
                .IsUnique();

            // Configure relationship between PatientMovementDetail and PatientMaster
            modelBuilder.Entity<PatientMovementDetail>()
                .HasOne(pmd => pmd.Patient)
                .WithMany()
                .HasForeignKey(pmd => pmd.HospNum)
                .OnDelete(DeleteBehavior.Restrict);

            // Configure relationship between PatientMovementDetail and PatientMovement
            modelBuilder.Entity<PatientMovementDetail>()
                .HasOne(pmd => pmd.PatientMovement)
                .WithMany()
                .HasForeignKey(pmd => pmd.MovementId);
        }
        public DbSet<IPCU.Models.ICRA> ICRA { get; set; } = default!;
        public DbSet<IPCU.Models.PostConstruction> PostConstruction { get; set; } = default!;
        public DbSet<IPCU.Models.TCSkillsChecklist> TCSkillsChecklist { get; set; } = default!;

        public DbSet<IPCU.Models.InfectionControl.Models.InfectionControlMonitoringForm> InfectionControlMonitoringForm { get; set; } = default!;
        public DbSet<IPCU.Models.VAEMonitoringChecklist> VAEMonitoringChecklist { get; set; } = default!;
        public DbSet<IPCU.Models.NoticeOfReferral> NoticeOfReferral { get; set; } = default!;
        public DbSet<IPCU.Models.Insertion> Insertion { get; set; } = default!;
        public DbSet<IPCU.Models.TCSkillsChecklistReal> TCSkillsChecklistReal { get; set; } = default!;
        public DbSet<IPCU.Models.PatientHAI> PatientHAI { get; set; }
        public object FitTestingForms { get; internal set; }

        //protected override void OnModelCreating(ModelBuilder modelBuilder)
        //{
        //    base.OnModelCreating(modelBuilder);

        //    // Optional: Configure relationships if needed
        //    modelBuilder.Entity<Patient>()
        //        .HasOne<PatientMaster>()
        //        .WithMany()
        //        .HasForeignKey(p => p.HospNum);
        //}

        //protected override void OnModelCreating(ModelBuilder modelBuilder)
        //{
        //    base.OnModelCreating(modelBuilder);

        //    // Configure the relationships for VitalSignsTable
        //    modelBuilder.Entity<VitalSigns>()
        //        .HasOne(v => v.Patient)
        //        .WithMany()
        //        .HasForeignKey(v => v.HospNum)
        //        .OnDelete(DeleteBehavior.NoAction);  // Use NoAction to avoid circular cascade delete

        //    modelBuilder.Entity<VitalSigns>()
        //        .HasOne(v => v.PatientMaster)
        //        .WithMany()
        //        .HasForeignKey(v => v.HospNum)
        //        .OnDelete(DeleteBehavior.NoAction);  // Use NoAction to avoid circular cascade delete
        //}

    }

}