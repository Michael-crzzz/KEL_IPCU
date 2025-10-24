using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IPCU.Models
{
    public class FitTestingForm
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public string EmployeeId { get; set; } // Identifier for other tables

        [Required]
        public string HCW_Name { get; set; }

        [Required] 
        public string DUO { get; set; }
        [Required]
        public string Professional_Category { get; set; }

        [Display(Name = "Limitation")]
        public string?  Limitation { get; set; }

        [Required]
        public string Fit_Test_Solution { get; set; }

        [Required]
        public string Sensitivity_Test { get; set; }

        [Required]
        public string Respiratory_Type { get; set; }

        [Required]
        public string Model { get; set; }

        [Required]
        public string Size { get; set; }

        [Required]
        public bool Normal_Breathing { get; set; }

        [Required]
        public bool Deep_Breathing { get; set; }

        [Required]
        public bool Turn_head_side_to_side { get; set; }

        [Required]
        public bool Move_head_up_and_down { get; set; }

        [Required]
        public bool Reading { get; set; }

        [Required]
        public bool Bending_Jogging { get; set; }

        [Required]
        public bool Normal_Breathing_2 { get; set; }

        // Test Results (computed dynamically)
        public string Test_Results
        {
            get
            {
                var booleanFields = new[]
                {
                Normal_Breathing,
                Deep_Breathing,
                Turn_head_side_to_side,
                Move_head_up_and_down,
                Reading,
                Bending_Jogging,
                Normal_Breathing_2
            };

                return booleanFields.All(b => b) ? "Passed" : "Failed";
            }
            set { }
        }

        [Required]
        public string Name_of_Fit_Tester { get; set; }

        [Required]
        public string DUO_Tester { get; set; }

        [Required]
        public DateTime SubmittedAt { get; set; } = DateTime.Now;

        public DateTime ExpiringAt { get; set; }


        public int SubmissionCount { get; set; } = 0;
        public int MaxRetakes { get; set; } = 2;
        public string Status
        {
            get
            {
                if (Test_Results == "Passed")
                {
                    return ExpiringAt >= DateTime.Now ? "Passed" : "Expired";
                }
                return "Failed";
            }
        }

    }


}
