namespace Quality_management_portal.Data.Entities
{
    public class MilestoneDeliverable
    {
        public long RecordId { get; set; }
        public int? SerialNumber { get; set; }
        public string MilestoneName { get; set; } = string.Empty;
        public string DeliverableName { get; set; } = string.Empty;
        public string? DeliverableType { get; set; }
        public DateTime? MilestoneTargetDate { get; set; }
        public DateTime? DeliverableTargetDate { get; set; }
        public string? AssignedTo { get; set; }
        public string? AnyScheduleDeviation { get; set; }
        public string? DeliverableAtWhichActivity { get; set; }
        public string? ReasonForScheduleDeviation { get; set; }
        public DateTime? DeliverableActualCompletedDate { get; set; }
        public DateTime? MilestoneAchievedDate { get; set; }
        public string Status { get; set; } = string.Empty;
        public string? Remarks { get; set; }
    }
}