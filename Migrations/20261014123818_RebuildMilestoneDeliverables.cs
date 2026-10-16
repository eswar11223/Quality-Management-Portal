using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Quality_management_portal.Migrations
{
    /// <inheritdoc />
    public partial class RebuildMilestoneDeliverables : Migration
    {
        /// <inheritdoc />
        
protected override void Up(MigrationBuilder migrationBuilder)
{
    migrationBuilder.DropTable(
        name: "MilestoneDeliverables");

    migrationBuilder.CreateTable(
        name: "MilestoneDeliverables",
        columns: table => new
        {
            RecordId = table.Column<long>(
                type: "bigint",
                nullable: false
            ).Annotation("MySql:AutoIncrement", true),

            SerialNumber = table.Column<int>(
                name: "S.No", type: "int", nullable: true),

            MilestoneName = table.Column<string>(
                name: "Milestone Name", type: "varchar(200)",
                maxLength: 200, nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4"),

            DeliverableName = table.Column<string>(
                name: "Deliverable Name", type: "varchar(200)",
                maxLength: 200, nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4"),

            DeliverableType = table.Column<string>(
                name: "Deliverable Type", type: "varchar(200)",
                maxLength: 200, nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4"),

            MilestoneTargetDate = table.Column<DateTime>(
                name: "Milestone Target Date", type: "date", nullable: true),

            DeliverableTargetDate = table.Column<DateTime>(
                name: "Deliverable Target Date", type: "date", nullable: true),

            AssignedTo = table.Column<string>(
                name: "Assigned To", type: "varchar(255)",
                maxLength: 255, nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4"),

            AnyScheduleDeviation = table.Column<string>(
                name: "Any Schedule Deviation", type: "varchar(100)",
                maxLength: 100, nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4"),

            DeliverableAtWhichActivity = table.Column<string>(
                name: "Deliverable at Which Activity", type: "varchar(500)",
                maxLength: 500, nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4"),

            ReasonForScheduleDeviation = table.Column<string>(
                name: "Reason for Schedule Deviation", type: "varchar(2000)",
                maxLength: 2000, nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4"),

            DeliverableActualCompletedDate = table.Column<DateTime>(
                name: "Deliverable Actual Completed Date",
                type: "date", nullable: true),

            MilestoneAchievedDate = table.Column<DateTime>(
                name: "Milestone Achieved Date", type: "date", nullable: true),

            Status = table.Column<string>(
                name: "Status", type: "varchar(100)",
                maxLength: 100, nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4"),

            Remarks = table.Column<string>(
                name: "Remarks", type: "varchar(2000)",
                maxLength: 2000, nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4")
        },
        constraints: table =>
        {
            table.PrimaryKey("PK_MilestoneDeliverables", x => x.RecordId);
        })
        .Annotation("MySql:CharSet", "utf8mb4");
}


        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "MilestoneDeliverables");
        }
    }
}
