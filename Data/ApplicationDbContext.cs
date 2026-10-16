
using Microsoft.EntityFrameworkCore;
using Quality_management_portal.Data.Entities;

namespace Quality_management_portal.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(
            DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<User> Users { get; set; }

        public DbSet<Role> Roles { get; set; }
        public DbSet<MilestoneDeliverable> MilestoneDeliverables { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<Role>(entity =>
            {
                entity.ToTable("Roles");

                entity.HasKey(r => r.RoleId);

                entity.Property(r => r.RoleName)
                    .HasMaxLength(100)
                    .IsRequired();

                entity.HasIndex(r => r.RoleName)
                    .IsUnique();

                entity.Property(r => r.Description)
                    .HasMaxLength(500);
            });

            modelBuilder.Entity<User>(entity =>
            {
                entity.ToTable("Users");

                entity.HasKey(u => u.UserId);

                entity.Property(u => u.FirstName)
                    .HasMaxLength(100)
                    .IsRequired();

                entity.Property(u => u.LastName)
                    .HasMaxLength(100)
                    .IsRequired();

                entity.Property(u => u.Email)
                    .HasMaxLength(255)
                    .IsRequired();

                entity.Property(u => u.Username)
                    .HasMaxLength(100)
                    .IsRequired();

                entity.Property(u => u.PasswordHash)
                    .HasMaxLength(255)
                    .IsRequired();

                entity.HasIndex(u => u.Email)
                    .IsUnique();

                entity.HasIndex(u => u.Username)
                    .IsUnique();

                entity.HasOne(u => u.Role)
                    .WithMany(r => r.Users)
                    .HasForeignKey(u => u.RoleId)
                    .OnDelete(DeleteBehavior.Restrict);
            });
            
            
            modelBuilder.Entity<MilestoneDeliverable>(entity =>
            {
                entity.ToTable("MilestoneDeliverables");

                entity.HasKey(m => m.RecordId);

                entity.Property(m => m.RecordId)
                    .HasColumnName("Record ID")
                    .ValueGeneratedOnAdd();

                entity.Property(m => m.SerialNumber)
                    .HasColumnName("S.No");

                entity.Property(m => m.MilestoneName)
                    .HasColumnName("Milestone Name")
                    .HasMaxLength(200)
                    .IsRequired();

                entity.Property(m => m.DeliverableName)
                    .HasColumnName("Deliverable Name")
                    .HasMaxLength(200)
                    .IsRequired();

                entity.Property(m => m.DeliverableType)
                    .HasColumnName("Deliverable Type")
                    .HasMaxLength(200);

                entity.Property(m => m.MilestoneTargetDate)
                    .HasColumnName("Milestone Target Date")
                    .HasColumnType("date");

                entity.Property(m => m.DeliverableTargetDate)
                    .HasColumnName("Deliverable Target Date")
                    .HasColumnType("date");

                entity.Property(m => m.AssignedTo)
                    .HasColumnName("Assigned To")
                    .HasMaxLength(255);

                entity.Property(m => m.AnyScheduleDeviation)
                    .HasColumnName("Any Schedule Deviation")
                    .HasMaxLength(100);

                entity.Property(m => m.DeliverableAtWhichActivity)
                    .HasColumnName("Deliverable at Which Activity")
                    .HasMaxLength(500);

                entity.Property(m => m.ReasonForScheduleDeviation)
                    .HasColumnName("Reason for Schedule Deviation")
                    .HasMaxLength(2000);

                entity.Property(m => m.DeliverableActualCompletedDate)
                    .HasColumnName("Deliverable Actual Completed Date")
                    .HasColumnType("date");

                entity.Property(m => m.MilestoneAchievedDate)
                    .HasColumnName("Milestone Achieved Date")
                    .HasColumnType("date");

                entity.Property(m => m.Status)
                    .HasColumnName("Status")
                    .HasMaxLength(100)
                    .IsRequired();

                entity.Property(m => m.Remarks)
                    .HasColumnName("Remarks")
                    .HasMaxLength(2000);
            });


                    }
                }
            }
