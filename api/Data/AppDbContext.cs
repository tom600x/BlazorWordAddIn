using Microsoft.EntityFrameworkCore;
using SnippetsApi.Models;

namespace SnippetsApi.Data;

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    public DbSet<TextSnippet> TextSnippets => Set<TextSnippet>();
    public DbSet<ImageSnippet> ImageSnippets => Set<ImageSnippet>();
    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<TextSnippet>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.OwnerUpn).HasMaxLength(320).IsRequired();
            entity.Property(e => e.Title).HasMaxLength(200).IsRequired();
            entity.Property(e => e.BodyText).IsRequired();
            entity.Property(e => e.Tags).HasMaxLength(500);
            entity.Property(e => e.UpdatedUtc).HasDefaultValueSql("SYSUTCDATETIME()");
            entity.HasIndex(e => e.OwnerUpn);
        });

        modelBuilder.Entity<ImageSnippet>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.OwnerUpn).HasMaxLength(320).IsRequired();
            entity.Property(e => e.Title).HasMaxLength(200).IsRequired();
            // ImageBase64 can be large; stored as nvarchar(max)
            entity.Property(e => e.ImageBase64).IsRequired();
            entity.Property(e => e.MimeType).HasMaxLength(50).IsRequired();
            entity.Property(e => e.Tags).HasMaxLength(500);
            entity.Property(e => e.UpdatedUtc).HasDefaultValueSql("SYSUTCDATETIME()");
            entity.HasIndex(e => e.OwnerUpn);
        });

    }
}
