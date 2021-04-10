using System;
using System.Configuration;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace EntregasRendir.Context.Models
{
    public partial class InterfaceDBContext : DbContext
    {
        public InterfaceDBContext()
        {
        }

        public InterfaceDBContext(DbContextOptions<InterfaceDBContext> options)
            : base(options)
        {
        }

        public virtual DbSet<EntregaRendir> EntregaRendir { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            var Conn = ConfigurationManager.ConnectionStrings["InterfaceDB"].ConnectionString;
            optionsBuilder.UseSqlServer(Conn);
            
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<EntregaRendir>(entity =>
            {
                entity.HasKey(e => new { e.CorrelativoHelm, e.Secuencial });
                entity.ToTable("EntregaRendir");
            });

        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
