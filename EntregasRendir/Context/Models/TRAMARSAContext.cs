using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace EntregasRendir.Context.Models
{
    public partial class TRAMARSAContext : DbContext
    {
        public TRAMARSAContext()
        {
        }

        public TRAMARSAContext(DbContextOptions<TRAMARSAContext> options)
            : base(options)
        {
        }

        public virtual DbSet<USR_CJRMVX> USR_CJRMVX { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            var Conn = ConfigurationManager.ConnectionStrings["TRAMARSA"].ConnectionString;
            optionsBuilder.UseSqlServer(Conn);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<USR_CJRMVX>(entity =>
            {
                entity.HasKey(e => new { e.USR_CJRMVX_CODEMP, e.USR_CJRMVX_CODFOR, e.USR_CJRMVX_IDHELM, e.USR_CJRMVX_NROITM });
                entity.ToTable("USR_CJRMVX");

                entity.Property(e => e.USR_CJ_TSTAMP)
                    .HasColumnName("USR_CJ_TSTAMP")
                    .IsRowVersion();
                entity.Property(e => e.USR_CJ_FECALT).IsRequired(false);
                entity.Property(e => e.USR_CJ_FECMOD).IsRequired(false);
                entity.Property(e => e.USR_CJ_USERID).IsRequired(false);
                entity.Property(e => e.USR_CJ_ULTOPR).IsRequired(false);
                entity.Property(e => e.USR_CJ_DEBAJA).IsRequired(false);
                entity.Property(e => e.USR_CJ_HORMOV).IsRequired(false);
                entity.Property(e => e.USR_CJ_MODULE).IsRequired(false);
                entity.Property(e => e.USR_CJ_OALIAS).IsRequired(false);
                entity.Property(e => e.USR_CJ_LOTTRA).IsRequired(false);
                entity.Property(e => e.USR_CJ_LOTREC).IsRequired(false);
                entity.Property(e => e.USR_CJ_LOTORI).IsRequired(false);
                entity.Property(e => e.USR_CJ_SYSVER).IsRequired(false);
                entity.Property(e => e.USR_CJ_CMPVER).IsRequired(false);
                entity.Property(e => e.USR_CJRMVX_VALIDACION).IsRequired(false);
                //entity.Property(e => e.USR_CJRMVX_FECHAENVIO).IsRequired(false);
            });

            modelBuilder.Entity<ListadoPeriodosDTO>().HasNoKey().ToView(null);

        }

        //partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
