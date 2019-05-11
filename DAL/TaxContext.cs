using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using TaxManager.Models.Database;

namespace TaxManager.DAL
{
    public class TaxContext : DbContext
    {
        public TaxContext(DbContextOptions<TaxContext> options) : base(options)
        {
        }

        public DbSet<Tax> Taxes { get; set; }
        public DbSet<Municipality> Municapilities { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Tax>().ToTable("Tax");
            modelBuilder.Entity<Municipality>().ToTable("Municipality");
        }
    }
}