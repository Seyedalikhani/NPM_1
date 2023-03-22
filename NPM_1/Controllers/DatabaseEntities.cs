using System.Data.Entity;
namespace NPM_1.Controllers
{
    public class DatabaseEntities : DbContext

    {   public DatabaseEntities() : base(@"Server=NAKPRG-NB1243\AHMAD; Database=AZWLL; Trusted_Connection=True;")
            {
            }


        public DbSet<Province> Departments;
      
    }
}