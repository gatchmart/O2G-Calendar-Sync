using System.Data.Entity;


namespace Outlook_Calendar_Sync.Models
{
    public class ArchiverContext : DbContext
    {
        public DbSet<SyncPair> Pairs { get; set; }

        public DbSet<Identifier> Identifiers { get; set; }

        public ArchiverContext(string conn) : base (conn)
        {
            
        }
    }
}
