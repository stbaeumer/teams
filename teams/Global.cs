using System.Data.OleDb;
using System.IO;

namespace teams
{
    public static class Global
    {
        public const string ConnectionStringUntis = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis9;uid=DBA";

        public static StreamWriter Streamwriter;

        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }
    }
}