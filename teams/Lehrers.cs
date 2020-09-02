using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace teams
{
    public class Lehrers : List<Lehrer>
    {
        public Lehrers()
        {
        }

        public Lehrers(string aktSj, string connectionString, Periodes periodes)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(connectionString))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
Teacher.Teacher_ID, 
Teacher.Name, 
Teacher.Longname, 
Teacher.FirstName,
Teacher.Email,
Teacher.PlannedWeek
FROM Teacher 
WHERE (((SCHOOLYEAR_ID)= " + aktSj + ") AND  ((TERM_ID)=" + periodes.Count + ") AND ((Teacher.SCHOOL_ID)=177659) AND (((Teacher.Deleted)=No))) ORDER BY Teacher.Name;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        Lehrer lehrer = new Lehrer()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Kürzel = Global.SafeGetString(oleDbDataReader, 1),
                            Mail = Global.SafeGetString(oleDbDataReader, 4)
                        };

                        if (lehrer.Mail.Contains("ertrud"))
                        {
                            lehrer.Mail = lehrer.Mail.Replace("ertrud", "erti");
                        }

                        this.Add(lehrer);
                    };

                    Console.WriteLine(("Lehrer*innen " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');
                    File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Anzahl Lehrer*innen : " + this.Count + "" }, Encoding.UTF8);

                    oleDbDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw new Exception(ex.ToString());
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }

    }
}