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

        public Lehrers(int periode)
        {
            Anrechnungs anrechnungen = new Anrechnungs(periode);

            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
Teacher.Teacher_ID, 
Teacher.Name, 
Teacher.Longname, 
Teacher.FirstName,
Teacher.Email,
Teacher.PlannedWeek,
Teacher.Flags
FROM Teacher 
WHERE (((SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] + ") AND  ((TERM_ID)=" + periode + ") AND ((Teacher.SCHOOL_ID)=177659) AND (((Teacher.Deleted)=No))) ORDER BY Teacher.Name;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {                        
                        Lehrer lehrer = new Lehrer()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Kürzel = Global.SafeGetString(oleDbDataReader, 1),
                            Nachname = Global.SafeGetString(oleDbDataReader, 2),
                            Vorname = Global.SafeGetString(oleDbDataReader, 3),
                            Mail = Global.SafeGetString(oleDbDataReader, 4),
                            Deputat = Convert.ToDouble(oleDbDataReader.GetInt32(5)) / 1000,
                            Geschlecht = Global.SafeGetString(oleDbDataReader, 6).Contains("W") ? "w" : "m",                            
                            Anrechnungen = (from a in anrechnungen where a.TeacherIdUntis == oleDbDataReader.GetInt32(0) select a).ToList()
                        };

                        lehrer.Mail = lehrer.Mail.Contains("ertrud") ? lehrer.Mail.Replace("ertrud", "erti") : lehrer.Mail;

                        if (lehrer.Mail != "")
                        {
                            if (!lehrer.Mail.ToUpper().StartsWith(lehrer.Vorname.Substring(0, 1)) || !lehrer.Mail.ToUpper().Contains("." + lehrer.Nachname.ToUpper().Substring(0, 1)))
                            {
                                throw new Exception("Die Mail des Lehrers " + lehrer.Kürzel + " lautet " + lehrer.Mail + ". Das kann nicht richtig sein.");
                            }
                            this.Add(lehrer);
                        }
                        if (lehrer.Mail != "")                            
                        {                            
                            this.Add(lehrer);
                        }                        
                    };

                    Global.WriteLine("Lehrer", this.Count);

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

        internal bool istLehrer(string istMember)
        {
            foreach (var item in this)
            {
                if (item.Mail == istMember)
                {
                    return true;
                }
            }
            return false;
        }
    }
}