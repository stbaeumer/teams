using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace teams
{
    public class Klasses : List<Klasse>
    {
        public Klasses(string aktSj, Lehrers lehrers, string connectionString, int periode)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(connectionString))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
Class.Class_ID, 
Class.Name,
Class.TeacherIds,
Class.Longname, 
Teacher.Name, 
Class.ClassLevel,
Class.PERIODS_TABLE_ID
FROM Class LEFT JOIN Teacher ON Class.TEACHER_ID = Teacher.TEACHER_ID WHERE (((Class.SCHOOLYEAR_ID)=" + aktSj + ") AND (((Class.TERM_ID)=" + periode + ")) AND ((Teacher.SCHOOLYEAR_ID)=" + aktSj + ") AND ((Teacher.TERM_ID)=" + periode + ")) OR (((Class.SCHOOLYEAR_ID)=" + aktSj + ") AND ((Class.TERM_ID)=" + periode + ") AND ((Class.SCHOOL_ID)=177659) AND ((Teacher.SCHOOLYEAR_ID) Is Null) AND ((Teacher.TERM_ID) Is Null)) ORDER BY Class.Name ASC;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        List<Lehrer> klassenleitungen = new List<Lehrer>();

                        foreach (var item in (Global.SafeGetString(oleDbDataReader, 2)).Split(','))
                        {
                            klassenleitungen.Add((from l in lehrers
                                                  where l.IdUntis.ToString() == item
                                                  select l).FirstOrDefault());
                        }

                        Klasse klasse = new Klasse()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            NameUntis = Global.SafeGetString(oleDbDataReader, 1),
                            Klassenleitungen = klassenleitungen
                        };

                        this.Add(klasse);
                    };

                    Console.WriteLine(("Klassen " + ".".PadRight(this.Count / 150, '.')).PadRight(48, '.') + (" " + this.Count).ToString().PadLeft(4), '.');

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