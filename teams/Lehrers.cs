﻿using System;
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
                            Mail = Global.SafeGetString(oleDbDataReader, 4),
                            Deputat = Convert.ToDouble(oleDbDataReader.GetInt32(5)) / 1000,
                            Geschlecht = Global.SafeGetString(oleDbDataReader, 6) == "W" ? "w" : "m",
                            Anrechnungen = (from a in anrechnungen where a.TeacherIdUntis == oleDbDataReader.GetInt32(0) select a).ToList()
                        };

                        if (lehrer.Mail.Contains("ertrud"))
                        {
                            lehrer.Mail = lehrer.Mail.Replace("ertrud", "erti");
                        }

                        if (lehrer.Deputat > 0 && lehrer.Mail != null && lehrer.Mail != "")
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

    }
}