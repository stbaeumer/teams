﻿using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace teams
{
    public class Klasses : List<Klasse>
    {
        public Klasses(int periode, Lehrers lehrers)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConnectionStringUntis))
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
FROM Class LEFT JOIN Teacher ON Class.TEACHER_ID = Teacher.TEACHER_ID WHERE (((Class.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] + ") AND (((Class.TERM_ID)=" + periode + ")) AND ((Teacher.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] + ") AND ((Teacher.TERM_ID)=" + periode + ")) OR (((Class.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] + ") AND ((Class.TERM_ID)=" + periode + ") AND ((Class.SCHOOL_ID)=177659) AND ((Teacher.SCHOOLYEAR_ID) Is Null) AND ((Teacher.TERM_ID) Is Null)) ORDER BY Class.Name ASC;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        List<Lehrer> klassenleitungen = new List<Lehrer>();

                        foreach (var klassenleitungIdUntis in (Global.SafeGetString(oleDbDataReader, 2)).Split(','))
                        {
                            var klassenleitung = (from l in lehrers
                                                  where l.IdUntis.ToString() == klassenleitungIdUntis
                                                  where l.Mail != null
                                                  where l.Mail != "" // Wer keine Mail hat, kann nicht Klassenleitung sein.
                                                  select l).FirstOrDefault();

                            if (klassenleitung != null)
                            {
                                klassenleitungen.Add(klassenleitung);
                            }                            
                        }

                        bool istVollzeit = istVollzeitKlasse(Global.SafeGetString(oleDbDataReader, 1));

                        Klasse klasse = new Klasse()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            NameUntis = Global.SafeGetString(oleDbDataReader, 1),
                            Klassenleitungen = klassenleitungen,
                            IstVollzeit = istVollzeit
                        };

                        this.Add(klasse);
                    };

                    Global.WriteLine("Klassen", this.Count);

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

        private bool istVollzeitKlasse(string klassenname)
        {
            var vollzeitBeginn = new List<string>() { "BS", "BW", "BT", "FM", "FS", "G", "HB" };
            
            foreach (var item in vollzeitBeginn)
            {
                if (klassenname.StartsWith(item))
                {
                    return true;
                }
            }
            return false;
        }
    }
}