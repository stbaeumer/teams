using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Text;

namespace teams
{
    public class Unterrichtsgruppes : List<Unterrichtsgruppe>
    {
        public Unterrichtsgruppes(string aktSj, string connectionString)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(connectionString))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT
LessonGroup.LESSON_GROUP_ID, 
LessonGroup.Name,
LessonGroup.DateFrom,
LessonGroup.DateTo,
LessonGroup.InrerruptionsFrom,
LessonGroup.InrerruptionsTo
FROM LessonGroup
WHERE (((SCHOOLYEAR_ID)= " + aktSj + ") AND ((LessonGroup.SCHOOL_ID)=177659)) ORDER BY LESSON_GROUP_ID;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    Console.WriteLine("");
                    Console.WriteLine("Unterrichtsgruppen");
                    Console.WriteLine("------------------");

                    while (oleDbDataReader.Read())
                    {
                        Interruption interruption = new Interruption();

                        foreach (var date in (Global.SafeGetString(oleDbDataReader, 4)).Split(','))
                        {
                            if (date != "")
                            {
                                interruption.von.Add(DateTime.ParseExact(date, "yyyyMMdd", CultureInfo.InvariantCulture));
                            }
                        }

                        foreach (var date in (Global.SafeGetString(oleDbDataReader, 5)).Split(','))
                        {
                            if (date != "")
                            {
                                interruption.bis.Add(DateTime.ParseExact(date, "yyyyMMdd", CultureInfo.InvariantCulture));
                            }
                        }

                        // Nach DateTo und vor DateFrom wird alles zur Interruption

                        interruption.von.Add(new DateTime(Convert.ToInt32(aktSj.Substring(0, 4)), 8, 1));
                        interruption.bis.Add(DateTime.ParseExact((oleDbDataReader.GetInt32(2)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture));

                        interruption.von.Add(DateTime.ParseExact((oleDbDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture));
                        interruption.bis.Add(new DateTime(Convert.ToInt32(aktSj.Substring(4, 4)), 7, 31));

                        Unterrichtsgruppe unterrichtsgruppe = new Unterrichtsgruppe()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Name = Global.SafeGetString(oleDbDataReader, 1),
                            Von = DateTime.ParseExact((oleDbDataReader.GetInt32(2)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Bis = DateTime.ParseExact((oleDbDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Interruption = interruption
                        };

                        Console.WriteLine(" " + unterrichtsgruppe.IdUntis.ToString().PadLeft(3) + " " + unterrichtsgruppe.Name.PadRight(8) + " " + unterrichtsgruppe.Von.ToShortDateString() + " - " + unterrichtsgruppe.Bis.ToShortDateString());

                        // Bei 1.HJ werden alle Unterrichte des 2.HJ als Interruption eingetragen

                        if (unterrichtsgruppe.Name == "1.HJ")
                        {
                            unterrichtsgruppe.Interruption.von.Add(unterrichtsgruppe.Bis);
                            unterrichtsgruppe.Interruption.bis.Add(new DateTime(unterrichtsgruppe.Bis.Year, 7, 31));
                        }

                        if (unterrichtsgruppe.Name == "2.HJ")
                        {
                            unterrichtsgruppe.Interruption.von.Add(new DateTime(unterrichtsgruppe.Von.AddYears(-1).Year, 8, 1));
                            unterrichtsgruppe.Interruption.bis.Add(unterrichtsgruppe.Von.AddDays(-1));
                        }

                        if (unterrichtsgruppe.Name == "U")
                        {
                            for (DateTime date = unterrichtsgruppe.Von; date.Date <= unterrichtsgruppe.Bis; date = date.AddDays(1))
                            {
                                int kw = (CultureInfo.CurrentCulture).Calendar.GetWeekOfYear(date, (CultureInfo.CurrentCulture).DateTimeFormat.CalendarWeekRule, (CultureInfo.CurrentCulture).DateTimeFormat.FirstDayOfWeek);

                                if (date.DayOfWeek == DayOfWeek.Monday)
                                {
                                    if (kw % 2 == 0)
                                    {
                                        unterrichtsgruppe.Interruption.von.Add(date);
                                    }
                                }
                                if (date.DayOfWeek == DayOfWeek.Sunday && unterrichtsgruppe.Interruption.von.Count > 0)
                                {
                                    if (kw % 2 == 0)
                                    {
                                        unterrichtsgruppe.Interruption.bis.Add(date);
                                    }
                                }
                            }
                        }

                        if (unterrichtsgruppe.Name == "G")
                        {
                            for (DateTime date = unterrichtsgruppe.Von; date.Date <= unterrichtsgruppe.Bis; date = date.AddDays(1))
                            {
                                int kw = (CultureInfo.CurrentCulture).Calendar.GetWeekOfYear(date, (CultureInfo.CurrentCulture).DateTimeFormat.CalendarWeekRule, (CultureInfo.CurrentCulture).DateTimeFormat.FirstDayOfWeek);

                                if (date.DayOfWeek == DayOfWeek.Monday)
                                {
                                    if (kw % 2 == 1)
                                    {
                                        unterrichtsgruppe.Interruption.von.Add(date);
                                    }
                                }
                                if (date.DayOfWeek == DayOfWeek.Sunday && unterrichtsgruppe.Interruption.von.Count > 0)
                                {
                                    if (kw % 2 == 1)
                                    {
                                        unterrichtsgruppe.Interruption.bis.Add(date);
                                    }
                                }
                            }
                        }

                        this.Add(unterrichtsgruppe);
                    };
                    oleDbDataReader.Close();
                    Console.WriteLine("");
                    File.AppendAllLines(Global.TeamsPs, new List<string>() { "# Anzahl Unterrichtsgruppen : " + this.Count + "" }, Encoding.UTF8);

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