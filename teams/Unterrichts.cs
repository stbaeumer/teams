
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace teams
{
    public class Unterrichts : List<Unterricht>
    {
        public Unterrichts()
        {
        }

        public Unterrichts(int periode, Klasses klasses, Lehrers lehrers, Fachs fachs, Raums raums, Unterrichtsgruppes unterrichtsgruppes)
        {
            DateTime datumMontagDerKalenderwoche = new DateTime(2020, 08, 10); //GetMondayDateOfWeek(kalenderwoche, DateTime.Now.Year);
            DateTime datumErsterTagDesPrüfZyklus = new DateTime(2020, 08, 10);

            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConnectionStringUntis))
            {
                int id = 0;

                try
                {
                    string queryString = @"SELECT DISTINCT 
Lesson_ID,
LessonElement1,
Periods,
Lesson.LESSON_GROUP_ID,
Lesson_TT,
Flags,
DateFrom,
DateTo
FROM LESSON
WHERE (((SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] + ") AND ((TERM_ID)=" + periode + ") AND ((Lesson.SCHOOL_ID)=177659) AND (((Lesson.Deleted)=No))) ORDER BY LESSON_ID;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    
                    while (oleDbDataReader.Read())
                    {
                        id = oleDbDataReader.GetInt32(0);

                        string wannUndWo = Global.SafeGetString(oleDbDataReader, 4);

                        var zur = wannUndWo.Replace("~~", "|").Split('|');

                        ZeitUndOrts zeitUndOrts = new ZeitUndOrts();

                        for (int i = 0; i < zur.Length; i++)
                        {
                            if (zur[i] != "")
                            {
                                var zurr = zur[i].Split('~');

                                int tag = 0;
                                int stunde = 0;
                                List<string> raum = new List<string>();

                                try
                                {
                                    tag = Convert.ToInt32(zurr[1]);
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Der Unterricht " + id + " hat keinen Tag.");
                                }

                                try
                                {
                                    stunde = Convert.ToInt32(zurr[2]);
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Der Unterricht " + id + " hat keine Stunde.");
                                }

                                try
                                {
                                    var ra = zurr[3].Split(';');

                                    foreach (var item in ra)
                                    {
                                        if (item != "")
                                        {
                                            raum.AddRange((from r in raums
                                                           where item.Replace(";", "") == r.IdUntis.ToString()
                                                           select r.Raumnummer));
                                        }
                                    }

                                    if (raum.Count == 0)
                                    {
                                        raum.Add("");
                                    }
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Der Unterricht " + id + " hat keinen Raum.");
                                }


                                ZeitUndOrt zeitUndOrt = new ZeitUndOrt(tag, stunde, raum);
                                zeitUndOrts.Add(zeitUndOrt);
                            }
                        }

                        string lessonElement = Global.SafeGetString(oleDbDataReader, 1);

                        int anzahlGekoppelterLehrer = lessonElement.Count(x => x == '~') / 21;

                        List<string> klassenKürzel = new List<string>();

                        for (int i = 0; i < anzahlGekoppelterLehrer; i++)
                        {
                            var lesson = lessonElement.Split(',');

                            var les = lesson[i].Split('~');
                            string lehrer = les[0] == "" ? null : (from l in lehrers where l.IdUntis.ToString() == les[0] select l.Kürzel).FirstOrDefault();

                            if (lehrer == "KU")
                            {
                                string a = "";
                            }
                            string fach = les[2] == "0" ? "" : (from f in fachs where f.IdUntis.ToString() == les[2] select f.KürzelUntis).FirstOrDefault();

                            string raumDiesesUnterrichts = "";
                            if (les[3] != "")
                            {
                                raumDiesesUnterrichts = (from r in raums where (les[3].Split(';')).Contains(r.IdUntis.ToString()) select r.Raumnummer).FirstOrDefault();
                            }

                            int anzahlStunden = oleDbDataReader.GetInt32(2);

                            var unterrichtsgruppeDiesesUnterrichts = (from u in unterrichtsgruppes where u.IdUntis == oleDbDataReader.GetInt32(3) select u).FirstOrDefault();

                            if (les.Count() >= 17)
                            {
                                foreach (var kla in les[17].Split(';'))
                                {
                                    Klasse klasse = new Klasse();

                                    if (kla != "")
                                    {
                                        if (!(from kl in klassenKürzel
                                              where kl == (from k in klasses
                                                           where k.IdUntis == Convert.ToInt32(kla)
                                                           select k.NameUntis).FirstOrDefault()
                                              select kl).Any())
                                        {
                                            klassenKürzel.Add((from k in klasses
                                                               where k.IdUntis == Convert.ToInt32(kla)
                                                               select k.NameUntis).FirstOrDefault());
                                        }
                                    }
                                }
                            }
                            else
                            {
                            }

                            if (lehrer != null)
                            {
                                for (int z = 0; z < zeitUndOrts.Count; z++)
                                {
                                    // Wenn zwei Lehrer gekoppelt sind und zwei Räume zu dieser Stunde gehören, dann werden die Räume entsprechend verteilt.

                                    string r = zeitUndOrts[z].Raum[0];
                                    try
                                    {
                                        if (anzahlGekoppelterLehrer > 1 && zeitUndOrts[z].Raum.Count > 1)
                                        {
                                            r = zeitUndOrts[z].Raum[i];
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        if (anzahlGekoppelterLehrer > 1 && zeitUndOrts[z].Raum.Count > 1)
                                        {
                                            r = zeitUndOrts[z].Raum[0];
                                        }
                                    }

                                    string k = "";

                                    foreach (var item in klassenKürzel)
                                    {
                                        k += item + ",";
                                    }

                                    // Nur wenn der tagDesUnterrichts innerhalb der Befristung stattfindet, wird er angelegt

                                    DateTime von = DateTime.ParseExact((oleDbDataReader.GetInt32(6)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                                    DateTime bis = DateTime.ParseExact((oleDbDataReader.GetInt32(7)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                                    DateTime tagDesUnterrichts = datumErsterTagDesPrüfZyklus.AddDays(zeitUndOrts[z].Tag - 1);

                                    if (von <= tagDesUnterrichts && tagDesUnterrichts <= bis)
                                    {
                                        Unterricht unterricht = new Unterricht(
                                            id,
                                            lehrer,
                                            fach,
                                            k.TrimEnd(','),
                                            r,
                                            "",
                                            zeitUndOrts[z].Tag,
                                            zeitUndOrts[z].Stunde,
                                            unterrichtsgruppeDiesesUnterrichts,
                                            datumErsterTagDesPrüfZyklus);
                                        this.Add(unterricht);
                                        try
                                        {
                                            string ugg = unterrichtsgruppeDiesesUnterrichts == null ? "" : unterrichtsgruppeDiesesUnterrichts.Name;
                                           // Console.WriteLine(unterricht.Id.ToString().PadLeft(4) + " " + unterricht.LehrerKürzel.PadRight(4) + unterricht.KlasseKürzel.PadRight(20) + unterricht.FachKürzel.PadRight(10) + " Raum: " + r.PadRight(10) + " Tag: " + unterricht.Tag + " Stunde: " + unterricht.Stunde + " " + ugg.PadLeft(3));
                                        }
                                        catch (Exception ex)
                                        {
                                            throw ex;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    Global.WriteLine("Unterrichte", this.Count);

                    oleDbDataReader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Fehler beim Unterricht mit der ID " + id + "\n" + ex.ToString());
                    throw new Exception("Fehler beim Unterricht mit der ID " + id + "\n" + ex.ToString());
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }

        internal List<Unterricht> SortierenUndKumulieren()
        {
            try
            {
                // Die Unterrichte werden chronologisch sortiert

                List<Unterricht> sortierteUnterrichts = (from u in this
                                                         orderby u.KlasseKürzel, u.FachKürzel, u.Raum, u.Tag, u.Stunde
                                                         select u).ToList();

                for (int i = 0; i < sortierteUnterrichts.Count; i++)
                {
                    // Wenn es einen nachfolgenden Unterricht gibt ...

                    if (i < sortierteUnterrichts.Count - 1)
                    {
                        // ... und dieser in allen Eigenschaften identisch ist ...

                        if (sortierteUnterrichts[i].KlasseKürzel == sortierteUnterrichts[i + 1].KlasseKürzel && sortierteUnterrichts[i].FachKürzel == sortierteUnterrichts[i + 1].FachKürzel && sortierteUnterrichts[i].Raum == sortierteUnterrichts[i + 1].Raum)
                        {
                            // ... und der nachfolgende Unterricht unmittelbar (nach der Pause) anschließt ... 

                            if (sortierteUnterrichts[i].Bis == sortierteUnterrichts[i + 1].Von || sortierteUnterrichts[i].Bis.AddMinutes(15) == sortierteUnterrichts[i + 1].Von || sortierteUnterrichts[i].Bis.AddMinutes(20) == sortierteUnterrichts[i + 1].Von)
                            {
                                // ... wird der Beginn des Nachfolgers nach vorne geschoben ...

                                sortierteUnterrichts[i + 1].Von = sortierteUnterrichts[i].Von;

                                // ... und der Vorgänger wird gelöscht.

                                sortierteUnterrichts.RemoveAt(i);

                                // Der Nachfolger bekommt den Index des Vorgängers.
                                i--;
                            }
                        }
                    }
                }
                return (
                    from s in sortierteUnterrichts
                    orderby s.Tag, s.Stunde, s.KlasseKürzel
                    select s).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception(ex.ToString());
            }
        }
        
        public Unterrichts Kumulieren()
        {
            try
            {
                for (int i = 0; i < this.Count; i++)
                {
                    // Wenn es einen nachfolgenden Unterricht gibt ...

                    if (i < this.Count - 1)
                    {
                        // ... und dieser in allen Eigenschaften identisch ist ...

                        if (this[i].KlasseKürzel == this[i + 1].KlasseKürzel && this[i].FachKürzel == this[i + 1].FachKürzel && this[i].Raum == this[i + 1].Raum)
                        {
                            // ... und der nachfolgende Unterricht unmittelbar (nach der Pause) anschließt ... 

                            if (this[i].Bis == this[i + 1].Von || this[i].Bis.AddMinutes(15) == this[i + 1].Von || this[i].Bis.AddMinutes(20) == this[i + 1].Von)
                            {
                                // ... wird der Beginn des Nachfolgers nach vorne geschoben ...

                                this[i + 1].Von = this[i].Von;

                                // ... und der Vorgänger wird gelöscht.

                                this.RemoveAt(i);

                                // Der Nachfolger bekommt den Index des Vorgängers.
                                i--;
                            }
                        }
                    }
                }
                return this;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception(ex.ToString());
            }
        }
    }
}