using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Text;

namespace teams
{
    public class Periodes : List<Periode>
    {
        public Periodes()
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT
Terms.TERM_ID, 
Terms.Name, 
Terms.Longname, 
Terms.DateFrom, 
Terms.DateTo
FROM Terms
WHERE (((Terms.SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] + ")  AND ((Terms.SCHOOL_ID)=177659)) ORDER BY Terms.TERM_ID;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        Periode periode = new Periode()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Name = Global.SafeGetString(oleDbDataReader, 1),
                            Langname = Global.SafeGetString(oleDbDataReader, 2),
                            Von = DateTime.ParseExact((oleDbDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Bis = DateTime.ParseExact((oleDbDataReader.GetInt32(4)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                        };

                        if (DateTime.Now > periode.Von && DateTime.Now < periode.Bis)
                            this.AktuellePeriode = periode.IdUntis;

                        this.Add(periode);
                    };

                    // Korrektur des Periodenendes

                    for (int i = 0; i < this.Count - 1; i++)
                    {
                        this[i].Bis = this[i + 1].Von.AddDays(-1);
                    }

                    oleDbDataReader.Close();

                    if (this.AktuellePeriode == 0)
                    {
                        Console.WriteLine("Es kann keine aktuelle Periode ermittelt werden. Das ist z. B. während der Sommerferien der Fall. Es wird die Periode " + this.Count + " als aktuelle Periode angenommen.");
                        this.AktuellePeriode = this.Count;
                    }
                    else
                    {
                        Global.WriteLine("Perioden", this.Count);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }

        public int AktuellePeriode { get; private set; }
    }
}