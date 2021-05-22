using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace teams
{
    public class Anrechnungsgrunds : List<Anrechnungsgrund>
    {
        public Anrechnungsgrunds(string aktSj)
        {
            string topic = "Anrechnungsgründe";
            string fehler = "";

            Console.WriteLine(topic + " ...");

            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT CV_Reason.CV_REASON_ID, CV_Reason.Name, CV_Reason.Longname, CV_Reason.DESCRIPTION_ID, CV_Reason.SortId
WHERE DESCRIPTION_ID = 99
FROM CV_Reason WHERE (((CV_Reason.SCHOOLYEAR_ID)= " + aktSj + ")) ORDER BY CV_Reason.SortId;";


                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        Anrechnungsgrund anrechnungsgrund = new Anrechnungsgrund()
                        {
                            UntisId = oleDbDataReader.GetInt32(0),
                            Nummer = Global.SafeGetString(oleDbDataReader, 1),
                            Beschreibung = Global.SafeGetString(oleDbDataReader, 2),
                            StatistikName = oleDbDataReader.GetInt32(3) == 51 ? 67 : oleDbDataReader.GetInt32(3) == 50 ? 66 : 65
                        };
                        this.Add(anrechnungsgrund);
                    };
                }
                catch (Exception ex)
                {
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }
    }
}