using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace teams
{
    public class Anrechnungs : List<Anrechnung>
    {
        public Anrechnungs(string aktSj, string connectionString, int periode)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT 
CountValue.TEACHER_ID,  
Description.Name, 
CountValue.Text
FROM Description INNER JOIN CountValue ON Description.DESCRIPTION_ID = CountValue.DESCRIPTION_ID
WHERE (((CountValue.SCHOOLYEAR_ID)=" + aktSj + @") AND ((CountValue.Deleted)=False) AND ((CountValue.Deleted)=False))
ORDER BY CountValue.TEACHER_ID;
";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {                        
                        Anrechnung anrechnung = new Anrechnung()
                        {
                            TeacherIdUntis = oleDbDataReader.GetInt32(0),                            
                            Kategorie = Global.SafeGetString(oleDbDataReader, 1), // Bildungsgangleitung usw.                            
                            Grund = Global.SafeGetString(oleDbDataReader, 2) // 500, 510 usw.
                        };

                        if (anrechnung.TeacherIdUntis != 0 && !(from t in this                              
                             where t.TeacherIdUntis == anrechnung.TeacherIdUntis                              
                             where t.Grund == anrechnung.Grund
                             where t.Kategorie == anrechnung.Kategorie
                             select t).Any())
                        {
                            this.Add(anrechnung);
                        }                            
                    };
                    oleDbDataReader.Close();
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