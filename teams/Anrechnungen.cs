using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace teams
{
    public class Anrechnungs : List<Anrechnung>
    {
        public Anrechnungs(int periode)
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
WHERE (((CountValue.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] + @") AND ((CountValue.Deleted)=False) AND ((CountValue.Deleted)=False))
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
                            Beschr = Global.SafeGetString(oleDbDataReader, 1), // Bildungsgangleitung usw.                            
                            Text = Global.SafeGetString(oleDbDataReader, 2) == null ? "" : Global.SafeGetString(oleDbDataReader, 2) // Vorsitz etc.

                        };

                        if (anrechnung.Text.Contains("("))
                        {
                            anrechnung.TextGekürzt = anrechnung.Text.Substring(0, anrechnung.Text.IndexOf('(')).Trim();
                        }
                        else
                        {
                            anrechnung.TextGekürzt = anrechnung.Text.Trim();
                        }

                        if (anrechnung.TeacherIdUntis != 0 && !(from t in this                              
                             where t.TeacherIdUntis == anrechnung.TeacherIdUntis                              
                             where t.Text == anrechnung.Text                             where t.Beschr == anrechnung.Beschr
                             select t).Any())
                        {
                            this.Add(anrechnung);
                        }                            
                    };
                    oleDbDataReader.Close();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }
    }
}