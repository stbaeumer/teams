using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace teams
{
    public class Raums : List<Raum>
    {
        public Raums()
        {
        }

        public Raums(int periode)
        {
            using (OleDbConnection oleDbConnection = new OleDbConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT Room.ROOM_ID, 
                                                    Room.Name,  
                                                    Room.Longname,
                                                    Room.Capacity
                                                    FROM Room
                                                    WHERE (((Room.SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] + ") AND ((Room.SCHOOL_ID)=177659) AND  ((Room.TERM_ID)=" + periode + "))";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        Raum raum = new Raum()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Raumnummer = Global.SafeGetString(oleDbDataReader, 1)
                        };

                        this.Add(raum);
                    };

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
                    Global.WriteLine("Räume", this.Count);
                }
            }
        }
    }
}