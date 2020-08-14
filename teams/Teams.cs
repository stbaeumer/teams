using System;
using System.Collections.Generic;
using System.Linq;

namespace teams
{
    public class Teams: List<Team>
    {
        internal void Kollegiumsteam(Lehrers lehrers)
        {
            List<Lehrer> kollegium = new List<Lehrer>();

            foreach (var lehrer in lehrers)
            {   
                kollegium.Add(lehrer);

                // Lehrer zur Office365-Gruppe hinzufügen

                Global.Streamwriter.WriteLine(@"Add-UnifiedGroupLinks -Identity 'Kollegium' -LinkType Members -Links '" + lehrer.Mail + "'");



            }

            this.Add(new Team("Kollegium", kollegium));
        }
    }
}