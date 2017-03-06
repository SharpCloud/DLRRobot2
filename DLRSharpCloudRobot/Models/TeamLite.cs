using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using SC.API.ComInterop.Models;

namespace DLRSharpCloudRobot.Models
{
    [DataContract]
    public class TeamLite 
    {
        [DataMember]
        public string Id { get; set; }
        [DataMember]
        public string ImageUrl { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public string Description { get; set; }

        public TeamLite(Team team)
        {
            Id = team.Id;
            ImageUrl = team.ImageUrl;
            Name = team.Name;
            Description = team.Description;
        }

        public TeamLite(TeamLite teamLite)
        {
            Id = teamLite.Id;
            ImageUrl = teamLite.ImageUrl;
            Name = teamLite.Name;
            Description = teamLite.Description;
        }
    }
}
