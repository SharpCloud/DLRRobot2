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
    public class StoryLite2 
    {
        [DataMember]
        public string Id { get; set; }
        [DataMember]
        public string ImageId { get; set; }
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public string Description { get; set; }  

        public bool IsSelected { get; set; }

        public StoryLite2(StoryLite storyLite)
        {
            Id = storyLite.Id;
            ImageId = storyLite.ImageId;
            Name = storyLite.Name;
            Description = storyLite.Description;
        }

    }
}
