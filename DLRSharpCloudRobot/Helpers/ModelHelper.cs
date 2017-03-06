using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;
using Microsoft.Win32;

namespace DLRSharpCloudRobot.Helpers
{
    public class ModelHelper
    {
        public interface IPostDeserializeAction<T>
        {
            void OnPostDeserialization(T model);
        }
        public static T DeepClone<T>(T obj)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                DataContractSerializer serializer = new DataContractSerializer(typeof(T));
                serializer.WriteObject(stream, obj);
                stream.Position = 0;
                T copyObj = (T)serializer.ReadObject(stream);
                if (copyObj is IPostDeserializeAction<T>)
                    ((IPostDeserializeAction<T>)copyObj).OnPostDeserialization(copyObj);

                return copyObj;
            }
        }

        public static string Serialize<T>(T obj)
        {
            byte[] array;
            using (MemoryStream stream = new MemoryStream())
            {
                DataContractSerializer serializer = new DataContractSerializer(typeof(T));
                serializer.WriteObject(stream, obj);
                array = stream.ToArray();
            }
            string data = System.Text.Encoding.UTF8.GetString(array, 0, array.Length);
            return data;
        }


        public static T Deserialize<T>(string xml)
        {
            using (StringReader sr = new StringReader(xml))
            {
                XmlReader xr = XmlReader.Create(sr);
                DataContractSerializer serializer = new DataContractSerializer(typeof(T));
                T obj = (T)serializer.ReadObject(xr);
                if (obj is IPostDeserializeAction<T>)
                    ((IPostDeserializeAction<T>)obj).OnPostDeserialization(obj);
                return obj;
            }
        }


        /// <summary>
        /// Serializes the JSON.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj">The obj.</param>
        /// <returns></returns>
        public static string SerializeJSON<T>(T obj)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));

                try
                {
                    serializer.WriteObject(stream, obj);
                    stream.Position = 0;
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
            return "";
        }

        /// <summary>
        /// Deserializes the JSON.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="json">The json.</param>
        /// <returns></returns>
        public static T DeserializeJSON<T>(string json)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                ms.Position = 0;
                T obj = (T)serializer.ReadObject(ms);
                if (obj is IPostDeserializeAction<T>)
                    ((IPostDeserializeAction<T>)obj).OnPostDeserialization(obj);
                return obj;
            }
        }

        const string RegKey = "SOFTWARE\\SharpCloud\\API\\SC.DLRSharpCloud";
        public static string RegRead(string KeyName, string defVal)
        {
            // Opening the registry key
            RegistryKey rk = Registry.CurrentUser;
            // Open a subKey as read-only
            RegistryKey sk1 = rk.OpenSubKey(RegKey);
            // If the RegistrySubKey doesn't exist -> (null)
            {
                try
                {
                    // If the RegistryKey exists I get its value
                    // or null is returned.
                    if (sk1 != null)
                    {
                        var ret = (string)sk1.GetValue(KeyName.ToUpper());
                        if (ret == null)
                            return defVal;
                        return ret;
                    }
                }
                catch (Exception e)
                {
                    // AAAAAAAAAAARGH, an error!
                    //ShowErrorMessage(e, "Reading registry " + KeyName.ToUpper());
                }
            }
            return defVal;
        }

        public static bool RegWrite(string KeyName, object Value)
        {
            try
            {
                // Setting
                RegistryKey rk = Registry.CurrentUser;
                // I have to use CreateSubKey 
                // (create or open it if already exits), 
                // 'cause OpenSubKey open a subKey as read-only
                RegistryKey sk1 = rk.CreateSubKey(RegKey);
                // Save the value
                sk1.SetValue(KeyName.ToUpper(), Value);

                return true;
            }
            catch (Exception e)
            {
                // AAAAAAAAAAARGH, an error!
                //ShowErrorMessage(e, "Writing registry " + KeyName.ToUpper());
                return false;
            }
        }


    }
}
