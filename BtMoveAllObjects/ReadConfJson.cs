
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading.Tasks;

namespace BtMoveAllObjects
{
    public struct JsonConfigs
    {
        public string Unit;
        public double X;
        public double Y;
    }


    static class ReadConfJson
    {
        static private string confUnits;
        static private double confX;
        static private double confY;

        static public JsonConfigs confUXY
        {
            get
            {
                return new JsonConfigs { Unit = confUnits, X = confX, Y = confY };
            }
        }

        public static async Task<JsonConfigs> ReadAsync()
        {
            await Task.Run(() =>
            {
                var currentDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\conf.json";
                using (var fs = new FileStream(currentDirectory, FileMode.Open, FileAccess.Read))
                {
                    var s = new DataContractJsonSerializer(typeof(Conf));
                    var p = s.ReadObject(fs) as Conf;

                    confUnits = p.Units;
                    confX = p.X;
                    confY = p.Y;
                }
            });
            
            return new JsonConfigs { Unit = confUnits, X = confX, Y = confY };
        }


        [DataContract]
        public class Conf
        {
            [DataMember]
            public string Units { get; set; }
            [DataMember]
            public double X { get; set; }
            [DataMember]
            public double Y { get; set; }
        }

    }

}
