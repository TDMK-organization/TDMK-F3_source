using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace OK2SHIP_SMT.Repositories
{
    public class TDMK_BinaryFile
    {
        public string _FILENAME { get; set; }
        public byte[] ConvertDataSetToByteArray(DataTable dataTable)
        {
            byte[] binaryDataResult = null;
            using (MemoryStream memStream = new MemoryStream())
            {
                BinaryFormatter brFormatter = new BinaryFormatter();
                dataTable.RemotingFormat = SerializationFormat.Binary;
                brFormatter.Serialize(memStream, dataTable);
                binaryDataResult = memStream.ToArray();
            }
            return binaryDataResult;
        }
        public bool ByteArrayToFile(byte[] byteArray)
        {
            System.IO.File.WriteAllBytes(_FILENAME, byteArray);
            return true;
        }
        public static void SaveByteArrayToFileWithFileStream(byte[] data, string filePath)
        {
            var stream = File.Create(filePath);
            stream.Write(data, 0, data.Length);
        }
    }
}
