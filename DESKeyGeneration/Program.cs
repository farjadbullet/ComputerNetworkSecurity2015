using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;

namespace DESKeyGeneration
{
    class Program
    {
        private static string _inputKey;
        private static string _initialPermutationResult;
        private static List<int> _ipTable = new List<int>();

        static void Main(string[] args)
        {
            // InitialPermutation();

            SplitKey().ForEach(Console.WriteLine);
        }

        private static List<string> SplitKey()
        {
            if (_initialPermutationResult?.Length > 0)
            {
                int length = _initialPermutationResult.Length;
                int halfLength = length / 2;
                string firstHalf = _initialPermutationResult.Substring(0, halfLength);
                string secondHalf = _initialPermutationResult.Substring(halfLength, halfLength);
                return new List<string>
                {
                    "\nFirst 32 Bits:" + firstHalf,
                    "\nLast 32 Bits:" + secondHalf
                };
                //return _ipTable.Select((x, i) => new { Index = i, Value = x }).GroupBy(x => x.Index / 2).Select(x => x.Select(v => v.Value).ToList()).ToList();
            }

            InitialPermutation();
            return SplitKey();
        }

        private static void InitialPermutation()
        {
            var excelData = new ExcelData("Resources/IP.xlsx");
            var ipTableData = excelData.GetData();

            foreach (var row in ipTableData)
            {
                _ipTable.AddRange(row.ItemArray.Select(t => int.Parse(t.ToString())));
            }
            Console.WriteLine("Please input 64 Bit Binary Key");
            _inputKey = Console.ReadLine();

            Console.WriteLine("Result from Initial Permutation");
            foreach (var index in _ipTable)
            {
                Console.Write(_inputKey?[index - 1]);
            }
            _ipTable.ForEach(m => _initialPermutationResult += _inputKey[m - 1]);
            Console.WriteLine();
        }
    }
    public class ExcelData
    {
        readonly string _path;

        public ExcelData(string path)
        {
            _path = path;
        }


        public IExcelDataReader GetExcelReader()
        {
            // ExcelDataReader works with the binary Excel file, so it needs a FileStream
            // to get started. This is how we avoid dependencies on ACE or Interop:
            FileStream stream = File.Open(_path, FileMode.Open, FileAccess.Read);

            // We return the interface, so that 
            IExcelDataReader reader = null;
            try
            {
                if (_path.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                if (_path.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                return reader;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public IEnumerable<string> GetWorksheetNames()
        {
            var reader = this.GetExcelReader();
            var workbook = reader.AsDataSet();
            var sheets = from DataTable sheet in workbook.Tables select sheet.TableName;
            return sheets;
        }
        public IEnumerable<DataRow> GetData(bool firstRowIsColumnNames = true)
        {
            var reader = this.GetExcelReader();
            reader.IsFirstRowAsColumnNames = firstRowIsColumnNames;
            //var workSheet = reader.AsDataSet().Tables[sheet];
            var workSheet = reader.AsDataSet().Tables[0];
            var rows = from DataRow a in workSheet.Rows select a;
            return rows;
        }
    }
}
