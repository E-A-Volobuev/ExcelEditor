using System.IO;
using System.Threading.Tasks;

namespace ExcelEditor
{
    public class Work
    {
        public async Task StartProcess(string template,string currentCatalog,string pathResult)
        {
            var files = Directory.GetFiles(currentCatalog);
            ExcelWork work = new ExcelWork();

            foreach(var file in files)
            {
                var list= await Task.Run(() => work.GetRawData(file));
                string fileName = NameFileHelper(file);
                await Task.Run(() => work.WriteBook(template, list, pathResult, fileName)); 
            }
        }

        private string NameFileHelper(string file)
        { 
            string ch = @"\";
            int indexOfChar = file.LastIndexOf(ch);
            string nameFile = file.Substring(indexOfChar);

            return nameFile;
        }

    }
}
