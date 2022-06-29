using System.Collections.Generic;

namespace ExcelEditor.Model
{
    public class ExcelHelperModel
    {
        public int IndexSheet { get; set; }
        public int IndexRowStart { get; set; }
        public List<int> IndexColumns { get; set; }
        public int IndexRowEnd { get; set; }
    }
}
