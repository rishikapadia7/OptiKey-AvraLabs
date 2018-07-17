using System.Reflection;
using log4net;

namespace JuliusSweetland.OptiKey.Models
{
    public class XmlKey
    {
        protected static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public XmlKey()
        {
            Width = 1;
            Height = 1;
            Label = "";
        }

        public int Row
        { get; set; }

        public int Col
        { get; set; }

        public string Label
        { get; set; }

        public string Symbol
        { get; set; }

        public int Width
        { get; set; }

        public int Height
        { get; set; }

    }
}