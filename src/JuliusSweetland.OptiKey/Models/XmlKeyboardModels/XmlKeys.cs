using System.Collections.Generic;
using System.Xml.Serialization;

namespace JuliusSweetland.OptiKey.Models
{
    public class XmlKeys
    {
        [XmlElement(ElementName = "ActionKey")]
        public List<XmlActionKey> ActionKeys
        { get; set; }

        [XmlElement(ElementName = "TextKey")]
        public List<XmlTextKey> TextKeys
        { get; set; }

        [XmlElement(ElementName = "ChangeKeyboardKey")]
        public List<XmlChangeKeyboardKey> ChangeKeyboardKeys
        { get; set; }

        [XmlIgnore]
        public int Count
        {
            get
            {
                return ActionKeys.Count + TextKeys.Count + ChangeKeyboardKeys.Count;
            }
        }
    }
}