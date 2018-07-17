﻿using log4net;
using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Xml.Serialization;

namespace JuliusSweetland.OptiKey.Models
{
    [XmlRoot(ElementName = "Keyboard")]
    public class XmlKeyboard
    {
        public XmlKeyboard()
        {
            Name = "";
        }

        public XmlGrid Grid
        { get; set; }

        public XmlKeys Keys
        { get; set; }

        // The following are all optional
        public double? Height
        { get; set; }

        public string Name
        { get; set; }

        public string Symbol
        { get; set; }

        public double? SymbolMargin
        { get; set; }

        protected static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        [XmlIgnore]
        public Thickness? BorderThickness
        { get; set; }

        [XmlElement("BorderThickness")]
        public string BorderThicknessAsString
        {
            get { return BorderThickness.ToString(); }
            set {
                try
                {
                    ThicknessConverter thicknessConverter = new ThicknessConverter();
                    BorderThickness = (Thickness)thicknessConverter.ConvertFromString(value);
                }
                catch (System.FormatException)
                {
                    Log.ErrorFormat("Cannot interpret \"{0}\" as thickness", value);                
                }
            }
        }        
        
        [XmlIgnore]
        public bool Hidden
        { get; set; }

        [XmlElement("HideFromKeyboardMenu")]
        public string HiddenBoolAsString
        {
            get { return this.Hidden ? "True" : "False"; }
            set { this.Hidden = XmlUtils.ConvertToBoolean(value); }
        }        
        
        public static XmlKeyboard ReadFromFile(string inputFilename)
        {
            XmlKeyboard keyboard;

            // If no extension given, try ".xml"
            string ext = Path.GetExtension(inputFilename);
            bool exists = File.Exists(inputFilename);
            if (!File.Exists(inputFilename) &&
                String.IsNullOrEmpty(Path.GetExtension(inputFilename)))
            {
                inputFilename += ".xml";
            } 

            // Read in XML file (may throw)
            XmlSerializer serializer = new XmlSerializer(typeof(XmlKeyboard));
            using (FileStream readStream = new FileStream(@inputFilename, FileMode.Open))
            {
                keyboard = (XmlKeyboard)serializer.Deserialize(readStream);
            }

            return keyboard;
        }
    }
}