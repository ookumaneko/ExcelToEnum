using System;
using System.Text;
using System.Collections.Generic;
using System.IO;

namespace ExcelToEnum
{
    class OutputEnumData
    {
        public struct Data
        {
            public int Value;
            public string Name;

            public Data(int value, string name)
            {
                Value = value;
                Name = name;
            }
        }

        const string _TEMPLATE_FILE_NAME = "Template.txt";
        const string _VAR_ENUM_NAME = "$NAME$";
        const string _VAR_CONTENTS = "$ENUM$";
        const string _VAR_NAMESPACE = "$NAMESPACE$";

        List<Data> m_data;
        string m_templateText;
        string m_outputDirectory;

        public OutputEnumData(string outputDirectory, string settingDirectory)
        {
            m_data = new List<Data>();
            m_templateText = File.ReadAllText(settingDirectory + _TEMPLATE_FILE_NAME);
            m_outputDirectory = outputDirectory;
        }

        public void Clear()
        {
            m_data.Clear();
        }

        public void AddData(int value, string name)
        {
            m_data.Add(new Data(value, name));
        }

        public void Write(string enumName, string namespaceName)
        {
            string templateText = m_templateText; // 無駄な気はするけど面倒だから今はこれで
            templateText = templateText.Replace(_VAR_NAMESPACE, namespaceName);
            templateText = templateText.Replace(_VAR_ENUM_NAME, enumName);
            templateText = templateText.Replace(_VAR_CONTENTS, CreateEnumString());

            string filename = CreateFileName(enumName);
            Logger.WriteLine("Writing file...{0}", filename);
            File.WriteAllText(filename, templateText);
        }

        private string CreateFileName(string enumName)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(m_outputDirectory);
            sb.Append(enumName);
            sb.Append(Defines._EXTENSION_CS);
            return sb.ToString();   
        }

        private string CreateEnumString()
        {
            StringBuilder sb = new StringBuilder();
            int count = m_data.Count;
            for (int i = 0; i < count; ++i)
            {
                if (string.IsNullOrEmpty(m_data[i].Name))
                {
                    continue;                
                }

                sb.Append(m_data[i].Name);
                sb.Append(" = ");
                sb.Append(m_data[i].Value);
                sb.AppendLine(",");
            }

            return sb.ToString();
        }
    }
}
