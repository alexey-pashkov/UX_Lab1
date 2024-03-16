using Aspose.Words;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Xml.Serialization;
using wfp_lab1.Services;
using System.Windows.Documents;
using System.Reflection;

namespace wfp_lab1.Models
{
    public class Participant : INotifyDataErrorInfo, IWordPrintable
    {
        
        private string _orgName;
        [PropertyName("Название организации")]
        public string OrgName 
        {
            get { return _orgName; }
            set 
            { 
                _orgName = value;

                RemoveError();
                if (_orgName == "")
                {
                    AddError("Название организации не должно быть пустым!");
                }
            }
        }
        private string _respPerson;
        [PropertyName("Ответственное лицо")]
        public string RespPerson
        {
            get { return _respPerson; }
            set
            {
                _respPerson = value;
            }
        }
        private string _country;
        [PropertyName("Страна")]
        public string Country
        {
            get { return _country; }
            set
            {
                _country = value;
            }
        }
        private string _phoneNum;
        [PropertyName("Номер телефона")]
        public string PhomeNum
        {
            get { return _phoneNum; }
            set
            {
                _phoneNum = value;
            }
        }
        private string _email;
        [PropertyName("Адрес эл. почты")]
        public string Email
        {
            get { return _email; }
            set
            {
                _email = value;
            }
        }
        private int _exhibArea;
        [PropertyName("Площадь для экспозиции")]
        public int ExhibArea
        {
            get { return _exhibArea; }
            set
            {
                _exhibArea = value;

                RemoveError();
                if (this._exhibArea <= 0 || this._exhibArea > 15)
                {
                    AddError("Указана недопустимая площадь для экспозиции!");
                }
            }
        }

        public static Participant LoadFromXml(string xmlFilePath)
        {
            Participant participant;

            XmlSerializer serializer = new XmlSerializer(typeof(Participant));

            using (FileStream fs = new FileStream(xmlFilePath, FileMode.Open))
            {
                try
                {
                    participant = (Participant)serializer.Deserialize(fs);
                }
                catch
                {
                    participant = new Participant();
                }
            }

            return participant;
        }

        public void SaveToXml(string xmlFilePath)
        {
            Participant participant;

            XmlSerializer serializer = new XmlSerializer(typeof(Participant));

            using (FileStream fs = new FileStream(xmlFilePath, FileMode.CreateNew))
            {
                serializer.Serialize(fs, this);
            }
        }

        #region INotifyDataErrorInfo

        private Dictionary<string, List<string>> _errors = new Dictionary<string, List<string>>();

        public event EventHandler<DataErrorsChangedEventArgs>? ErrorsChanged;

        public bool HasErrors => _errors.Any();

        public IEnumerable GetErrors([CallerMemberName] string propertyName = "")
        {
            if (this._errors.ContainsKey(propertyName))
            {
                return this._errors[propertyName];
            }
            return Array.Empty<string>();
        }

        private void AddError(string errorMessage, [CallerMemberName] string propertyName = "")
        {
            if (!_errors.ContainsKey(propertyName))
            {
                _errors[propertyName] = new List<string>();
            }

            _errors[propertyName].Add(errorMessage);

            OnErrorsChanged(propertyName);

        }

        private void RemoveError([CallerMemberName] string propertyName = "")
        {
            if (_errors.ContainsKey(propertyName))
            {
                _errors.Remove(propertyName);
                OnErrorsChanged(propertyName);
            }
        }

        private void OnErrorsChanged([CallerMemberName] string propertyName = "")
        {
            ErrorsChanged?.Invoke(this, new DataErrorsChangedEventArgs(propertyName));
        }

        #endregion

        #region IWordPrintable

        public Document ToWordDoc()
        {
            Document doc = new Document();

            DocumentBuilder docBuilder = new DocumentBuilder(doc);

            docBuilder.Font.Size = 14;
            docBuilder.Font.Bold = true;
            docBuilder.Font.AllCaps = true;

            docBuilder.Writeln("Сведения о участнике");

            Aspose.Words.Paragraph para = doc.LastSection.Body.LastParagraph;

            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            return doc;
        }

        #endregion
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class PropertyNameAttribute : Attribute
    {
        public string PropertyName { get; private set; }
        public PropertyNameAttribute(string name) => PropertyName = name;
    }
}
