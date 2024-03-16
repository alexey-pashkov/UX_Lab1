using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using System.Windows.Input;
using Aspose.Words;
using wfp_lab1.Models;

namespace wfp_lab1.ViewModels
{
    internal class MainWindowViewModel : INotifyPropertyChanged
    {
        private Participant _particpant;
        public Participant Participant 
        {
            get { return this._particpant; }
            set
            {
                this._particpant = value;

                OnPropertyChanged();
            }
        }

        private XmlFileDialogService _dialogService;

        public CustomCommand Import
        {
            get
            {
                return new CustomCommand
                    ((param) =>
                        {
                            if (_dialogService.Show(XmlFileDialogMode.OpenFile))
                            {
                                this.Participant = Participant.LoadFromXml(this._dialogService.FilePath);
                            }
                        }
                    );
            }
        }

        public CustomCommand Export
        {
            get
            {
                return new CustomCommand
                    ((param) =>
                        {
                            if (_dialogService.Show(XmlFileDialogMode.SaveFile))
                            {
                                this._particpant.SaveToXml(this._dialogService.FilePath);
                            }
                        },
                        (param) =>
                        {
                            return Participant.HasErrors;
                        }
                    );
            }

        }

        public AsyncCommand Print
        {
            get
            {
                return new AsyncCommand
                    ((param) =>
                        {
                            Random rand = new Random();

                            string myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                            string filePath = myDocumentsPath + "\\temp" + rand.Next() + ".docx";

                            Aspose.Words.Document doc = Participant.ToWordDoc();

                            doc.Save(filePath);

                            ProcessStartInfo procInfo = new ProcessStartInfo();
                            procInfo.FileName = "winword.exe";
                            procInfo.Arguments = filePath;
                            Process.Start(procInfo);
                        }
                        //(param) =>
                        //{
                        //    return Participant.HasErrors;
                        //}
                    );
            }
        }

        public MainWindowViewModel()
        {
            this._particpant = new Participant();
            this._dialogService = new XmlFileDialogService();
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }

    public class CustomCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool>? _canExecute;

        public event EventHandler? CanExecuteChanged;

        public CustomCommand(Action<object> execute, Func<object, bool>? canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object? parameter)
        {
            return this._canExecute == null ? true : this._canExecute(parameter);
        }

        public void Execute(object? parameter)
        {
            this._execute(parameter);
        }
    }

    class AsyncCommand : ICommand
    {
        private readonly Action<object?> _execute;
        private readonly Func<bool>? _canExecute;

        public event EventHandler? CanExecuteChanged;

        public AsyncCommand(Action<object?> execute, Func<bool> canExecute = null)
        {
            this._execute = execute;
            if (canExecute != null)
            {
                this._canExecute = canExecute;
            }
        }

        public bool CanExecute(object? parameter)
        {
            return _canExecute == null ? true : _canExecute();
        }

        public async void Execute(object? parameter)
        {
            Task task = new Task(this._execute, parameter);

            task.Start();
        }
    }
}
