using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ProductDataBaseMVVM.MainWindow
{
    class MainWindowViewModel : INotifyPropertyChanged {

        private MainWindowModel m_Model;

        public MainWindowViewModel() {
            m_Model = new MainWindowModel();
        }

        /// <summary>
        /// 通知イベント
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// プロパティの変更通知を起動する
        /// </summary>
        /// <param name="propertyName">プロパティ名</param>
        protected void NotifyPropertyChanged([CallerMemberName] string propertyName = "") {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }






    }
}
