using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace TPDSSDataManager.Models
{
    public class TreeNode : INotifyPropertyChanged
    {
        public string Name { get; set; } = string.Empty;
        public ObservableCollection<TreeNode> Children { get; set; } = new ObservableCollection<TreeNode>();
        public TreeNode? Parent { get; set; }

        private bool? _isChecked = true;
        public bool? IsChecked
        {
            get => _isChecked;
            set { SetIsChecked(value, true, true); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public void SetIsChecked(bool? value, bool updateChildren, bool updateParent)
        {
            if (_isChecked == value) return;
            _isChecked = value;
            OnPropertyChanged(nameof(IsChecked));

            if (updateChildren && _isChecked.HasValue)
            {
                foreach (var child in Children)
                    child.SetIsChecked(_isChecked.Value, true, false);
            }

            if (updateParent && Parent != null)
                Parent.VerifyCheckState();
        }

        public void VerifyCheckState()
        {
            bool? state = null;
            if (Children.Count > 0)
            {
                bool allChecked = Children.All(c => c.IsChecked == true);
                bool allUnchecked = Children.All(c => c.IsChecked == false);

                if (allChecked) state = true;
                else if (allUnchecked) state = false;
                else state = null;
            }
            SetIsChecked(state, false, true);
        }
    }
}