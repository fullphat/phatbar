using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace PhatBar
{
    public class Entry
    {
        private string _guid = "";
        private AddOn _owner = null;
        public string Text { get; set; }
        public Image Icon { get; set; }
        public string Command { get; set; }
        public string Data { get; set; }

        public Entry(AddOn Owner)
        {
            _owner = Owner;
            _guid = Guid.NewGuid().ToString();
        }

        public string Key
        {
            get
            {
                return _guid;
            }
        }

        public AddOn Owner
        {
            get
            {
                return _owner;
            }
        }


    }
    
    
    public class AddOn
    {

        public virtual string Name()
        {
            return "";
        }

        public virtual string Icon()
        {
            return "";
        }

        public virtual List<Entry> QueryChanged(string text)
        {
            return new List<Entry>();
        }

        public virtual string Handle(Entry entry, string text)
        {
            return "";
        }

    }
}
