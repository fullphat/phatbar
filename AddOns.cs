using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Drawing;

namespace PhatBar
{
    public class CalcAddOn : AddOn
    {

        private Image _icon;

        public CalcAddOn()
        {
            _icon = Image.FromFile(Path.Combine(Environment.CurrentDirectory, "calc.png"));
        }

        public override string Name()
        {
            return "Calc";
        }

        public override List<Entry> QueryChanged(string text)
        {
            List<Entry> results = new List<Entry>();
            if (text.StartsWith("="))
            {
                Entry result = new Entry(this);
                result.Icon = this._icon;
                string calc = _calc(text);
                if (calc != "")
                {
                    result.Text = calc;
                    results.Add(result);
                }
            }
            return results;
        }

        public override string Handle(Entry entry, string text)
        {
            return "=" + _calc(text);
        }

        private string _calc(string text)
        {
            if (text.StartsWith("="))
            {
                DataTable dt = new DataTable();
                try
                {
                    var v = dt.Compute(text.TrimStart('='), "");
                    return v.ToString();
                }
                //catch (SyntaxErrorException e)
                //{
                //    result.Text = "{syntax error}";
                //}
                //catch (EvaluateException e)
                //{
                //    result.Text = "{logic error}";
                //}
                catch (Exception e)
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }


    }

    public class GoogleSearchAddOn : AddOn
    {
        public override string Name()
        {
            return "Google";
        }

        public override List<Entry> QueryChanged(string text)
        {
            List<Entry> results = new List<Entry>();

            Entry result = new Entry(this);
            result.Text = "Search Google for '" + text + "'";
            results.Add(result);

            return results;
        }
    }

    public class FilesystemSearchAddOn : AddOn
    {
        public override string Name()
        {
            return "Filesys";
        }

        public override List<Entry> QueryChanged(string text)
        {
            List<Entry> results = new List<Entry>();

            Entry result = new Entry(this);
            result.Text = "Search computer for '" + text + "'";
            results.Add(result);

            return results;
        }

    }

    public class RunnerAddOn : AddOn
    {
        private Image _folderIcon = null;
        private Image _fileIcon = null;

        public RunnerAddOn()
        {
            _folderIcon = Image.FromFile(Path.Combine(Environment.CurrentDirectory, "folder.png"));
            _fileIcon = Image.FromFile(Path.Combine(Environment.CurrentDirectory, "file.png"));
        }

        public override string Name()
        {
            return "Runner";
        }

        public override List<Entry> QueryChanged(string text)
        {
            List<Entry> results = new List<Entry>();
            string path = "";

            try
            {
                path = Path.GetFullPath(text);
            }
            catch (Exception e)
            {
                Debug.WriteLine("'" + text + "' is not a vaild path");
            }

            Entry result = new Entry(this);

            if (path != "")
            {
                if (Directory.Exists(path))
                {
                    // add ".." entry
                    result.Command = "go_up";
                    result.Data = path;
                    result.Text = "..";
                    result.Icon = _folderIcon;
                    results.Add(result);
                    
                    string[] contents = Directory.GetFileSystemEntries(path);
                    foreach (string item in contents)
                    {
                        result = new Entry(this);
                        result.Command = "exec";
                        result.Data = item;
                        result.Text = Path.GetFileName(item);
                        result.Icon = _folderIcon;
                        results.Add(result);
                    }

                }
                else if (File.Exists(path))
                {
                    result.Command = "exec";
                    result.Text = "Open '" + text + "'...";
                    result.Icon = _fileIcon;
                    results.Add(result);
                    //c:\program files\winuae\winuae.exe
                }
            }

            return results;
        }

        public override string Handle(Entry entry, string text)
        {
            if (entry.Command == "go_up")
            {
                return Path.GetDirectoryName(entry.Data);
            }

            if (Directory.Exists(entry.Data))
            {
                return entry.Data;
            }
            else if (File.Exists(entry.Data))
            {
                System.Diagnostics.Process.Start(entry.Data);
            }
            else
            {
                Debug.WriteLine("runner>" + entry.Data);
            }
            return "";
        }


    }

    public class QuitAddOn : AddOn
    {
        public override string Name()
        {
            return "Quitter";
        }

        public override List<Entry> QueryChanged(string text)
        {
            List<Entry> results = new List<Entry>();

            if (text == "!!")
            {
                Application.Exit();
            }

            return results;
        }

    }

}
