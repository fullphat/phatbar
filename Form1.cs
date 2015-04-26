using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing.Drawing2D;

namespace PhatBar
{
    public partial class Form1 : Form
    {

        List<AddOn> _AddOns = new List<AddOn>();
        List<Entry> _results = new List<Entry>();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _AddOns.Add(new CalcAddOn());
            _AddOns.Add(new RunnerAddOn());
            _AddOns.Add(new FilesystemSearchAddOn());
            _AddOns.Add(new GoogleSearchAddOn());
            _AddOns.Add(new QuitAddOn());
        }

        private void tbInput_TextChanged(object sender, EventArgs e)
        {
            lvResults.Items.Clear();
            ilDynamic.Images.Clear();
            _results.Clear();

            ilDynamic.Images.Add("na", imageList1.Images["na"]);

            if (tbInput.Text != "")
            {
                Debug.WriteLine(tbInput.Text + " / " + _AddOns.Count);

                foreach (AddOn addon in _AddOns)
                {
                    List<Entry> tmp = addon.QueryChanged(tbInput.Text);
                    Debug.WriteLine(tmp.Count);

                    foreach (Entry result in tmp)
                    {
                        Debug.WriteLine(">>" + result.Text);
                        _results.Add(result);

                        if (result.Icon != null)
                        {
                            ilDynamic.Images.Add(result.Key, result.Icon);
                            lvResults.Items.Add(result.Key, result.Text, result.Key);
                        }
                        else
                        {
                            lvResults.Items.Add(result.Key, result.Text, "na");
                        }

                    }
                }

                if (lvResults.Items.Count > 0)
                {
                    lvResults.Items[0].Selected = true;
                }
            }
        }

        private void tbInput_KeyDown(object sender, KeyEventArgs e)
        {
            int i = -99;
            switch (e.KeyCode)
            {
                case Keys.Tab:
                    Debug.WriteLine("TAB!");
                    tbInput.SelectAll();
                    e.SuppressKeyPress = true;
                    break;

                case Keys.Escape:
                    tbInput.Text = "";
                    e.SuppressKeyPress = true;
                    break;

                case Keys.Return:
                    string selected = lvResults.SelectedItems[0].Name;
                    Entry entry = _results.Find(k => k.Key == selected);
                    Debug.WriteLine("asking '" + entry.Owner.Name() + "' to perform '" + entry.Command + "' on '" + entry.Data + "'...");
                    string result = entry.Owner.Handle(entry, tbInput.Text);
                    if (!string.IsNullOrEmpty(result))
                    {
                        tbInput.Text = result;
                        tbInput.Select(tbInput.Text.Length, 0);
                    }
                    break;

                case Keys.Up:
                    i = lvResults.SelectedIndices[0];
                    if (i > 0)
                    {
                        lvResults.Items[i - 1].Selected = true;
                        lvResults.EnsureVisible(i - 1);
                    }
                    e.SuppressKeyPress = true;
                    break;

                case Keys.Down:
                    i = lvResults.SelectedIndices[0];
                    if (i < (lvResults.Items.Count - 1))
                    {
                        lvResults.Items[i + 1].Selected = true;
                        lvResults.EnsureVisible(i + 1);
                    }
                    e.SuppressKeyPress = true;
                    break;

                default:
                    Debug.WriteLine("key: " + e.KeyCode);
                    break;
                   
            }

        }

        private void lvResults_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            Debug.WriteLine(e.Item.Text + "> " + e.State.ToString());
            if (e.Item.Selected)
            {
                // Draw the background and focus rectangle for a selected item.
                Brush br = new SolidBrush(Color.FromArgb(51, 51, 51));
                e.Graphics.FillRectangle(br, e.Bounds);
                //e.DrawFocusRectangle();
            }
            else
            {
                Brush br = new SolidBrush(Color.FromArgb(228, 228, 228));
                e.Graphics.FillRectangle(br, e.Bounds);
            }

            Image img;

            if (e.Item.ImageKey != "")
            {
                img = e.Item.ImageList.Images[e.Item.ImageKey];
            }
            else
            {
                img = e.Item.ImageList.Images[e.Item.ImageIndex];
            }

            if (img != null)
            {
                Rectangle pr = e.Bounds;
                pr.Width = 32;
                pr.Height = 32;
                pr.Offset(0, (e.Bounds.Height - pr.Height) / 2);
                e.Graphics.DrawImage(img, pr);
            }

            Rectangle pr2 = e.Bounds;
            pr2.X += 32;

            Color fg;
            if (e.Item.Selected)
            {
                fg = Color.FromArgb(223, 223, 191);
            }
            else
            {
                fg = Color.FromArgb(0, 0, 0);
            }

            TextRenderer.DrawText(e.Graphics, e.Item.Text, lvResults.Font, pr2, fg,
                                    TextFormatFlags.Left | TextFormatFlags.SingleLine | TextFormatFlags.GlyphOverhangPadding | TextFormatFlags.VerticalCenter | TextFormatFlags.WordEllipsis);

            //e.DrawText();

            //// Draw the item text for views other than the Details view. 
            //if (listView1.View != View.Details)
            //{
            //}
        }

        private void lvResults_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tbInput_Leave(object sender, EventArgs e)
        {
        }
    }
}
