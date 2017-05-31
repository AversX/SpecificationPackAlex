namespace SpecificationPackAlex
{
    using Microsoft.CSharp.RuntimeBinder;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Data.OleDb;
    using System.Drawing;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public class MainForm : Form
    {
        private System.Windows.Forms.Button addSpecBtn;
        private System.Windows.Forms.Button clearSpecBtn;
        private IContainer components = null;
        private System.Windows.Forms.Button deleteSpecBtn;
        private Microsoft.Office.Interop.Excel.Application excel;
        private System.Windows.Forms.Button formBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox specListBox;
        private List<Unit> Units;

        public MainForm()
        {
            this.InitializeComponent();
        }

        private void addSpecBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog {
                Multiselect = true,
                Filter = "(*.xlsx); (*.xls)|*.xlsx; *.xls"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string str in dialog.FileNames)
                {
                    this.specListBox.Items.Add(str);
                }
            }
        }

        private void clearSpecBtn_Click(object sender, EventArgs e)
        {
            this.specListBox.Items.Clear();
        }

        private List<Unit> Consolidate(List<Unit> units)
        {
            for (int i = 0; i < units.Count; i++)
            {
                for (int j = i + 1; j < units.Count; j++)
                {
                    if (units[i].Code == units[j].Code)
                    {
                        Unit unit = units[i];
                        for (int k = 0; k < unit.Num.Length; k++)
                        {
                            unit.Num[k] += units[j].Num[k];
                        }
                        if (units[j].Pos != "")
                        {
                            unit.PosNum = "";
                            unit.Pos = unit.Pos + ", " + units[j].Pos;
                        }
                        else
                        {
                            unit.PosNum = unit.PosNum + ", " + units[j].PosNum;
                        }
                        units.RemoveAt(j);
                        j--;
                        units[i] = unit;
                    }
                }
            }
            return units;
        }

        private void deleteSpecBtn_Click(object sender, EventArgs e)
        {
            if (this.specListBox.SelectedIndex >= 0)
            {
                this.specListBox.Items.RemoveAt(this.specListBox.SelectedIndex);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private List<Unit> findPos(List<Unit> units)
        {
            for (int i = 0; i < units.Count; i++)
            {
                if ((units[i].Pos != "") && !units[i].Pos.Contains("XT"))
                {
                    string str;
                    string str2;
                    if (units[i].Pos.Contains("."))
                    {
                        str2 = this.getPosWithDot(units[i].Pos, out str);
                    }
                    else
                    {
                        str2 = this.getPos(units[i].Pos, out str);
                    }
                    int num2 = int.Parse(str);
                    if ((num2 % 2) != 0)
                    {
                        num2++;
                        units.Insert(i + 1, new Unit(units[i].PosNum, str2 + num2, units[i].Code, units[i].Name, units[i].Manufacture, units[i].Num));
                        i++;
                    }
                }
            }
            return units;
        }

        private void formBtn_Click(object sender, EventArgs e)
        {
            this.Units = new List<Unit>();
            for (int i = 0; i < this.specListBox.Items.Count; i++)
            {
                this.Units.AddRange(this.loadDataSpec(this.specListBox.Items[i].ToString(), i));
            }
            this.Units = this.findPos(this.Units);
            this.Units = this.Consolidate(this.Units);
            this.posGroup();
            this.Units = this.loadDataPack(this.Units);
            this.Units = this.loadDataPackCoeffs(this.Units);
            this.uploadData();
        }

        private string getNum(string str)
        {
            string str2 = "0";
            for (int i = 1; i < str.Length; i++)
            {
                if (char.IsDigit(str[i]))
                {
                    str2 = str2 + str[i];
                }
                else if (char.IsPunctuation(str[i]))
                {
                    return str2;
                }
            }
            return str2;
        }

        private string getNumWithDot(string str, out string phase)
        {
            string str2 = "0";
            bool flag = false;
            int num = 1;
            for (int i = 1; i < str.Length; i++)
            {
                if (flag)
                {
                    if (char.IsDigit(str[i]))
                    {
                        str2 = str2 + str[i];
                        num = i;
                    }
                    else if (char.IsPunctuation(str[i]))
                    {
                        num = i;
                        break;
                    }
                }
                else if (char.IsPunctuation(str[i]))
                {
                    flag = true;
                }
            }
            if (num < (str.Length - 1))
            {
                phase = str.Substring(num + 1);
                return str2;
            }
            phase = "0";
            return str2;
        }

        private string getPos(string str, out string number)
        {
            string str2 = str[0].ToString();
            int num = 0;
            for (int i = 1; i < str.Length; i++)
            {
                if (char.IsLetterOrDigit(str[i]))
                {
                    str2 = str2 + str[i];
                    num = i;
                }
                else if (char.IsPunctuation(str[i]))
                {
                    str2 = str2 + str[i];
                    num = i;
                    break;
                }
            }
            if (num == (str.Length - 1))
            {
                number = "0";
                return str2;
            }
            number = str.Substring(num + 1);
            return str2;
        }

        private string getPosWithDot(string str, out string postDotNumber)
        {
            string str2 = str[0].ToString();
            int num = 0;
            for (int i = 1; i < str.Length; i++)
            {
                if (str[i] == '.')
                {
                    str2 = str2 + str[i];
                    num = i;
                    break;
                }
                else
                {
                    str2 = str2 + str[i];
                    num = i;
                }
            }
            if (num == (str.Length - 1))
            {
                postDotNumber = "0";
                return str2;
            }
            postDotNumber = str.Substring(num + 1);
            return str2;
        }

        private string getPosNum(string str, out string number)
        {
            string str2 = str[0].ToString();
            int num = 0;
            for (int i = 1; i < str.Length; i++)
            {
                if (char.IsLetter(str[i]))
                {
                    str2 = str2 + str[i];
                    num = i;
                }
            }
            if (num == (str.Length - 1))
            {
                number = "0";
                return str2;
            }
            number = str.Substring(num + 1);
            return str2;
        }

        private void InitializeComponent()
        {
            this.specListBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.addSpecBtn = new System.Windows.Forms.Button();
            this.deleteSpecBtn = new System.Windows.Forms.Button();
            this.formBtn = new System.Windows.Forms.Button();
            this.clearSpecBtn = new System.Windows.Forms.Button();
            base.SuspendLayout();
            this.specListBox.AllowDrop = true;
            this.specListBox.FormattingEnabled = true;
            this.specListBox.Location = new System.Drawing.Point(12, 0x1c);
            this.specListBox.Name = "specListBox";
            this.specListBox.Size = new Size(0x199, 0x93);
            this.specListBox.TabIndex = 0;
            this.specListBox.DragDrop += new DragEventHandler(this.specListBox_DragDrop);
            this.specListBox.DragEnter += new DragEventHandler(this.specListBox_DragEnter);
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 9);
            this.label1.Name = "label1";
            this.label1.Size = new Size(0x52, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Спецификации";
            this.addSpecBtn.Location = new System.Drawing.Point(0x142, 0xb1);
            this.addSpecBtn.Name = "addSpecBtn";
            this.addSpecBtn.Size = new Size(0x63, 0x24);
            this.addSpecBtn.TabIndex = 2;
            this.addSpecBtn.Text = "Добавить спецификацию";
            this.addSpecBtn.UseVisualStyleBackColor = true;
            this.addSpecBtn.Click += new EventHandler(this.addSpecBtn_Click);
            this.deleteSpecBtn.Location = new System.Drawing.Point(12, 0xb1);
            this.deleteSpecBtn.Name = "deleteSpecBtn";
            this.deleteSpecBtn.Size = new Size(0x63, 0x24);
            this.deleteSpecBtn.TabIndex = 2;
            this.deleteSpecBtn.Text = "Удалить";
            this.deleteSpecBtn.UseVisualStyleBackColor = true;
            this.deleteSpecBtn.Click += new EventHandler(this.deleteSpecBtn_Click);
            this.formBtn.Location = new System.Drawing.Point(0x12f, 0xe0);
            this.formBtn.Name = "formBtn";
            this.formBtn.Size = new Size(0x76, 0x22);
            this.formBtn.TabIndex = 3;
            this.formBtn.Text = "Сформировать";
            this.formBtn.UseVisualStyleBackColor = true;
            this.formBtn.Click += new EventHandler(this.formBtn_Click);
            this.clearSpecBtn.Location = new System.Drawing.Point(0x75, 0xb1);
            this.clearSpecBtn.Name = "clearSpecBtn";
            this.clearSpecBtn.Size = new Size(0x63, 0x24);
            this.clearSpecBtn.TabIndex = 2;
            this.clearSpecBtn.Text = "Очистить";
            this.clearSpecBtn.UseVisualStyleBackColor = true;
            this.clearSpecBtn.Click += new EventHandler(this.clearSpecBtn_Click);
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x1b1, 0x106);
            base.Controls.Add(this.formBtn);
            base.Controls.Add(this.clearSpecBtn);
            base.Controls.Add(this.deleteSpecBtn);
            base.Controls.Add(this.addSpecBtn);
            base.Controls.Add(this.label1);
            base.Controls.Add(this.specListBox);
            base.Name = "MainForm";
            this.Text = "Расчёт количества упаковок";
            base.ResumeLayout(false);
            base.PerformLayout();
        }

        private List<Unit> loadDataPack(List<Unit> units)
        {
            DataSet dataSet = new DataSet("EXCEL");
            OleDbConnection selectConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + System.Windows.Forms.Application.StartupPath + @"\Data\Pack.xlsx;Extended Properties='Excel 12.0;IMEX=0'");
            selectConnection.Open();
            object[] restrictions = new object[4];
            restrictions[3] = "TABLE";
            string str2 = (string) selectConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, restrictions).Rows[0].ItemArray[2];
            OleDbDataAdapter adapter = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", str2), selectConnection);
            adapter.FillSchema(dataSet, SchemaType.Source);
            adapter.Fill(dataSet);
            selectConnection.Close();
            List<UnitPack> list = new List<UnitPack>();
            for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
            {
                list.Add(new UnitPack(dataSet.Tables[0].Rows[i][0].ToString(), dataSet.Tables[0].Rows[i][1].ToString(), dataSet.Tables[0].Rows[i][2].ToString()));
            }
            for (int j = 0; j < units.Count; j++)
            {
                for (int k = 0; k < list.Count; k++)
                {
                    if (units[j].Code == list[k].Code)
                    {
                        Unit unit = units[j];
                        unit.PackNum = list[k].PackNum;
                        units[j] = unit;
                    }
                }
            }
            return units;
        }

        private List<Unit> loadDataPackCoeffs(List<Unit> units)
        {
            DataSet dataSet = new DataSet("EXCEL");
            OleDbConnection selectConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + System.Windows.Forms.Application.StartupPath + @"\Data\PackCoeffs.xlsx;Extended Properties='Excel 12.0;IMEX=0'");
            selectConnection.Open();
            object[] restrictions = new object[4];
            restrictions[3] = "TABLE";
            string str2 = (string) selectConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, restrictions).Rows[0].ItemArray[2];
            OleDbDataAdapter adapter = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", str2), selectConnection);
            adapter.FillSchema(dataSet, SchemaType.Source);
            adapter.Fill(dataSet);
            selectConnection.Close();
            List<UnitCoeffs> list = new List<UnitCoeffs>();
            for (int i = 0; i < dataSet.Tables[0].Rows.Count; i++)
            {
                list.Add(new UnitCoeffs(dataSet.Tables[0].Rows[i][0].ToString(), dataSet.Tables[0].Rows[i][1].ToString(), dataSet.Tables[0].Rows[i][2].ToString()));
            }
            for (int j = 0; j < units.Count; j++)
            {
                for (int k = 0; k < list.Count; k++)
                {
                    if (units[j].Code == list[k].Code)
                    {
                        Unit unit = units[j];
                        unit.Coeffs = list[k].Coeffs;
                        units[j] = unit;
                    }
                }
            }
            return units;
        }

        private List<Unit> loadDataSpec(string path, int index)
        {
            List<Unit> list = new List<Unit>();
            DataSet dataSet = new DataSet("EXCEL");
            OleDbConnection selectConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;IMEX=0'");
            selectConnection.Open();
            object[] restrictions = new object[4];
            restrictions[3] = "TABLE";
            string str2 = (string) selectConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, restrictions).Rows[0].ItemArray[2];
            new OleDbDataAdapter(string.Format("SELECT * FROM [{0}]", str2), selectConnection).Fill(dataSet);
            selectConnection.Close();
            for (int i = 1; i < dataSet.Tables[0].Rows.Count; i++)
            {
                if ((dataSet.Tables[0].Rows[i][3].ToString().Length > 0) && (dataSet.Tables[0].Rows[i][1].ToString().Length > 0))
                {
                    Unit unit = new Unit();
                    unit.PosNum = dataSet.Tables[0].Rows[i][0].ToString().Trim();
                    unit.PosNum = unit.PosNum.Replace("-", "");
                    unit.Pos = dataSet.Tables[0].Rows[i][1].ToString().Trim();
                    unit.Pos = unit.Pos.Replace("-", "");
                    unit.Code = dataSet.Tables[0].Rows[i][2].ToString().Trim();
                    unit.Name = dataSet.Tables[0].Rows[i][3].ToString().Trim();
                    unit.Manufacture = dataSet.Tables[0].Rows[i][4].ToString().Trim();
                    int[] numArray = new int[this.specListBox.Items.Count];
                    for (int j = 0; j < numArray.Length; j++)
                    {
                        if (j == index)
                        {
                            numArray[index] = int.Parse(dataSet.Tables[0].Rows[i][5].ToString().Trim());
                        }
                        else
                        {
                            numArray[j] = 0;
                        }
                    }
                    unit.Num = numArray;
                    list.Add(unit);
                }
            }
            return list;
        }

        private void posGroup()
        {
            for (int i = 0; i < this.Units.Count; i++)
            {
                List<string> list;
                int index;
                List<posUnit> list2;
                string str2;
                int num3;
                string str4;
                int num4;
                string str5;
                string str6;
                string str7;
                int num5;
                string str8;
                string str9;
                string str10;
                int num6;
                int num7;
                posUnit unit;
                Unit unit2;
                string posNum;
                if (this.Units[i].Pos == "")
                {
                    list = this.Units[i].PosNum.Split(new char[] { ',' }).ToList<string>();
                    index = 0;
                    while (index < list.Count)
                    {
                        list[index] = list[index].Trim();
                        index++;
                    }
                    if (list.Count > 0)
                    {
                        list.Sort(new NaturalStringComparer());
                        list2 = new List<posUnit>();
                        if (list[0] != "")
                        {
                            list[0] = list[0].Replace(" ", "");
                            if (list[0].Contains(".") || list[0].Contains(":"))
                            {
                                string str;
                                str2 = this.getPos(list[0], out str);
                                num3 = int.Parse(this.getNum(list[0]));
                                list2.Add(new posUnit(list[0], "", num3));
                            }
                            else
                            {
                                str2 = this.getPosNum(list[0], out str4);
                                num4 = int.Parse(str4);
                                list2.Add(new posUnit(list[0], "", num4));
                            }
                            index = 1;
                            while (index < list.Count)
                            {
                                if (list[index] != "")
                                {
                                    if (list2[list2.Count - 1].right != "")
                                    {
                                        if (list[index].Contains(".") || list[index].Contains(":"))
                                        {
                                            str6 = this.getPos(list[index], out str5);
                                            str7 = this.getNum(list[index]);
                                            num3 = int.Parse(str7);
                                            num5 = int.Parse(str5);
                                            if (list2[list2.Count - 1].right.Contains(".") || list2[list2.Count - 1].right.Contains(":"))
                                            {
                                                str9 = this.getPos(list2[list2.Count - 1].right, out str8);
                                                str10 = this.getNum(list2[list2.Count - 1].right);
                                                num6 = int.Parse(str10);
                                                num7 = int.Parse(str8);
                                                if ((num5 == (num7 + 1)) && (num3 == num6))
                                                {
                                                    unit = list2[list2.Count - 1];
                                                    unit.right = list[index];
                                                    unit.lastNum = num5;
                                                    list2[list2.Count - 1] = unit;
                                                }
                                                else
                                                {
                                                    list2.Add(new posUnit(list[index], "", num5));
                                                }
                                            }
                                            else
                                            {
                                                list2.Add(new posUnit(list[index], "", num5));
                                            }
                                        }
                                        else if (list2[list2.Count - 1].right.Contains(".") || list2[list2.Count - 1].right.Contains(":"))
                                        {
                                            str6 = this.getPosNum(list[index], out str7);
                                            num3 = int.Parse(str7);
                                            list2.Add(new posUnit(list[index], "", num3));
                                        }
                                        else if (list[index].Length > 2)
                                        {
                                            str6 = this.getPosNum(list[index], out str7);
                                            str9 = this.getPosNum(list2[list2.Count - 1].left, out str10);
                                            num3 = int.Parse(str7);
                                            num6 = int.Parse(str10);
                                            if ((str6 == str9) && ((num3 == (list2[list2.Count - 1].lastNum + 1)) || (num3 == list2[list2.Count - 1].lastNum)))
                                            {
                                                unit = list2[list2.Count - 1];
                                                unit.right = list[index];
                                                unit.lastNum = num3;
                                                list2[list2.Count - 1] = unit;
                                            }
                                            else
                                            {
                                                list2.Add(new posUnit(list[index], "", num3));
                                            }
                                        }
                                        else if (list[index] != list2[list2.Count - 1].left)
                                        {
                                            list2.Add(new posUnit(list[index], "", 0));
                                        }
                                    }
                                    else if (list[index].Contains(".") || list[index].Contains(":"))
                                    {
                                        str6 = this.getPos(list[index], out str5);
                                        str7 = this.getNum(list[index]);
                                        num3 = int.Parse(str7);
                                        num5 = int.Parse(str5);
                                        if (list2[list2.Count - 1].left.Contains(".") || list2[list2.Count - 1].left.Contains(":"))
                                        {
                                            str9 = this.getPos(list2[list2.Count - 1].left, out str8);
                                            str10 = this.getNum(list2[list2.Count - 1].left);
                                            num6 = int.Parse(str10);
                                            num7 = int.Parse(str8);
                                            if ((num5 == (num7 + 1)) && (num3 == num6))
                                            {
                                                unit = list2[list2.Count - 1];
                                                unit.right = list[index];
                                                unit.lastNum = num5;
                                                list2[list2.Count - 1] = unit;
                                            }
                                            else
                                            {
                                                list2.Add(new posUnit(list[index], "", num5));
                                            }
                                        }
                                        else
                                        {
                                            list2.Add(new posUnit(list[index], "", num5));
                                        }
                                    }
                                    else if (list2[list2.Count - 1].left.Contains(".") || list2[list2.Count - 1].left.Contains(":"))
                                    {
                                        str6 = this.getPosNum(list[index], out str7);
                                        num3 = int.Parse(str7);
                                        list2.Add(new posUnit(list[index], "", num3));
                                    }
                                    else if (list[index].Length > 2)
                                    {
                                        str6 = this.getPosNum(list[index], out str7);
                                        str9 = this.getPosNum(list2[list2.Count - 1].left, out str10);
                                        num3 = int.Parse(str7);
                                        num6 = int.Parse(str10);
                                        if ((str6 == str9) && ((num3 == (list2[list2.Count - 1].lastNum + 1)) || (num3 == list2[list2.Count - 1].lastNum)))
                                        {
                                            unit = list2[list2.Count - 1];
                                            unit.right = list[index];
                                            unit.lastNum = num3;
                                            list2[list2.Count - 1] = unit;
                                        }
                                        else
                                        {
                                            list2.Add(new posUnit(list[index], "", num3));
                                        }
                                    }
                                    else if (list[index] != list2[list2.Count - 1].left)
                                    {
                                        list2.Add(new posUnit(list[index], "", 0));
                                    }
                                }
                                index++;
                            }
                            unit2 = this.Units[i];
                            if (list2[0].right != string.Empty)
                            {
                                unit2.PosNum = list2[0].left + "-" + list2[0].right;
                            }
                            else
                            {
                                unit2.PosNum = list2[0].left;
                            }
                            index = 1;
                            while (index < list2.Count)
                            {
                                if (list2[index].right != string.Empty)
                                {
                                    posNum = unit2.PosNum;
                                    unit2.PosNum = posNum + ", " + list2[index].left + "-" + list2[index].right;
                                }
                                else
                                {
                                    unit2.PosNum = unit2.PosNum + ", " + list2[index].left;
                                }
                                index++;
                            }
                            this.Units[i] = unit2;
                        }
                    }
                }
                else
                {
                    list = this.Units[i].Pos.Split(new char[] { ',' }).ToList<string>();
                    index = 0;
                    while (index < list.Count)
                    {
                        list[index] = list[index].Trim();
                        index++;
                    }
                    if (list.Count > 0)
                    {
                        list2 = new List<posUnit>();
                        if (list[0] != "")
                        {
                            if (list[0].Contains("XT"))
                            {
                                list = this.strSort(list);
                                list[0] = list[0].Replace(" ", "");
                                if (list[0].Contains("."))
                                {
                                    string str11;
                                    num4 = int.Parse(this.getNumWithDot(list[0], out str11));
                                    list2.Add(new posUnit(list[0], "", num4));
                                }
                                else
                                {
                                    str2 = this.getPos(list[0], out str4);
                                    num4 = int.Parse(str4);
                                    list2.Add(new posUnit(list[0], "", num4));
                                }
                                index = 1;
                                while (index < list.Count)
                                {
                                    if (list[index] != "")
                                    {
                                        int num8;
                                        if (list[index].Contains("."))
                                        {
                                            int num9;
                                            if (list2[list2.Count - 1].right != "")
                                            {
                                                str6 = this.getPos(list[index], out str5);
                                                num8 = int.Parse(this.getNumWithDot(list[index], out str5));
                                                num5 = int.Parse(str5);
                                                if (list2[list2.Count - 1].right.Contains("."))
                                                {
                                                    str9 = this.getPos(list2[list2.Count - 1].right, out str8);
                                                    num9 = int.Parse(this.getNumWithDot(list2[list2.Count - 1].right, out str8));
                                                    num7 = int.Parse(str8);
                                                    if ((num5 == num7) && (num8 == (num9 + 1)))
                                                    {
                                                        unit = list2[list2.Count - 1];
                                                        unit.right = list[index];
                                                        unit.lastNum = num5;
                                                        list2[list2.Count - 1] = unit;
                                                    }
                                                    else
                                                    {
                                                        list2.Add(new posUnit(list[index], "", num5));
                                                    }
                                                }
                                                else
                                                {
                                                    list2.Add(new posUnit(list[index], "", num5));
                                                }
                                            }
                                            else
                                            {
                                                str6 = this.getPos(list[index], out str5);
                                                num8 = int.Parse(this.getNumWithDot(list[index], out str5));
                                                num5 = int.Parse(str5);
                                                if (list2[list2.Count - 1].left.Contains("."))
                                                {
                                                    str9 = this.getPos(list2[list2.Count - 1].left, out str8);
                                                    num9 = int.Parse(this.getNumWithDot(list2[list2.Count - 1].left, out str8));
                                                    num7 = int.Parse(str8);
                                                    if ((num5 == num7) && (num8 == (num9 + 1)))
                                                    {
                                                        unit = list2[list2.Count - 1];
                                                        unit.right = list[index];
                                                        unit.lastNum = num5;
                                                        list2[list2.Count - 1] = unit;
                                                    }
                                                    else
                                                    {
                                                        list2.Add(new posUnit(list[index], "", num5));
                                                    }
                                                }
                                                else
                                                {
                                                    list2.Add(new posUnit(list[index], "", num5));
                                                }
                                            }
                                        }
                                        else if (list2[list2.Count - 1].right != "")
                                        {
                                            if (list2[list2.Count - 1].right.Contains("."))
                                            {
                                                str6 = this.getPos(list[index], out str5);
                                                num8 = int.Parse(this.getNumWithDot(list[index], out str5));
                                                num5 = int.Parse(str5);
                                                list2.Add(new posUnit(list[index], "", num5));
                                            }
                                            else
                                            {
                                                list[index] = list[index].Replace(" ", "");
                                                str6 = this.getPos(list[index], out str5);
                                                str9 = this.getPos(list2[list2.Count - 1].right, out str8);
                                                str7 = this.getNum(list[index]);
                                                str10 = this.getNum(list2[list2.Count - 1].right);
                                                num3 = int.Parse(str7);
                                                num6 = int.Parse(str10);
                                                num5 = int.Parse(str5);
                                                num7 = int.Parse(str8);
                                                if ((num5 == num7) && (num3 == (num6 + 1)))
                                                {
                                                    unit = list2[list2.Count - 1];
                                                    unit.right = list[index];
                                                    unit.lastNum = num5;
                                                    list2[list2.Count - 1] = unit;
                                                }
                                                else
                                                {
                                                    list2.Add(new posUnit(list[index], "", num5));
                                                }
                                            }
                                        }
                                        else if (list2[list2.Count - 1].left.Contains("."))
                                        {
                                            str6 = this.getPos(list[index], out str5);
                                            num8 = int.Parse(this.getNumWithDot(list[index], out str5));
                                            num5 = int.Parse(str5);
                                            list2.Add(new posUnit(list[index], "", num5));
                                        }
                                        else
                                        {
                                            list[index] = list[index].Replace(" ", "");
                                            str6 = this.getPos(list[index], out str5);
                                            str9 = this.getPos(list2[list2.Count - 1].left, out str8);
                                            str7 = this.getNum(list[index]);
                                            str10 = this.getNum(list2[list2.Count - 1].left);
                                            num3 = int.Parse(str7);
                                            num6 = int.Parse(str10);
                                            num5 = int.Parse(str5);
                                            num7 = int.Parse(str8);
                                            if ((num5 == num7) && (num3 == (num6 + 1)))
                                            {
                                                unit = list2[list2.Count - 1];
                                                unit.right = list[index];
                                                unit.lastNum = num5;
                                                list2[list2.Count - 1] = unit;
                                            }
                                            else
                                            {
                                                list2.Add(new posUnit(list[index], "", num5));
                                            }
                                        }
                                    }
                                    index++;
                                }
                                unit2 = this.Units[i];
                                if (list2[0].right != string.Empty)
                                {
                                    unit2.PosNum = list2[0].left + "-" + list2[0].right;
                                }
                                else
                                {
                                    unit2.PosNum = list2[0].left;
                                }
                                index = 1;
                                while (index < list2.Count)
                                {
                                    if (list2[index].right != string.Empty)
                                    {
                                        posNum = unit2.PosNum;
                                        unit2.PosNum = posNum + ", " + list2[index].left + "-" + list2[index].right;
                                    }
                                    else
                                    {
                                        unit2.PosNum = unit2.PosNum + ", " + list2[index].left;
                                    }
                                    index++;
                                }
                                this.Units[i] = unit2;
                            }
                            else
                            {
                                list.Sort(new NaturalStringComparer());
                                List<string> posWithDot = new List<string>();
                                for (int j=0; j<list.Count; j++)
                                {
                                    if (list[j].Contains("."))
                                    {
                                        posWithDot.Add(list[j]);
                                        list.RemoveAt(j);
                                        j--;
                                    }
                                }
                                list[0] = list[0].Replace(" ", "");
                                str2 = this.getPos(list[0], out str4);
                                num4 = int.Parse(str4);
                                list2.Add(new posUnit(list[0], "", num4));
                                index = 1;
                                while (index < list.Count)
                                {
                                    if (list[index] != "")
                                    {
                                        list[index] = list[index].Replace(" ", "");
                                        if (list[index].Length > 2)
                                        {
                                            str6 = this.getPos(list[index], out str7);
                                            str9 = this.getPos(list2[list2.Count - 1].left, out str10);
                                            num3 = int.Parse(str7);
                                            num6 = int.Parse(str10);
                                            if ((str6 == str9) && (num3 == (list2[list2.Count - 1].lastNum + 1)))
                                            {
                                                unit = list2[list2.Count - 1];
                                                unit.right = list[index];
                                                unit.lastNum = num3;
                                                list2[list2.Count - 1] = unit;
                                            }
                                            else
                                            {
                                                list2.Add(new posUnit(list[index], "", num3));
                                            }
                                        }
                                        else if (list[index] != list2[list2.Count - 1].left)
                                        {
                                            list2.Add(new posUnit(list[index], "", 0));
                                        }
                                    }
                                    index++;
                                }
                                unit2 = this.Units[i];
                                if (list2[0].right != string.Empty)
                                {
                                    unit2.PosNum = list2[0].left + "-" + list2[0].right;
                                }
                                else
                                {
                                    unit2.PosNum = list2[0].left;
                                }
                                for (index = 1; index < list2.Count; index++)
                                {
                                    if (list2[index].right != string.Empty)
                                    {
                                        posNum = unit2.PosNum;
                                        unit2.PosNum = posNum + ", " + list2[index].left + "-" + list2[index].right;
                                    }
                                    else
                                    {
                                        unit2.PosNum = unit2.PosNum + ", " + list2[index].left;
                                    }
                                }
                                posWithDot[0] = posWithDot[0].Replace(" ", "");
                                string postNumStr;
                                str2 = this.getPosWithDot(posWithDot[0], out postNumStr);
                                int postNum = int.Parse(postNumStr);
                                List<posUnit> pwdList = new List<posUnit>();
                                pwdList.Add(new posUnit(posWithDot[0], "", postNum));
                                index = 1;
                                while (index < posWithDot.Count)
                                {
                                    if (posWithDot[index] != "")
                                    {
                                        posWithDot[index] = posWithDot[index].Replace(" ", "");
                                        if (posWithDot[index].Length > 2)
                                        {
                                            string present = this.getPosWithDot(posWithDot[index], out postNumStr);
                                            postNum = int.Parse(postNumStr);
                                            string postNumStr1;
                                            string past = this.getPosWithDot(pwdList[pwdList.Count - 1].left, out postNumStr1);
                                            if ((present == past) && (postNum == (pwdList[pwdList.Count - 1].lastNum + 1)))
                                            {
                                                unit = pwdList[pwdList.Count - 1];
                                                unit.right = posWithDot[index];
                                                unit.lastNum = postNum;
                                                pwdList[pwdList.Count - 1] = unit;
                                            }
                                            else
                                            {
                                                pwdList.Add(new posUnit(posWithDot[index], "", postNum));
                                            }
                                        }
                                        else if (posWithDot[index] != pwdList[pwdList.Count - 1].left)
                                        {
                                            pwdList.Add(new posUnit(posWithDot[index], "", 0));
                                        }
                                    }
                                    index++;
                                }
                                unit2 = this.Units[i];
                                if (pwdList[0].right != string.Empty)
                                {
                                    unit2.PosNum = pwdList[0].left + "-" + pwdList[0].right;
                                }
                                else
                                {
                                    unit2.PosNum = pwdList[0].left;
                                }
                                for (index = 1; index < pwdList.Count; index++)
                                {
                                    if (pwdList[index].right != string.Empty)
                                    {
                                        posNum = unit2.PosNum;
                                        unit2.PosNum = posNum + ", " + pwdList[index].left + "-" + pwdList[index].right;
                                    }
                                    else
                                    {
                                        unit2.PosNum = unit2.PosNum + ", " + pwdList[index].left;
                                    }
                                }
                                this.Units[i] = unit2;
                            }
                        }
                    }
                }
            }
        }

        private void specListBox_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && (e.Effect == DragDropEffects.Move))
            {
                string[] data = (string[]) e.Data.GetData(DataFormats.FileDrop);
                for (int i = 0; i < data.Length; i++)
                {
                    this.specListBox.Items.Add(data[i]);
                }
            }
        }

        private void specListBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && ((e.AllowedEffect & DragDropEffects.Move) == DragDropEffects.Move))
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        private List<string> strSort(List<string> str)
        {
            List<string> list = new List<string>();
            for (int i = 0; i < str.Count; i++)
            {
                List<string> list2;
                string str2;
                string str3;
                int num2;
                string str4;
                string str5;
                if (!str[i].Contains("."))
                {
                    list2 = new List<string>();
                    str3 = this.getPos(str[i], out str2);
                    list2.Add(str[i]);
                    num2 = i + 1;
                    while (num2 < str.Count)
                    {
                        str5 = this.getPos(str[num2], out str4);
                        if (str4 == str2)
                        {
                            list2.Add(str[num2]);
                            str.RemoveAt(num2);
                            num2--;
                        }
                        num2++;
                    }
                    str.RemoveAt(i);
                    i--;
                    list2.Sort(new NaturalStringComparer());
                    list.AddRange(list2);
                }
                else
                {
                    list2 = new List<string>();
                    str3 = this.getPos(str[i], out str2);
                    int num3 = int.Parse(this.getNumWithDot(str[i], out str2));
                    list2.Add(str[i]);
                    for (num2 = i + 1; num2 < str.Count; num2++)
                    {
                        if (str[num2].Contains("."))
                        {
                            str5 = this.getPos(str[num2], out str4);
                            int num4 = int.Parse(this.getNumWithDot(str[num2], out str4));
                            if ((str4 == str2) && (str5 == str3))
                            {
                                list2.Add(str[num2]);
                                str.RemoveAt(num2);
                                num2--;
                            }
                        }
                    }
                    str.RemoveAt(i);
                    i--;
                    list2.Sort(new NaturalStringComparer());
                    list.AddRange(list2);
                }
            }
            return list;
        }

        private void uploadData()
        {
            int row;
            this.excel = (Microsoft.Office.Interop.Excel.Application) Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
            this.excel.SheetsInNewWorkbook = 1;
            this.excel.Workbooks.Add(System.Type.Missing);
            Worksheet worksheet = (Worksheet) this.excel.Sheets.get_Item(1);
            worksheet.Cells[1, 1] = "Поз. обозн";
            ((dynamic) worksheet.Columns[1, Missing.Value]).NumberFormat = "@";
            worksheet.Cells[1, 2] = "Код";
            ((dynamic) worksheet.Columns[2, Missing.Value]).NumberFormat = "@";
            worksheet.Cells[1, 3] = "Наименование";
            ((dynamic) worksheet.Columns[3, Missing.Value]).NumberFormat = "@";
            worksheet.Cells[1, 4] = "Завод изготовитель";
            ((dynamic) worksheet.Columns[4, Missing.Value]).NumberFormat = "@";
            int column = 4;
            for (row = 0; row < this.Units.Count; row++)
            {
                worksheet.Cells[row + 2, 1] = this.Units[row].PosNum;
                worksheet.Cells[row + 2, 2] = this.Units[row].Code;
                worksheet.Cells[row + 2, 3] = this.Units[row].Name;
                worksheet.Cells[row + 2, 4] = this.Units[row].Manufacture;
                for (int i = 0; i < this.Units[row].Num.Length; i++)
                {
                    worksheet.Cells[row + 2, 5 + i] = this.Units[row].Num[i];
                    if ((5 + i) > column)
                    {
                        column++;
                        ((dynamic) worksheet.Cells[1, column]).NumberFormat = "#";
                        ((dynamic) worksheet.Cells[1, column]).Value2 = i + 1;
                    }
                }
            }
            worksheet.Cells[1, column + 1] = "В упаковке, шт.";
            worksheet.Cells[1, column + 2] = "Сумма по шкафам";
            worksheet.Cells[1, column + 3] = "Коэффицент";
            worksheet.Cells[1, column + 4] = "Расчёт";
            worksheet.Cells[1, column + 5] = "Упаковок, шт.";
            for (row = 0; row < this.Units.Count; row++)
            {
                ((dynamic) worksheet.Cells[row + 2, column + 1]).NumberFormat = "#";
                ((dynamic) worksheet.Cells[row + 2, column + 1]).Value2 = this.Units[row].PackNum;

                Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 2];

                Microsoft.Office.Interop.Excel.Range range2 = (Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, 4];
                Microsoft.Office.Interop.Excel.Range range3 = (Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column];
                Microsoft.Office.Interop.Excel.Range range4 = worksheet.get_Range(range2, range3);
                ((dynamic) worksheet.Cells[row + 2, column + 2]).NumberFormat = "#";
                
                //if (<uploadData>o__SiteContainer0.<>p__Sitee.Target(<uploadData>o__SiteContainer0.<>p__Sitee, <uploadData>o__SiteContainer0.<>p__Sitef.Target(<uploadData>o__SiteContainer0.<>p__Sitef, ((dynamic) worksheet.Cells[row + 2, column + 1]).Value2, null)))
                {
                    ((Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 2]).FormulaLocal = "=СУММ(" + range4.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString() + ")";
                }
                //else
                //{
                //    ((Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 2]).FormulaLocal = "=СУММ(" + range4.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString() + ")";
                //    ((Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 1]).Interior.Color = ColorTranslator.ToOle(Color.Red);
                //    ((dynamic) worksheet.Cells[row + 2, column + 1]).Value2 = "шт.";
                //}
                ((dynamic) worksheet.Cells[row + 2, column + 3]).NumberFormat = "0,00";
                if (this.Units[row].Coeffs != null)
                {
                    ((dynamic) worksheet.Cells[row + 2, column + 3]).Value2 = double.Parse(this.Units[row].Coeffs);
                }
                ((dynamic) worksheet.Cells[row + 2, column + 4]).NumberFormat = "0,00";
                if (this.Units[row].Manufacture == "HELUKABEL")
                {
                    range = (Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 2];
                    if (this.Units[row].Coeffs != null)
                    {
                        ((Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 4]).Formula = "=" + this.Units[row].Coeffs.Replace(',', '.') + " * (10 * " + range.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString() + ")/" + range.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString();
                    }
                    else
                    {
                        ((Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 4]).Formula = "=(10 * " + range.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString() + ")/" + range.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString();
                    }
                }
                else
                {
                    range = (Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 2];
                    if (this.Units[row].Coeffs != null)
                    {
                        ((Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 4]).Formula = "=" + this.Units[row].Coeffs.Replace(',', '.') + " * (" + range.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString() + ")/" + range.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString();
                    }
                    else
                    {
                        ((Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 4]).Formula = "=(" + range.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString() + ")/" + range.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString();
                    }
                }
                Microsoft.Office.Interop.Excel.Range range6 = (Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 4];
                ((dynamic) worksheet.Cells[row + 2, column + 5]).NumberFormat = "#";
                ((Microsoft.Office.Interop.Excel.Range) worksheet.Cells[row + 2, column + 5]).FormulaLocal = "=ОКРУГЛВВЕРХ(" + range6.get_Address(Missing.Value, Missing.Value, XlReferenceStyle.xlA1, Missing.Value, Missing.Value).ToString() + ";0)";
            }
            this.excel.Visible = true;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct posUnit
        {
            public string left;
            public string right;
            public int lastNum;
            public posUnit(string l, string r, int num)
            {
                this.left = l;
                this.right = r;
                this.lastNum = num;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Unit
        {
            public string PosNum;
            public string Pos;
            public string Code;
            public string Name;
            public string Manufacture;
            public int[] Num;
            public string PackNum;
            public string Coeffs;
            public Unit(string posNum, string pos, string code, string name, string manufacture, int[] num)
            {
                this.PosNum = posNum;
                this.Pos = pos;
                this.Code = code;
                this.Name = name;
                this.Manufacture = manufacture;
                this.Num = num;
                this.PackNum = "";
                this.Coeffs = "";
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct UnitCoeffs
        {
            public string Code;
            public string Name;
            public string Coeffs;
            public UnitCoeffs(string code, string name, string coeffs)
            {
                this.Code = code;
                this.Name = name;
                this.Coeffs = coeffs;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct UnitPack
        {
            public string Code;
            public string Name;
            public string PackNum;
            public UnitPack(string code, string name, string packNum)
            {
                this.Code = code;
                this.Name = name;
                this.PackNum = packNum;
            }
        }
    }
}

