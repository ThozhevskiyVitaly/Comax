using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading;
using System.IO;
namespace Comax
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public Microsoft.Office.Interop.Excel.Application XL;
        private void Form1_Load(object sender, EventArgs e)
        {
        }
        public void PrintTreeView()
        {
            TreeNode MainArticle = new TreeNode("Articles");
            foreach(NewArticle art in (from a in newArticles.Values  orderby a.wireElSize() select a).ToArray<NewArticle>())
            {
                TreeNode artic = new TreeNode("[NewArticle] "+art.ArticleKey);
                List<TreeNode> propertyarticle = (from l in art.PrintArticle() select new TreeNode(l)).ToList();
                artic.Nodes.AddRange(propertyarticle.ToArray());
                List<TreeNode> newleadsets = new List<TreeNode>();
                int i = 1;
                foreach (NewLeadSet leadset in art.Items())
                {
                    List<TreeNode> propertyleadset = (from l in leadset.PrintNewLeadSet() select new TreeNode(l)).ToList();
                    TreeNode lead = new TreeNode("[NewLeadSet "+i+"]");
                    lead.Nodes.AddRange(propertyleadset.ToArray());
                    newleadsets.Add(lead);
                    i++;
                }
                artic.Nodes.AddRange(newleadsets.ToArray());
                MainArticle.Nodes.Add(artic);
            }
            treeView1.Nodes.Add(MainArticle);

        }CancellationTokenSource cancel = new CancellationTokenSource();
        bool token = true;
        private async void button1_Click(object sender, EventArgs e)
        { 
            OpenFileDialog xlsFile = new OpenFileDialog();
            IProgress<int> onChangeProgress = new Progress<int>((i) => 
            {
                if (token)
                {
                    token = false;
                    progressBar1.Maximum = i;
                }
                progressBar1.Value = i;
                textBox1.Text = i.ToString();
            });
            if (xlsFile.ShowDialog() == DialogResult.OK)
            {
                XL = new Microsoft.Office.Interop.Excel.Application();
                XL.Workbooks.Open(xlsFile.FileName);
                XL.Visible = true;
                Workbook wb = XL.ActiveWorkbook;
                button2.Click += delegate { cancel.Cancel(); };
                await Process(wb,onChangeProgress,cancel.Token);
               // textBox1.Text = t.ToString();
            }
            else return;
            PrintTreeView();
        }
       public static Dictionary<string, NewArticle> newArticles = new Dictionary<string, NewArticle>();
       enum ColumnNumbers {NumberByGraf1=4, NumberByGraf2 = 5,WireS=6,WireColor=7,WireLenght=8,StrippingLenght1=16, StrippingLenght2 = 29,PullOfLength1=17, PullOfLength2 = 30,WireMark=36,Pieces=40};
        public  Task<int> Process(Workbook wb,IProgress<int> progress,CancellationToken cancel)
        { return Task.Run(delegate
         {
             Worksheet t1 = wb.Sheets["А407-088Т1"];
             int NumberOfFirstRow = 8, CurrentRow = NumberOfFirstRow;
             int AmountRows = 0;
             while(t1.Cells[CurrentRow+AmountRows,ColumnNumbers.StrippingLenght1].Value!=null)
             {
                 AmountRows++;
             }
             progress.Report(AmountRows);
             while (t1.Cells[CurrentRow, ColumnNumbers.WireS].Value != null)
             {

                 if (cancel.IsCancellationRequested)
                 {
                     Wire wr = new Wire();
                     wr.ElectricalSize = t1.Cells[CurrentRow, ColumnNumbers.WireS].Value;
                     wr.Color = WireColor.GetColor(t1.Cells[CurrentRow, ColumnNumbers.WireColor].Value.ToString());
                     wr.WireKey = String.Format("{0}_{1}__PVA__{2} U_660", wr.ElectricalSize, wr.Color, t1.Cells[CurrentRow, ColumnNumbers.WireMark].Value.ToString());
                     string name = "\"" + t1.Cells[CurrentRow, ColumnNumbers.NumberByGraf1].Value + "   " + t1.Cells[CurrentRow, ColumnNumbers.NumberByGraf2].Value + "\"";
                     double wireLength = t1.Cells[CurrentRow, ColumnNumbers.WireLenght].Value;
                     double?[] strippinglength = { t1.Cells[CurrentRow, ColumnNumbers.StrippingLenght1].Value as double?, t1.Cells[CurrentRow, ColumnNumbers.StrippingLenght2].Value as double? };
                     double?[] pulloflength = { t1.Cells[CurrentRow, ColumnNumbers.PullOfLength1].Value as double?, t1.Cells[CurrentRow, ColumnNumbers.PullOfLength2].Value as double? };
                     int pieces = (int)t1.Cells[CurrentRow, ColumnNumbers.Pieces].Value;
                     NewLeadSet newleadset = new NewLeadSet(name, new string[] { wr.WireKey }, new double[] { wireLength }, strippinglength, pulloflength, pieces);
                     string ArticleKey = wr.ElectricalSize.ToString() + wr.Color;
                     if (newArticles.ContainsKey(ArticleKey)) newArticles[ArticleKey].AddItem = newleadset;
                     else
                     {
                         NewArticle newArticle = new NewArticle(wr, new List<NewLeadSet> { newleadset });
                         newArticle.ArticleKey = ArticleKey;
                         newArticles.Add(newArticle.ArticleKey, newArticle);
                     }
                     CurrentRow++;
                     progress.Report(CurrentRow - NumberOfFirstRow);
                 }
             }
             return CurrentRow - NumberOfFirstRow;
         },cancel);
        }

        private void button2_Click(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog savefile = new SaveFileDialog())
            {
                if(savefile.ShowDialog()==DialogResult.OK)
                File.WriteAllText(savefile.FileName, "");
            }
        }
    }
    public class NewLeadSet
    {
        string name;
        public string Name
        {
            get { return name; }
            set
            {
                if (value.Length > 50) name = value.Substring(0, 50);
                else name = value;
            }
        }
        public List<string> PrintNewLeadSet()
        {
            List<string> listnewleadset = new List<string>();
            Type type = this.GetType();
            foreach(PropertyInfo pi in type.GetProperties())
            {
                if(!(pi.GetValue(this) is Array))
                listnewleadset.Add(pi.Name+" = "+ pi.GetValue(this));
                else
                {
                    string data = "";
                    var array =( pi.GetValue(this) as Array);
                    foreach(var b in array)
                    {
                        double? fakeB = b as double?;
                        if(b!=null&fakeB!=0) data += b.ToString() + ", ";
                        if (data.Length>=2&&fakeB==0&&data.Substring(data.Length - 2, 2) == ", ") data = data.Remove(data.Length - 2);
                    }
                    if(data.Length >= 2 && data.Substring(data.Length - 2, 2) == ", ")data = data.Remove(data.Length - 2, 2);
                    listnewleadset.Add(pi.Name + " = " + data);
                }
            }
            return listnewleadset;
        }
        string[] wirekey=new string[2];
        public string[] WireKey
        {
            get { return wirekey; }
            set
            {
                for(int i=0;i<value.Length;i++)
                {
                    if (value[i].Length > 25) wirekey[i] = value[i].Substring(0, 25);
                    else wirekey[i] = value[i];
                }
            }
        }
        double [] wirelength = new double[2];
        public double[] WireLength
        {
            get { return wirelength; }
            set
            {
                for (int i = 0; i <value.Length; i++)
                {
                    wirelength[i] = value[i];
                }
            }
        }
        double?[] strippinglength = new double?[3];
        public double?[] StrippingLength
        {
            get { return strippinglength; }
            set
            {
                if (value != null)
                {
                    for (int i = 0; i < value.Length; i++)
                    {
                        strippinglength[i] = value[i];
                    }
                }
            
            }
        }
        double?[] pulloflength = new double?[3];
        public double?[] PullOffLength
        {
            get { return pulloflength; }
            set
            {
                if (value != null)
                {
                    for (int i = 0; i < value.Length; i++)
                    {
                        pulloflength[i] = value[i];
                    }
                }
            }
        }
        int pieces;
        public int Pieces
        {
            get { return pieces; }
            set { pieces = value; }
        }
        public NewLeadSet(string name,string[] wk,double[] wl)
        {
            Name = name;
            WireKey = wk;
            WireLength = wl;
        }
        public NewLeadSet(string name,string[] wk, double[] wl,double?[] sl):this(name,wk,wl)
        {
            StrippingLength = sl;
        }
        public NewLeadSet(string name, string[] wk, double[] wl, double?[] sl,double? [] pl) : this(name, wk, wl,sl)
        {
            PullOffLength = pl;
        }
        public NewLeadSet(string name, string[] wk, double[] wl, double?[] sl, double?[] pl,int p) : this(name, wk, wl, sl,pl)
        {
            Pieces = p;
        }
    }
    public class NewArticle
    {
        List<NewLeadSet> items = new List<NewLeadSet>();
        Wire wire = new Wire();
        public double wireElSize()
        {
             return wire.ElectricalSize;
        }
        public List<string> PrintArticle()
        {
            List<string> listnewarticle = new List<string>();
            Type type = this.GetType();
            
            foreach (PropertyInfo pi in type.GetProperties())
            {
                try
                {
                    listnewarticle.Add(pi.Name + " = " + pi.GetValue(this));
                }
                catch { }
            }
            return listnewarticle;

        }
        public List<NewLeadSet> Items()
        {
             return items;
        }
        public NewLeadSet AddItem
        {
            set
            {
                items.Add(value);
                items = (from it in items orderby it.WireLength[0] descending select it).ToList();
            }
        }
        string articlekey;
        public string ArticleKey
        {
            get { return articlekey; }
            set
            {
                if (value.Length >25) articlekey = value.Substring(0, 25);
                else articlekey = value;
            }
        }
        public int NumberOfLeadSets
        {
            get { return items.Count; }
        }
        public NewArticle(Wire wire)
        {
            this.wire.ElectricalSize = wire.ElectricalSize;
            ArticleKey = wire.ElectricalSize.ToString() + wire.Color;
        }
        public NewArticle(Wire wire,List<NewLeadSet> newleadset):this(wire)
        {
            items.AddRange(newleadset);
        }
    }
    public class Wire
    {
        string wirekey;
        WireColor.engColor color;
        public WireColor.engColor Color
        {
            get { return color; }
            set { color = value; }
        }
        public string WireKey
        {
            get { return wirekey; }
            set
            {
                if (value.Length > 25) wirekey = value.Substring(0, 25);
                else wirekey = value;
            }
        }
        double electricalsize;
        public double ElectricalSize
        {
            get { return electricalsize; }
            set { electricalsize = value; }
        }
    }
    public class WireColor
    {
        public enum ukrColor { Б, БГ,БЖ,БЗ,  БК,  БО,  БП,  БР,  БС , БФ , БЧ,  Г,   ГБ,  ГЖ , ГЗ , ГК , ГО , ГП , ГР,  ГС,  ГФ , ГЧ , Ж   ,ЖБ , ЖГ , ЖЗ,  ЖК , ЖО  ,ЖП , ЖР , ЖС , ЖФ , ЖЧ,  З  , ЗБ , ЗГ , ЗЖ  ,ЗК,  ЗО  ,ЗП,  ЗР , ЗС , ЗФ,  ЗЧ,  К,   КБ , КГ , КЖ , КЗ,  КО , КП , КР,  КС,  КФ,  КЧ , О ,  ОБ , ОГ,  ОЖ,  ОЗ,  ОК,  ОП , ОР , ОС , ОФ,  ОЧ,  П  , ПБ  ,ПГ , ПЖ , ПЗ,  ПК,  ПО,  ПР , ПС , ПФ , ПЧ,  Р ,  РБ,  РГ , РЖ,  РЗ , РК,  РО , РП,  РС,  РФ  ,РЧ,  С  , СБ , СГ,  СЖ , СЗ,  СК , СО , СП , СР , СФ,  СЧ , Ф,   ФБ , ФГ , ФЖ,  ФЗ , ФК , ФО , ФП , ФР , ФС , ФЧ , Ч  , ЧБ , ЧГ,  ЧЖ,  ЧЗ , ЧК , ЧО , ЧП,  ЧР,  ЧС,  ЧФ };
        public enum engColor {WT,    WT_BI,   WT_YW ,  WT_GR  , WT_BR ,  WT_OR ,  WT_RD ,  WT_PR  , WT_GY ,  WT_VT ,  WT_BK  , BI  ,BI_WT ,  BI_YW ,  BI_GR ,  BI_BR ,  BI_OR,   BI_RD,   BI_PR,   BI_GY,   BI_VT,   BI_BK  , YW , YW_WT,   YW_BI,   YW_GR,   YW_BR  , YW_OR ,  YW_RD ,  YW_PR ,  YW_GY,   YW_VT  , YW_BK  , GR , GR_WT ,  GR_BI ,  GR_YW ,  GR_BR ,  GR_OR ,  GR_RD,   GR_PR ,  GR_GY  , GR_VT  , GR_BK ,  BR , BR_WT,   BR_BI,   BR_YW ,  BR_GR ,  BR_OR ,  BR_RD,   BR_PR ,  BR_GY ,  BR_VT ,  BR_BK  , OR , OR_WT,   OR_BI ,  OR_YW ,  OR_GR ,  OR_BR,   OR_RD,   OR_PR ,  OR_GY,   OR_VT ,  OR_BK ,  RD , RD_WT ,  RD_BI,   RD_YW  , RD_GR ,  RD_BR ,  RD_OR ,  RD_PR  , RD_GY ,  RD_VT,   RD_BK ,  PR,  PR_WT ,  PR_BI ,  PR_YW ,  PR_GR ,  PR_BR ,  PR_OR,   PR_RD,   PR_GY  , PR_VT ,  PR_BK  , GY , GY_WT ,  GY_BI,   GY_YW ,  GY_GR ,  GY_BR ,  GY_OR,   GY_RD ,  GY_PR ,  GY_VT,   GY_BK,   VT , VT_WT,   VT_BI,   VT_YW ,  VT_GR,   VT_BR ,  VT_OR,   VT_RD,   VT_PR ,  VT_GY ,  VT_BK ,  BK , BK_WT ,  BK_BI,   BK_YW ,  BK_GR ,  BK_BR ,  BK_OR ,  BK_RD,   BK_PR ,  BK_GY ,  BK_VT};
        public static engColor GetColor(string ukrColor)
        {
            ukrColor tempUkrColor;
            if (Enum.TryParse<ukrColor>(ukrColor, out tempUkrColor)) return (engColor)tempUkrColor;
            else
            {
                Exception WireColorException = new Exception("Color {ukrColor} wasn`t recognize");
                throw WireColorException;
            }
        }
    }
}
