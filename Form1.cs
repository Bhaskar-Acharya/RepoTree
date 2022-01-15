using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RepoTree
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
            InitializeTree();
        }

        public void InitializeTree()
        {
            ReadExcel();
            TreeNode root = new TreeNode("DigitalTech     ");
            root.NodeFont = new Font(treeView1.Font, FontStyle.Bold);
            treeView1.Nodes.Add(root);
            CreateTree("1", root);
            treeView1.ExpandAll();
        }


        private void CreateTree(string ParentId, TreeNode tn)
        {
            var result = from s in listRepoItem where s.ParentId == ParentId select s;
            foreach(RepoItem x in result)
            {
                TreeNode newTn = new TreeNode(x.DisplayText);
                if (x.RepoType == "Repo") { newTn.ForeColor = Color.Blue; }
                tn.Nodes.Add(newTn);
                CreateTree(x.Id, newTn);
            }
        }

        public class RepoItem
        {
            public string Id;
            public string DisplayText;
            public string RepoType;
            public string ParentId;
        }

        public Collection<RepoItem> listRepoItem;
        private void ReadExcel()
        {
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=.\RepoTree.xlsx;Extended Properties='Excel 12.0;HDR=YES;'");
            try
            {
                conn.Open();
                string sqlQuery = "SELECT * FROM [Repo$] WHERE Active=1";
                OleDbCommand cmdRepo = new OleDbCommand(sqlQuery, conn);
                OleDbDataReader drRepo = cmdRepo.ExecuteReader();
                listRepoItem = new Collection<RepoItem>();
                while (drRepo.Read())
                {
                    RepoItem repoItem = new RepoItem()
                    {
                        Id = drRepo["ID"].ToString(),
                        DisplayText = drRepo["Name"].ToString(),
                        RepoType = drRepo["Type"].ToString(),
                        ParentId = drRepo["ParentID"].ToString(),
                    };
                    listRepoItem.Add(repoItem);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();    
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            InitializeTree();
        }
    }
}
