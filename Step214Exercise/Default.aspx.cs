using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Step214Exercise
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                PopulateData();
                lblMessage.Text = "Current Database Data:";
            }
        }

        private void PopulateData()
        {
            using (Step214dbEntities dc = new Step214dbEntities())
            {
                gvData.DataSource = dc.Players.ToList();
                gvData.DataBind();
            }
        }

        protected void btnImport_Click(object sender, EventArgs e)
        {
            if (FileUpload1.PostedFile.ContentType == "application/vnd.ms-excel" || 
                FileUpload1.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                try
                {
                    string fileName = Path.Combine(Server.MapPath("~/ImportDocument"), Guid.NewGuid().ToString() + Path.GetExtension(FileUpload1.PostedFile.FileName));
                    FileUpload1.PostedFile.SaveAs(fileName);

                    string conString = "";
                    string ext = Path.GetExtension(FileUpload1.PostedFile.FileName);
                    if (ext.ToLower() == ".xls")
                    {
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\""; ;
                    }
                    else if (ext.ToLower() == ".xlsx")
                    {
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\""; 
                    }

                    string query = "SELECT [Player Id], [First Name], [Last Name], [Team Name], [Jersey Number], [Points] FROM [Sheet1$]";
                    OleDbConnection con = new OleDbConnection(conString);
                    if (con.State == System.Data.ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    OleDbCommand cmd = new OleDbCommand(query, con);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    da.Dispose();
                    con.Close();
                    con.Dispose();

                    //Import to Database
                    using (Step214dbEntities dc = new Step214dbEntities())
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            int playID = Convert.ToInt32(dr["Player ID"]);
                            var v = dc.Players.Where(a => a.PlayerID.Equals(playID)).FirstOrDefault();
                            if (v != null)
                            {
                                // Update here
                                v.FirstName = dr["First Name"].ToString();
                                v.LastName = dr["Last Name"].ToString();
                                v.TeamName = dr["Team Name"].ToString();
                                v.JerseyNumber = Convert.ToInt32(dr["Jersey Number"]);
                                v.Points = Convert.ToInt32(dr["Points"]);
                            }
                            else
                            {
                                // Insert
                                dc.Players.Add(new Player
                                {
                                    PlayerID = Convert.ToInt32(dr["Player ID"]),
                                    FirstName = dr["First Name"].ToString(),
                                    LastName = dr["Last Name"].ToString(),
                                    TeamName = dr["Team Name"].ToString(),
                                    JerseyNumber = Convert.ToInt32(dr["Jersey Number"]),
                                    Points = Convert.ToInt32(dr["Points"])

                                });
                            }
                        }

                        dc.SaveChanges();
                    }
                    PopulateData();
                    lblMessage.Text = "Data successfully imported.";
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
    }
}