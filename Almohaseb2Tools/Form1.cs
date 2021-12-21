using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace Almohaseb2Tools
{
    public partial class Form1 : Form
    {
        string ConnectionString = @"Server=.\SQLEXPRESS; Database=AlmohasebSQL; Integrated Security=True";
        string BossDbConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=D:\WORK\almohaseb2\boss.mdb; Jet OLEDB:Database Password=636305";
        DataTable BossDataTable;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, @EventArgs e)
        {

        }


        private async void button1_Click(object sender, @EventArgs e)
        {
            //if (MessageBox.Show())
            //{

            //}
            
            using (OleDbConnection acc = new OleDbConnection(BossDbConnectionString))
            {
                OleDbDataAdapter da = new OleDbDataAdapter("Select * from ggg1", acc);
                DataSet ds = new DataSet();
                da.Fill(ds);
                BossDataTable = ds.Tables[0];
            }

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    progressBar1.Maximum = BossDataTable.Rows.Count;
                    for (int i = 0; i < BossDataTable.Rows.Count; i++)
                    {
                        //Insert Into The_Items
                        SqlCommand cmd = new SqlCommand(@"insert into the_items(scientific_name, out_quantitative, type_validity, redial, item_status, item_kind, formpackaging_no, group_no, last_movement, item_add, localhostname, item_sn, bar_print, person_no, newitem_add, show_menu, quntity_rounded, old_unit) 
                        Values(@scientific_name, @out_quantitative, @type_validity, @redial, @item_status, @item_kind, @formpackaging_no, @group_no, @last_movement, @item_add, @localhostname, @item_sn, @bar_print, @person_no, @newitem_add, @show_menu, @quntity_rounded, @old_unit)", conn);
                        cmd.Parameters.Add("@scientific_name", @SqlDbType.NVarChar).Value = BossDataTable.Rows[i][2].ToString();
                        cmd.Parameters.Add("@out_quantitative", @SqlDbType.Real).Value = 0;
                        cmd.Parameters.Add("@type_validity", @SqlDbType.Bit).Value = 0;
                        cmd.Parameters.Add("@redial", @SqlDbType.Real).Value = 0;
                        cmd.Parameters.Add("@item_status", @SqlDbType.Bit).Value = 0;
                        cmd.Parameters.Add("@item_kind", @SqlDbType.SmallInt).Value = 0;
                        cmd.Parameters.Add("@formpackaging_no", @SqlDbType.Int).Value = 2;
                        cmd.Parameters.Add("@group_no", @SqlDbType.Int).Value = 1;
                        cmd.Parameters.Add("@last_movement", @SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.Parameters.Add("@item_add", @SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.Parameters.Add("@localhostname", @SqlDbType.NVarChar).Value = Environment.MachineName;
                        cmd.Parameters.Add("@item_sn", @SqlDbType.Bit).Value = 0;
                        cmd.Parameters.Add("@bar_print", @SqlDbType.Bit).Value = 0;
                        cmd.Parameters.Add("@person_no", @SqlDbType.Int).Value = 1;
                        cmd.Parameters.Add("@newitem_add", @SqlDbType.NVarChar).Value = DateTime.Now;
                        cmd.Parameters.Add("@show_menu", @SqlDbType.Bit).Value = 0;
                        cmd.Parameters.Add("@quntity_rounded", @SqlDbType.Bit).Value = 0;
                        cmd.Parameters.Add("@old_unit", @SqlDbType.Real).Value = 1;

                        await conn.OpenAsync();
                        await cmd.ExecuteNonQueryAsync();
                        conn.Close();

                        int Bond = GetLastItemNumber();

                        //Insert into TheItem_Details
                        SqlCommand cmd2 = new SqlCommand(@"insert into the_itemdetails(item_no, item_quantity, item_reserved, item_cost, placeexchange_no, placeoffer_no, last_movement, exp_date, item_add, localhostname, newitem_add)" +
                        " Values(@item_no, @item_quantity, @item_reserved, @item_cost, @placeexchange_no, @placeoffer_no, @last_movement, @exp_date, @item_add, @localhostname, @newitem_add)", conn);
                        cmd2.Parameters.Add("@item_no", SqlDbType.Int).Value = Bond;
                        cmd2.Parameters.Add("@item_quantity", SqlDbType.Int).Value = 0;
                        cmd2.Parameters.Add("@item_reserved", SqlDbType.Int).Value = 0;
                        cmd2.Parameters.Add("@item_cost", SqlDbType.Real).Value = 0;
                        cmd2.Parameters.Add("@placeexchange_no", SqlDbType.Int).Value = 1;
                        cmd2.Parameters.Add("@placeoffer_no", SqlDbType.Int).Value = 1;
                        cmd2.Parameters.Add("@last_movement", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd2.Parameters.Add("@exp_date", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd2.Parameters.Add("@item_add", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd2.Parameters.Add("@localhostname", SqlDbType.NVarChar).Value = Environment.MachineName;
                        cmd2.Parameters.Add("@newitem_add", SqlDbType.NVarChar).Value = DateTime.Now;

                        conn.Open();
                        cmd2.ExecuteNonQuery();
                        conn.Close();

                        int LastItemDetials = GetLastItemDetailNumber();


                        //Insert into The_Charge
                        SqlCommand cmd3 = new SqlCommand("insert into the_charge(itemdetails_no, charge_value, charge_value2, charge_rate, default_charge) Values(@itemdetails_no, @charge_value, @charge_value2, @charge_rate, @default_charge)", conn);
                        cmd3.Parameters.Add("@itemdetails_no", SqlDbType.Int).Value = LastItemDetials;
                        cmd3.Parameters.Add("@charge_value", SqlDbType.Real).Value = 0;
                        cmd3.Parameters.Add("@charge_value2", SqlDbType.Real).Value = 0;
                        cmd3.Parameters.Add("@charge_rate", SqlDbType.Real).Value = 0;
                        cmd3.Parameters.Add("@default_charge", SqlDbType.Int).Value = 1;


                        //Insert into the_barcode
                        SqlCommand cmd4 = new SqlCommand("insert into the_barcode(item_no, barcode, unit_no) Values(@item_no, @barcode, @unit_no)", conn);
                        cmd4.Parameters.Add("@item_no", SqlDbType.Int).Value = Bond;
                        cmd4.Parameters.Add("@barcode", SqlDbType.NVarChar).Value = BossDataTable.Rows[i][0].ToString();
                        cmd4.Parameters.Add("@unit_no", SqlDbType.Int).Value = 0;

                        //insert into The_Units
                        if (Convert.ToInt32(BossDataTable.Rows[i][3]) > 1)
                        {
                            AddUnits(Bond, "ÚáÈÉ", Convert.ToInt32(BossDataTable.Rows[i][3]), 1, 1);
                            AddUnits(Bond, "ÔÑíØ", 1, Convert.ToInt32(BossDataTable.Rows[i][3]), 0);
                        } else
                        {
                            AddUnits(Bond, "ÚáÈÉ", 3, 1, 1);
                        }



                        //Insert Into The_Trade
                        SqlCommand cmd6 = new SqlCommand("insert into the_trade(item_no, trade_name) Values(@item_no, @trade_name)", conn);
                        cmd6.Parameters.Add("@item_no", SqlDbType.Int).Value = Bond;
                        cmd6.Parameters.Add("@trade_name", SqlDbType.NVarChar).Value = BossDataTable.Rows[i][2].ToString();


                        conn.Open();
                        cmd3.ExecuteNonQuery();
                        cmd4.ExecuteNonQuery();
                        cmd6.ExecuteNonQuery();
                        label1.Text = "ÇáÕäÝ: " +  BossDataTable.Rows[i][2].ToString();
                        progressBar1.Value += 1;
                        conn.Close();
                    }
                }
                catch (Exception)
                {

                    throw;
                }
            }
        }

        private int GetLastItemNumber()
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                int res = 0;
                SqlDataAdapter da = new SqlDataAdapter("Select Max(Item_No) from The_Items", @conn);
                DataSet ds = new DataSet();
                da.Fill(ds);
                DataTable dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    res = Convert.ToInt32(dt.Rows[0][0]);
                }
                else
                {
                    res = 0;
                }
                return res;
            }
        }
        private int GetLastItemDetailNumber()
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                int res = 0;
                SqlDataAdapter da = new SqlDataAdapter("Select Max(ItemDetails_No) from The_ItemDetails", @conn);
                DataSet ds = new DataSet();
                da.Fill(ds);
                DataTable dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    res = Convert.ToInt32(dt.Rows[0][0]);
                }
                else
                {
                    res = 0;
                }
                return res;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(GetLastItemNumber().ToString());
            MessageBox.Show(GetLastItemDetailNumber().ToString());
        }

        private void AddUnits(int bond, string unitName, int unitQuantity, int unitInverted, int defUnit)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                try
                {
                    SqlCommand cmd5 = new SqlCommand("insert into the_units(item_no, unit_type, unit_quantity, unit_oldquantity, unit_inverted, default_unit) Values(@item_no, @unit_type, @unit_quantity, @unit_oldquantity, @unit_inverted, @default_unit)", conn);
                    cmd5.Parameters.Add("@item_no", SqlDbType.Int).Value = bond;
                    cmd5.Parameters.Add("@unit_type", SqlDbType.NVarChar).Value = unitName;
                    cmd5.Parameters.Add("@unit_quantity", SqlDbType.Real).Value = unitQuantity;
                    cmd5.Parameters.Add("@unit_oldquantity", SqlDbType.Real).Value = unitQuantity;
                    cmd5.Parameters.Add("@unit_inverted", SqlDbType.Int).Value = unitInverted;
                    cmd5.Parameters.Add("@default_unit", SqlDbType.Int).Value = defUnit;

                    conn.Open();
                    cmd5.ExecuteNonQuery();
                    conn.Close();
                }
                catch (Exception)
                {

                    throw;
                }
            }


        }
    }
}