using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GenerateSchedule
{
    public static class DBTrans
    {
        public static SqlConnection conn = new SqlConnection("Data Source=EMAD-LAP;Initial Catalog=HR900;User ID=sa;Password=ccsystem*360;Connect Timeout=30000");
        public static void Exec(this string strCommand)
        {
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            var sqlTransaction = conn.BeginTransaction("Trans");
            try
            {
                using (SqlCommand cmd = new SqlCommand(strCommand, conn))
                {
                    cmd.Transaction = sqlTransaction;
                    cmd.ExecuteNonQuery();
                }
                sqlTransaction.Commit();

            }
            catch (Exception exception)
            {
                sqlTransaction.Rollback();
                MessageBox.Show("فشل الاتصال");
            }
            finally
            {
                sqlTransaction.Dispose();
                conn.Close();
            }
        }
        public static void ExecProc(this string Stored, params SqlParameter[] prms)
        {
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            var sqlTransaction = conn.BeginTransaction("Transp");
            try
            {
                using (SqlCommand cmd = new SqlCommand(Stored, conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddRange(prms);
                    cmd.Transaction = sqlTransaction;
                    cmd.ExecuteNonQuery();
                }
                sqlTransaction.Commit();

            }
            catch (Exception exception)
            {
                sqlTransaction.Rollback();
                MessageBox.Show("فشل الاتصال");
            }
            finally
            {
                sqlTransaction.Dispose();
                conn.Close();
            }
        }
        public static void ExecP(this string Stored, int m, string userid, int tag)
        {
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
            var sqlTransaction = conn.BeginTransaction("Transp");
            try
            {
                using (SqlCommand cmd = new SqlCommand(Stored, conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddWithValue("@QuizId", tag);
                    cmd.Parameters.AddWithValue("@userid", userid);
                    cmd.Parameters.AddWithValue("@Minutes", m);
                    cmd.Transaction = sqlTransaction;
                    cmd.ExecuteNonQuery();
                }
                sqlTransaction.Commit();

            }
            catch (Exception exception)
            {
                sqlTransaction.Rollback();
                MessageBox.Show("فشل الاتصال");
            }
            finally
            {
                sqlTransaction.Dispose();
                conn.Close();
            }
        }

        public static DataTable GetTable(this string sql)
        {
            var dt = new DataTable();
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, conn))
            {
                try
                {
                    adapter.Fill(dt);
                    return dt;
                }
                catch (Exception)
                {

                    return null;
                }

            }
        }

        public static DataRow GetRow(this string sql)
        {
            var dt = new DataTable();
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, conn))
            {
                adapter.Fill(dt);
                return dt.Rows[0];
            }
        }
    }
}
