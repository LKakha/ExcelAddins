using ExcelDna.Integration;
using System;
using System.Data;
using System.Data.SqlClient;

public static class MyFunctions
{
	private const string ConnString = "Data Source=SQL-SERVER; Initial Catalog=TaxInfo; Persist Security Info=True; User ID=TIDB_User; Password=X7Vo5t8K63hXfM2Z; Network=DBMSSOCN";
	public static SqlConnection Conn;
	private static SqlCommand Comm;
	private static bool ConnectionIsValid = false;

	public static void SetConnection()
	{
		ConnectionIsValid = ServerIsPingable();
		if (ConnectionIsValid)
		{
			try
			{
				Conn = new SqlConnection(ConnString);
				Conn.Open();
				Comm = new SqlCommand("usp_TIDB", Conn) { CommandType = CommandType.StoredProcedure };
				Comm.Prepare();
				SqlParameter P;
				P = new SqlParameter("@C", SqlDbType.Int); P.IsNullable = true; Comm.Parameters.Add(P);
				P = new SqlParameter("@Par1", SqlDbType.Variant); P.IsNullable = true; Comm.Parameters.Add(P);
				P = new SqlParameter("@Par2", SqlDbType.Variant); P.IsNullable = true; Comm.Parameters.Add(P);
				P = new SqlParameter("@Par3", SqlDbType.Variant); P.IsNullable = true; Comm.Parameters.Add(P);
				P = new SqlParameter("@Par4", SqlDbType.Variant); P.IsNullable = true; Comm.Parameters.Add(P);
				P = new SqlParameter("@Par5", SqlDbType.Variant); P.IsNullable = true; Comm.Parameters.Add(P);
			}
			catch (Exception ex)
			{
				ConnectionIsValid = false;
				System.Windows.Forms.MessageBox.Show(ex.Message);
			}
		}
	}

	[ExcelFunction(Description = "ინფორმაცია TaxInfo-დან")]
	public static object TIDB(int C, string Par1 = null, string Par2 = null, string Par3 = null, string Par4 = null, string Par5 = null)
	{
		if (ConnectionIsValid)
		{
			if (string.IsNullOrEmpty(Par1)) Par1 = null;
			if (string.IsNullOrEmpty(Par2)) Par2 = null;
			if (string.IsNullOrEmpty(Par3)) Par3 = null;
			if (string.IsNullOrEmpty(Par4)) Par4 = null;
			if (string.IsNullOrEmpty(Par5)) Par5 = null;

			try
			{
				Comm.Parameters[0].SqlValue = C;
				Comm.Parameters[1].SqlValue = Par1;
				Comm.Parameters[2].SqlValue = Par2;
				Comm.Parameters[3].SqlValue = Par3;
				Comm.Parameters[4].SqlValue = Par4;
				Comm.Parameters[5].SqlValue = Par5;
				var R = Comm.ExecuteScalar();
				if (R == null || R == DBNull.Value) return ExcelError.ExcelErrorNA;
				return R;
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}
		else
			return "სერვერთან კავშირი არ არის";
	}

	private static bool ServerIsPingable()
	{
		var ConBld = new SqlConnectionStringBuilder(ConnString);
		var Server = ConBld.DataSource;
		var P = new System.Net.NetworkInformation.Ping();
		try
		{
			var R = P.Send(Server, 1000).Status;
			return R == System.Net.NetworkInformation.IPStatus.Success;
		}
		catch (Exception ex)
		{
			System.Windows.Forms.MessageBox.Show(ex.Message);
			return false;
		}
	}
}
