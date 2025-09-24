using System;
using System.Data.OleDb;
using System.IO;

class ExtractOLE
{
  static void Main()
  {
    string connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\db\mydb.accdb;";
    using (OleDbConnection conn = new OleDbConnection(connStr))
    {
      conn.Open();
      OleDbCommand cmd = new OleDbCommand("SELECT ID, OLEField FROM MyTable WHERE OLEField IS NOT NULL", conn);
      OleDbDataReader reader = cmd.ExecuteReader();

      while (reader.Read())
      {
        byte[] data = (byte[])reader["OLEField"];
        string filePath = $@"C:\ExportedFiles\Doc_{reader["ID"]}.bin";
        File.WriteAllBytes(filePath, data);
        Console.WriteLine("Saved: " + filePath);
      }
    }
  }
} 
