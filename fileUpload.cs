 using (SqlConnection con = new SqlConnection(constr))
 {
	cmd = new SqlCommand("sms_downloadIncentive", con);
	cmd.CommandTimeout = 5000;
	cmd.CommandType = CommandType.StoredProcedure;
	da = new SqlDataAdapter(cmd);
	 try
     {
	 String filename ="Template-Incentive-Upload.xlsx";
	 System.Data.DataTable dt = new System.Data.DataTable("Overall - MIS");
	 da.Fill(dt);
	 ds = new DataSet();  
	 ds.Tables.Add(dt);  
	 if (dt.Rows.Count > 0)
	 {             
		ClosedXML.Excel.XLWorkbook wb = new ClosedXML.Excel.XLWorkbook();
		wb.Worksheets.Add(ds);
		wb.Worksheet(1).Row(1).InsertRowsAbove(1);
		Response.Clear();
		Response.Buffer = true;
		Response.Charset = "";
		Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
		Response.AddHeader("content-disposition", "attachment;filename="+filename+".xlsx");
		MemoryStream MyMemoryStream = new MemoryStream();
		wb.SaveAs(MyMemoryStream);
		MyMemoryStream.WriteTo(Response.OutputStream);
		Response.Flush();
		Response.End();
		}
		}catch(Exception e)
		{
			e.throws();
		}
	}
