/**
 @author pankaj.singh
 **/


import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.Collection;


import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.google.gson.Gson;



public class FetchDataInExcel {

	private static final long serialVersionUID = 1L;
	
	
	
	 


	Connection con = null;
	PreparedStatement stmt = null;

	

	public void fillDataInExcel() {


		try {
			con = DataSourceConfiguration.getDatabaseConnection(); //Make sure your Data source is configured, you can use your own configuration here.

			stmt = con.prepareStatement(SQLConstants.YOUR_QUERY);  //Keep your query in a file 

			ResultSet rs = stmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			int count = rsmd.getColumnCount();
			if (rs != null) {

				try {

					HSSFWorkbook workbook = new HSSFWorkbook();

					HSSFSheet sheet = workbook.createSheet("Policies"); //Give sheet name
					HSSFRow rowhead = sheet.createRow((short) 0);

					for (int k = 1; k <= count; k++) {

						String name = rsmd.getColumnName(k);

						rowhead.createCell((short) (k - 1)).setCellValue(name); //creates column name in 1st column which matches with DB columns.
					}

					int i = 1;
					while (rs.next()) {

						HSSFRow row = sheet.createRow((short) i);
						for (int k = 1; k <= count; k++) {

							String name = rsmd.getColumnName(k);

							row.createCell((short) (k - 1)).setCellValue(rs.getString(name));
						}
						i++;

					}

					OutputStream out = response.getOutputStream();
					String fileName = "Policies(" + fromDate + "to"+toDate+").xls";  //Generated File name 
					//String fileName = "Policies.xls";
					response.setContentType("application/vnd.ms-excel");
					response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
					workbook.write(out);
					out.flush();
					out.close();

				} catch (SQLException e1) {
					e1.printStackTrace();
				} catch (FileNotFoundException e1) {
					e1.printStackTrace();
				} catch (IOException e1) {
					e1.printStackTrace();
				}

			}

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
