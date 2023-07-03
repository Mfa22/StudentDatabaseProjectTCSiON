package StudentDatabaseCreation;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class studentdatabase {

	public static void main(String[] args)throws Exception {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook();
			FileOutputStream out = new FileOutputStream(new File("D:/StudentDatabase.xlsx"));
			XSSFSheet Spreadsheet = workbook.createSheet("Database");
			
			Object empdata[][]= { {"Student Id","Student Name","Father Name","Mother Name","Address","Country","State","Pin Code","Mobile No.","Blood Group"},
					              {101,"Anas","Pervej","Khatoon","Delhi","India","Delhi",110025,98765432,"B positive"},
					              {102,"Akash","Pandit","Safia","Delhi","India","Delhi",110025,98765432,"B positive"},
					              {103,"Rashid","Saif","Sazma","Delhi","India","Delhi",110025,98765432,"B positive"},
					              {104,"Babu","Hussain","Shaiqa","Delhi","India","Delhi",110025,98765432,"B positive"}
			                    };
			
			int rows=empdata.length;
			int cols=empdata[0].length;
			
			System.out.println(rows);
			System.out.println(cols);
			
			for(int r=0;r<rows;r++)
			{
				XSSFRow row=Spreadsheet.createRow(r);
				
				for(int c=0;c<cols;c++)
				{
					XSSFCell cell=row.createCell(c);
					Object value=empdata[r][c];
					
					if(value instanceof String)
						cell.setCellValue((String)value);
					if(value instanceof Integer)
						cell.setCellValue((Integer)value);
					if(value instanceof Boolean)
						cell.setCellValue((Boolean)value);
				}
			}
			
			workbook.write(out);
			out.close();
			
		}
		// TODO Auto-generated method stub
        catch(Exception e) {
        	System.out.println(e);
        }
		
		
		System.out.println("Student Database file written successfully...");
	}

}
