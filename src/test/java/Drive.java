import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Drive {

	public static void main(String[] args) throws FileNotFoundException {
		// TODO Auto-generated method stub
		 FileInputStream fis=new FileInputStream("C://Users//sagar//OneDrive//Documents//DemoData.xlsx");
			XSSFWorkbook workbook =new XSSFWorkbook();
			int sheets = workbook.getNumberOfSheets();
			
			for(int i=0; i<sheets; i++)
			{
				if(workbook.getSheetName(i).equalsIgnoreCase("first"))
						{
					XSSFSheet sheet = workbook.getSheetAt(i);
					Iterator<Row> row = sheet.rowIterator();
					Row firstrow = row.next();
					Iterator<Cell> cell = firstrow.cellIterator();
					int k=0;
					int column=0;
					while(cell.hasNext())
					{
						Cell value = cell.next();
						if(value.getStringCellValue().equalsIgnoreCase("testcase"))
								{
							column=k;
								}
					}
					System.out.println(column);
				
						}

	}
	}
}


