import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataAccessPractice {
	
	public void accessExcelFile() throws Exception {
		File filePath = new File("C:\\Users\\ssi49\\Documents\\practicefile.xlsx");
		FileInputStream stramedFile = new FileInputStream(filePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(stramedFile);
		int sheetsCount = workbook.getNumberOfSheets(); //Get Sheet Number Size
		//System.out.println(sheetsCount);
		for(int i=0; i<sheetsCount;i++) {
			
			if(workbook.getSheetName(i).equalsIgnoreCase("main")) {
				XSSFSheet retrivedSheet = workbook.getSheetAt(i); // Get access to a specific sheet
				Iterator<Row> rows = retrivedSheet.iterator(); // Get access to all rows of that retrieved or specific sheet
				Row firstRow = rows.next(); // get access to first row
				Iterator<Cell> cells = firstRow.cellIterator(); // Get access to all Cell of that specific row
				
					// Cell firstCell = cells.next(); // get access to first cell of the first row		
					// String firstCellValue = firstCell.getStringCellValue(); // get the value of that cell
				    // System.out.println(firstCellValue);
					//Cell has a method call getCellTypeEnum which will provide what value format the cell has. if its a string or int or any other format it will provide it.
				
				
				
				//***********************
				
					while(cells.hasNext()) {
						Cell value = cells.next();
						if (value.getStringCellValue().equalsIgnoreCase("islam")) {
							System.out.println("islam");
						} else {
							System.out.println("Not found what you are looking for");
						}
					}
				
				//**********************
					//***********************
					ArrayList<String> arr = new ArrayList<String>();
					while(cells.hasNext()) {
						Cell value = cells.next();
						// getCellTypeEnum()==CellType.STRING check either cell has String value or not
						if(value.getCellTypeEnum()==CellType.STRING) {
							arr.add(value.getStringCellValue());
						}
						else {
							// NumberToTextConverter.toText(value.getNumericCellValue()) convert number to string and its is Apache Poi Api class and method
							arr.add(NumberToTextConverter.toText(value.getNumericCellValue()));
						}
						
					}
				
				//**********************
				
			}
		}
		
		//[NOTE : Iterator<Row> gets all the rows and Iterator<Cell> gets all the the cells]
		//[NOTE : Cell has a method call getCellTypeEnum which will provide what value format the cell has. if its a string or int or any other format it will provide it.]
		
	}

}
