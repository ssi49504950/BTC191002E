import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Practice {

	public void practice() throws Exception {
		File filePath = new File("C:\\Users\\ssi49\\Documents\\practicefile.xlsx");
		FileInputStream stramedFile = new FileInputStream(filePath);
		
		XSSFWorkbook workbook  = new XSSFWorkbook(stramedFile);
		int size = workbook.getNumberOfSheets();
		int rowNumber = 0;
		int cellNumber =0;
		boolean flag = false;
		for (int i=0; i<size;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("stringonly")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				while(rows.hasNext()) {
					Row gotRow = rows.next();
					Iterator<Cell> cells = gotRow.cellIterator();
					
					while(cells.hasNext()) {
						
						if(cells.next().getStringCellValue().equalsIgnoreCase("gsan")) {
							System.out.println("Found gsan on row index number: " + rowNumber+", Cell Index Number = " + cellNumber);
							flag = true;
							break;
						} cellNumber++;
						
					}
					if(flag) {
						break;
					}
					rowNumber++;
				}
			}
		}
	}
}
