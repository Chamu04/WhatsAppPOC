package whatsApp;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;

import resources.base;

public class excelDriven extends base{	
	
	public String month;
	public String mobile;
	public String name;
	
	public excelDriven() throws IOException
	{
		super.initializeDriver();
	}
	
	public String get_Month(){
		   Date date = new Date();
		   SimpleDateFormat formatter = new SimpleDateFormat("MMMMMMMMM");
		   return month = formatter.format(date);
		 }

	public void readXLSXFile(String fileName) {
		InputStream XlsxFileToRead = null;
		XSSFWorkbook workbook = null;
		try {
			XlsxFileToRead = new FileInputStream(fileName);
			
			//Getting the workbook instance for xlsx file
			workbook = new XSSFWorkbook(XlsxFileToRead);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		//getting the first sheet from the workbook using sheet name. 
		// We can also pass the index of the sheet which starts from '0'.
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		XSSFRow row;
		XSSFCell cell;
		
		//Iterating all the rows in the sheet
		Iterator rows = sheet.rowIterator();

		while (rows.hasNext()) {
			row = (XSSFRow) rows.next();
			
			//Iterating all the cells of the current row
			Iterator cells = row.cellIterator();

			while (cells.hasNext()) {
				cell = (XSSFCell) cells.next();

				
				if (cell.getCellType() == CellType.STRING) {
					System.out.print(cell.getStringCellValue() + " ");
					name = cell.getStringCellValue();
				} //else if (cell.getCellType() == CellType.NUMERIC) {
//					System.out.print(NumberToTextConverter.toText(cell.getNumericCellValue()) + " ");
//				} else if (cell.getCellType() == CellType.BOOLEAN) {
//					System.out.print(cell.getBooleanCellValue() + " ");

				//} 
			else { // //Here if require, we can also add below methods to
							// read the cell content
							// XSSFCell.CELL_TYPE_BLANK
							// XSSFCell.CELL_TYPE_FORMULA
							// XSSFCell.CELL_TYPE_ERROR
					
					mobile = NumberToTextConverter.toText(cell.getNumericCellValue());
					System.out.println(mobile);
				}
		}
			excelDriven.driver.get("https://api.whatsapp.com/send?phone="+mobile+"&text=Hi"+name+" Please submit the timesheet for the month of "+month+"&source=&data=");
			//https://api.whatsapp.com/send?phone=&text=urlencodedtext&source=&data=
			excelDriven.driver.manage().window().maximize();
			
			System.out.println();
			try {
				XlsxFileToRead.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
			
		}
		
	}

public static void main(String [] args) throws IOException {
	
	excelDriven readXlsx = new excelDriven();
	readXlsx.readXLSXFile("D:\\Pramod\\Projects\\BRT Portal\\WhatsApp\\EmpWhatsApp.xlsx");	
	
}}