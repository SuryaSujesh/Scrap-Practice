package utility;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {

	public String path;

	FileInputStream Fis;
	FileOutputStream Fos;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	Row r;
	Cell c;

	public ExcelUtility(String path) { this.path = path; } 

	public int getrowcount(String Sheetname) throws IOException {
		Fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(Fis);
		sheet = workbook.getSheet(Sheetname);
		int rowcount = sheet.getLastRowNum();
		workbook.close();
		Fis.close();
		return rowcount;
	}

	public int getcellcount(String Sheetname, int rownum) throws IOException {
		Fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(Fis);
		sheet = workbook.getSheet(Sheetname);
		Row row = sheet.getRow(rownum);
		int cellcount = row.getLastCellNum();
		workbook.close();
		Fis.close();
		return cellcount;

	}

	public String setcelldata(String sheetname, int rownum, int cellnum, String data) throws IOException 
	{
		Fis= new FileInputStream(path);
		workbook = new XSSFWorkbook(Fis);
		if(workbook.getSheetIndex(sheetname)==-1)

		workbook.createSheet(sheetname); 
		sheet=workbook.getSheet(sheetname);


		if(sheet.getRow(rownum)==null)

		sheet.createRow(rownum);
		sheet.getRow(rownum).createCell(cellnum).setCellValue(data);

		Fos = new FileOutputStream(path);
		workbook.write(Fos); 
		workbook.close(); 
		Fis.close();
		Fos.close(); 

		return data;
	}

}
