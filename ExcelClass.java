package com.mavenproject.ExcelReadandWrite;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelClass {
	public static void main(String[] args) throws Exception {
		writeMethod();
		readMethod();
	}
	public static void writeMethod(){
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Sheet1");
		Row hrow = sheet.createRow(0);
		hrow.createCell(0).setCellValue("Name");
		hrow.createCell(1).setCellValue("Age");
		hrow.createCell(2).setCellValue("Email");
		Object[][] data = {
				{"JohnDoe",30,"john@test.com"},
				{"JaneDoe",28,"jane@test.com"},
				{"BobSmith",35,"jackey@example.com"},
				{"Swapnil",37,"swapnil@example.com"},
		};
		int rowNum=1;
		for (Object[] rowData : data) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue((String)rowData[0]);
			row.createCell(1).setCellValue((int)rowData[1]);
			row.createCell(2).setCellValue((String)rowData[2]);
		}
		try (FileOutputStream file = new FileOutputStream("Dummy.xlsx")) {
			workbook.write(file);
			System.out.println("File Created SuccessFully");
		} 
		catch (IOException e) {
			e.printStackTrace();
		}
		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println();
	}
	public static void readMethod() {
		FileInputStream file;
		try {
			file = new FileInputStream("C:\\Users\\mukil\\eclipse-workspace\\ExcelReadandWrite\\Dummy.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		Sheet sheet = workbook.getSheetAt(0);
		for (Row row : sheet) {
			for (Cell cell : row) {
				switch (cell.getCellType()) {
				case STRING:
					System.out.println(cell.getStringCellValue()+"\t\t");
					break;
				case NUMERIC:
					System.out.println(cell.getNumericCellValue()+"\t\t");
				default:
					break;
				}
			}
			System.out.println();
		}
		workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}