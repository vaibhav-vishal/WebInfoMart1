/*
 * @Author Vaibhav Vishal
 */
package com.webInfoMart.util;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Simple class to write data into excel file
 */
public class ExcelWriter extends ExcelReader {
	
	public static void exportDataToExcel(String fileName, String[][] data) throws FileNotFoundException, IOException {
		// Create new workbook
		Workbook wb = new XSSFWorkbook();
		FileOutputStream fileOut = new FileOutputStream(fileName);
		Sheet sheet = wb.createSheet();

		// Create 2D Cell Array
		Row[] row = new Row[data.length];
		Cell[][] cell = new Cell[row.length][];

		// Define and Assign Cell Data from Given
		for (int i = 0; i < row.length; i++) {
			row[i] = sheet.createRow(i);
			cell[i] = new Cell[data[i].length];

			for (int j = 0; j < cell[i].length; j++) {
				cell[i][j] = row[i].createCell(j);
				cell[i][j].setCellValue(data[i][j]);
			}

		}

		// Export Data
		wb.write(fileOut);
		wb.close();
		fileOut.close();
		System.out.println("File exported successfully");
	}
}
