/*
 * @Autor Vaibhav Vishal
 */
package com.webInfoMart.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Takes two excel files as input along with sequence for final Excel
 * Stores then manipulates the data in two excels to a 2D array to be fed to ExcelWriter class
 */
public class ExcelReader {
	/*
	 * main method for running and sample/test purpose
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {

		// sample input
		File excel1 = new File("C:/Users/vaibhav/Desktop/Employee1.xlsx");
		File excel2 = new File("C:/Users/vaibhav/Desktop/Employee2.xlsx");

		String[][] file1op = readFile(excel1);
		for (int row = 0; row < file1op.length; row++) {
			for (int col = 0; col < file1op[row].length; col++) {
				if (file1op[row][col] == null)
					break;
				System.out.println(file1op[row][col]);
			}
		}
		String[][] file2op = readFile(excel2);
		for (int row = 0; row < file2op.length; row++) {
			for (int col = 0; col < file2op[row].length; col++) {
				if (file2op[row][col] == null)
					break;
				System.out.println(file2op[row][col]);
			}
		}

		System.out.println("\n \n \n");
		String[] sequence = { "file2 col1", "file1 col2", "file1 col3", "file1 col5", "file1 col1", "file2 col2",
				"file2 col3", "file1 col4" };
		String[][] abc = finalResult(file1op, file2op, sequence);
		for (int row = 0; row < abc.length; row++) {
			for (int col = 0; col < abc[row].length; col++) {
				if (abc[row][col] == null)
					break;
				System.out.println(abc[row][col]);
			}
		}

		ExcelWriter.exportDataToExcel("FinalData.xlsx", abc);

	}

	// reads and returns excel into a 2D array
	public static String[][] readFile(File excel1) {
		try {

			FileInputStream fis = new FileInputStream(excel1);
			XSSFWorkbook book = new XSSFWorkbook(fis);
			XSSFSheet sheet = book.getSheetAt(0);

			Iterator<Row> itr = sheet.iterator();
			String[][] file1op = new String[100][100];
			int i = 0, j = 0;
			// Iterating over Excel file in Java
			while (itr.hasNext()) {
				Row row = itr.next();

				// Iterating over each column of Excel file
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();
					String temp = "";

					switch (cell.getCellType()) {

					case Cell.CELL_TYPE_STRING:
						temp += cell.getStringCellValue();
						break;
					case Cell.CELL_TYPE_NUMERIC:
						temp += cell.getNumericCellValue();
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						temp += cell.getBooleanCellValue();
						break;
					default:
						temp = "";

					}
					file1op[i][j] = temp;
					j++;
				}
				i++;
				j = 0;
			}
			book.close();
			fis.close();
			return file1op;
		} catch (FileNotFoundException fe) {
			fe.printStackTrace();
		} catch (IOException ie) {
			ie.printStackTrace();
		}
		return null;
	}

	// maniplutes data from two excels into a single 2D array according to sequence
	public static String[][] finalResult(String[][] file1op, String[][] file2op, String[] sequence) {


		String[][] resFile = new String[file1op.length][sequence.length];
		for (int i = 0; i < sequence.length; i++) {
			String[] arrOfStr = sequence[i].split(" ");
			String firstPart = arrOfStr[0];
			String secondPart = arrOfStr[1];
			String[] fileno = firstPart.split("file");
			String[] colno = secondPart.split("col");
			int fileNumber = Integer.parseInt(fileno[1]);
			int colNumber = Integer.parseInt(colno[1]);
			String[][] t = file1op;
			if (fileNumber == 2)
				t = file2op;

			for (int j = 0; j < t.length; j++) {
				resFile[j][i] = t[j][colNumber - 1];
			}
		}
		return resFile;
	}

}
