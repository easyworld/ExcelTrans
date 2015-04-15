package com.test.ExcelTransformer;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	public static void main(String[] args) {
		if (args.length != 2) {
			System.out.println("Wrong number of args");
			return;
		}
		try {
			XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(args[0]));
			XSSFWorkbook wb2 = new XSSFWorkbook();
			int numberOfSheets = wb.getNumberOfSheets();
			XSSFSheet sheet;
			XSSFRow row;
			XSSFCell cell;
			double value;
			double[][] matrix;
			for (int i = 0; i < numberOfSheets; i++) {
				sheet = wb.getSheetAt(i);
				int counter = 0;
				while (sheet.getRow(counter) != null
						&& sheet.getRow(counter).getCell(0) != null
						&& sheet.getRow(counter).getCell(0).getRawValue() != null)
					counter++;
				int rowNum = counter;
				counter = 0;
				while (sheet.getRow(0).getCell(counter) != null
						&& sheet.getRow(0).getCell(counter).getRawValue() != null)
					counter++;
				int cellNum = counter;
				System.out.println(rowNum + "\t" + cellNum);
				matrix = new double[rowNum][cellNum];
				for (int j = 0; j < rowNum; j++) {
					row = sheet.getRow(j);
					for (int k = 0; k < cellNum; k++) {
						cell = row.getCell(k);
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							value = cell.getNumericCellValue();
							break;
						case Cell.CELL_TYPE_STRING:
							value = Double.parseDouble(cell
									.getStringCellValue());
							break;
						default:
							value = 0;
							break;
						}
						matrix[j][k] = value;
						// System.out.print(value + " ");
					}
					// System.out.println();
				}
				// 读取数据完成
				System.out.println("读取" + wb.getSheetName(i) + "数据完成");
				// 写入新的excel

				sheet = wb2.createSheet();
				wb2.setSheetName(i, wb.getSheetName(i));
				for (int j = 0; j < cellNum; j++) {
					for (int k = 0; k < rowNum; k++) {
						row = sheet.createRow(j * rowNum + k);
						cell = row.createCell(0);
						cell.setCellValue(matrix[k][j]);
					}
				}
			}
			wb2.write(new FileOutputStream(args[1]));
			System.out.println("写入" + args[1] + "完成");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
