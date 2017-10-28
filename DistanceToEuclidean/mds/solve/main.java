package mds.solve;

import mdsj.MDSJ;
import mdsj.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.Iterator;
import java.util.ListIterator;
import java.util.Vector;

/*
 * Author: Davis Smith
 * Date: October 10, 2017
 * Purpose: This project was created for two purposes. The first is to read in a compressed matrix
 * from an excel file (with zeros as a delimiter) and output a half matrix in proper format. 
 * The second (and main) purpose is to input a full matrix and use multidimensional scaling method 
 * to create euclidean coordinates from a distance matrix. The method for multidimensional scaling is taken from the MDSJ library. 
 * Half matrices were transposed and turned into full matrices by hand, using tools in Microsoft Excel. 
 */

public class main {

	public static void main(String[] args) throws IOException {
		
		String file = "bayg29";
		int dim = 29;
		double distanceMatrix[][] = new double[dim][dim];
		
		//Vector<Integer> tempMatrix = new Vector<>();
		//tempMatrix = readCompressedMatrix(file);
		//writeHalfMatrix(file, tempMatrix);
		
		distanceMatrix = readMatrixFromExcel(file, dim);
		writeToExcel(distanceMatrix, file, dim);
	}
	
	//-------------------------------------------------------------------------
	
	@SuppressWarnings("null")
	public static Vector<Integer> readCompressedMatrix(String file) {
		
		Vector<Integer> distMatrix = new Vector<>(); //Store distances in vector
		XSSFWorkbook workbook = new XSSFWorkbook();    
		XSSFSheet sheet = null;
		XSSFRow curRow = null;
		int rowCounter = 0; //initialize row counter
		int cellCounter = 0;
		String FILE_NAME = "./data/" + file + ".xlsx";

		try {
			FileInputStream fis = new FileInputStream(new File(FILE_NAME));
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
		}
		catch (Exception e) {
			System.out.println("File is not present");
			e.printStackTrace();
		}
		
		rowCounter = 7;
		curRow = sheet.getRow(rowCounter); //set current row to 8th row			
		int cell = (int)curRow.getCell(0).getNumericCellValue(); //set cell to 0
		Cell c = curRow.getCell(0);

		try {			
			while(!c.toString().equals("EOF") && !c.toString().equals("DISPLAY_DATA_SECTION")) {
				
				cellCounter = 0; //reinitialize cell to 0
				c = curRow.getCell(cellCounter);
					
				//loop through columns until null cell is reached
				while(c != null) {
					cell = (int)curRow.getCell(cellCounter).getNumericCellValue();
					distMatrix.add(cell);
					cellCounter++;
					c = curRow.getCell(cellCounter);
				}
				rowCounter++;
				curRow = sheet.getRow(rowCounter);
				c = curRow.getCell(0);
			}
		}
		catch(Exception e) {
			e.printStackTrace();
		}
				
		return distMatrix;
	}
	
	//------------------------------------------------------------------------
	
	public static void writeHalfMatrix(String file, Vector<Integer> distMatrix) throws IOException {
			//write newly formatted matrix
			
			XSSFWorkbook workbook = new XSSFWorkbook();    
			XSSFSheet sheet = workbook.createSheet("Sheet1");
			XSSFRow curRow = null;
			int rowCounter = 7; //initialize row counter
			int cellCounter = 0;
			int i = 0;
			Iterator<Integer> iterator = distMatrix.iterator();
			
			System.out.println("Printing half matrix...\n");
			
			curRow = sheet.createRow(rowCounter); //starting at row 7
					
					while(iterator.hasNext()) {
					
						if(distMatrix.elementAt(i) != 0) {
							curRow.createCell(cellCounter).setCellValue(distMatrix.elementAt(i));
							cellCounter++;
							System.out.print(distMatrix.elementAt(i) + " ");
						}
						else {
							curRow.createCell(cellCounter).setCellValue(distMatrix.elementAt(i));
							cellCounter = 0;
							rowCounter++;
							curRow = sheet.createRow(rowCounter); //starting at row 7
							System.out.println(distMatrix.elementAt(i));
						}
						iterator.next();
						i++;						
					}
					
					try {
						FileOutputStream outputStream = new FileOutputStream("./half-matrix-data/" + file + "-half.xlsx");
						workbook.write(outputStream);
					}
					catch(Exception e) {
						e.printStackTrace();
					}
					
					System.out.println("success");
					workbook.close();
					
			return;
	}
	
	//-------------------------------------------------------------------------
	
	public static double[][] readMatrixFromExcel(String file, int dim) throws IOException {
		
		double distMatrix[][] = new double[dim][dim];
		
		XSSFWorkbook workbook = new XSSFWorkbook();    
		XSSFSheet sheet = null;
		XSSFRow curRow;
		int rowCounter = 0; //initialize row counter
		String FILE_NAME = "./data/" + file + ".xlsx";
		
		try {
			FileInputStream fis = new FileInputStream(new File(FILE_NAME));
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
		}
		catch (Exception e) {
			System.out.println("File is not present");
			e.printStackTrace();
		}
				
		rowCounter = 7;
		curRow = sheet.getRow(rowCounter); //set current row to 8th row	
		double cell = curRow.getCell(0).getNumericCellValue(); //set cell to 0
		int i = 0;
		int k = 0;
		
		System.out.println("Printing full matrix...\n");
		for(i = 0; i < dim; i++) {
			cell = (double) curRow.getCell(0).getNumericCellValue();
			
			for(k = 0; k < dim; k++) {
				distMatrix[i][k] = cell;
				cell = (double) curRow.getCell(k).getNumericCellValue();
				System.out.print(cell + " ");
			}
			System.out.println("");
			curRow = sheet.getRow(++rowCounter);
		}
		
		return distMatrix;
	}
	
	//-------------------------------------------------------------------------
	
	@SuppressWarnings("null")
	public static void writeToExcel(double distMatrix[][], String file, int dim) {
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();    
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		XSSFRow curRow;
		int rowCounter = 0; //initialize row counter

		//write to excel
		System.out.println("\nWriting euclidean coordinates to excel...\n");
		
				workbook = new XSSFWorkbook(); // create a book
				sheet = workbook.createSheet("Sheet1");// create a sheet		
				
				double[][] output = MDSJ.classicalScaling(distMatrix); // apply MDS
				
				curRow = sheet.createRow(0);
				curRow.createCell(0).setCellValue("NAME");
				curRow.createCell(1).setCellValue(file);
				curRow = sheet.createRow(1);
				curRow.createCell(0).setCellValue("COMMENT");
				curRow = sheet.createRow(2);
				curRow.createCell(0).setCellValue("TYPE");
				curRow.createCell(1).setCellValue("TSP");
				curRow = sheet.createRow(3);
				curRow.createCell(0).setCellValue("DIMENSION");
				curRow.createCell(1).setCellValue(dim);
				curRow = sheet.createRow(4);
				curRow.createCell(0).setCellValue("EDGE_WEIGHT_TYPE");
				curRow.createCell(1).setCellValue("EUC_2D");
				curRow = sheet.createRow(5);
				curRow.createCell(0).setCellValue("NODE_COORD_SECTION");
				
				rowCounter = 6;
				
				System.out.println("Printing euclidean coordinates: \n");
				
				for(int j = 0; j<dim; j++) {  // output all coordinates
					curRow = sheet.createRow(rowCounter++);
				    System.out.println(output[0][j]+" "+output[1][j]);
				    curRow.createCell(0).setCellValue(j + 1);
				    curRow.createCell(1).setCellValue(output[0][j]);
				    curRow.createCell(2).setCellValue(output[1][j]);
				}
				
				try {
					FileOutputStream outputStream = new FileOutputStream("./euc-data/" + file + "-euc.xlsx");
					workbook.write(outputStream);

				}
				catch(Exception e) {
					e.printStackTrace();
				}
				
				return;
	}

}

