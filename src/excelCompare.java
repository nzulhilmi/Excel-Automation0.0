import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;

import org.apache.poi.hssf.record.OldCellRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelCompare {
	
	final static String[] columnsNeeded = {"TAM", "SSR", "Organization Name", "Contract Title", 
											"Contract Number", "Schedule Name", "Value", "TPID",
											"Service Name", "Start Date", "End Date"};
	
	final static String[] miaColumns = {"TAM", "SISR", "Customer", "Contract Title",
										"Contract", "Schedule Name", "RM Value", "TPID",
										"Service", "Start Date", "End Date"};
	
	final static String[] TAMs = {"abusmt", "alea", "amazahar", "colee",
								  "easonlau", "taufiqo", "tuchong", "mseng", "gurushr",
								  "huzaidim", "iansu", "jhew", "kansiva", "kkphoon",
								  "nabinti", "paerun", "sivask", "superuma"};
	
	//The number represent the indexes (starts at 0)
	final static int[] numberColumns = {6, 7};
			
	public static void main(String[] args) throws IOException {
		//check for number of arguments
		//if 2, proceed,
		//else, abort.
		
		String filepath1 = "C:/Users/t-ninikf/Downloads/ExcelSheets/HanimBook.xlsx";
		String filepath2 = "C:/Users/t-ninikf/Downloads/ExcelSheets/Book2.xlsx";
		
		File file1 = new File(filepath1);
		File file2 = new File(filepath2);
		long file1Size = file1.length();
		long file2Size = file2.length();
		
		//Check which file size is bigger. Bigger file size is Unicorn.
		if(file1Size > file2Size) {
			execute(file1, file2, filepath2);
		}
		else {
			execute(file2, file1, filepath2);
		}
	}
	
	//Read excel file
	public static void read(File file, ArrayList<ArrayList<String>> excelSheet, String sheetName) {
		try {
	        //Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(file);
	        
	        workbook.setMissingCellPolicy(MissingCellPolicy.RETURN_BLANK_AS_NULL);
	        DataFormatter fmt = new DataFormatter();
	        
	        XSSFSheet sheet;
	        if(sheetName.equals("null")) {
	        	sheet = workbook.getSheetAt(0);
	        }
	        else {
	        	sheet = workbook.getSheet(sheetName);
	        }
        	
        	for(int rn = sheet.getFirstRowNum(); rn <= sheet.getLastRowNum(); rn++) {
        		//Create 2nd-dimension array list
	            ArrayList<String> columns = new ArrayList<String>();
	            
        		Row row = sheet.getRow(rn);
        		if(row == null) {
        			//There is no data in this row. Need to handle appropriately
        			columns.add("null");
        		}
        		else {
        			for(int cn = 0; cn < row.getLastCellNum(); cn++) {
        				Cell cell = row.getCell(cn);
        				if(cell == null) {
        					//This cell is empty/blank/unused, handle appropriately
        					columns.add("null");
        				}
        				else {
        					String cellStr = fmt.formatCellValue(cell);
        					//Do something with the value
        					columns.add(cellStr);
        				}
        			}
        		}
        		excelSheet.add(columns);
        	}
        	//file.close();
        	workbook.close();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	//Write the excel file
	public static void write(File file, ArrayList<ArrayList<String>> excelSheet, String sheetName, 
			String filename, String filepath2) {
		
        try {
        	//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Tracker");
			
			int rowNum = excelSheet.size();
			int colNum = excelSheet.get(0).size();
			int rowCount = 0;
			
			for(int i = 0; i < rowNum; i++) {
				XSSFRow row = sheet.createRow(rowCount++);
				
				int colCount = 0;
				
				for(int j = 0; j < colNum; j++) {
					XSSFCell cell = row.createCell(colCount++);
				
					if((j == 6 || j == 7) && i > 0) {
						cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
						int value = Integer.parseInt(excelSheet.get(i).get(j));
						cell.setCellValue(value);
					}
					else {
						cell.setCellValue((String) excelSheet.get(i).get(j));
					}
				}
			}
			
			FileOutputStream outputStream = new FileOutputStream(filename);
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	//Execute the whole program
	public static void execute(File file1, File file2, String filepath2) {
		ArrayList<ArrayList<String>> excelSheet1 = new ArrayList<ArrayList<String>>();
		ArrayList<ArrayList<String>> excelSheet2 = new ArrayList<ArrayList<String>>();
		
		String sheetName1 = "Detailed Schedule";
		
		read(file1, excelSheet1, "Sheet1");
		
		//check excel sheet
		ArrayList<ArrayList<String>> output = filter(excelSheet1);
		
		int rowSize = output.size();
		int columnSize = output.get(0).size();
		
		sortColumns(output);
		renameColumns(output);
		
		for(int i = 1; i < rowSize; i++) {
			//formatDate_(output.get(i), columnSize - 1);
			correctPrice(output.get(i));
			convertLowerCase(output.get(i));
		}
		
		read(file2, excelSheet2, sheetName1);
		ArrayList<ArrayList<String>> output2 = filter2(excelSheet2);
		
		int rowSize2 = output2.size();
		//int columnSize2 = output2.get(0).size();
		for(int i = 1; i < rowSize2; i++) {
			//formatDate_(output2.get(i), columnSize2 - 1);
		}
		
		addRows(output, output2);
		removeComma(output2);
		//output2 is the file that contains all of the data

		
		write(file2, output2, sheetName1, "output1.xlsx", filepath2);
	}
	
	//Traverse through the columns, get the indexes of the columns needed
	public static ArrayList<Integer> checkColumns(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> columnNumbers = new ArrayList<Integer>();
		
		int columnSize = array.get(0).size();
		
		for(int i = 0; i < columnSize; i++) {
			if(Arrays.asList(columnsNeeded).contains(array.get(0).get(i))) {
				columnNumbers.add(i);
			}
		}
		
		return columnNumbers;
	}
	
	//Traverse through the rows, get the indexes of the rows needed
	public static ArrayList<Integer> checkRows(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> rowNumbers = new ArrayList<Integer>();
		
		int rowSize = array.size();
		
		for(int i = 0; i < rowSize; i++) {
			if(Arrays.asList(TAMs).contains(array.get(i).get(0))) {
				rowNumbers.add(i);
			}
		}
		
		return rowNumbers;
	}
	
	//Produce an output array with filtered columns and rows
	public static ArrayList<ArrayList<String>> filter(ArrayList<ArrayList<String>> array) {
		ArrayList<Integer> columnNumbers = checkColumns(array);
		ArrayList<Integer> rowNumbers = checkRows(array);
		ArrayList<ArrayList<String>> output = new ArrayList<ArrayList<String>>();
		
		int numberOfRows = array.size();
		int numberOfColumns = array.get(0).size();
		
		for(int i = 0; i < numberOfRows; i++) {
			//Check for row numbers
			if(rowNumbers.contains(i) || i == 0) {
				ArrayList<String> tempRow = new ArrayList<String>();
			
				for(int j = 0; j < numberOfColumns; j++) {
					if(columnNumbers.contains(j)) {
						tempRow.add(array.get(i).get(j));
					}
				}
				
				output.add(tempRow);
			}
		}
		
		return output;
	}
	
	//Helper method for filter
	public static ArrayList<ArrayList<String>> filter2(ArrayList<ArrayList<String>> array) {
		ArrayList<ArrayList<String>> output = new ArrayList<ArrayList<String>>();
		
		int rowNumbers = array.size();
		
		for(int i = 0; i < rowNumbers; i++) {
			ArrayList<String> temp = new ArrayList<String>();
			for(int j = 0; j < 11; j++) {
				temp.add(array.get(i).get(j));
			}
			output.add(temp);
		}
		
		return output;
	}
	
	//Sort/swap the columns
	public static void sortColumns(ArrayList<ArrayList<String>> array) {
		int columnSize = array.get(0).size();
		int rowSize = array.size();
		
		for(int i = 0; i < columnSize - 1; i++) {	
			if(!(array.get(0).get(i).equals(columnsNeeded[i]))) {	
				//look for index number of searched column/string
				for(int j = i + 1; j < columnSize; j++) {
					if(array.get(0).get(j).equals(columnsNeeded[i])) {
						for(int m = 0; m < rowSize; m ++) {
							Collections.swap(array.get(m), i, j);
							
						}
					}
				}
			}
		}
	}
	
	
	//Helper method for formatDate
	public static String formatDate(String date) {
		int dateLength = date.length();
		String temp = date.substring(0, dateLength - 2) + 
				"20" + date.substring(dateLength - 2, dateLength);
		
		return temp;
	}
	
	//Change the format of the date
	public static void formatDate_(ArrayList<String> array, int columnSize) {
		//int columnSize = array.get(0).size() - 1;
		
		String temp1 = formatDate(array.get(columnSize));
		String temp2 = formatDate(array.get(columnSize - 1));
		array.set(columnSize, temp1);
		array.set(columnSize - 1, temp2);
	}
	
	//Rename the columns according to Mia's tracker sheet
	public static void renameColumns(ArrayList<ArrayList<String>> array) {
		int columnSize = array.get(0).size();
		
		for(int i = 1; i < columnSize; i++) {
			if(!array.get(0).get(i).equals(miaColumns[i])) {
				array.get(0).set(i, miaColumns[i]);
			}
		}
	}
	
	//Convert all the names in column 1 to lowercase
	public static void convertLowerCase(ArrayList<String> array) {
		String temp = array.get(1);
		array.set(1, temp.toLowerCase());
	}
	
	//Remove trailing zeros (cents) in the price value. i.e RM400.00 to RM400
	public static void correctPrice(ArrayList<String> array) {
		String temp = array.get(6);
		int length = temp.length();
		array.set(6, temp.substring(0, length - 3));
	}
	
	//Compare rows, if they are equal, returns true. False otherwise
	public static boolean compareRows(ArrayList<String> array1, ArrayList<String> array2) {
		int columnSize = array2.size();
		
		boolean b1 = true;
		for(int i = 0; i < columnSize; i++) {
			if(!array1.get(i).equals(array2.get(i))) {
				b1 = false;
			}
		}
		
		return b1;
	}
	
	//Check if the row is in the excel sheet array1, returns true if it's there. false otherwise
	public static boolean checkExistence(ArrayList<ArrayList<String>> array1, ArrayList<String> array2) {
		int rowSize = array1.size();
		
		boolean b1 = false;
		
		for(int i = 0; i < rowSize; i++) {
			if(compareRows(array1.get(i), array2)) {
				b1 = true;
			}
		}
		
		return b1;
	}
	
	//Add missing rows to array2.
	public static void addRows(ArrayList<ArrayList<String>> array1, ArrayList<ArrayList<String>> array2) {
		int rowSize = array1.size();
		
		for(int i = 0; i < rowSize; i++) {
			if(!checkExistence(array2, array1.get(i))) {
				array2.add(array1.get(i));
			}
		}
	}
	
	//Remove commas, i.e. 20,000 to 20000
	public static void removeComma(ArrayList<ArrayList<String>> array) {
		int rowNum = array.size();
		int colNum = array.get(0).size();
		
		for(int i = 0; i < rowNum; i++) {
			String value = array.get(i).get(6);
			value = value.replace(",", "");
		
			array.get(i).set(6, value);
		}
	}
	
	//Create a text file with error message
	public void printError(String errorMessage) {
		
	}
	
	//Print array to system log output
	public static void printArray(ArrayList<ArrayList<String>> array) {
		int rows = array.size();
		int columns = array.get(0).size();
		
		for(int i = 0; i < rows; i++) {
			for(int j = 0; j < columns; j++) {
				System.out.print(array.get(i).get(j) + " | ");
			}
			System.out.println();
		}
	}
}
