package com.olympus.exceltest;

 

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.logging.Handler;
import java.util.logging.Logger;

 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.olympus.olyutil.Olyutil;
import com.olympus.olyutil.log.OlyLog;

@WebServlet("/xlread")
public class ReadExcelFile extends HttpServlet {
	private final Logger logger = Logger.getLogger(ReadExcelFile.class.getName()); // define logger
	/*****************************************************************************************************************************************************/
	public static ArrayList<String> readXlsFile(String xlsFileName) throws IOException {
		ArrayList<String> tgtArr = new ArrayList<String>();
		StringBuilder sb = new StringBuilder();
		int numCols = 0;
		try {
			File file = new File(xlsFileName); // creating a new file instance
			FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file
			// creating Workbook instance that refers to .xlsx file
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object
			Iterator<Row> itr = sheet.iterator(); // iterating over excel file
			FormulaEvaluator objFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) wb);
			
			
			DataFormatter objDefaultFormat = new DataFormatter();
			System.out.println("SheetName " + sheet.getSheetName());
			
			// This will find empty cells
			int j = 1;
			for(Row row : sheet) {
				   for(int cn=0; cn<row.getLastCellNum(); cn++) {
				       // If the cell is missing from the file, generate a blank one
				       // (Works by specifying a MissingCellPolicy)
				       Cell cell = row.getCell(cn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				       // Print the cell for debugging
				       if (j > 2 &&  j < 4) {
				             System.out.println("CELL: " + cn + " --> " + cell.toString());
				       }
				   }
				   j++;
				}
			
			   
			 // This will not find empty cells. 
			int r = 1;
			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator(); // iterating over each column
				int k = 0;
				//numCols = sheet.getRow(r).getLastCellNum();
				int firstRow = sheet.getFirstRowNum();
				int lastRow = sheet.getLastRowNum();
				
				short minColIdx = row.getFirstCellNum();
				short maxColIdx = row.getLastCellNum();
				
				
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

				//System.out.print("*** " + cell.getColumnIndex() + "--^^");
					Cell cellValue = row.getCell(k++);
				    objFormulaEvaluator.evaluate(cellValue); // This will evaluate the cell, And any type of cell will return string value
				    
				    
				    //String cellValueStr = objDefaultFormat.formatCellValue(cellValue,objFormulaEvaluator);
				    //System.out.println("r=" + r + "--" + cell.getColumnIndex() + "--" + cellValueStr);

				    //System.out.print(cellValueStr + "^"); 
				    if (r > 2 &&  r < 4) {
				    	/*
				    	if (cellValueStr.equals("") ||cellValueStr == null ) {
				    		cellValueStr = "MT";
				    	}
				    	
				    	*/
				    	 //System.out.println("SheetCellsMin:" + minColIdx + "-- SheetCellsMin:" + maxColIdx);
					    //System.out.println("** TEnum=" +  cell.getCellTypeEnum() + "** Type=" + 
				    	// objFormulaEvaluator.evaluate(cellValue)   + "-- r=" + r + "-- Col=" + cell.getColumnIndex() + "-- Val=" + cellValueStr);

				    	
				   
				    /*
				 
							switch (cell.getCellTypeEnum()) {
							case BOOLEAN:
								 System.out.print("-- Col=" + cell.getColumnIndex() + "--- " + cell.getBooleanCellValue() );
								//sb.append(cell.getBooleanCellValue());
								break;
							case STRING:
								 System.out.print("-- Col=" + cell.getColumnIndex() +  "--- " + cell.getRichStringCellValue().getString());
								//String cValue = cell.getRichStringCellValue().getString();
								//cValue = (cValue == null || cValue.equals("")) ? "--MT--" : cValue;
								
								//sb.append(cValue);
								
								break;
							case NUMERIC:
								if (DateUtil.isCellDateFormatted(cell)) {
									System.out.print(cell.getDateCellValue());
									//sb.append(cell.getDateCellValue());
								} else {
									 System.out.print("-- Col=" + cell.getColumnIndex() +  "--- " + cell.getNumericCellValue()  );
									//sb.append(cell.getNumericCellValue());
								}
								break;
							case FORMULA:
								 System.out.print(cell.getCellFormula());
								//sb.append(cell.getCellFormula());
								break;
							case BLANK:
								 System.out.print("-- Col=" + cell.getColumnIndex() +  "--- " +  "--MT--");
								
								//sb.append(cell.getStringCellValue().replaceAll("", "-MT-"));
								
								break;
							default:
								 System.out.print("-- Col=" + cell.getColumnIndex() +  "--- " + "-DF-");
								//sb.append(cell.getRichStringCellValue().getString());
							} // End Switch
							
					*/
				    }
					 
				} // End inner while
				r++;
				//tgtArr.add(sb.toString());
				//sb = null;

			} // End outer while
		} catch (Exception e) {
			e.printStackTrace();
		}

		return (tgtArr);

	}
	/*****************************************************************************************************************************************************/
	 
		// Service method
		@Override
		protected void doGet(HttpServletRequest req, HttpServletResponse res) throws ServletException, IOException {
			String fileName = "C:\\TEMP\\SAP_Asset_Upload\\SAP.xlsx";
			String sep = "\\^";
			String[] strSplit = null;
			String[] ss = null;
			DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss.SSS");  
			ArrayList<String> strArr = new ArrayList<String>();

			LocalDateTime now = null;
	 
			String logFileName = "sapDataUpload.log";
			String directoryName = "D:/Kettle/logfiles/sapDataUpload";
			Handler fileHandler =  OlyLog.setAppendLog(directoryName, logFileName, logger );
			strArr = readXlsFile(fileName);
			
			// 101-0010622-002
			//String tgtID = "101-0010622-002";
			String tgtID = "101-0010095-033";
			//System.out.println("arrSZ:" + strArr.size());
			 //Olyutil.printStrArray(strArr);
		}
	/*****************************************************************************************************************************************************/

}


 