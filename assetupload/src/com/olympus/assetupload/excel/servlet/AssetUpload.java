package com.olympus.assetupload.excel.servlet;

 // Run: http://localhost:8181/assetupload/assetup?id=101-0008740-006

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.logging.Handler;
import java.util.logging.LogManager;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.olympus.olyutil.Olyutil;
import com.olympus.olyutil.log.OlyLog;
 
 
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.extractor.XSSFBEventBasedExcelExtractor;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlException;
import org.xml.sax.SAXException;

 
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;

@WebServlet("/assetup")
public class AssetUpload extends HttpServlet {
	
	private final Logger logger = Logger.getLogger(AssetUpload.class.getName()); // define logger

	public enum colLetters {
	    A, B, C, D, F, G, H, I, J, O, P, Q, V, W, Z, AN, AP
	}
	
	//public static String contractID = "";
	
	private List<String> getSheets(String testFileName) throws Exception {
	    OPCPackage pkg = OPCPackage.open( testFileName);
	    List<String> sheetTexts = new ArrayList<String>();
	    XSSFBReader r = new XSSFBReader(pkg);
	    //        assertNotNull(r.getWorkbookData());
	    //      assertNotNull(r.getSharedStringsData());
	    
	    
	    XSSFBSharedStringsTable strings = new XSSFBSharedStringsTable(pkg);
	    XSSFBReader xssfbReader = new XSSFBReader(pkg);
	    XSSFBStylesTable styles = xssfbReader.getXSSFBStylesTable();
	    XSSFBReader.SheetIterator it = (XSSFBReader.SheetIterator) xssfbReader.getSheetsData();

	    return sheetTexts;
	}

	/*****************************************************************************************************************************************************/
	public static void printStrArray_2(ArrayList<String> strArr, String tag) {
		 int i = 0;
		for (String str : strArr) { // iterating ArrayList
			 if ( i > 172321 && i < 173340) {
				
			//if ( i > 86420 && i < 86440) {	
				System.out.println(tag + str);	
			 }
			 i++;		
		}
		// System.out.println(names[index]);
	}
	/*****************************************************************************************************************************************************/
	public static ArrayList<String> readXlsbFile(String xlsbFileName) throws IOException {
		ArrayList<String> tgtArr = new ArrayList<String>();
       
        ExcelSheetHandler  exlSheetHandler = new ExcelSheetHandler();       
        OPCPackage pkg;
        XSSFEventBasedExcelExtractor ext = null;
        try {
            pkg = OPCPackage.open(xlsbFileName);
           
           
            /*****************************************************************************************************/
            
            XSSFBReader r = new XSSFBReader(pkg);
            XSSFBSharedStringsTable sst = new XSSFBSharedStringsTable(pkg);
            XSSFBStylesTable xssfbStylesTable = r.getXSSFBStylesTable();
            XSSFBReader.SheetIterator it = (XSSFBReader.SheetIterator) r.getSheetsData();
     
          

            	//System.out.println("PTH:" + r.getAbsPathMetadata().toString());

            List<String> sheetTexts = new ArrayList<>();
            InputStream is = null;
            
            while (it.hasNext()) {
                is = it.next();
                String name = it.getSheetName();
                 
                 
                if (! name.equals("Asset Upload")) {
                	continue;
                }
              
          
				// SheetHandler exlSheetHandler = new SheetHandler();
				exlSheetHandler.startSheet(name);
				XSSFBSheetHandler sheetHandler = new XSSFBSheetHandler(is, xssfbStylesTable, it.getXSSFBSheetComments(),
						sst, exlSheetHandler, new DataFormatter(), false);

				sheetHandler.parse();
				exlSheetHandler.endSheet();

				// sheetTexts.add(exlSheetHandler.toString());

				//System.out.println("ST:" + sheetTexts.size());
				//System.out.println("STL:" + sheetTexts.add(exlSheetHandler.toString()));
                
                /*-----------------------------------------------------------------------------------------------------------*/

            }
            
			// System.out.println("output text:"+sheetTexts);
/*
			String[] arr = exlSheetHandler.toString().split("\n");
			//System.out.println("arrSZ:" + arr.length);
			int j = 0;
			for (String s : arr) {
				tgtArr.add(s);
				if (j > 86438 && j < 86435) {
					System.out.println("********" + s +" ");
				}
				
				j++;
			}
			
	*/		
        	String[] arr = exlSheetHandler.toString().split("\n");
            
           // System.out.println("arrSZ:" + arr.length);
            
        	int m = 1;
            for(String s1: arr) { // add String to tgtArr
            	tgtArr.add(s1);
            	if (m > 110795 && m < 110801) {
            		//System.out.println("**!!!***" + s1 +" ");
            	}
            	m++;
            }
		 if (is != null) {
			 is.close();
		 }		
        } catch (InvalidFormatException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (OpenXML4JException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (SAXException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    
       // System.out.println("tgtArr SZ=" + tgtArr.size());
		return(tgtArr);
	}
	/*****************************************************************************************************************************************************/
	public static List<Map<String, String>> findDateinData(ArrayList<String> strArr, String date) {
		ArrayList<String> rtnArr = new ArrayList<String>();
		 List<Map<String, String>> newMapArr  = new ArrayList<Map<String,String>>();
		 HashMap<String, String> rMap = new HashMap<String, String>();
		 
		 
		 
		 System.out.println("*** Query for Date=" + date + "--");
		 
		 
		 return (newMapArr); 
	}

	/*****************************************************************************************************************************************************/

	
	public static List<Map<String, String>> findIDinData(ArrayList<String> strArr, String contractID) {
		ArrayList<String> rtnArr = new ArrayList<String>();
		 List<Map<String, String>> newMapArr  = new ArrayList<Map<String,String>>();
		 HashMap<String, String> rMap = new HashMap<String, String>();

		String cID = "";
		String contract = "";
		int n = 0;
		int k = 0;
		for (String s1 : strArr) { // add String to tgtArr
			if (s1.contains("Order Type") || s1.contains("sheet name")) {
				//System.out.println("^^^ Skip Header");
				n++;
				continue;
			}
			/*
			if (n < 10) { // debug
				System.out.println("***^^^*** LineNo=" + n + "-- S=" + s1 + "--");
			}
			 */
			contract = getContractID(s1, n);
			/*
			if (n > 110795 && n < 110800) { // debug
				System.out.println("***^FID^^*** LNum="  + n + "-- contract RTN=" + contract + "--");
			}
			*/
			if ((!Olyutil.isNullStr(contract)) && contract.equals(contractID)) {
				/*
				if (n > 110795 && n < 110800) { // debug
					System.out.println("***^FID^^*** LineNo=" + n + "-- S=" + s1 + "--");
				}
				*/
				rtnArr.add(s1);
				rMap = getColData(s1, n);
				newMapArr.add(k++, rMap);
			}

			
			n++;
		}
		return (newMapArr);
	}
	
	/*****************************************************************************************************************************************************/
	
	/*****************************************************************************************************************************************************/
	public static String getContractID(String s, int n) {
		//System.out.println("*** n=" + n + "-- S=" +s );
		String idVal = "";
		String idxKey = "C";
		if (! s.contains("Order Type")) {
			//System.out.println("^^^ Skip Header");
			//String[] parts = s.split("\\^");
			String[] parts = s.split("^");
			
			int pSZ = parts.length;

			int lineNo = new Scanner(parts[0]).useDelimiter("\\D+").nextInt();
			idxKey += String.valueOf(lineNo);

			int indexOfSubStr = s.toUpperCase().indexOf(idxKey.toUpperCase());
			
			//if (n >9841 && n < 9850) {
				//.out.println("*** indexOfSubStr=" + indexOfSubStr);
			//}
			
			if (indexOfSubStr > 0) {
				int nextDelim = s.indexOf('^', indexOfSubStr);
				String sub = s.substring(indexOfSubStr, nextDelim);
				String[] id = sub.split("%");
				idVal = id[1];
			} else {
				idVal = "";
			}
			
			// System.out.println("*** ID=" + idVal + "-- LN=" + lineNo + "-- idxKey=" +
			// idxKey + "-- idx=" + indexOfSubStr + "-- sub=" + sub + "--");
		}
		return (idVal);
	}
	/*****************************************************************************************************************************************************/

	public static HashMap<String, String> parseRowData( String str, int lineNo) {
		
		 List<Map<String , String>> newMap  = new ArrayList<Map<String,String>>();
		 HashMap<String, String> colMap = new HashMap<String, String>();
		 String cellValue_t = "";
		 String cellValue = "";
		int  nextIdx = -1;
		String colTag = "";
		int idx = -1;
		//if (lineNo > 110795 && lineNo < 110800) {	
			//System.out.println(str);
		//}
		int k = 0;
		for (colLetters tag : colLetters.values()) { 
			colTag = tag.toString() + lineNo;
			idx = str.indexOf(colTag);
			if (idx < 0 ) {
				//System.out.println("** Col DNE:" + colTag);
				colMap.put(colTag, "");
			} else {
				nextIdx = str.indexOf('^', idx);
				cellValue_t = str.substring(idx, nextIdx);
				String[] cells = cellValue_t.split("%");
				cellValue = cells[1];
				
				//if (lineNo > 110795 && lineNo < 110800) {	
					//System.out.println(str);
					 //System.out.println("** LN=" + lineNo + "-- NI=" + nextIdx + "-- CV=" + cellValue + "--");
				//}
				
				colMap.put(colTag, cellValue);
				cellValue = "";
			}
				
			/* 
			if (lineNo > 110795 && lineNo < 110800) {	// debug
				//System.out.println("**PRD*** LN=" + lineNo + "-- col=" + tag + "-- colTag=" + colTag + "-- idx=" + idx + "--"); 
			}
			*/
			;
			colTag = "";
		} // End for
	
		
	
		
		/*
		for (Map.Entry<String, String> entry : colMap.entrySet()) {
			   System.out.println("*** colMao --  Key:" + entry.getKey()+" --> Value:" +entry.getValue() + "--");
		}
		*/
		return(colMap);
	}

	/*****************************************************************************************************************************************************/
	public static HashMap<String, String> getColData(String s, int lineNo) {
		// List<Map<String , String>> map  = new ArrayList<Map<String,String>>();
		 
		 HashMap<String, String> cMap = new HashMap<String, String>();
		 
		String idVal = "";
		String idxDelim = "";
		String idxKey = "";
		
		//char charAt0 = parts[1].charAt(0);
		
		char charDelim = s.charAt(0);
		/*if (lineNo > 110795 && lineNo < 110800) {
			System.out.println("***GCD*** LN=" + lineNo + "-- Process row:" + s);
		}*/

		idxDelim = Character.toString(charDelim); // get first char
		//int idxOfDelim = sp[1].indexOf('^');
		String[] sp = s.split("%");
			 
		char p_charAt1 = sp[0].charAt(1);
		idxKey = Character.toString(p_charAt1); // get first char

		int pSZ = sp.length;
		//System.out.println("Line=" + s);
		//int lineNo = new Scanner(sp[1]).useDelimiter("\\D+").nextInt();  // get digits from string
		 //System.out.println("*^^* LN=" + lineNo + "-- p_charAt1" + p_charAt1 + "-- sp[0]=" + sp[0] + "--");
		/*********************************************************************************************************************************/
		idxKey += String.valueOf(lineNo);
		//System.out.println("*** idxDelim:" + idxDelim  + "-- ColChar=" + p_charAt1  +  "-- idxKey:" + idxKey + "-- s:" + s);
		cMap = parseRowData(s, lineNo);
		
		
		return(cMap);
	}
	
	/*****************************************************************************************************************************************************/

	public static void displayMapArr(List<Map<String, String>> rowMap) {
		int i = 0;
		for (Map<String, String> map : rowMap) {
			//if (i < 2) {

				Map<String, String> treeMap = new TreeMap<>(map);
				for (Map.Entry<String, String> entry : treeMap.entrySet()) {
					System.out.println("*** Key:" + entry.getKey() + " --> Value:" + entry.getValue() + "--");
				}
				//System.out.println("*****************************************************************************************************");
			//}
			//i++;
		}    
	}
	
	/*****************************************************************************************************************************************************/

	/*****************************************************************************************************************************************************/
 
	// Service method
	@Override
	protected void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		boolean processDate = false;
		boolean proceed = false;
		String idParam = "";
		String processType = "";
		String xlsbFileName = "C:\\Java_Dev\\props\\SAP_Upload\\SAP_Asset_Upload\\new asset upload cumulative through current.xlsb";
		HashMap<String, String> dataMap = new HashMap<String, String>();
		List<Map<String, String>> rowMap = new ArrayList<Map<String, String>>();
		String dispatchJSP = "/assetuploaddetail.jsp";

		String dispatchJSP_Err = "/assetuperror.jsp";
		String contractID = "";
		contractID = ""; // 102-0017647-001
		ArrayList<String> strArr = new ArrayList<String>();
		ArrayList<String> tgtArr = new ArrayList<String>();
		// String xlsbFileName =
		String sep = "\\^";
		String[] strSplit = null;
		String[] ss = null;
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss.SSS");
		LocalDateTime now = null;
		String idVal = "";
		String dateParam = "";

		String logFileName = "sapAssetUpload.log";
		String directoryName = "D:/Kettle/logfiles/sapAssetUpload";
		Handler fileHandler = OlyLog.setAppendLog(directoryName, logFileName, logger);
		String paramName = "id";
		String paramValue = request.getParameter(paramName);
		
		String paramDate2 = request.getParameter("date2");
		//System.out.println("%%%%%% DATE=" + paramDate2 + "--");
		
		/*
		 * if ( ! Olyutil.isNullStr(paramValue) ){ contractID= paramValue.trim();
		 * //System.out.println("*** contractID:" + contractID + "--"); } else {
		 * 
		 * }
		 */

		if (request.getParameterMap().containsKey("id")) {
			 
			idParam = "id";
			contractID = request.getParameter(idParam).trim();
			System.out.println("%%%%%% IDVal=" + contractID + "--");
			processType = "id";

		} else if (request.getParameterMap().containsKey("date2")) {
			processDate = true;
			idParam = "date2";
			dateParam = request.getParameter(idParam);
			System.out.println("*****%%%%%% DateParam=" + dateParam + "--");
		}

		if (processType.equals("id")) { // process contractID
			
			// check ID size
			if (contractID.length() != 15) {
				System.out.println("%%%%%% ID is not valid -- ID SZ=" + contractID.length());
				
				now = LocalDateTime.now();
				logger.info(dtf.format(now) + ": " + "------------------ Forward to Error JSP: " + dispatchJSP_Err);
				request.getRequestDispatcher(dispatchJSP_Err).forward(request, response);
				return;
			} else {
				proceed = true;
				 
			}
			
			
			

		} else if (processDate) {
			proceed = true;
			System.out.println("** Processing date -- PD=" + processDate );
		} else {
			
			System.out.println("** IN Else -- PT=" + processDate );
		}

		if (proceed) {
			System.out.println("** Begin reading XLSB");
			now = LocalDateTime.now();
			// System.out.println("Begin SQL:" + dtf.format(now));
			logger.info(dtf.format(now) + ": " + "------------------ Begin reading XLSB file");

			strArr = readXlsbFile(xlsbFileName);
			now = LocalDateTime.now();
			// System.out.println("Begin SQL:" + dtf.format(now));
			logger.info(dtf.format(now) + ": " + "------------------ End reading XLSB file");
			
			
			
			if (processType.equals("id")) {
				rowMap = findIDinData(strArr, contractID);
			} else {
				logger.info(dtf.format(now) + ": " + "------------------ Begin processing date param");
				String dateParam_t = Olyutil.formatDate(dateParam,  "yyyy-MM-dd","yyyyMdd");
				
				rowMap = findDateinData(strArr, dateParam_t);		
			}
		}
/*
		now = LocalDateTime.now();
		// System.out.println("Begin SQL:" + dtf.format(now));
		logger.info(dtf.format(now) + ": " + "------------------ Begin reading XLSB file");

		strArr = readXlsbFile(xlsbFileName);
		now = LocalDateTime.now();
		// System.out.println("Begin SQL:" + dtf.format(now));
		logger.info(dtf.format(now) + ": " + "------------------ End reading XLSB file");

		// System.out.println("SZ=" + strArr.size());
		// Olyutil.printStrArray(strArr);
		 * 
		 * 
		 * */
		
	
		
		
		
		//rowMap = findIDinData(strArr, contractID);
		
		
		
		// Olyutil.printStrArray(tgtArr);

		// Map<String, String> treeMap1 = new TreeMap<>(dataMap);

		// for (String str : treeMap.keySet()) {
		// System.out.println(str);
		// }

		// System.out.println("*****************************************************************************************************");

		// for (Map.Entry<String, String> entry : dataMap.entrySet()) {
		// System.out.println("*** Key:" + entry.getKey()+" --> Value:"
		// +entry.getValue() + "--");
		// }
		// System.out.println("rmSZ=" + rowMap.size());
		// displayMapArr( rowMap) ;

		// System.out.println("tgtArrSZ:" + tgtArr.size());
		
		
		
		request.getSession().setAttribute("rowMap", rowMap);

		request.getSession().setAttribute(paramName, paramValue);
		now = LocalDateTime.now();
		// System.out.println("Begin SQL:" + dtf.format(now));
		logger.info(dtf.format(now) + ": " + "------------------ Forward to JSP: " + dispatchJSP);
		LogManager.getLogManager().reset();

		for (Handler h : logger.getHandlers()) {
			System.out.println("*** Closing loggers");
			h.close(); // must call h.close or a .LCK file will remain.
		}

		//request.getRequestDispatcher(dispatchJSP).forward(request, response);
		// System.out.println("** End:");
	}

	/*****************************************************************************************************************************************************/

	// Service method
	@Override
	protected void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {

	}

	/*****************************************************************************************************************************************************/
		
}

