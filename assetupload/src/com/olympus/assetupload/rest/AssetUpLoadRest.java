package com.olympus.assetupload.rest;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;
import java.util.logging.Handler;
import java.util.logging.Logger;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.json.JSONArray;
import org.xml.sax.SAXException;

import com.olympus.assetupload.excel.servlet.AssetUpload.colLetters;
import com.olympus.olyutil.Olyutil;
import com.olympus.olyutil.log.OlyLog;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

//import com.olympus.assetupload.excel.servlet.ExcelSheetHandler;

@WebServlet("/assetupws")
public class AssetUpLoadRest extends HttpServlet {
	
	private final Logger logger = Logger.getLogger(AssetUpLoadRest.class.getName()); // define logger

	public enum colLetters {
	    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, 
	    AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP
	}
	
	/*****************************************************************************************************************************************************/

	
	/*****************************************************************************************************************************************************/

	
	
	/*****************************************************************************************************************************************************/
	public static ArrayList<String> readXlsbFile(String xlsbFileName) throws IOException {
		ArrayList<String> tgtArr = new ArrayList<String>();
       
        XLSBSheetHandler  exlSheetHandler = new XLSBSheetHandler();       
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

				 
                
                /*-----------------------------------------------------------------------------------------------------------*/

            }
            
			// System.out.println("output text:"+sheetTexts);
		
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
		 
		 
		 if (pkg != null) {
			 pkg.close();
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
	public static List<Map<String, String>> parseData_HOLD(ArrayList<String> dataArr, ArrayList<String>  hArr) throws IOException {
		
		List<Map<String, String>> rMap = new ArrayList<Map<String, String>>();

		 

		int n = 0;
		int k = 0;
		char charAt0;
		char charAt1;
		String idxKey = "";
		String key = "";
		String colTag = "";
		int lineNo = 0;
		String hashKey = "";
		String colValue = "";
		for (String s : dataArr) { // add String to tgtArr
			if (s.contains("Order Type") || s.contains("sheet name")) {
				// System.out.println("^^^ Skip Header");
				n++;
				continue;
			}
			HashMap<String, String> colMap = new HashMap<String, String>();
			JsonObject obj = new JsonObject();
			String[] parts = s.split("\\^");
			
			int pSZ = parts.length;
			String[] parts2 = s.split("%");
			
			
			
				//System.out.println("^^^ Line=" + s);
				//System.out.println("^^^P=" + parts2[0] + "--");
				lineNo = new Scanner(parts2[0]).useDelimiter("\\D+").nextInt();
				
				if (n < 5) {
					System.out.println("^^^ LN=" + lineNo + "--");
				}
				for (int i = 1; i < pSZ; i++) {	// process split string
					int m = 0;
					for (colLetters tag : colLetters.values()) { // search for data in each column				
						colTag = tag.toString();
						//System.out.println("^^^ Tag=" + colTag + "--");
						charAt0 = parts[i].charAt(0);
						charAt1 = parts[i].charAt(1);
						if (!Character.isDigit(charAt1)) {
							key = Character.toString(charAt0) + Character.toString(charAt1);
							
						} else {
							key = Character.toString(charAt0);
						}
						hashKey = colTag + lineNo;
						if (! colTag.equals(key)) {
							colValue = "";
							colMap.put(hashKey, colValue);
						} else {
							String[] colData = s.split("%");
							colValue = colData[1];
							colMap.put(hashKey, colValue);
						}
						//System.out.println("^^^DNE  Tag=" + colTag + "-- LN=" + lineNo + "-- hashKey=" + hashKey + "--");
						 if (n < 5) { // DEBUG 
						System.out.println("^^^ LN=" + lineNo + "-- HDR/Val=" + hArr.get(m) + "->" + colValue +"-- HK=" + hashKey + "--");
						
						  } // End if DEBUG
						m++;
						
					} // end inner for loop ColEnum
	
					//System.out.println("^^^ p[" + i +"] =" + parts[i]   + "-- Char0=" + charAt0 + "-- Char1=" + charAt1 + "--");
					//System.out.println("^^^ p[" + i +"] =" + parts[i]   + "-- key=" + key + "-- indKey=" + idxKey + "-- pSZ=" + pSZ + "--");
				}	// end for processing split string	
				rMap.add(colMap);
			
			
			n++;
		} // end for loop
		
		return(rMap);
	}
	
	
	
	/*****************************************************************************************************************************************************/

	public static HashMap<String, String> initHashMapHeader(ArrayList<String>  hArr) {
		HashMap<String, String> hMap = new HashMap<String, String>();
		int i = 0;
		ArrayList<String> headerArr = new ArrayList<String>();
	    for (colLetters c : colLetters.values()) { 	
	    	hMap.put(c.toString(), hArr.get(i));  
	    	i++;
	    }
	    return hMap;
	}
	/*****************************************************************************************************************************************************/

	public static HashMap<String, String> initHashMap() {
		HashMap<String, String> dataMap = new HashMap<String, String>();

	    for (colLetters c : colLetters.values()) { 	
	    	dataMap.put(c.toString(), "");      
	    }
	    return dataMap;
	}
	/*****************************************************************************************************************************************************/

	public static List<Map<String, String>> parseData(ArrayList<String> dataArr, ArrayList<String> hArr)
			throws IOException {

		List<Map<String, String>> rMap = new ArrayList<Map<String, String>>();

		int n = 0;
		int k = 0;
		char charAt0;
		char charAt1;
		String idxKey = "";
		String key = "";
		String colTag = "";
		int lineNo = 0;
		String hashKey = "";
		String colValue = "";
		for (String s : dataArr) { // add String to tgtArr
			if (s.contains("Order Type") || s.contains("sheet name")) {
				// System.out.println("^^^ Skip Header");
				n++;
				continue;
			}
			HashMap<String, String> colMap = new HashMap<String, String>();
			colMap = initHashMap();
			
			String[] parts = s.split("\\^");

			int pSZ = parts.length;
			String[] parts2 = s.split("%");

			// System.out.println("^^^ Line=" + s);
			// System.out.println("^^^P=" + parts2[0] + "--");
			lineNo = new Scanner(parts2[0]).useDelimiter("\\D+").nextInt();

			//if (n < 7) {
				// System.out.println("^^^ LN=" + lineNo + "-- partsSZ=" + pSZ + "--");
			//}

			for (int i = 1; i < pSZ; i++) { // process split string
				int m = 0;
				int j = 1;
				String[] parts3 = parts[i].split("%");
				
				charAt0 = parts[i].charAt(0);
				charAt1 = parts[i].charAt(1);
				if (!Character.isDigit(charAt1)) {
					key = Character.toString(charAt0) + Character.toString(charAt1);

				} else {
					key = Character.toString(charAt0);
				}
				hashKey = key + lineNo;
			 /*
				if (n < 8) {
					System.out.println("^^^ Key =" + key + "-- LN=" + lineNo + "-- hashKey=" + hashKey + "-- parts=" + parts[i] + "-- parts3=" + parts3[1] + "--");
				}
				*/ 		
				

				colMap.put(key, parts3[1]);
				
				
				
				
				
				/*
				 * if (! colTag.equals(key)) { colValue = ""; colMap.put(hashKey, colValue); }
				 * else { String[] colData = s.split("%"); colValue = colData[1];
				 * colMap.put(hashKey, colValue); }
				 */

				/*
				 * if (n < 5) { System.out.println("^^^ Key (DNE)=" + key + "-- LN=" + lineNo +
				 * "-- hashKey=" + hashKey + "--"); }
				 * 
				 * //System.out.println("^^^DNE  Key=" + key + "-- LN=" + lineNo + "-- hashKey="
				 * + hashKey + "--"); }
				 */

				//n++;
			} // end inner for loop
			rMap.add(colMap);
			n++;
		} // end outer for loop
		return (rMap);
	}
	/*****************************************************************************************************************************************************/

	public static void displayMapArrSorted(List<Map<String, String>> rowMap) {
		int i = 0;
		for (Map<String, String> map : rowMap) {
			 
			 if (i > 3 && i < 6) {

				Map<String, String> treeMap = new TreeMap<>(map); // sort hash by key
				for (Map.Entry<String, String> entry : treeMap.entrySet()) {
					System.out.println("*** Key:" + entry.getKey() + " --> Value:" + entry.getValue() + "--");
				}
				System.out.println("*****************************************************************************************************");
			 }
			 i++;
		}    
	}
	/*****************************************************************************************************************************************************/
	public static void displayMapArr(List<Map<String, String>> rowMap) {
		int i = 0;
		for (Map<String, String> map : rowMap) {
			 
			 if (i > 3 && i < 6) {

				//Map<String, String> treeMap = new TreeMap<>(map); // sort hash by key
				for (Map.Entry<String, String> entry : map.entrySet()) {
					System.out.println("*** Key:" + entry.getKey() + " --> Value:" + entry.getValue() + "--");
				}
				System.out.println("*****************************************************************************************************");
			 }
			 i++;
		}    
	}
	
	/*****************************************************************************************************************************************************/

	public static void displayHashMap(HashMap<String, String> map) {

		// Map<String, String> treeMap = new TreeMap<>(map); // sort hash by key
		for (Map.Entry<String, String> entry : map.entrySet()) {

			System.out.println("*** Key:" + entry.getKey() + " --> Value:" + entry.getValue() + "--");
		}
		System.out.println("*****************************************************************************************************");
	}
	
	/*****************************************************************************************************************************************************/
	public static JsonArray  writeJSON( List<Map<String, String>> dataMap, HashMap<String, String> hdrMap) throws IOException {
		JsonArray jsonArr = new JsonArray();
		System.out.println("**** dataMapSZ=:" + dataMap.size() + "--");
		String hdr = "";
		int i = 0;
		int k = 0;
		for (Map<String, String> map : dataMap) {
			JsonObject obj = new JsonObject();
			 //if (i > 3 && i < 6) {

				//Map<String, String> treeMap = new TreeMap<>(map); // sort hash by key
				for (Map.Entry<String, String> entry : map.entrySet()) {
					//System.out.println("*** Key:" + entry.getKey() + " --> Value:" + entry.getValue() + "--");
					
					if (hdrMap.containsKey(entry.getKey())) {
						hdr = hdrMap.get(entry.getKey());
						obj.addProperty(hdr, entry.getValue());
					}
					
					//System.out.println("*** Key:" + hdr + " --> Value:" + entry.getValue() + "--");
					hdr = "";
				}
				jsonArr.add(obj);
				//System.out.println("*****************************************************************************************************");
			 //} // end if DEBUG
			 i++;
			 k++;
		}    
		return(jsonArr);
	}
	/*****************************************************************************************************************************************************/
	/***********************************************************************************************************************************/
	public static void writeToFile(List<Map<String, String>> dataMap, HashMap<String, String> hdrMap, String fileName, PrintWriter out) throws IOException {
	    BufferedWriter writer = new BufferedWriter(new FileWriter(fileName));
	     
	    String hdr = "";
	    writer.write("[ "); 
	    out.print("[");
	    for (Map<String, String> map : dataMap) { // iterating map
	    	 
	    	int i = 0;
	    	writer.write("{" + "\n"); 
	    	out.print("{" + "\n");
	    	for (Map.Entry<String, String> entry : map.entrySet())  {
	    		 
	    		if (hdrMap.containsKey(entry.getKey())) {
					hdr = hdrMap.get(entry.getKey());
					writer.write(hdr + ": \t"   +  "\"" + entry.getValue()   +  "\"" );
					out.print(hdr + ": \t"   +  "\"" + entry.getValue()   +  "\"" );
					//obj.addProperty(hdr, entry.getValue());
				}
				
	    		
	    		
				//System.out.println("*** Key:" + hdr + " --> Value:" + entry.getValue() + "--");
				hdr = "";
				 
				if (i < map.entrySet().size() -1 ) {
					writer.write(",\n");
					out.print(",\n");
				}
				
			}
	    	writer.write("\n },");
	    	out.print("\n },");
	    	
	    	writer.newLine();
	    	out.print("\n");
	   
	    	i++;		
	    	//writer.newLine();
	    }
	    writer.write("]"); 
	    out.print("]");
	    writer.close(); 
	   
	}
	/***********************************************************************************************************************************/

	/***********************************************************************************************************************************/

	// Service method
	@Override
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		String headerFile = "C:\\Java_Dev\\props\\headers\\assetUploadWSHdr.txt";
		 String xlsbFileName = "C:\\Java_Dev\\props\\SAP_Upload\\SAP_Asset_Upload\\new asset upload cumulative through current.xlsb";
		
		//String xlsbFileName = "C:\\Java_Dev\\props\\SAP_Upload\\SAP_Asset_Upload\\AU.xlsb";
		String jsonFile2 = "C:\\Java_Dev\\JSON\\assetupload2.json";
		
		String Jsonfile = "C:\\Java_Dev\\JSON\\assetupload.json";
		PrintWriter out = response.getWriter();
		HashMap<String, String> dataMap = new HashMap<String, String>();
		HashMap<String, String> hdrMap = new HashMap<String, String>();
		List<Map<String, String>> rowMap = new ArrayList<Map<String, String>>();
		ArrayList<String> strArr = new ArrayList<String>();
		JsonArray jsonArr = new JsonArray();
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss.SSS");
		LocalDateTime now = null;
		String logFileName = "AssetUploadWS.log";
		String directoryName = "D:/Kettle/logfiles/AssetUploadWS";
		Handler fileHandler = OlyLog.setAppendLog(directoryName, logFileName, logger);
		
		ArrayList<String> headerArr = new ArrayList<String>();
		
		
		headerArr = Olyutil.readInputFile(headerFile);
		System.out.println("** HDRSZ=" + headerArr.size() + "--");
		hdrMap = initHashMapHeader(headerArr);
		//displayHashMap(hdrMap);
		//Olyutil.printStrArray(headerArr);
		//System.out.println("** Begin reading XLSB");
		now = LocalDateTime.now();
		// System.out.println("Begin reading XLSB file:" + dtf.format(now));
		logger.info(dtf.format(now) + ": " + "------------------ Begin reading XLSB file");

		strArr = readXlsbFile(xlsbFileName);
		
		now = LocalDateTime.now();
		// System.out.println("End reading XLSB file:" + dtf.format(now));
		logger.info(dtf.format(now) + ": " + "------------------ End reading XLSB file");
		
		System.out.println("*** arrSZ=" + strArr.size() + "--");
		
		
		rowMap = parseData(strArr, headerArr);
		//displayMapArr(rowMap);
		//jsonArr = writeJSON(rowMap, hdrMap);
		now = LocalDateTime.now();
		logger.info(dtf.format(now) + ": " + "------------------ End parse data -- Begin jsonWriter"); 
		
		writeToFile(rowMap, hdrMap, jsonFile2, out);
		
	    //Olyutil.jsonWriter(jsonArr, Jsonfile);
		
		now = LocalDateTime.now();
		logger.info(dtf.format(now) + ": " + "------------------ End jsonWriter -- Begin jsonWriterReaponse"); 
		//Olyutil.jsonWriterResponse(jsonArr, out);
		now = LocalDateTime.now();
		//logger.info(dtf.format(now) + ": " + "------------------ End jsonWriterReaponse"); 
		
		logger.info(dtf.format(now) + ": " + "------------------ Close fileHandle"); 
		fileHandler.close();
	}

}
