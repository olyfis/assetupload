package com.olympus.assetupload.excel;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.xml.sax.SAXException;
 
public class ReadXlsbFile {

	
	/***********************************************************************************************************************/
	
	static void callXLToList(String xlsbFileName){
        OPCPackage pkg;
        try {
            pkg = OPCPackage.open(xlsbFileName);
            XSSFBReader r = new XSSFBReader(pkg);
            XSSFBSharedStringsTable sst = new XSSFBSharedStringsTable(pkg);
            XSSFBStylesTable xssfbStylesTable = r.getXSSFBStylesTable();
            XSSFBReader.SheetIterator it = (XSSFBReader.SheetIterator) r.getSheetsData();

            List<XLSB2Lists> workBookAsList = new ArrayList<>();
            int sheetNr = 1;
            while (it.hasNext()) {
                InputStream is = it.next();
                String name = it.getSheetName();

                System.out.println("Begin parsing sheet "+sheetNr+": "+name);

                XLSB2Lists testSheetHandler = new XLSB2Lists();
                testSheetHandler.startSheet(name);
                XSSFBSheetHandler sheetHandler = new XSSFBSheetHandler(is,
                        xssfbStylesTable,
                        it.getXSSFBSheetComments(),
                        sst, testSheetHandler,
                        new DataFormatter(),
                        false);
                sheetHandler.parse();
                testSheetHandler.endSheet();

                System.out.println("End parsing sheet "+sheetNr+": "+name);
                sheetNr++;

                // Add parsed sheet to workbook list
                workBookAsList.add(testSheetHandler);
            }

            // For every sheet in Workbook
            System.out.println("\nShort Report:");
            for(XLSB2Lists sheet:workBookAsList){
                // sheet content
                System.out.println("Size of content: " +sheet.getSheetContentAsList().size());
                // sheet comment
                System.out.println("Size fo comment: "+sheet.getSheetCommentAsList().size());
                // sheet extra info
                System.out.println("Extra info.: "+sheet.getMapOfInfo().toString());                
            }

        } catch (InvalidFormatException e) {
            // TODO Please do your catch hier
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Please do your catch hier
            e.printStackTrace();
        } catch (OpenXML4JException e) {
            // TODO Please do your catch hier
            e.printStackTrace();
        } catch (SAXException e) {
            // TODO Please do your catch hier
            e.printStackTrace();
        }
	}
	
	
	/***********************************************************************************************************************/

	
	
	/***********************************************************************************************************************/

	
	public static void main(String[] args) throws OpenXML4JException {
		OPCPackage pkg;
	    String xlsbFileName = "C:\\Java_Dev\\props\\SAP_Upload\\SAP_Asset_Upload\\sap_AU.xlsb";
	    pkg = OPCPackage.open(xlsbFileName);
       
	    
	    
	    callXLToList(xlsbFileName);  
	    
	    
	    
	    
	    
	    
	    
	    
	    

    } // End main
} // End Clas
 
