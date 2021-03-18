package com.olympus.assetupload.excel;

import java.io.File;
import java.io.InputStream;
import java.text.MessageFormat;
import java.util.Iterator;

import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

public class MultiTagTest {
    public static void main(final String[] args) throws Exception {
		 String xlsbFileName = "C:\\Java_Dev\\props\\SAP_Upload\\SAP_Asset_Upload\\new asset upload cumulative through current.xlsb";
		 String xlsxFileName = "C:\\Java_Dev\\props\\SAP_Upload\\SAP_Asset_Upload\\masterList.xlsx";

    final File file = new File(xlsbFileName);

    try (OPCPackage xlsxPackage = OPCPackage.open(file, PackageAccess.READ)) {
        final XSSFReader reader = new XSSFReader(xlsxPackage);

        final Iterator<InputStream> iter = reader.getSheetsData();

        try (InputStream stream = iter.next()) {
        final SAXParserFactory saxParserFactory = SAXParserFactory.newInstance();
        saxParserFactory.setNamespaceAware(true);

        final XMLReader sheetParser = saxParserFactory.newSAXParser().getXMLReader();

        sheetParser.setContentHandler(new XSSFSheetXMLHandler(reader.getStylesTable(),
            new ReadOnlySharedStringsTable(xlsxPackage), new SheetContentsHandler() {
                @Override
                public void startRow(final int rowNum) {
                }

                @Override
                public void endRow(final int rowNum) {
                }

                @Override
                public void cell(final String cellReference, final String formattedValue,
                    final XSSFComment comment) {
                System.out.println(MessageFormat.format(
                    "XSSFSheetXMLHandler Cell - cellReference={0}, formattedValue={1}, comment={2}",
                    cellReference, formattedValue, comment));
                }
            }, true));

        sheetParser.parse(new InputSource(stream));
        }
    }

    try (Workbook workbook = WorkbookFactory.create(file, null, true)) {
        final Row row = workbook.getSheetAt(0).getRow(0);

        for (int col = row.getFirstCellNum(); col < row.getLastCellNum(); col++) {
        System.out.println(MessageFormat.format("WorkbookFactory Cell - {0}", row.getCell(col)));
        }
    }
    }
}