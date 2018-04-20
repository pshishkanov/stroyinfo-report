package io.pshishkanov;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.*;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {

    private static final boolean debug = false;

//    private static final String PATH_TO_FILE_FROM = "resources/2018q1/from-matherial.xlsx";
    private static final String PATH_TO_FILE_FROM = "resources/2018q1/from-machinery.xlsx";

    private static final String PATH_TO_FILE_TO = "resources/2018q1/to.xls";

//    private static final Integer OFFSET_FILE_FROM = 3;
//    private static final String PATTERN = "\\d{3}-\\d{4}";

    private static final Integer OFFSET_FILE_FROM = 2;
    private static final String PATTERN = "\\d{5,6}";

    private static final Integer OFFSET_FILE_TO = 4;

    private static Map<String, String> map = new HashMap<>();

    private static ReportHandler handler;

    public void processAllSheets(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename);
        XSSFReader r = new XSSFReader( pkg );
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst);

        Iterator<InputStream> sheets = r.getSheetsData();
        while(sheets.hasNext()) {
            InputStream sheet = sheets.next();
            System.out.println("Processing sheet");
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
    }

    public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        handler = new ReportHandler(Pattern.compile(PATTERN), OFFSET_FILE_FROM, sst);
        parser.setContentHandler(handler);
        return parser;
    }

    public static void main(String[] args) throws Exception {
        Main example = new Main();
        example.processAllSheets(PATH_TO_FILE_FROM);

        map = handler.getMap();

        FileInputStream file = new FileInputStream(new File(PATH_TO_FILE_TO));
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);

        System.out.println("Start search...");
        for (Row row : sheet) {
            Cell keyCell = row.getCell(0);
            if (keyCell != null && Cell.CELL_TYPE_STRING == keyCell.getCellType()) {
                String key = keyCell.getStringCellValue();
                Matcher m = Pattern.compile(PATTERN).matcher(key);
                if (m.matches()) {
                    String value = map.get(key.replaceFirst("^0+(?!$)", ""));
                    System.out.println(String.join(" - ", key.replaceFirst("^0+(?!$)", ""), value));

                    if (!debug) {
                        Cell valueCell = row.getCell(OFFSET_FILE_TO);
                        valueCell.setCellValue(stringToFloat(value));
                    }
                }
            }
        }
        System.out.println("Finish search...");

        file.close();

        if (!debug) {
            FileOutputStream out = new FileOutputStream(new File(PATH_TO_FILE_TO));
            workbook.write(out);
            out.close();
        }
    }

    private static float stringToFloat(String value) {
        float result = 0;
        try {
            result = Float.parseFloat(value);
        } catch (Exception e) {
            System.out.println("Error during convert string to float.");
            e.printStackTrace(System.out);
        }
        return result;
    }

}
