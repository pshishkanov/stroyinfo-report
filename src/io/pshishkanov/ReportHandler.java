package io.pshishkanov;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by pshishkanov on 12/06/16.
 */

public class ReportHandler extends DefaultHandler {
    private SharedStringsTable sst;
    private String lastContents;
    private boolean nextIsString;
    private Integer step = 0;
    private String key;
    private Boolean found = false;

    private Pattern pattern;
    private Integer offset;
    private Map<String, String> map = new HashMap<>();

    public ReportHandler(Pattern pattern, Integer offset, SharedStringsTable sst) {
        this.pattern = pattern;
        this.offset = offset;
        this.sst = sst;
    }

    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {
        if(name.equals("c")) {
            String cellType = attributes.getValue("t");
            nextIsString = cellType != null && cellType.equals("s");
        }
        lastContents = "";
    }

    public void endElement(String uri, String localName, String name)
            throws SAXException {
        if(nextIsString) {
            int idx = Integer.parseInt(lastContents);
            lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            nextIsString = false;
        }

        if(name.equals("v")) {
            Matcher m = pattern.matcher(lastContents);
            if (m.matches()) {
                key = lastContents;
                step = 0;
                found = false;
            } else if (step.equals(offset)) {
                map.put(key, lastContents);
                step = 0;
                found = true;
            } else {
                if (!found)
                    step ++;
            }
        }
    }

    public void characters(char[] ch, int start, int length)
            throws SAXException {
        lastContents += new String(ch, start, length);
    }

    public Map<String, String> getMap() {
        return map;
    }
}