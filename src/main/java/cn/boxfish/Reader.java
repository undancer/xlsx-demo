package cn.boxfish;

import com.google.common.collect.Table;
import com.google.common.collect.TreeBasedTable;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.ArrayList;
import java.util.Collection;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * Created by undancer on 2016/10/22.
 */
public class Reader implements XSSFSheetXMLHandler.SheetContentsHandler {


    public static class Sheet {

        private Integer id;
        private String name;
        private String relId;

        private Sheet(Integer id, String name, String relId) {
            this.id = id;
            this.name = name;
            this.relId = relId;
        }

        public Integer getId() {
            return id;
        }

        public String getName() {
            return name;
        }

        public String getRelId() {
            return relId;
        }
    }

    private OPCPackage container;

    private XSSFReader reader;
    private StylesTable styles;

    private ReadOnlySharedStringsTable strings;
    private WorkbooksTable workbooks;

    private Table<Integer, Integer, String> table;

    DataFormatter dataFormatter = new DataFormatter();

    public Reader(String file) throws OpenXML4JException, IOException, SAXException {
        this(OPCPackage.open(file, PackageAccess.READ));
    }

    public Reader(OPCPackage container) throws IOException, OpenXML4JException, SAXException {
        this.container = container;
        reader = new XSSFReader(container);
        styles = reader.getStylesTable();
        workbooks = new WorkbooksTable(container);
        strings = new ReadOnlySharedStringsTable(container);
    }

    public Collection<Sheet> getSheets() {
        return workbooks.getSheets();
    }

    public Table<Integer, Integer, String> getSheet(String relId) {
        InputStream is = null;

        try {
            is = this.reader.getSheet(relId);
            table = TreeBasedTable.create();
            XMLReader xmlReader = SAXHelper.newXMLReader();
            InputSource source = new InputSource(is);
            XSSFSheetXMLHandler handler = new XSSFSheetXMLHandler(styles, null, strings, this, dataFormatter, false);
            xmlReader.setContentHandler(handler);
            xmlReader.parse(source);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(is);
        }
        return table;
    }

    public void startRow(int rowNum) {

    }

    public void endRow(int rowNum) {

    }

    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        CellReference ref = new CellReference(cellReference);
        table.put(ref.getRow(), (int) ref.getCol(), formattedValue);
    }

    public void headerFooter(String text, boolean isHeader, String tagName) {

    }

    static class WorkbooksTable extends DefaultHandler {

        private Collection<Sheet> sheets = new ArrayList<>();

        public WorkbooksTable(OPCPackage pkg) throws IOException {
            ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.WORKBOOK.getContentType());
            if (parts.size() > 0) {
                PackagePart sstPart = parts.get(0);
                readFrom(sstPart.getInputStream());
            }
        }

        public void readFrom(InputStream is) throws IOException {
            PushbackInputStream pis = new PushbackInputStream(is, 1);
            int emptyTest = pis.read();
            if (emptyTest > -1) {
                pis.unread(emptyTest);
                InputSource sheetSource = new InputSource(pis);
                try {
                    XMLReader sheetParser = SAXHelper.newXMLReader();
                    sheetParser.setContentHandler(this);
                    sheetParser.parse(sheetSource);
                } catch (ParserConfigurationException e) {
                    throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
                } catch (SAXException e) {
                    e.printStackTrace();
                }
            }
        }

        public Collection<Sheet> getSheets() {
            return sheets;
        }

        public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {

            if (uri != null && !uri.equals(NS_SPREADSHEETML)) {
                return;
            }

            if ("sheet".equals(localName)) {
                Integer sheetId = Integer.valueOf(attributes.getValue("sheetId"));
                String sheetName = attributes.getValue("name");
                String relId = attributes.getValue("r:id");
                Sheet sheet = new Sheet(sheetId, sheetName, relId);
                sheets.add(sheet);
            }

        }
    }

}
