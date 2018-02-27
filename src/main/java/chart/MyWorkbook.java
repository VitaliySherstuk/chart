package chart;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;

public class MyWorkbook extends POIXMLDocumentPart {

    private XSSFWorkbook workbook;

    public MyWorkbook(PackagePart part) throws Exception {
        super(part);
        workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

    }

    public XSSFWorkbook getXSSFWorkbook() {
        return workbook;
    }

    @Override
    protected void commit() throws IOException {
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        workbook.write(out);
        workbook.close();
        out.close();
    }
}
