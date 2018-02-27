package chart;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTExternalData;
import org.openxmlformats.schemas.drawingml.x2006.chart.ChartSpaceDocument;
import javax.xml.namespace.QName;
import java.io.IOException;
import java.io.OutputStream;
import java.util.regex.Pattern;
import static org.apache.poi.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

public class MyXSLFChart extends POIXMLDocumentPart {

    private CTChartSpace chartSpace;
    private MyWorkbook myWorkbook;
    private PackagePart myXlsxpart;

    public MyXSLFChart(PackagePart part) throws Exception {
        super(part);
        myXlsxpart = createOPCPackage(part);
        myWorkbook = new MyWorkbook(myXlsxpart);
        chartSpace = createChartSpace(myWorkbook);

    }

    public CTChartSpace getChartSpace() {
        return chartSpace;
    }

    private PackagePart createOPCPackage(PackagePart part) throws Exception {
        OPCPackage oPCPackage = part.getPackage();
        int chartCount = oPCPackage.getPartsByName(Pattern.compile("/ppt/embeddings/.*.xlsx")).size() + 1;
        PackagePartName partName = PackagingURIHelper.createPartName("/ppt/embeddings/Microsoft_Excel_Worksheet" + chartCount + ".xlsx");
        PackagePart xlsxpart = oPCPackage.createPart(partName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        return xlsxpart;
    }

    private CTChartSpace createChartSpace(MyWorkbook workbook)
    {
        String rId = "rId" + (this.getRelationParts().size()+1);
        MyRelation xSLFXSSFRelationPACKAGE = new MyRelation(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");

        this.addRelation(rId, xSLFXSSFRelationPACKAGE, workbook);

        chartSpace = ChartSpaceDocument.Factory.newInstance().addNewChartSpace();
        CTExternalData cTExternalData = chartSpace.addNewExternalData();
        cTExternalData.setId(rId);
        return chartSpace;
    }

    public PackagePart getMyXlsxpart()
    {
        return myXlsxpart;
    }

    public MyWorkbook getWorkbook() {
        return myWorkbook;
    }

    @Override
    protected void commit() throws IOException {
        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        xmlOptions.setSaveSyntheticDocumentElement(new QName(CTChartSpace.type.getName().getNamespaceURI(), "chartSpace", "c"));
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        chartSpace.save(out, xmlOptions);
        out.close();
    }
}
