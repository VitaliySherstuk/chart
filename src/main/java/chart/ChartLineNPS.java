package chart;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;

import java.util.List;

public class ChartLineNPS {

    public void setLineGoal(XSLFSlide slid2) {
        XSLFShape shape = slid2.getShapes().get(0);

        XSLFChart chartPPt = null;
        CTChart ctChrtPPT = null;

        POIXMLDocumentPart docPart = shape.getSheet().getRelations().get(2);

        chartPPt = (XSLFChart) docPart;
        ctChrtPPT = chartPPt.getCTChart();

        List<CTNumVal> ctNums = ctChrtPPT.getPlotArea().getLineChartArray(0).getSerList().get(0).getVal().getNumRef().getNumCache().getPtList();
        for (CTNumVal ctNumVal : ctNums) {
            ctNumVal.setV("70");
        }
    }
}
