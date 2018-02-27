package chart;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class GrabChart {

    public CTChart getChart(XSSFSheet sheet) throws IOException, InvalidFormatException {


        List<PackagePart> parts = new ArrayList<>();
        for(PackagePart pp : sheet.getPackagePart().getPackage().getParts()) {

            if(pp.getPartName().getName().contains("charts"))
            {

                parts.add(pp);
            }
        }
        XSSFChart xssfChart = sheet.getDrawingPatriarch().getCharts().get(0);
        CTChart chart = xssfChart.getCTChart();

        return chart;
    }


    public void setChart(XSLFSlide slideToChart, CTChart chart)
    {
        XmlObject xO = chart.getPlotArea().copy();
        CTLegend legendTemp = chart.getLegend();
        CTBoolean plotVisible = chart.getPlotVisOnly();
        CTBoolean overMax = chart.getShowDLblsOverMax();

        XSLFChart chartPPt = null;
        CTChart ctChrtPPT = null;
        for (XSLFShape shape : slideToChart.getShapes()) {

            if (shape.getShapeName().contains("Chart")) {

                for (POIXMLDocumentPart docPart : shape.getSheet().getRelations()) {

                    if (docPart instanceof XSLFChart) {
                        chartPPt = (XSLFChart) docPart;

                        ctChrtPPT = chartPPt.getCTChart();
                        ctChrtPPT.getPlotArea().set(xO);
                        ctChrtPPT.setLegend(legendTemp);
                        ctChrtPPT.setPlotVisOnly(plotVisible);
                        ctChrtPPT.setShowDLblsOverMax(overMax);
                        ctChrtPPT.unsetLegend();
                        List<CTScatterSer> ctScatterSers = ctChrtPPT.getPlotArea().getScatterChartList().get(0).getSerList();
                        for(CTScatterSer ctScatterSer : ctScatterSers)
                        {

                            String fRef = ctScatterSer.getTx().getStrRef().getF().split("!")[1];
                            String cValX = ctScatterSer.getXVal().getNumRef().getF().split("!")[1];
                            String cValY = ctScatterSer.getYVal().getNumRef().getF().split("!")[1];

                            ctScatterSer.getTx().getStrRef().setF("Sheet1!" + fRef);
                            ctScatterSer.getXVal().getNumRef().setF("Sheet1!" + cValX);
                            ctScatterSer.getYVal().getNumRef().setF("Sheet1!" + cValY);
                        }

                    }
                }
            }
        }

    }

}
