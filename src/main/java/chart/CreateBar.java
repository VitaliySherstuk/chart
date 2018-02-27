package chart;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrame;
import java.awt.geom.Rectangle2D;
import java.util.*;
import java.util.regex.Pattern;

public class CreateBar {

    private XSLFSlide templiteSlide;
    private XSLFShape chartShape;
    private Rectangle2D templeteAnchor;
    private CTShapeProperties ctShapePropertiesPloatArea;
    private CTGapAmount ctGapAmount;
    private CTOverlap ctOverlap;
    private List<CTUnsignedInt> ctUnsignedInt;
    private CTBarSer[] ctBarSers;
    private List<CTCatAx> ctCatAxes;
    private List<CTValAx> ctValAxes;
    private CTLegend ctLegendOld;
    private CTPositiveSize2D ctPositiveSize2DExt;
    private CTPoint2D ctPoint2D;
    private List<CTLineChart> ctLineChartList;
    private CTDLbls ctdLbl;
    private CTDLbls ctdLblsT;


    public CreateBar(XSLFSlide templiteSlide, XSLFShape chartShape) throws InvalidFormatException {
        this.templiteSlide = templiteSlide;
        this.chartShape = chartShape;
        parseChartTemplite();
    }

    public List<DataChart> getDataChart(Workbook workbook, String diagramKind)
    {
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = sheet.iterator();

        Row rowKey = sheet.getRow(0);

        List<DataChart> dataChartList = new ArrayList<>();

        for(int i=1; i<sheet.getLastRowNum()+1; i++)
        {
            Map<String, String> mapAllData = new HashMap<>();
            for(int y=1; y<rowKey.getLastCellNum()-1; y++)
            {
                String value = "";
               switch(sheet.getRow(i).getCell(y).getCellTypeEnum())
                {
                    case NUMERIC:

                        if(Math.abs(sheet.getRow(i).getCell(y).getNumericCellValue())>=10)
                        {
                            double formatValue = Math.round(sheet.getRow(i).getCell(y).getNumericCellValue());
                            value = String.valueOf(formatValue);
                        }
                        else
                        {
                            double formatValue = Math.round(sheet.getRow(i).getCell(y).getNumericCellValue() * 100);
                            value = String.valueOf(formatValue/100);
                        }

                        break;
                    case STRING:
                        value = String.valueOf(sheet.getRow(i).getCell(y).getStringCellValue());
                        break;
                            //TODO: thrown exception for other format;
                }

                mapAllData.put(rowKey.getCell(y).getStringCellValue().substring(1).replace("_", " "),
                       value);
            }

            Set<String> keys = mapAllData.keySet();
            String round = mapAllData.get("csat round");

            Map<String, String> chartData = new HashMap<>();

            switch(diagramKind)
            {
                case "bar":
                    for(String key : keys)
                    {
                        if(!key.equals("csat round") & !key.equals("nps") & !key.equals("response rate"))
                        {
                            chartData.put(key, mapAllData.get(key));

                        }
                    }
                    break;
                case "line":
                    for(String key : keys)
                    {
                        if(key.equals("nps") | key.equals("response rate"))
                        {
                            chartData.put(key, mapAllData.get(key));
                        }
                    }
                    break;
            }
            dataChartList.add(new DataChart(round, chartData));
        }
            return  dataChartList;
    }


    private void parseChartTemplite() throws InvalidFormatException {
        List<XSLFChart> charts = new ArrayList<>();
        XSLFSheet sheet = chartShape.getSheet();
        XSLFChart chartGraf=null;
        for (POIXMLDocumentPart docPart : chartShape.getSheet().getRelations()) {

            if (docPart instanceof XSLFChart) {

                chartGraf = (XSLFChart) docPart;
                if(chartGraf.getCTChart().getPlotArea().getBarChartList().size()!=0)
                {
                    charts.add(chartGraf);
                }

            }
        }

        CTPlotArea ctPlotAreaOld = charts.get(0).getCTChart().getPlotArea();
        templeteAnchor = chartShape.getAnchor();
        ctShapePropertiesPloatArea = ctPlotAreaOld.getSpPr();
        ctdLblsT = ctPlotAreaOld.getBarChartList().get(0).getSerArray()[0].getDLbls();
        //data for bar
        ctGapAmount = ctPlotAreaOld.getBarChartList().get(0).getGapWidth();
        ctOverlap = ctPlotAreaOld.getBarChartList().get(0).getOverlap();
        ctdLbl = ctPlotAreaOld.getBarChartList().get(0).getDLbls();
        ctUnsignedInt = ctPlotAreaOld.getBarChartList().get(0).getAxIdList();
        ctBarSers = ctPlotAreaOld.getBarChartList().get(0).getSerArray();
        ctCatAxes = ctPlotAreaOld.getCatAxList();
        ctValAxes = ctPlotAreaOld.getValAxList();
        ctLegendOld = charts.get(0).getCTChart().getLegend();

        //data for Line
        ctLineChartList = ctPlotAreaOld.getLineChartList();

        //delete the old diagrams
        List<CTGraphicalObjectFrame> ctGraphicalObjectFrame =
                templiteSlide.getXmlObject().getCSld().getSpTree().getGraphicFrameList();

        for(int i=0; i<ctGraphicalObjectFrame.size(); i++)
        {
            CTGraphicalObjectFrame frame = ctGraphicalObjectFrame.get(i);
            if(frame.getNvGraphicFramePr().getCNvPr().getName().contains("Chart 8"))
            {
                ctPositiveSize2DExt = ctGraphicalObjectFrame.get(i).getXfrm().getExt();
                ctPoint2D = ctGraphicalObjectFrame.get(i).getXfrm().getOff();
                ctGraphicalObjectFrame.remove(i);
            }
        }

    }

    public MyXSLFChart createChart() throws Exception {

        OPCPackage oPCPackage = templiteSlide.getSlideShow().getPackage();
        int chartCount = oPCPackage.getPartsByName(Pattern.compile("/ppt/charts/chart.*")).size() + 1;
        PackagePartName partName = PackagingURIHelper.createPartName("/ppt/charts/chart" + chartCount + ".xml");
        PackagePart part = oPCPackage.createPart(partName, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");

        MyXSLFChart myXSLFChart = new MyXSLFChart(part);

        return myXSLFChart;
    }

    public MyXSLFChartShape createChartShape(MyXSLFChart myXSLFChart) throws XmlException {

        MyXSLFChartShape myXSLFChartShape = new MyXSLFChartShape(templiteSlide, myXSLFChart, templeteAnchor);
        return myXSLFChartShape;
    }

    public void drawBar(MyXSLFChartShape myXSLFChartShape, List<DataChart> listDataBar, List<DataChart> listDataLine)
    {
        //create excel in the pptx file for edit data
        XSSFWorkbook workbook = myXSLFChartShape.getMyXSLFChart().getWorkbook().getXSSFWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        String[] nameColumn = {"Detractors", "Passives", "Promoters", "NPS", "Response Rate"};
        sheet.createRow(0);
        for(int i=0; i<nameColumn.length; i++)
        {
            sheet.getRow(0).createCell(i+1).setCellValue(nameColumn[i]);
        }
        for(int i=0; i<listDataBar.size(); i++)
        {
            DataChart dataChartBar = listDataBar.get(i);
            Map<String, String> valuesBar = dataChartBar.getData();

            DataChart dataChartLine = listDataLine.get(i);
            Map<String, String> valuesLine = dataChartLine.getData();

            XSSFCell cellCatalog = sheet.createRow(i+1).createCell(0);
            cellCatalog.setCellValue(dataChartBar.getPeriod());

            XSSFCellStyle stylePersent = workbook.createCellStyle();
            DataFormat format = workbook.createDataFormat();
            stylePersent.setDataFormat(format.getFormat("0%"));

            //values for Bar
            for(int y = 0; y< dataChartBar.getData().size(); y++)
            {
                String parameter = sheet.getRow(0).getCell(y+1).getStringCellValue();
                XSSFCell cellValue = sheet.getRow(i+1).createCell(y+1);
                cellValue.setCellStyle(stylePersent);
                cellValue.setCellValue(Double.valueOf(valuesBar.get(parameter.toLowerCase())));
            }

            //values for Line
            for(int z = 0; z< dataChartLine.getData().size(); z++)
            {
                String parameter = sheet.getRow(0).getCell(dataChartBar.getData().size()+(z+1)).getStringCellValue();
                XSSFCell cellValueLine = sheet.getRow(i+1).createCell(dataChartBar.getData().size()+(z+1));
                if(parameter.equals("Response Rate"))
                {
                    cellValueLine.setCellStyle(stylePersent);
                }

                cellValueLine.setCellValue(Double.valueOf(valuesLine.get(parameter.toLowerCase())));
            }
        }

        //create Column Bar diagrams
        CTChartSpace ctChartSpace = myXSLFChartShape.getMyXSLFChart().getChartSpace();
        CTChart ctChart = ctChartSpace.addNewChart();
        ctChart.addNewPlotVisOnly().setVal(true);
        ctChart.addNewDispBlanksAs().setVal(STDispBlanksAs.GAP);
        ctChart.addNewShowDLblsOverMax().setVal(false);
        ctChart.addNewAutoTitleDeleted();

        CTPlotArea ctPlotArea = ctChart.addNewPlotArea();
        ctPlotArea.addNewLayout();
        ctPlotArea.addNewSpPr();
        ctPlotArea.setSpPr(ctShapePropertiesPloatArea);

        CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
        ctBarChart.addNewVaryColors().setVal(false);
        ctBarChart.addNewBarDir().setVal(STBarDir.COL);
        ctBarChart.addNewGrouping().setVal(STBarGrouping.PERCENT_STACKED);
        ctBarChart.addNewGapWidth().setVal(ctGapAmount.getVal());
        ctBarChart.addNewOverlap().setVal(ctOverlap.getVal());
        ctBarChart.setDLbls(ctdLbl);

        for(CTUnsignedInt axId : ctUnsignedInt)
        {
            ctBarChart.addNewAxId().setVal(axId.getVal());
        }

        //create Bar
        int amountSer = listDataBar.get(0).getData().size();
        int amountPeriod = listDataBar.size();

            ctBarChart.setSerArray(ctBarSers);

            CTBarSer[] ctBarSers = ctBarChart.getSerArray();

            Character[] letters = new Character[26];
            int n=0;
            for(char q ='A'; q<='Z'; q++)
            {
                letters[n]=q;
                n++;
            }

            for(int i=0; i<amountSer; i++)
            {
                CTBarSer ctBarSer = ctBarSers[i];
                ctBarSer.unsetDLbls();
                ctBarSer.setDLbls(ctdLblsT);

                ctBarSer.getDLbls().addNewDLblPos().setVal(STDLblPos.CTR);
                CTStrRef ctStrRef = ctBarSer.getCat().getStrRef();
                String fStr = null;
                if(amountPeriod==1)
                {
                    fStr = "Sheet1!$A$2";
                }
                else
                {
                    fStr = "Sheet1!$A$2:$A$" + String.valueOf(amountPeriod+1);
                }
                ctStrRef.setF(fStr);
                ctStrRef.unsetStrCache();
                CTStrData ctStrData = ctStrRef.addNewStrCache();
                ctStrData.addNewPtCount().setVal(listDataBar.size());
                CTNumRef ctNumRef = ctBarSer.getVal().getNumRef();

                ctNumRef.unsetNumCache();
                CTNumData ctNumData = ctNumRef.addNewNumCache();
                ctNumData.setFormatCode("0%");
                ctNumData.addNewPtCount().setVal(listDataBar.size());
                ctNumRef.setF("Sheet1!$" + letters[i+1] +"$2:$" + letters[i+1] +"$" + String.valueOf(amountPeriod+1));
                for(int y=0; y<amountPeriod; y++)
                {
                    CTStrVal ctStrVal = ctStrData.addNewPt();
                    ctStrVal.setIdx(y);
                    ctStrVal.setV(listDataBar.get(y).getPeriod());
                    CTNumVal ctNumVal = ctNumData.addNewPt();
                    ctNumVal.setIdx(y);


                    ctNumVal.setV(listDataBar.get(y).getData()
                            .get((ctBarSer.getTx().getStrRef().getStrCache().getPtList().get(0).getV()).toLowerCase()));

                }
            }

        //axis
        CTCatAx[] ctCatAxesArray = new CTCatAx[ctCatAxes.size()];
        ctCatAxes.toArray(ctCatAxesArray);
        ctPlotArea.setCatAxArray(ctCatAxesArray);
        CTValAx[] ctValAxesArray = new CTValAx[ctValAxes.size()];
        ctValAxes.toArray(ctValAxesArray);
        ctPlotArea.setValAxArray(ctValAxesArray);

        //create Line
        CTLineChart[] ctLineChartsArray = new CTLineChart[ctLineChartList.size()];
        ctLineChartList.toArray(ctLineChartsArray);
        ctPlotArea.setLineChartArray(ctLineChartsArray);
        CTLineChart[] ctLineCharts = ctPlotArea.getLineChartArray();
        for(int i=0; i<ctLineCharts.length; i++)
        {
            CTLineChart ctLineChart = ctPlotArea.getLineChartArray()[i];
            CTSerTx ctSerTx = ctLineChart.getSerList().get(0).getTx();


            CTStrRef ctStrRef = ctLineChart.getSerList().get(0).getCat().getStrRef();
            String fStrLine = null;
            if(amountPeriod==1)
            {
                fStrLine = "Sheet1!$A$2";
            }
            else
            {
                fStrLine = "Sheet1!$A$2:$A$" + String.valueOf(amountPeriod+1);
            }
            ctStrRef.setF(fStrLine);
            ctStrRef.unsetStrCache();
            CTStrData ctStrData = ctStrRef.addNewStrCache();
            ctStrData.addNewPtCount().setVal(listDataLine.size());

            CTNumRef ctNumRef = ctLineChart.getSerList().get(0).getVal().getNumRef();
            if(ctSerTx.getStrRef().getStrCache().getPtList().get(0).getV().equals("NPS"))
            {
                ctNumRef.setF("Sheet1!$E$2:$E$" + String.valueOf(amountPeriod+1));
            }
            else if(ctSerTx.getStrRef().getStrCache().getPtList().get(0).getV().equals("Response Rate"))
            {
                ctNumRef.setF("Sheet1!$F$2:$F$" + String.valueOf(amountPeriod+1));
            }
            else
            {
                ctNumRef.setF("Sheet1!$" + letters[i+4] +"$2:$" + letters[i+4] +"$" + String.valueOf(amountPeriod+1));
            }
            ctNumRef.unsetNumCache();
            CTNumData ctNumData = ctNumRef.addNewNumCache();
            if(ctSerTx.getStrRef().getStrCache().getPtList().get(0).getV().equals("Response Rate"))
            {
                ctNumData.setFormatCode("0%");
            }
            ctNumData.addNewPtCount().setVal(listDataLine.size());

            for(int y=0; y<listDataLine.size(); y++)
            {
                CTStrVal ctStrVal = ctStrData.addNewPt();
                ctStrVal.setIdx(y);
                ctStrVal.setV(listDataLine.get(y).getPeriod());

                CTNumVal ctNumVal = ctNumData.addNewPt();
                ctNumVal.setIdx(y);
                ctNumVal.setV(listDataLine.get(y).getData().get((ctSerTx.getStrRef().getStrCache().getPtList().get(0).getV()).toLowerCase()));
            }

        }

        //Legend
        ctChart.setLegend(ctLegendOld);

    }

    public Rectangle2D getCoordinate(){
        return new Rectangle2D.Double(
                Units.toPoints(ctPoint2D.getX()), Units.toPoints(ctPoint2D.getY()),
                Units.toPoints(ctPositiveSize2DExt.getCx()), Units.toPoints(ctPositiveSize2DExt.getCy()));
    }
}


