package chart;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import java.awt.geom.Rectangle2D;
import java.util.*;
import java.util.regex.Pattern;

public class CreateRadar {

    private XSLFSlide templiteSlide;
    private XSLFShape chartShape;
    private Rectangle2D templeteAnchor;

    //old diagrams parameters
    private List<CTValAx> ctValAxes;
    private List<CTCatAx> ctCatAxes;
    private CTTextBody legendPropertiesText;
    private CTShapeProperties legendPropertiesShape;
    private CTShapeProperties ctShapePropertiesPloatArea;
    private CTRadarSer ctRadarSerTempliteRefOne;
    private CTRadarSer ctRadarSerTempliteRefTwo;
    private CTRadarSer ctRadarSerTempliteOne;
    private CTRadarSer ctRadarSerTempliteTwo;

    public CreateRadar(XSLFSlide templiteSlide, XSLFShape chartShape){
        this.templiteSlide = templiteSlide;
        this.chartShape = chartShape;
        parseChartTemplite();
    }

    public List<DataChart> getDataRadar(Workbook workbook) {
        List<DataChart> dataRadar = new ArrayList<>();

        Sheet sheet = workbook.getSheetAt(0);
        Row rowKey = sheet.getRow(0);
        for(int i=1; i<sheet.getLastRowNum()+1; i++)
        {
            Map<String, String> allData = new HashMap<>();
            Map<String, String> chartData = new HashMap<>();
            String round = "";
            Row rowValue = sheet.getRow(i);
            for(int y=0; y<rowKey.getLastCellNum(); y++)
            {

                if(!rowKey.getCell(y).getStringCellValue().equals("system_id") & !rowKey.getCell(y).getStringCellValue().equals("_presentation_id"))
                {

                    String value = "";
                    switch(rowValue.getCell(y).getCellTypeEnum())
                    {
                        case NUMERIC:
                            if(Math.abs(rowValue.getCell(y).getNumericCellValue())>=10)
                            {
                                double formatValue = Math.round(rowValue.getCell(y).getNumericCellValue());
                                value = String.valueOf(formatValue);
                            }
                            else
                            {
                                double formatValue = Math.round(rowValue.getCell(y).getNumericCellValue()*100);
                                value = String.valueOf(formatValue/100);
                            }
                            break;
                        case  STRING:
                            value = rowValue.getCell(y).getStringCellValue();
                            break;
                    }
                    String keyWord = "";
                    if(rowKey.getCell(y).getStringCellValue().substring(1).replace("_", " ").equals("agreed upon timeline"))
                    {
                        keyWord = "agreed-upon timeline";
                    }
                    else
                    {
                        keyWord = rowKey.getCell(y).getStringCellValue().substring(1).replace("_", " ");
                    }

                    allData.put(keyWord, value);
                }
                Set<String> keys = allData.keySet();
                round = allData.get("csat round");

                for(String key : keys)
                {
                    if(!key.equals("csat round"))
                    {
                        chartData.put(key, allData.get(key));
                    }
                }
            }
            dataRadar.add(new DataChart(round, chartData));
        }

        return dataRadar;
    }
    private void parseChartTemplite() {
        List<XSLFChart> charts = new ArrayList<>();
         XSLFSheet sheet = chartShape.getSheet();
        XSLFChart chartGraf=null;
        for (POIXMLDocumentPart docPart : chartShape.getSheet().getRelations()) {

            if (docPart instanceof XSLFChart) {

               chartGraf = (XSLFChart) docPart;
                charts.add(chartGraf);
            }
        }

        //get old diagrams properties
        ctValAxes = charts.get(0).getCTChart().getPlotArea().getValAxList();
        ctCatAxes = charts.get(0).getCTChart().getPlotArea().getCatAxList();
        templeteAnchor = chartShape.getAnchor();
        legendPropertiesText = charts.get(0).getCTChart().getLegend().getTxPr();
        legendPropertiesShape = charts.get(0).getCTChart().getLegend().getSpPr();
        ctShapePropertiesPloatArea = charts.get(0).getCTChart().getPlotArea().getSpPr();
        ctRadarSerTempliteRefOne = charts.get(0).getCTChart().getPlotArea().getRadarChartArray(0).getSerArray(0);
        ctRadarSerTempliteRefTwo = charts.get(0).getCTChart().getPlotArea().getRadarChartArray(0).getSerArray(1);
        ctRadarSerTempliteOne = charts.get(0).getCTChart().getPlotArea().getRadarChartArray(0).getSerArray(2);
        ctRadarSerTempliteTwo = charts.get(0).getCTChart().getPlotArea().getRadarChartArray(0).getSerArray(3);

        //delete the old diagrams
        templiteSlide.getXmlObject().getCSld().getSpTree().getGraphicFrameList().clear();



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

    public void drawRadar(MyXSLFChartShape myXSLFChartShape, List<DataChart> dataRadar)
    {
        //create excel in the pptx file for edit data
        String[] mas = {"Adaptability", "Agreed-upon Timeline", "Appropriate Team", "Communication", "Comprehensive Capabilities", "Innovation", "Process", "Technical Excellence", "Value"};
        myXSLFChartShape.getMyXSLFChart().getWorkbook().getXSSFWorkbook();
        XSSFWorkbook workbook = myXSLFChartShape.getMyXSLFChart().getWorkbook().getXSSFWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);


        sheet.createRow(0);
        sheet.getRow(0).createCell(0).setCellValue("CSAT round");
       for(int i=0; i<mas.length; i++)
       {
           sheet.getRow(0).createCell(i+1).setCellValue(mas[i]);
       }

        for(int i=0; i<dataRadar.size(); i++)
        {
            sheet.createRow(i+1);
            sheet.getRow(i+1).createCell(0).setCellValue(dataRadar.get(i).getPeriod());
            for (int c = 1; c < mas.length+1; c++) {
                sheet.getRow(i+1).createCell(c).setCellValue(Double.valueOf(dataRadar.get(i).getData().get((mas[c-1]).toLowerCase())));
            }

        }

        //create Radar chart
        CTChartSpace ctChartSpace = myXSLFChartShape.getMyXSLFChart().getChartSpace();
        CTChart ctChart = ctChartSpace.addNewChart();
        ctChart.addNewPlotVisOnly().setVal(true);
        ctChart.addNewDispBlanksAs().setVal(STDispBlanksAs.GAP);
        ctChart.addNewShowDLblsOverMax().setVal(false);
        ctChart.addNewAutoTitleDeleted();

        CTPlotArea ctPlotArea = ctChart.addNewPlotArea();
        ctPlotArea.addNewLayout();
        ctPlotArea.setSpPr(ctShapePropertiesPloatArea);

        CTRadarChart ctRadarChart = ctPlotArea.addNewRadarChart();
        ctRadarChart.addNewVaryColors().setVal(false);
        ctRadarChart.addNewRadarStyle().setVal(STRadarStyle.MARKER);

        Character[] letters = new Character[26];
        int n=0;
        for(char q ='A'; q<='Z'; q++)
        {
            letters[n]=q;
            n++;
        }

        CTRadarSer[] arrayRadarSer = null;
        if(dataRadar.size()>2)
        {
            arrayRadarSer = new CTRadarSer[dataRadar.size() + 2];
            arrayRadarSer[0] = ctRadarSerTempliteRefOne;
            arrayRadarSer[1] = ctRadarSerTempliteRefTwo;
            arrayRadarSer[2] = ctRadarSerTempliteOne;
            arrayRadarSer[3] = ctRadarSerTempliteTwo;
            for(int i=0; i<dataRadar.size()-2; i++)
            {
               arrayRadarSer[3+(i+1)] = ctRadarSerTempliteTwo;
            }

        }
        else if(dataRadar.size()==1)
        {
            arrayRadarSer = new CTRadarSer[3];
            arrayRadarSer[0] = ctRadarSerTempliteRefOne;
            arrayRadarSer[1] = ctRadarSerTempliteRefTwo;
            arrayRadarSer[2] = ctRadarSerTempliteOne;
        }
        else
        {
            arrayRadarSer = new CTRadarSer[4];
            arrayRadarSer[0] = ctRadarSerTempliteRefOne;
            arrayRadarSer[1] = ctRadarSerTempliteRefTwo;
            arrayRadarSer[2] = ctRadarSerTempliteOne;
            arrayRadarSer[3] = ctRadarSerTempliteTwo;
        }

        //create RadarSer !REF
        ctRadarChart.setSerArray(arrayRadarSer);
        for(int i=0; i<2; i++)
        {
            CTRadarSer ctRadarSerOne = ctRadarChart.getSerArray()[i];
            ctRadarSerOne.getIdx().setVal(i);
            ctRadarSerOne.getOrder().setVal(i);
            CTStrRef ctStrRef_cat = ctRadarSerOne.getCat().getStrRef();
            ctStrRef_cat.setF("Sheet1!$B$1:$" + letters[dataRadar.get(0).getData().size()] + "$1");
            CTStrData ctStrData1_cat = ctStrRef_cat.getStrCache();
            ctStrData1_cat.getPtCount().setVal(dataRadar.get(0).getData().size());
            CTStrVal[] ctStrVal = ctStrData1_cat.getPtArray();
            for(int y=0; y<dataRadar.get(0).getData().size(); y++)
            {
                ctStrVal[y].setIdx(y);
                ctStrVal[y].setV(mas[y]);
            }
        }

        //create RadarSer with value
        for(int i=0; i<dataRadar.size(); i++)
        {
            CTRadarSer ctRadarSerTwo = ctRadarChart.getSerArray()[2+i];
            ctRadarSerTwo.getIdx().setVal(2+i);
            ctRadarSerTwo.getOrder().setVal(2+i);
            ctRadarSerTwo.getTx().getStrRef().setF("Sheet1!$A$" + String.valueOf(i+2));
            ctRadarSerTwo.getTx().getStrRef().getStrCache().getPtCount().setVal(1);
            ctRadarSerTwo.getTx().getStrRef().getStrCache().getPtList().get(0).setV(dataRadar.get(i).getPeriod());

            switch(i)
            {
                case 0:
                    ctRadarSerTwo.getSpPr().getLn().getSolidFill().getSchemeClr().setVal(STSchemeColorVal.ACCENT_2);
                    break;
                case 1:
                    ctRadarSerTwo.getSpPr().getLn().getSolidFill().getSchemeClr().setVal(STSchemeColorVal.ACCENT_3);
                    break;
                case 2:
                    ctRadarSerTwo.getSpPr().getLn().getSolidFill().getSchemeClr().setVal(STSchemeColorVal.ACCENT_6);
                    ctRadarSerTwo.getSpPr().getLn().getSolidFill().getSchemeClr().addNewLumOff().setVal(15000);
                    break;
                case 3:
                    ctRadarSerTwo.getSpPr().getLn().getSolidFill().getSchemeClr().setVal(STSchemeColorVal.ACCENT_6);
                    break;
            }

            ctRadarSerTwo.getCat().getStrRef().setF("Sheet1!$B$1:$" + letters[dataRadar.get(0).getData().size()] + "$1");
            CTStrRef ctStrRef_cat_Two = ctRadarSerTwo.getCat().getStrRef();
            ctStrRef_cat_Two.unsetStrCache();

            CTStrData ctStrData1_cat_Two = ctStrRef_cat_Two.addNewStrCache();
            ctStrData1_cat_Two.addNewPtCount().setVal(dataRadar.get(0).getData().size());

            CTNumRef ctNumRefSerTwo = ctRadarSerTwo.getVal().getNumRef();
            ctNumRefSerTwo.setF("Sheet1!$B$" + String.valueOf(i+2) + ":$" + letters[dataRadar.get(0).getData().size()] + "$" + String.valueOf(i+2));
            ctNumRefSerTwo.unsetNumCache();
            CTNumData ctNumData_val_Two = ctNumRefSerTwo.addNewNumCache();
            ctNumData_val_Two.setFormatCode("General");
            ctNumData_val_Two.addNewPtCount().setVal(dataRadar.get(0).getData().size());

            for(int y=0; y<dataRadar.get(0).getData().size(); y++)
            {
                CTStrVal ctStrValTwo = ctStrData1_cat_Two.addNewPt();
                ctStrValTwo.setIdx(y);
                ctStrValTwo.setV(mas[y]);

                CTNumVal ctNumValTwo= ctNumData_val_Two.addNewPt();
                ctNumValTwo.setIdx(y);
                ctNumValTwo.setV(dataRadar.get(i).getData().get((mas[y]).toLowerCase()));
            }

        }
        ctRadarChart.addNewAxId().setVal(490351680);
        ctRadarChart.addNewAxId().setVal(490346688);

        //create cat axis
        CTCatAx[] arrayCTCatAx = new CTCatAx[ctCatAxes.size()];
        ctCatAxes.toArray(arrayCTCatAx);
        ctPlotArea.setCatAxArray(arrayCTCatAx);


        //val axis
        CTValAx[] arrayCTValAx = new CTValAx[ctValAxes.size()];
        ctValAxes.toArray(arrayCTValAx);
        ctPlotArea.setValAxArray(arrayCTValAx);
        ctPlotArea.getValAxArray(0).getScaling().getMax().setVal(5);


        //legend
        CTLegend cTLegend = ctChart.addNewLegend();
        cTLegend.addNewLegendPos().setVal(STLegendPos.T);

        CTLegendEntry ctLegendEntryOne = cTLegend.addNewLegendEntry();
        ctLegendEntryOne.addNewIdx().setVal(0);
        ctLegendEntryOne.addNewDelete().setVal(true);

        CTLegendEntry ctLegendEntryTwo = cTLegend.addNewLegendEntry();
        ctLegendEntryTwo.addNewIdx().setVal(1);
        ctLegendEntryTwo.addNewDelete().setVal(true);

        cTLegend.addNewOverlay().setVal(false);
        cTLegend.setTxPr(legendPropertiesText);
        cTLegend.setSpPr(legendPropertiesShape);

    }

}
