package chart;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;

import java.io.*;
import java.util.List;

public class Main {

    public static void main(String[] args) throws Exception {


        String filenameExcel = args[1];
        String filePptx = args[0];
        String fileExcelForColumnDiagram = args[2];
        String fileExcelForRadardDiagram = args[3];

        /*String filenameExcel = "d:\\TestBot\\ChartPointss.xlsx";
        String filePptx = "d:\\TestBot\\template.pptx";
        String fileExcelForColumnDiagram = "d:\\TestBot\\csat_chart_bar.xlsx";
        String fileExcelForRadardDiagram = "d:\\TestBot\\csat_chart_radar.xlsx";*/

        XMLSlideShow slideShow = null;
        List<XSLFSlide> slideList = null;

            slideShow = new XMLSlideShow(new FileInputStream(filePptx));
            slideList = slideShow.getSlides();

            //Create Bar chart
            XSLFSlide slideTwo = slideList.get(1);
            XSLFShape chartShapeBar = null;
            List<XSLFShape> shapes = slideTwo.getShapes();
            for (XSLFShape shape : shapes) {

                if (shape.getShapeName().contains("Chart 8")) {
                    chartShapeBar = shape;
                }
            }

            CreateBar createBar = new CreateBar(slideTwo, chartShapeBar);

            List<DataChart> dataBarList = createBar.getDataChart(new XSSFWorkbook(fileExcelForColumnDiagram), "bar");
            List<DataChart> dataLineList = createBar.getDataChart(new XSSFWorkbook(fileExcelForColumnDiagram), "line");

            MyXSLFChart myXSLFChartBar = createBar.createChart();
            MyXSLFChartShape myXSLFChartBarShape = createBar.createChartShape(myXSLFChartBar);
            createBar.drawBar(myXSLFChartBarShape, dataBarList, dataLineList);


            //create Radar chart
            XSLFSlide slideFive = slideList.get(4);
            XSLFShape chartShapeRadar = null;
            List<XSLFShape> shapesRadar = slideFive.getShapes();
            for (XSLFShape sh : shapesRadar) {
                if (sh.getShapeName().contains("Chart")) {

                    chartShapeRadar = sh;
                }
            }
            CreateRadar createRadar = new CreateRadar(slideFive, chartShapeRadar);

            List<DataChart> dataRadar = createRadar.getDataRadar(new XSSFWorkbook(fileExcelForRadardDiagram));

            MyXSLFChart myXSLFChartRadar = createRadar.createChart();
            MyXSLFChartShape myXSLFChartRadarShape = createRadar.createChartShape(myXSLFChartRadar);

            createRadar.drawRadar(myXSLFChartRadarShape, dataRadar);

            //Chart points
            InputStream fs = new FileInputStream(filenameExcel);

            XSSFWorkbook book = new XSSFWorkbook(fs);
            XSSFSheet sheetOneH = book.getSheet("2x2 1H");
            XSSFSheet sheetTwoH = book.getSheet("2x2 2H");

            XSLFSlide slidePointsFirstRound = slideList.get(5);
            XSLFSlide slidePointsSecondRound = slideList.get(6);

            GrabChart grabChart = new GrabChart();
            CTChart ctChartFirstRound = grabChart.getChart(sheetOneH);
            CTChart ctChartSecondRound = grabChart.getChart(sheetTwoH);

            grabChart.setChart(slidePointsFirstRound, ctChartFirstRound);
            grabChart.setChart(slidePointsSecondRound, ctChartSecondRound);


            ChartLineNPS npsGoal = new ChartLineNPS();
            npsGoal.setLineGoal(slideList.get(1));
            try (FileOutputStream out = new FileOutputStream(filePptx)) {

                slideShow.write(out);
            }

    }
}
