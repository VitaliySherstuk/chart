package chart;

import org.apache.poi.util.Units;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrame;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrameNonVisual;
import java.awt.geom.Rectangle2D;

public class MyXSLFChartShape {

    private XSLFSlide newSlide;
    private Rectangle2D anchor;
    private MyXSLFChart myXSLFChart;
    private CTGraphicalObjectFrame graphicalObjectFrame;

    public MyXSLFChartShape(XSLFSlide newSlide, MyXSLFChart myXSLFChart, Rectangle2D anchor) throws XmlException {
        this.newSlide = newSlide;
        this.myXSLFChart = myXSLFChart;
        this.anchor = anchor;
        createGraficalObjectFrame();
        setAnchor(anchor);
    }

    private void createGraficalObjectFrame() throws XmlException {
        String rId = "rId" + (newSlide.getRelationParts().size()+1);
        newSlide.addRelation(rId, XSLFRelation.CHART, myXSLFChart);

        long cNvPrId = 1;
        String cNvPrName = "MyChart";
        int cNvPrNameCount = 1;

        for (CTGraphicalObjectFrame currGraphicalObjectFrame : newSlide.getXmlObject().getCSld().getSpTree().getGraphicFrameList()) {
            if (currGraphicalObjectFrame.getNvGraphicFramePr() != null) {
                if (currGraphicalObjectFrame.getNvGraphicFramePr().getCNvPr() != null) {
                    cNvPrId++;
                    if (currGraphicalObjectFrame.getNvGraphicFramePr().getCNvPr().getName().startsWith(cNvPrName)) {
                        cNvPrNameCount++;
                    }
                }
            }
        }

        graphicalObjectFrame = newSlide.getXmlObject().getCSld().getSpTree().addNewGraphicFrame();
        CTGraphicalObjectFrameNonVisual cTGraphicalObjectFrameNonVisual = graphicalObjectFrame.addNewNvGraphicFramePr();
        cTGraphicalObjectFrameNonVisual.addNewCNvGraphicFramePr();
        cTGraphicalObjectFrameNonVisual.addNewNvPr();

        CTNonVisualDrawingProps cTNonVisualDrawingProps = cTGraphicalObjectFrameNonVisual.addNewCNvPr();
        cTNonVisualDrawingProps.setId(cNvPrId);
        cTNonVisualDrawingProps.setName("MyChart" + cNvPrNameCount);

        CTGraphicalObject graphicalObject = graphicalObjectFrame.addNewGraphic();
        CTGraphicalObjectData graphicalObjectData = CTGraphicalObjectData.Factory.parse(
                "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
                        +"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                        +"r:id=\"" + rId + "\"/>"
        );
        graphicalObjectData.setUri("http://schemas.openxmlformats.org/drawingml/2006/chart");
        graphicalObject.setGraphicData(graphicalObjectData);

    }

    private void setAnchor(Rectangle2D anchor)
    {
        CTTransform2D xfrm = (graphicalObjectFrame.getXfrm() != null) ? graphicalObjectFrame.getXfrm() : graphicalObjectFrame.addNewXfrm();
        CTPoint2D off = xfrm.isSetOff() ? xfrm.getOff() : xfrm.addNewOff();
        long x = Units.toEMU(anchor.getX());
        long y = Units.toEMU(anchor.getY());
        off.setX(x);
        off.setY(y);
        CTPositiveSize2D ext = xfrm.isSetExt() ? xfrm.getExt() : xfrm.addNewExt();
        long cx = Units.toEMU(anchor.getWidth());
        long cy = Units.toEMU(anchor.getHeight());
        ext.setCx(cx);
        ext.setCy(cy);
    }

    public MyXSLFChart getMyXSLFChart()
    {
        return myXSLFChart;
    }
}
