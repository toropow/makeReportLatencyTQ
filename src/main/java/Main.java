import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.charts.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

import java.io.FileOutputStream;

/**
 * Created by Семья on 15.04.2017.
 */
public class Main {
    public static void main(String[] args) throws Exception {
        String path_wr="C:/analyze/my_test_romka.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet mySheet = wb.createSheet("hello");

        for (int j = 0; j < 100 ; j++) {


            XSSFRow row = mySheet.createRow(j);
            for (int i = 1; i < 5; i++) {
                XSSFCell cell = row.createCell(i);
                cell.setCellValue(j*3+i);
            }
        }

        Drawing drow = mySheet.createDrawingPatriarch();
        ClientAnchor anchor = drow.createAnchor(0,0,0,0,5,5,15,15);

        Chart chart = drow.createChart(anchor);
        ChartLegend legend = chart.getOrCreateLegend();
        legend.setPosition(LegendPosition.BOTTOM);
        LineChartData data = chart.getChartDataFactory().createLineChartData();
        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);

        ValueAxis valueAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        setValueAxisTitle((XSSFChart) chart,0,"title of bottom axis");
        setCatAxisTitle((XSSFChart) chart,0, "title of left axis");
        ChartDataSource<Number> xs = DataSources.fromNumericCellRange(mySheet,new CellRangeAddress(1,10,1,1));
        ChartDataSource<Number> ys = DataSources.fromNumericCellRange(mySheet,new CellRangeAddress(1,10,2,2));
       LineChartSeries chartSeries = data.addSeries(xs,ys);
       chartSeries.setTitle("Hello");



        chart.plot(data,bottomAxis,valueAxis);

        wb.write(new FileOutputStream(path_wr));
        wb.close();


    }

    public static void setCatAxisTitle(XSSFChart chart, int axisIdx, String title) {
        CTCatAx valAx = chart.getCTChart().getPlotArea().getCatAxArray(axisIdx);
        CTTitle ctTitle = valAx.addNewTitle();
        ctTitle.addNewLayout();
        ctTitle.addNewOverlay().setVal(false);
        CTTextBody rich = ctTitle.addNewTx().addNewRich();
        rich.addNewBodyPr();
        rich.addNewLstStyle();
        CTTextParagraph p = rich.addNewP();
        p.addNewPPr().addNewDefRPr();
        p.addNewR().setT(title);
        p.addNewEndParaRPr();
    }


    public static void setValueAxisTitle(XSSFChart chart, int axisIdx, String title) {
        CTValAx valAx = chart.getCTChart().getPlotArea().getValAxArray(axisIdx);
        CTTitle ctTitle = valAx.addNewTitle();
        ctTitle.addNewLayout();
        ctTitle.addNewOverlay().setVal(false);
        CTTextBody rich = ctTitle.addNewTx().addNewRich();
        rich.addNewBodyPr();
        rich.addNewLstStyle();
        CTTextParagraph p = rich.addNewP();
        p.addNewPPr().addNewDefRPr();
        p.addNewR().setT(title);
        p.addNewEndParaRPr();
    }
}
