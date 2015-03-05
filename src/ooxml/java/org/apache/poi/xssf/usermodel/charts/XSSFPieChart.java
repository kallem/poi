/**
 * Created by IntelliJ IDEA.
 * User: Kalle
 * Date: 20.11.2013
 * Time: 11:44
 * Copyright Surveypal Ltd 2013
 */

package org.apache.poi.xssf.usermodel.charts;

import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.IPieChart;
import org.apache.poi.util.Beta;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

@Beta
public class XSSFPieChart extends AXSSFChart implements IPieChart {
    public XSSFPieChart() {
    }

    public void initialize(final Chart chart) {
        super.initialize(chart);

        final CTPieChart pieChart = getPieChart();
        pieChart.addNewVaryColors().setVal(true);
        pieChart.addNewFirstSliceAng().setVal(0);

        final CTDLbls labels = pieChart.addNewDLbls();
        labels.addNewShowLegendKey().setVal(false);
        labels.addNewShowVal().setVal(false);
        labels.addNewShowCatName().setVal(false);
        labels.addNewShowSerName().setVal(false);
        labels.addNewShowPercent().setVal(false);
        labels.addNewShowBubbleSize().setVal(false);
        labels.addNewShowLeaderLines().setVal(true);
    }

    private CTPieChart getPieChart() {
        final CTPlotArea plotArea = chart.getCTChart().getPlotArea();
        return plotArea.getPieChartList().isEmpty() ? plotArea.addNewPieChart() : plotArea.getPieChartArray(0);
    }

    public void setVaryColors(final boolean varyColors) {
        final CTBoolean ctVaryColors = getPieChart().getVaryColors() != null ? getPieChart().getVaryColors() : getPieChart().addNewVaryColors();
        ctVaryColors.setVal(varyColors);
    }

    public void addSerie(final ChartDataSource<?> cat, final ChartDataSource<? extends Number> val) {
        final CTPieChart pieChart = getPieChart();

        final int serId = pieChart.sizeOfSerArray();

        final CTPieSer serie = pieChart.addNewSer();
        serie.addNewIdx().setVal(serId);
        serie.addNewOrder().setVal(serId);

        final CTDLbls labels = serie.addNewDLbls();
        final CTNumFmt ctNumFmt = labels.addNewNumFmt();
        ctNumFmt.setFormatCode("0%;;;");
        ctNumFmt.setSourceLinked(false);
        labels.addNewDLblPos().setVal(STDLblPos.CTR);
        labels.addNewShowLegendKey().setVal(false);
        labels.addNewShowVal().setVal(false);
        labels.addNewShowCatName().setVal(false);
        labels.addNewShowSerName().setVal(false);
        labels.addNewShowPercent().setVal(true);
        labels.addNewShowBubbleSize().setVal(true);
        labels.addNewShowLeaderLines().setVal(true);

        final CTAxDataSource catDataSource = serie.addNewCat();
        org.apache.poi.xssf.usermodel.charts.XSSFChartUtil.buildAxDataSource(catDataSource, cat);

        final CTNumDataSource numDataSource = serie.addNewVal();
        org.apache.poi.xssf.usermodel.charts.XSSFChartUtil.buildNumDataSource(numDataSource, val);
        numDataSource.getNumRef().getNumCache().setFormatCode("0.0%");
    }
}
