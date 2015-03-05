/**
 * Created by IntelliJ IDEA.
 * User: Kalle
 * Date: 21.11.2013
 * Time: 15:23
 * Copyright Surveypal Ltd 2013
 */

package org.apache.poi.xssf.usermodel.charts;

import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.IBarChart;
import org.apache.poi.util.Beta;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

@Beta
public class XSSFBarChart extends AXSSFChart implements IBarChart {
    public XSSFBarChart() {
    }

    @Override
    public void initialize(final Chart chart) {
        super.initialize(chart);

        final CTBarChart barChart = getBarChart();
        barChart.addNewBarDir().setVal(STBarDir.COL);
        barChart.addNewGrouping().setVal(STBarGrouping.CLUSTERED);
        barChart.addNewVaryColors().setVal(false);
        barChart.addNewGapWidth().setVal(100);

        final CTDLbls labels = barChart.addNewDLbls();
        labels.addNewShowLegendKey().setVal(false);
        labels.addNewShowVal().setVal(false);
        labels.addNewShowCatName().setVal(false);
        labels.addNewShowSerName().setVal(false);
        labels.addNewShowPercent().setVal(false);
        labels.addNewShowBubbleSize().setVal(false);
    }

    private CTBarChart getBarChart() {
        final CTPlotArea plotArea = chart.getCTChart().getPlotArea();
        return plotArea.getBarChartList().isEmpty() ? plotArea.addNewBarChart() : plotArea.getBarChartArray(0);
    }

    public void setVaryColors(final boolean varyColors) {
        final CTBoolean ctVaryColors = getBarChart().getVaryColors() != null ? getBarChart().getVaryColors() : getBarChart().addNewVaryColors();
        ctVaryColors.setVal(varyColors);
    }

    public void addSerie(final ChartDataSource<?> cat, final ChartDataSource<? extends Number> val) {
        final CTBarChart barChart = getBarChart();

        final int serId = barChart.sizeOfSerArray();

        final CTBarSer serie = barChart.addNewSer();
        serie.addNewIdx().setVal(serId);
        serie.addNewOrder().setVal(serId);
        serie.addNewInvertIfNegative().setVal(true);

        final CTAxDataSource catDataSource = serie.addNewCat();
        org.apache.poi.xssf.usermodel.charts.XSSFChartUtil.buildAxDataSource(catDataSource, cat);

        final CTNumDataSource numDataSource = serie.addNewVal();
        org.apache.poi.xssf.usermodel.charts.XSSFChartUtil.buildNumDataSource(numDataSource, val);

        barChart.addNewAxId().setVal(0);
        barChart.addNewAxId().setVal(1);

        final CTPlotArea ctPlotArea = getPlotArea();

        final CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(1);
        final CTScaling ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctScaling.addNewMax().setVal(1);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        ctValAx.addNewMajorGridlines();
        final CTNumFmt ctNumFmt = ctValAx.addNewNumFmt();
        ctNumFmt.setFormatCode("0%");
        ctNumFmt.setSourceLinked(false);
        ctValAx.addNewMajorTickMark().setVal(STTickMark.NONE);
        ctValAx.addNewMinorTickMark().setVal(STTickMark.NONE);
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        ctValAx.addNewCrossAx().setVal(0);
        ctValAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO);
        ctValAx.addNewCrossBetween().setVal(STCrossBetween.BETWEEN);

        final CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(0);
        ctCatAx.addNewScaling().addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewMajorTickMark().setVal(STTickMark.NONE);
        ctCatAx.addNewMinorTickMark().setVal(STTickMark.NONE);
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
        ctCatAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO);
        ctCatAx.addNewCrossAx().setVal(1);
        ctCatAx.addNewAuto().setVal(true);
        ctCatAx.addNewLblAlgn().setVal(STLblAlgn.CTR);
        ctCatAx.addNewLblOffset().setVal(100);
        ctCatAx.addNewNoMultiLvlLbl().setVal(true);
    }
}
