/**
 * Created by IntelliJ IDEA.
 * User: Kalle
 * Date: 21.11.2013
 * Time: 15:26
 * Copyright Surveypal Ltd 2013
 */

package org.apache.poi.xssf.usermodel.charts;

import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.util.Beta;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTx;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

@Beta
public class AXSSFChart {
    protected XSSFChart chart;

    public void initialize(final Chart chart) {
        if (!(chart instanceof XSSFChart)) {
            throw new IllegalArgumentException("Chart must be instance of XSSFChart");
        }

        this.chart = (XSSFChart) chart;
    }

    protected CTPlotArea getPlotArea() {
        return chart.getCTChart().getPlotArea();
    }

    public void setTitle(final String title) {
        CTTitle ctTitle = chart.getCTChart().getTitle();
        if (ctTitle == null) {
            ctTitle = chart.getCTChart().addNewTitle();
            ctTitle.addNewLayout();
            ctTitle.addNewOverlay().setVal(false);
        }

        CTTx ctTx = ctTitle.getTx();
        if (ctTx == null) {
            ctTx = ctTitle.addNewTx();
        }

        CTTextBody ctTextBody = ctTx.getRich();
        if (ctTextBody == null) {
            ctTextBody = ctTx.addNewRich();
            ctTextBody.addNewBodyPr();
        }

        CTTextParagraph ctTextParagraph = ctTextBody.getPList().isEmpty() ? ctTextBody.addNewP() : ctTextBody.getPArray(0);

        CTRegularTextRun ctRegularTextRun = ctTextParagraph.getRList().isEmpty() ? ctTextParagraph.addNewR() : ctTextParagraph.getRArray(0);
        ctRegularTextRun.setT(title);
    }
}
