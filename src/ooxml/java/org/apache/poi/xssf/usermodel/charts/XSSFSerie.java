/**
 * Created by IntelliJ IDEA.
 * User: Kalle
 * Date: 20.11.2013
 * Time: 12:30
 * Copyright Surveypal Ltd 2013
 */

package org.apache.poi.xssf.usermodel.charts;

import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ISerie;
import org.apache.poi.util.Beta;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

@Beta
public class XSSFSerie implements ISerie {
    private int id;
    private int order;
    private ChartDataSource<?> categorySource;
    private ChartDataSource<? extends Number> valueSource;

    public XSSFSerie(final int id, final int order, final ChartDataSource<?> categorySource, final ChartDataSource<? extends Number> valueSource) {
        this.id = id;
        this.order = order;
        this.categorySource = categorySource;
        this.valueSource = valueSource;
    }

    /**
     * Returns data source used for categories.
     *
     * @return data source used for categories
     */
    public ChartDataSource<?> getCategories() {
        return categorySource;
    }

    /**
     * Returns data source used for values.
     *
     * @return data source used for values
     */
    public ChartDataSource<? extends Number> getValues() {
        return valueSource;
    }

    public void addToChart(final CTPieChart pieChart) {
        final CTPieSer serie = pieChart.addNewSer();
        serie.addNewIdx().setVal(this.id);
        serie.addNewOrder().setVal(this.order);

        final CTAxDataSource catDataSource = serie.addNewCat();
        buildCatDataSource(catDataSource);

        final CTNumDataSource numDataSource = serie.addNewVal();
        buildValDataSource(numDataSource);
    }

    /**
     * Builds CTAxDataSource object content from POI ChartDataSource.
     */
    private void buildCatDataSource(final CTAxDataSource dataSource) {
        if (categorySource.isNumeric()) {
            if (categorySource.isReference()) {
                buildNumRef(dataSource.addNewNumRef(), categorySource);
            } else {
                buildNumLit(dataSource.addNewNumLit(), categorySource);
            }
        } else {
            if (categorySource.isReference()) {
                buildStrRef(dataSource.addNewStrRef(), categorySource);
            } else {
                buildStrLit(dataSource.addNewStrLit(), categorySource);
            }
        }
    }

    /**
     * Builds CTNumDataSource object content from POI ChartDataSource
     */
    private void buildValDataSource(final CTNumDataSource valDataSource) {
        if (valueSource.isReference()) {
            buildNumRef(valDataSource.addNewNumRef(), valueSource);
        } else {
            buildNumLit(valDataSource.addNewNumLit(), valueSource);
        }
    }

    private void buildNumRef(CTNumRef ctNumRef, ChartDataSource<?> dataSource) {
        ctNumRef.setF(dataSource.getFormulaString());
        CTNumData cache = ctNumRef.addNewNumCache();
        fillNumCache(cache, dataSource);
    }

    private void buildNumLit(CTNumData ctNumData, ChartDataSource<?> dataSource) {
        fillNumCache(ctNumData, dataSource);
    }

    private void buildStrRef(CTStrRef ctStrRef, ChartDataSource<?> dataSource) {
        ctStrRef.setF(dataSource.getFormulaString());
        CTStrData cache = ctStrRef.addNewStrCache();
        fillStringCache(cache, dataSource);
    }

    private void buildStrLit(CTStrData ctStrData, ChartDataSource<?> dataSource) {
        fillStringCache(ctStrData, dataSource);
    }

    private void fillStringCache(CTStrData cache, ChartDataSource<?> dataSource) {
        int numOfPoints = dataSource.getPointCount();
        cache.addNewPtCount().setVal(numOfPoints);
        for (int i = 0; i < numOfPoints; ++i) {
            Object value = dataSource.getPointAt(i);
            if (value != null) {
                CTStrVal ctStrVal = cache.addNewPt();
                ctStrVal.setIdx(i);
                ctStrVal.setV(value.toString());
            }
        }
    }

    private void fillNumCache(CTNumData cache, ChartDataSource<?> dataSource) {
        int numOfPoints = dataSource.getPointCount();
        cache.addNewPtCount().setVal(numOfPoints);
        for (int i = 0; i < numOfPoints; ++i) {
            Number value = (Number) dataSource.getPointAt(i);
            if (value != null) {
                CTNumVal ctNumVal = cache.addNewPt();
                ctNumVal.setIdx(i);
                ctNumVal.setV(value.toString());
            }
        }
    }
}
