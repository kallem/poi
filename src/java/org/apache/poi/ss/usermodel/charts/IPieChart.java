/**
 * Created by IntelliJ IDEA.
 * User: Kalle
 * Date: 20.11.2013
 * Time: 11:41
 * Copyright Surveypal Ltd 2013
 */

package org.apache.poi.ss.usermodel.charts;

import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.util.Beta;

@Beta
public interface IPieChart {
    void initialize(Chart chart);

    void setTitle(String title);

    void setVaryColors(boolean varyColors);

    void addSerie(ChartDataSource<?> cat, ChartDataSource<? extends Number> val);
}
