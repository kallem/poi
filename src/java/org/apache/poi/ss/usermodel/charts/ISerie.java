/**
 * Created by IntelliJ IDEA.
 * User: Kalle
 * Date: 20.11.2013
 * Time: 12:32
 * Copyright Surveypal Ltd 2013
 */

package org.apache.poi.ss.usermodel.charts;

import org.apache.poi.util.Beta;

@Beta
public interface ISerie {
    ChartDataSource<?> getCategories();

    ChartDataSource<? extends Number> getValues();
}
