/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package org.apache.poi.xslf.usermodel;

import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.util.Beta;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.drawingml.x2006.chart.ChartSpaceDocument;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrame;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrameNonVisual;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * Represents a Chart in a .pptx presentation
 */
@Beta
public final class XSLFChartFrame extends XSLFGraphicFrame {
    public final static String CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    public final static String RELATIONSHIPS_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private PackagePart chartPart;

    static CTGraphicalObjectFrame prototype(int shapeId){
        final CTGraphicalObjectFrame ctGraphicalObjectFrame = CTGraphicalObjectFrame.Factory.newInstance();
        final CTGraphicalObjectFrameNonVisual ctGraphicalObjectFrameNonVisual = ctGraphicalObjectFrame.addNewNvGraphicFramePr();

        final CTNonVisualDrawingProps ctNonVisualDrawingProps = ctGraphicalObjectFrameNonVisual.addNewCNvPr();
        ctNonVisualDrawingProps.setName("Chart " + shapeId);
        ctNonVisualDrawingProps.setId(shapeId + 1);
        ctGraphicalObjectFrameNonVisual.addNewCNvGraphicFramePr().addNewGraphicFrameLocks().setNoGrp(true);
        ctGraphicalObjectFrameNonVisual.addNewNvPr();

        // Create container for the chart
        final CTGraphicalObjectData ctGraphicalObjectData = ctGraphicalObjectFrame.addNewGraphic().addNewGraphicData();
        ctGraphicalObjectData.setUri(CHART_URI);
        return ctGraphicalObjectFrame;
    }

    public XSLFChartFrame(CTGraphicalObjectFrame shape, XSLFSheet sheet) {
        super(shape, sheet);
    }

    public void setChart(final CTChartSpace chartSpace, final XSLFRelation dataRelation, final byte[] dataAY) throws IOException {

        final XMLSlideShow slideShow = getSheet().getSlideShow();

        // Create data part
        final PackagePart dataPart = slideShow.addPart(dataAY, dataRelation);

        // Create chart part
        chartPart = slideShow.addPart(XSLFRelation.CHART);
        // Create Excel part for the chart part
        final PackageRelationship dataRel = chartPart.addRelationship(dataPart.getPartName(), TargetMode.INTERNAL, XSLFRelation.EXCEL.getRelation());
        // Attach chart space to Excel part
        chartSpace.addNewExternalData().setId(dataRel.getId());

        // Write chart data into raw format to be red into chart part
        final ByteArrayOutputStream chartOut = new ByteArrayOutputStream();
        final ChartSpaceDocument doc = ChartSpaceDocument.Factory.newInstance();
        doc.setChartSpace(chartSpace);
        doc.save(chartOut);
        chartOut.close();

        // Load chart space into chart part
        try {
            final ByteArrayInputStream chartIn = new ByteArrayInputStream(chartOut.toByteArray());
            chartPart.load(chartIn);
            chartIn.close();
        } catch (final InvalidFormatException e) {
            throw new POIXMLException("Failed to load chart part from xml input", e);
        }

        // Create chart relationship
        final PackageRelationship chartRel = getSheet().getPackagePart().addRelationship(chartPart.getPartName(), TargetMode.INTERNAL, XSLFRelation.CHART.getRelation());

        // Append chart element
        setChartRelation(chartRel);
    }

    /**
     * The low level code to insert {@code <c:chart>} tag into
     * {@code<a:graphicData>}.
     * <p/>
     * Here is the schema (ECMA-376):
     * <pre>
     * {@code
     * <complexType name="CT_GraphicalObjectData">
     *   <sequence>
     *     <any minOccurs="0" maxOccurs="unbounded" processContents="strict"/>
     *   </sequence>
     *   <attribute name="uri" type="xsd:token"/>
     * </complexType>
     * }
     * </pre>
     */
    private void setChartRelation(final PackageRelationship chartRel) {
        final CTGraphicalObject ctGraphicalObject = getXmlObject().getGraphic();
        // Create the graphical object if necessary
        CTGraphicalObjectData ctGraphicalObjectData = ctGraphicalObject.getGraphicData();
        if (ctGraphicalObjectData ==  null) {
            ctGraphicalObjectData = ctGraphicalObject.addNewGraphicData();
            ctGraphicalObjectData.setUri(CHART_URI);
        }

        final Node graphicNode = ctGraphicalObject.getGraphicData().getDomNode();
        final Document document = graphicNode.getOwnerDocument();

        // Create chart node
        final Element chartNode = document.createElementNS(CHART_URI, "c:chart");
        chartNode.setAttributeNS(RELATIONSHIPS_URI, "id", chartRel.getId());
        graphicNode.appendChild(chartNode);
    }
}
