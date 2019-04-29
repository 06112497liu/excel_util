package com.tellyes.platform.toolkit.freemarker;

import freemarker.core.Environment;
import freemarker.template.TemplateBooleanModel;
import freemarker.template.TemplateDirectiveBody;
import freemarker.template.TemplateDirectiveModel;
import freemarker.template.TemplateException;
import freemarker.template.TemplateModel;

import java.io.IOException;
import java.util.Map;
import java.util.Optional;

/**
 * excel freemarker header自定义指令
 * @author xiehai
 * @date 2018/07/06 09:31
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
public class ExcelDirective implements TemplateDirectiveModel {
    /**
     * 自定义样式闭合标签
     */
    static final String CLOSE_STYLES = "</Styles>\r\n";
    /**
     * 扩展样式属性
     */
    private static final String EXPAND_STYLE = "expandStyle";
    private static final String HEADER = "<?xml version=\"1.0\"?>\r\n"
        + "<?mso-application progid=\"Excel.Sheet\"?>\r\n"
        + "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n"
        + "xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n"
        + "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n"
        + "xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n"
        + "xmlns:html=\"http://www.w3.org/TR/REC-html40\">\r\n"
        + "<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">\r\n" + "<Version>12.00</Version>\r\n"
        + "</DocumentProperties>\r\n" + "<ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\"> \r\n"
        + "<WindowHeight>11700</WindowHeight>\r\n" + "<WindowWidth>27735</WindowWidth>\r\n"
        + "<WindowTopX>600</WindowTopX>\r\n" + "<WindowTopY>570</WindowTopY>\r\n"
        + "<ProtectStructure>False</ProtectStructure>\r\n" + "<ProtectWindows>False</ProtectWindows>\r\n"
        + "</ExcelWorkbook>\r\n" + "<Styles>\r\n" + "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n"
        + "<Alignment ss:Vertical=\"Center\"/>\r\n" + "<Borders/>\r\n"
        + "<Font ss:FontName=\"宋体\" x:CharSet=\"134\" ss:Size=\"11\" ss:Color=\"#000000\"/>\r\n" + "<Interior/>\r\n"
        + "<NumberFormat/>\r\n" + "<Protection/>\r\n" + "</Style>\r\n" + "<Style ss:ID=\"s1\">\r\n"
        + "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\r\n" + "<Borders>\r\n"
        + "<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\r\n"
        + "<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\r\n"
        + "<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\r\n"
        + "<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\r\n" + "</Borders>\r\n"
        + "<Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Bold=\"1\"/>\r\n"
        + "<Interior ss:Color=\"#00CCFF\" ss:Pattern=\"Solid\"/>\r\n" + "</Style>\r\n" + "<Style ss:ID=\"s2\">\r\n"
        + "<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>\r\n" + "<Borders>\r\n"
        + "<Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\r\n"
        + "<Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\r\n"
        + "<Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\r\n"
        + "<Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/>\r\n" + "</Borders>\r\n"
        + "<Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\"/>\r\n" + "</Style>\r\n";

    @SuppressWarnings("rawtypes")
    @Override
    public void execute(Environment environment, Map map, TemplateModel[] templateModels,
                        TemplateDirectiveBody templateDirectiveBody) throws TemplateException, IOException {
        environment.getOut().append(HEADER);
        // 若非自定义样式 加上样式闭合标签 否则在@style标签中加上</Styles>闭合标签
        boolean expandStyle = Optional.ofNullable(map.get(EXPAND_STYLE))
            .map(TemplateBooleanModel.class::cast)
            .map(value -> {
                try {
                    return value.getAsBoolean();
                } catch (Exception e) {
                    return false;
                }
            }).orElse(false);
        if (!expandStyle) {
            environment.getOut().append(CLOSE_STYLES);
        }
        templateDirectiveBody.render(environment.getOut());
        environment.getOut().append("</Workbook>");
    }
}