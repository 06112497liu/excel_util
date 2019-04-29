package com.tellyes.platform.toolkit.freemarker;

import freemarker.core.Environment;
import freemarker.template.TemplateDirectiveBody;
import freemarker.template.TemplateDirectiveModel;
import freemarker.template.TemplateException;
import freemarker.template.TemplateModel;

import java.io.IOException;
import java.util.Map;

/**
 * freemarker excel导出模版自定义样式标签
 * @author xiehai
 * @date 2019/04/29 11:15
 * @Copyright(c) tellyes tech. inc. co.,ltd
 */
public class ExcelStyleDirective implements TemplateDirectiveModel {
    @Override
    public void execute(Environment env, Map params, TemplateModel[] loopVars, TemplateDirectiveBody body)
        throws TemplateException, IOException {
        body.render(env.getOut());
        // 闭合标签
        env.getOut().append(ExcelDirective.CLOSE_STYLES);
    }
}
