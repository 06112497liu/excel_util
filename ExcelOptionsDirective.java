package com.tellyes.platform.toolkit.freemarker;

import freemarker.core.Environment;
import freemarker.template.TemplateDirectiveBody;
import freemarker.template.TemplateDirectiveModel;
import freemarker.template.TemplateModel;
import freemarker.template.TemplateModelException;
import freemarker.template.TemplateScalarModel;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.experimental.FieldDefaults;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.Arrays;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;

/**
 * excel freemarker 冻结选项指令
 * @author xiehai
 * @date 2018/07/06 09:50 @Copyright(c) tellyes tech. inc. co.,ltd
 */
public class ExcelOptionsDirective implements TemplateDirectiveModel {
    /**
     * 默认分隔数量
     */
    private static final String DEFAULT_SPLIT = "1";
    private static final Function<String, String> FOOTER = value ->
        "<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\r\n"
            + "<PageSetup>\r\n" + "<Header x:Margin=\"0.3\"/>\r\n" + "<Footer x:Margin=\"0.3\"/>\r\n"
            + "<PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/>\r\n" + "</PageSetup>\r\n"
            + "<Unsynced/>\r\n" + "<Print>\r\n" + "<ValidPrinterInfo/>\r\n" + "<PaperSizeIndex>9</PaperSizeIndex>\r\n"
            + "<HorizontalResolution>600</HorizontalResolution>\r\n"
            + "<VerticalResolution>600</VerticalResolution>\r\n" + "</Print>\r\n" + "<Selected/>\r\n"
            + "<FreezePanes/>\r\n" + "<FrozenNoSplit/>\r\n" + FooterTag.SPLIT_HORIZONTAL.format(value) + "\r\n"
            + FooterTag.TOP_ROW_BOTTOM_PANE.format(value) + "\r\n" + FooterTag.SPLIT_VERTICAL.format(value) + "\r\n"
            + FooterTag.LEFT_COLUMN_RIGHT_PANE.format(value) + "\r\n" + "<ActivePane>0</ActivePane>\r\n" + "<Panes>\r\n"
            + "<Pane>\r\n" + "<Number>3</Number>\r\n" + "</Pane>\r\n" + "<Pane>\r\n" + "<Number>1</Number>\r\n"
            + "</Pane>\r\n" + "<Pane>\r\n" + "<Number>2</Number>\r\n" + "<ActiveCol>1</ActiveCol>\r\n" + "</Pane>\r\n"
            + "<Pane>\r\n" + "<Number>0</Number>\r\n" + "</Pane>\r\n" + "</Panes>\r\n"
            + "<ProtectObjects>False</ProtectObjects>\r\n" + "<ProtectScenarios>False</ProtectScenarios>\r\n"
            + "</WorksheetOptions>\r\n";
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelOptionsDirective.class);

    @SuppressWarnings("rawtypes")
    @Override
    public void execute(Environment environment, Map map, TemplateModel[] templateModels,
                        TemplateDirectiveBody templateDirectiveBody) throws IOException {
        AtomicReference<String> footer = new AtomicReference<>(FOOTER.apply(DEFAULT_SPLIT));
        Arrays.stream(FooterTag.values())
            .forEach(item ->
                Optional.ofNullable(map.get(item.outerTag))
                    .map(value -> (TemplateScalarModel) value)
                    .map(value -> {
                        try {
                            return value.getAsString();
                        } catch (TemplateModelException e) {
                            LOGGER.error(e.getMessage(), e);
                        }

                        return null;
                    })
                    // 如果传入了参数 进行替换
                    .ifPresent(value -> footer.set(footer.get().replace(item.format(), item.format(value))))
            );

        environment.getOut().append(footer.get());
    }

    @Getter
    @FieldDefaults(makeFinal = true, level = AccessLevel.PRIVATE)
    @RequiredArgsConstructor
    enum FooterTag {
        /**
         * SplitHorizontal Tag
         */
        SPLIT_HORIZONTAL("<SplitHorizontal>%s</SplitHorizontal>", "row"),
        TOP_ROW_BOTTOM_PANE("<TopRowBottomPane>%s</TopRowBottomPane>", "row"),
        SPLIT_VERTICAL("<SplitVertical>%s</SplitVertical>", "column"),
        LEFT_COLUMN_RIGHT_PANE("<LeftColumnRightPane>%s</LeftColumnRightPane>", "column");

        String tag;
        String outerTag;

        String format() {
            // 默认均为1
            return this.format(DEFAULT_SPLIT);
        }

        String format(String value) {
            return String.format(this.tag, value);
        }
    }
}
