package com.aici.compatibility;

import java.util.Map;
import java.util.List;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map.Entry;
import java.io.OutputStream;
import java.util.stream.Collectors;
import java.io.ByteArrayOutputStream;

import javax.imageio.ImageIO;

import lombok.extern.slf4j.Slf4j;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFTextbox;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.record.TextObjectRecord;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;

import com.aici.compatibility.dto.CompatibilityMachineDTO;
import com.aici.compatibility.dto.CompatibilityCommandStatisticItem;

/**
 * Excel生成
 *
 * @author <a href="mailto:472546172@qq.com">张天赐</a>
 */
@Slf4j
public class Excel {
    private Excel() {}

    private static final String FONT_MICROSOFT_YAHEI = "Microsoft YaHei";

    /**
     * 生成Excel
     *
     * @param data   数据项
     * @param stream 结果流
     */
    public static void generate(
        List<CompatibilityCommandStatisticItem> data, OutputStream stream)
        throws IOException {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            HSSFSheet sheet0 = workbook.createSheet("首页");
            generateHome(sheet0);
            HSSFSheet sheet1 = workbook.createSheet("兼容测试综述");
            generateOverview(sheet1, data);
            HSSFSheet sheet2 = workbook.createSheet("测试用例");
            generateCase(sheet2, data);
            HSSFSheet sheet3 = workbook.createSheet("Android终端结果汇总");
            generateAndroid(sheet3, data);
            HSSFSheet sheet4 = workbook.createSheet("iOS终端结果汇总");
            generateI(sheet4, data);
            HSSFSheet sheet5 = workbook.createSheet("HarmonyOS终端结果汇总");
            generateHarmony(sheet5, data);
            HSSFSheet sheet6 = workbook.createSheet("性能测试统计");
            generatePerformanceStatistics(sheet6, data);
            HSSFSheet sheet7 = workbook.createSheet("问题汇总");
            generateProblemSummary(sheet7, data);
            workbook.setActiveSheet(7);
            workbook.write(stream);
        }
    }

    /**
     * 生成首页
     */
    private static void generateHome(HSSFSheet sheet) throws IOException {
        sheet.setDefaultRowHeight((short)405);
        sheet.setDefaultColumnWidth((short)7);
        HSSFPatriarch drawingPatriarch = sheet.createDrawingPatriarch();
        // 画背景
        try (InputStream resource = Excel.class.getResourceAsStream("/background.png")) {
            if (resource == null) {return;}
            try (ByteArrayOutputStream memoryStream = new ByteArrayOutputStream()) {
                ImageIO.write(ImageIO.read(resource), "PNG", memoryStream);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)0, 0, (short)20, 26);
                drawingPatriarch.createPicture(anchor, sheet.getWorkbook().addPicture(memoryStream.toByteArray(), Workbook.PICTURE_TYPE_PNG));
            }
        }
        // 画LOGO
        try (InputStream resource = Excel.class.getResourceAsStream("/logo.png")) {
            if (resource == null) {return;}
            try (ByteArrayOutputStream memoryStream = new ByteArrayOutputStream()) {
                ImageIO.write(ImageIO.read(resource), "PNG", memoryStream);
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, (short)7, 6, (short)14, 12);
                drawingPatriarch.createPicture(anchor, sheet.getWorkbook().addPicture(memoryStream.toByteArray(), Workbook.PICTURE_TYPE_PNG));
            }
        }
        // 文本框 - title
        HSSFClientAnchor anchorTitle = new HSSFClientAnchor(0, 0, 0, 0, (short)4, 12, (short)15, 15);
        HSSFTextbox titleBox = drawingPatriarch.createTextbox(anchorTitle);
        HSSFRichTextString richTextTitle = new HSSFRichTextString("[ 随需而用的互联网测试平台 ]");
        HSSFFont fontTile = sheet.getWorkbook().createFont();
        fontTile.setBold(false);
        fontTile.setFontName(FONT_MICROSOFT_YAHEI);
        fontTile.setFontHeightInPoints((short)20);
        richTextTitle.applyFont(fontTile);
        titleBox.setString(richTextTitle);
        titleBox.setVerticalAlignment(TextObjectRecord.VERTICAL_TEXT_ALIGNMENT_CENTER);
        titleBox.setHorizontalAlignment(TextObjectRecord.HORIZONTAL_TEXT_ALIGNMENT_CENTERED);
        titleBox.setNoFill(true);
        titleBox.setLineStyle(HSSFShape.LINESTYLE_DEFAULT);
        titleBox.setLineWidth(0);
        // 文本框  - 关于
        HSSFClientAnchor anchorAbout = new HSSFClientAnchor(0, 0, 0, 0, (short)3, 16, (short)16, 20);
        HSSFTextbox aboutBox = drawingPatriarch.createTextbox(anchorAbout);
        HSSFRichTextString richTextAbout = new HSSFRichTextString("网址：https://www.fairytest.com\n 客服电话：400-180-0381    客服QQ：3411384093");
        HSSFFont fontAbout = sheet.getWorkbook().createFont();
        fontAbout.setBold(false);
        fontAbout.setFontName(FONT_MICROSOFT_YAHEI);
        fontAbout.setFontHeightInPoints((short)16);
        richTextAbout.applyFont(fontAbout);
        aboutBox.setString(richTextAbout);
        aboutBox.setVerticalAlignment(TextObjectRecord.VERTICAL_TEXT_ALIGNMENT_CENTER);
        aboutBox.setHorizontalAlignment(TextObjectRecord.HORIZONTAL_TEXT_ALIGNMENT_CENTERED);
        aboutBox.setNoFill(true);
        aboutBox.setLineStyle(HSSFShape.LINESTYLE_DEFAULT);
        aboutBox.setLineWidth(0);
    }

    /**
     * 生成概览
     */
    private static void generateOverview(HSSFSheet sheet, List<CompatibilityCommandStatisticItem> data) {
        // TODO 概览信息
        log.debug("概览信息:{},{}", sheet, data);
    }

    /**
     * 生成测试用例
     */
    private static void generateCase(HSSFSheet sheet, List<CompatibilityCommandStatisticItem> data) {
        // TODO 测试用例
        log.debug("测试用例:{},{}", sheet, data);
    }

    /**
     * 生成Android
     */
    private static void generateAndroid(HSSFSheet sheet, List<CompatibilityCommandStatisticItem> data) {
        List<CompatibilityCommandStatisticItem> filteredData = data.stream()
            .filter(t -> "ANDROID".equalsIgnoreCase(t.getMachine().getOsType())).collect(Collectors.toList());
        generateGroupByOS(sheet, filteredData);
    }

    /**
     * 生成iOS
     */
    private static void generateI(HSSFSheet sheet, List<CompatibilityCommandStatisticItem> data) {
        List<CompatibilityCommandStatisticItem> filteredData = data.stream()
            .filter(t -> "IOS".equalsIgnoreCase(t.getMachine().getOsType())).collect(Collectors.toList());
        generateGroupByOS(sheet, filteredData);
    }

    /**
     * 生成HarmonyOS
     */
    private static void generateHarmony(HSSFSheet sheet, List<CompatibilityCommandStatisticItem> data) {
        List<CompatibilityCommandStatisticItem> filteredData = data.stream()
            .filter(t -> "HARMONY".equalsIgnoreCase(t.getMachine().getOsType())).collect(Collectors.toList());
        generateGroupByOS(sheet, filteredData);
    }

    private static void generateGroupByOS(HSSFSheet sheet, List<CompatibilityCommandStatisticItem> data) {
        // 字体大小是10
        sheet.setDefaultRowHeightInPoints(10);
        // 正文字体
        HSSFFont font = sheet.getWorkbook().createFont();
        font.setFontName(FONT_MICROSOFT_YAHEI);
        // 标题字体
        HSSFFont titleFont = sheet.getWorkbook().createFont();
        titleFont.setBold(true);
        titleFont.setFontName(FONT_MICROSOFT_YAHEI);
        titleFont.setFontHeightInPoints((short)11);
        titleFont.setColor(HSSFColorPredefined.WHITE.getIndex());
        // 第一行
        HSSFRow row1 = sheet.createRow(0);
        for (int i = 0; i < 11; i++) {
            HSSFCell cell = row1.createCell(i);
            HSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyle.setFont(titleFont);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(HSSFColorPredefined.ROYAL_BLUE.getIndex());
            cell.setCellStyle(cellStyle);
        }
        row1.getCell(0).setCellValue("机型信息");
        row1.getCell(6).setCellValue("测试结果");
        // 第二行
        HSSFRow row2 = sheet.createRow(1);
        sheet.setColumnWidth(10, 255 * 35);
        for (int i = 0; i < 6; i++) {
            sheet.setColumnWidth(i, 255 * 15);
            HSSFFont titleFont2 = sheet.getWorkbook().createFont();
            titleFont2.setBold(true);
            titleFont2.setFontName(FONT_MICROSOFT_YAHEI);
            titleFont2.setFontHeightInPoints((short)11);
            titleFont2.setColor(HSSFColorPredefined.BLACK.getIndex());
            HSSFCell cell = row2.createCell(i);
            HSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
            cellStyle.setFont(titleFont2);
            cell.setCellStyle(cellStyle);
        }
        for (int i = 6; i < 11; i++) {
            HSSFCell cell = row2.createCell(i);
            HSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyle.setFont(titleFont);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(HSSFColorPredefined.BLACK.getIndex());
            cell.setCellStyle(cellStyle);
        }
        row2.getCell(0).setCellValue("手机编号");
        row2.getCell(1).setCellValue("品牌");
        row2.getCell(2).setCellValue("型号");
        row2.getCell(3).setCellValue("系统版本");
        row2.getCell(4).setCellValue("屏幕尺寸");
        row2.getCell(5).setCellValue("分辨率");
        row2.getCell(6).setCellValue("安装");
        row2.getCell(7).setCellValue("启动");
        row2.getCell(8).setCellValue("卸载");
        row2.getCell(9).setCellValue("问题类型");
        row2.getCell(10).setCellValue("问题描述");
        CellRangeAddress a1 = new CellRangeAddress(0, 0, 0, 5);
        CellRangeAddress g1 = new CellRangeAddress(0, 0, 6, 10);
        sheet.addMergedRegion(a1);
        sheet.addMergedRegion(g1);
        // 绘制数据
        HSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFont(font);
        for (int i = 0; i < data.size(); i++) {
            HSSFRow dataRow = sheet.createRow(2 + i);
            dataRow.setHeight((short)500);
            CompatibilityMachineDTO machine = data.get(i).getMachine();
            dataRow.createCell(0).setCellValue(machine.getAgentId());
            dataRow.createCell(1).setCellValue(machine.getBrand());
            dataRow.createCell(2).setCellValue(machine.getTerminalsModel());
            dataRow.createCell(3).setCellValue(machine.getOsType() + " " + machine.getOsVersion());
            dataRow.createCell(4).setCellValue(machine.getResolution());
            dataRow.createCell(5).setCellValue(machine.getResolution());
            dataRow.createCell(6).setCellValue("-");
            dataRow.createCell(7).setCellValue("-");
            dataRow.createCell(8).setCellValue("-");
            dataRow.createCell(9).setCellValue(data.get(i).getResult());
            dataRow.createCell(10).setCellValue(data.get(i).getMessage());
            for (int j = 0; j < 11; j++) {dataRow.getCell(j).setCellStyle(cellStyle);}
        }
    }

    /**
     * 生成性能统计
     */
    private static void generatePerformanceStatistics(HSSFSheet sheet, List<CompatibilityCommandStatisticItem> data) {
        // TODO 性能报告
        log.debug("性能报告:{},{}", sheet, data);
    }

    /**
     * 生成问题汇总
     */
    private static void generateProblemSummary(HSSFSheet sheet, List<CompatibilityCommandStatisticItem> data) {
        // 根据系统类型分组
        Map<String, List<CompatibilityCommandStatisticItem>> collect = data.stream()
            .filter(t -> !t.getResult())
            .collect(Collectors.groupingBy(t -> t.getMachine().getOsType()));

        // ----  ----  ----  ----  绘制表头
        // 标题字体
        HSSFFont titleFont = sheet.getWorkbook().createFont();
        titleFont.setBold(true);
        titleFont.setFontName(FONT_MICROSOFT_YAHEI);
        titleFont.setFontHeightInPoints((short)11);
        titleFont.setColor(HSSFColorPredefined.WHITE.getIndex());
        HSSFRow row1 = sheet.createRow(0);
        for (int i = 0; i < 7; i++) {
            HSSFCell cell = row1.createCell(i);
            HSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyle.setFont(titleFont);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(HSSFColorPredefined.ROYAL_BLUE.getIndex());
            cell.setCellStyle(cellStyle);
        }
        row1.getCell(0).setCellValue("系统类型");
        row1.getCell(1).setCellValue("问题类型");
        row1.getCell(2).setCellValue("问题描述");
        row1.getCell(3).setCellValue("问题步骤");
        row1.getCell(4).setCellValue("问题图");
        row1.getCell(5).setCellValue("问题机型数");
        row1.getCell(6).setCellValue("问题机型");
        // ----  ----  ----  ----  绘制数据
        // 字体大小是10
        sheet.setDefaultRowHeightInPoints(10);
        // 正文字体
        HSSFFont font = sheet.getWorkbook().createFont();
        font.setFontName(FONT_MICROSOFT_YAHEI);
        HSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFont(font);
        int index = 0;
        for (Entry<String, List<CompatibilityCommandStatisticItem>> entry : collect.entrySet()) {
            List<CompatibilityCommandStatisticItem> v = entry.getValue();
            for (int i = 0; i < v.size(); i++) {
                HSSFRow dataRow = sheet.createRow(++index);
                dataRow.setHeight((short)500);
                CompatibilityMachineDTO machine = v.get(i).getMachine();
                dataRow.createCell(0).setCellValue(machine.getOsType());
                dataRow.createCell(1).setCellValue(machine.getOsVersion());
                dataRow.createCell(2).setCellValue(v.get(i).getResult());
                dataRow.createCell(3).setCellValue(v.get(i).getMessage());
                dataRow.createCell(4).setCellValue("-");
                dataRow.createCell(5).setCellValue(v.size());
                dataRow.createCell(6).setCellValue(machine.getTerminalsModel());
                for (int j = 0; j < 7; j++) {dataRow.getCell(j).setCellStyle(cellStyle);}
            }
        }
    }
}
