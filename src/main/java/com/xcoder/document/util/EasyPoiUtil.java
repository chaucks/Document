package com.xcoder.document.util;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.word.WordExportUtil;
import org.apache.http.util.Asserts;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.util.Assert;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

/**
 * Easy poi util
 *
 * @author Chuck Lee
 */
public class EasyPoiUtil {

    public static final String EXCEL_SUFFIX = ".xls,.xlsx";

    public static final String WORD_SUFFIX = ".doc,.docx";

    /**
     * importExcel
     *
     * @param file file
     * @return list
     */
    public static List<Map> importExcel(File file) {
        ImportParams importParams = new ImportParams();
        importParams.setTitleRows(0);
        importParams.setHeadRows(1);
        return importExcel(file, importParams);
    }

    /**
     * importExcel
     *
     * @param file         file
     * @param importParams importParams
     * @return list
     */
    public static List<Map> importExcel(File file, ImportParams importParams) {
        return ExcelImportUtil.importExcel(file, Map.class, importParams);
    }


    /**
     * easyPoiTemplateExport
     *
     * @param fileName    fileName
     * @param templateUrl templateUrl
     * @param objects     objects
     */
    public static void templateExport(final String fileName, final String templateUrl, final Object... objects) {
        final String fileSuffix = fileName.split("\\.")[1];
        Asserts.check(EasyPoiUtil.isFileSuffix(fileSuffix), "文件后缀名错误，点号开头");
        if (EasyPoiUtil.isExcelFileSuffix(fileSuffix)) {
            EasyPoiUtil.excelTemplateExport(fileName, templateUrl, objects);
        }
        if (EasyPoiUtil.isWordFileSuffix(fileSuffix)) {
            EasyPoiUtil.word07TemplateExport(fileName, templateUrl, objects);
        }
    }

    /**
     * excelTemplateExport
     *
     * @param fileName    fileName
     * @param templateUrl templateUrl
     * @param objects     objects
     */
    public static void excelTemplateExport(String fileName, String templateUrl, Object... objects) {
        Assert.notNull(templateUrl, "Poi template url can not be null.");

        Map<String, Object> map = getMap(objects);
        String absolutePath = SpringWebUtil.getAbsolutePath(templateUrl);
        TemplateExportParams params = new TemplateExportParams(absolutePath);
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        easyPoiExport(fileName, out -> {
            try {
                workbook.write(out);
                return null;
            } catch (IOException ioe) {
                throw new RuntimeException(ioe);
            }
        });
    }

    /**
     * word07TemplateExport
     *
     * @param fileName    fileName
     * @param templateUrl templateUrl
     * @param objects     objects
     */
    public static void word07TemplateExport(String fileName, String templateUrl, Object... objects) {
        try {
            Assert.notNull(templateUrl, "Poi template url can not be null.");

            Map<String, Object> map = getMap(objects);
            String absolutePath = SpringWebUtil.getAbsolutePath(templateUrl);
            XWPFDocument xwpfDocument = WordExportUtil.exportWord07(absolutePath, map);
            easyPoiExport(fileName, out -> {
                try {
                    xwpfDocument.write(out);
                    return null;
                } catch (IOException ioe) {
                    throw new RuntimeException(ioe);
                }
            });
        } catch (Throwable t) {
            throw new RuntimeException(t);
        }
    }

    /**
     * Document export.
     *
     * @param fileName fileName
     * @param function function
     */
    private static void easyPoiExport(String fileName, Function<OutputStream, Void> function) {
        Assert.notNull(fileName, "Poi file name can not be null.");
        HttpServletRequest request = SpringWebUtil.getHttpServletRequest();
        HttpServletResponse response = SpringWebUtil.getHttpServletResponse();
        WebUtil.poiWrite(fileName, function, request, response);
    }

    /**
     * Get poi map
     *
     * @param objects objects
     * @return map
     */
    private static Map<String, Object> getMap(Object... objects) {
        Map<String, Object> map = new HashMap<>(128);
        for (int i = 0, l = objects.length; i < l; i++) {
            map.put((String) objects[i], objects[++i]);
        }
        return map;
    }

    /**
     * isFileSuffix
     *
     * @param fileSuffix fileSuffix
     * @return boolean
     */
    public static boolean isFileSuffix(String fileSuffix) {
        return 0 == fileSuffix.indexOf(".") && 0 == fileSuffix.lastIndexOf(".");
    }

    /**
     * isExcelFileSuffix
     *
     * @param suffix suffix
     * @return boolean
     */
    public static boolean isExcelFileSuffix(String suffix) {
        return EXCEL_SUFFIX.contains(suffix.toLowerCase());
    }

    /**
     * isWordFileSuffix
     *
     * @param suffix suffix
     * @return boolean
     */
    public static boolean isWordFileSuffix(String suffix) {
        return WORD_SUFFIX.contains(suffix.toLowerCase());
    }
}