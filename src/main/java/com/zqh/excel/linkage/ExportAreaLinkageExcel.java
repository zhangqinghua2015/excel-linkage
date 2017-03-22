package com.zqh.excel.linkage;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExportAreaLinkageExcel {

    private static final String OUTPUT_PATH = "D:\\dev\\document\\template.xlsx";

    /**
     * 创建带省市区街道四级联动名称管理器的excel
     *
     * @param provinces 省list
     * @param citysList 市list
     * @param countysList 区list
     * @param streetsList 街道list
     */
    public static void export(List<AreaInfo> provinces, List<List<AreaInfo>> citysList, List<List<AreaInfo>> countysList, List<List<AreaInfo>> streetsList) {

        // 创建一个excel
        Workbook book = new XSSFWorkbook();
        XSSFFont font = (XSSFFont) book.createFont();
        font.setBold(true);
        XSSFCellStyle titleStyle = (XSSFCellStyle) book.createCellStyle();
        // 设置单元格填充颜色
        titleStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 设置单元格字体
        titleStyle.setFont(font);

        XSSFCellStyle idValueStyle = (XSSFCellStyle) book.createCellStyle();
        idValueStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        idValueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 创建需要用户填写的数据页
        // 设计表头
        Sheet sheet1 = book.createSheet("sheet1");
        Row row0 = sheet1.createRow(0);
        row0.createCell(0).setCellValue("省");
        row0.getCell(0).setCellStyle(titleStyle);
        row0.createCell(1).setCellValue("省ID");
        row0.getCell(1).setCellStyle(titleStyle);
        row0.createCell(2).setCellValue("市");
        row0.getCell(2).setCellStyle(titleStyle);
        row0.createCell(3).setCellValue("市ID");
        row0.getCell(3).setCellStyle(titleStyle);
        row0.createCell(4).setCellValue("区");
        row0.getCell(4).setCellStyle(titleStyle);
        row0.createCell(5).setCellValue("区ID");
        row0.getCell(5).setCellStyle(titleStyle);
        row0.createCell(6).setCellValue("街道");
        row0.getCell(6).setCellStyle(titleStyle);
        row0.createCell(7).setCellValue("街道ID");
        row0.getCell(7).setCellStyle(titleStyle);

        //创建一个专门用来存放地区信息的隐藏sheet页
        //因此也不能在现实页之前创建，否则无法隐藏。
        Sheet hideSheet = book.createSheet("site");
        book.setSheetHidden(book.getSheetIndex(hideSheet), true);
        int rowId = 0;
        Row proviRow = hideSheet.createRow(rowId++);
        Row proviRowValue = hideSheet.createRow(rowId++);
        String sheetName = hideSheet.getSheetName();
        String rowIndex;
        String endColumnIndex;
        for (int i = 0; i < provinces.size(); i++) {
            Cell proviCellValue = proviRowValue.createCell(i);
            Cell proviCell = proviRow.createCell(i);
            proviCell.setCellValue(provinces.get(i).getName());
            proviCellValue.setCellValue(provinces.get(i).getId());
            if (i == provinces.size() - 1) {
                rowIndex = (rowId - 1) + "";
                endColumnIndex = CellReference.convertNumToColString(proviCellValue.getColumnIndex());
                String provinceName = sheetName + "!$A$" + rowIndex + ":$" + endColumnIndex + "$" + rowIndex;//"Sheet2!$A$7:$AI$7"
                String provinceValue = sheetName + "!$A$" + rowIndex + ":$" + endColumnIndex + "$" + rowId;
                // 添加省名称
                Name nameProvince = book.createName();
                nameProvince.setNameName("province");
                nameProvince.setRefersToFormula(provinceName);
                Name nameProvinceValue = book.createName();
                nameProvinceValue.setNameName("provinceValue");
                nameProvinceValue.setRefersToFormula(provinceValue);
            }
        }
        // 设置市、区县、乡镇街道数据以及名称
        List<List<List<AreaInfo>>> areaInfos = new ArrayList<List<List<AreaInfo>>>();
        areaInfos.add(citysList);
        areaInfos.add(countysList);
        areaInfos.add(streetsList);
        for (List<List<AreaInfo>> areaInfo : areaInfos) {
            for (int i = 0; i < areaInfo.size(); i++) {
                Row nameRow = hideSheet.createRow(rowId++);
                Row valueRow = hideSheet.createRow(rowId++);
                List<AreaInfo> areas = areaInfo.get(i);
                for (int j = 0; j < areas.size(); j++) {
                    Cell valueCell = valueRow.createCell(j);
                    Cell nameCell = nameRow.createCell(j);
                    nameCell.setCellValue(areas.get(j).getName());
                    valueCell.setCellValue(areas.get(j).getId());
                    if (j == areas.size() - 1) {
                        rowIndex = (rowId - 1) + "";
                        endColumnIndex = CellReference.convertNumToColString(valueCell.getColumnIndex());
                        String nameStr = sheetName + "!$A$" + rowIndex + ":$" + endColumnIndex + "$" + rowIndex;//"Sheet2!$A$7:$AI$7"
                        String valueStr = sheetName + "!$A$" + rowIndex + ":$" + endColumnIndex + "$" + rowId;
                        // 添加名称
                        Name name = book.createName();
                        name.setNameName("_" + areas.get(j).getParentId());
                        name.setRefersToFormula(nameStr);
                        Name value = book.createName();
                        value.setNameName("_" + areas.get(j).getParentId() + "Value");
                        value.setRefersToFormula(valueStr);
                    }
                }
            }
        }
        // 设置数据验证
        Row row1 = sheet1.createRow(1);
        row1.createCell(0).setCellValue("");
        // 设置公式（不知道为什么，id的值单元格必须这样设置公式才能自动填充）
        row1.createCell(1).setCellFormula("HLOOKUP(A2,provinceValue,2,FALSE)");
        row1.getCell(1).setCellStyle(idValueStyle);
        row1.createCell(2).setCellValue("");
        row1.createCell(3).setCellFormula("HLOOKUP(sheet1!C2,INDIRECT(CONCATENATE(\"_\",sheet1!B2,\"Value\")),2,FALSE)");
        row1.getCell(3).setCellStyle(idValueStyle);
        row1.createCell(4).setCellValue("");
        row1.createCell(5).setCellFormula("HLOOKUP(sheet1!E2,INDIRECT(CONCATENATE(\"_\",sheet1!D2,\"Value\")),2,FALSE)");
        row1.getCell(5).setCellStyle(idValueStyle);
        row1.createCell(6).setCellValue("");
        row1.createCell(7).setCellFormula("HLOOKUP(sheet1!G2,INDIRECT(CONCATENATE(\"_\",sheet1!F2,\"Value\")),2,FALSE)");
        row1.getCell(7).setCellStyle(idValueStyle);
        // 设置名称
        String promptBoxStr = "请使用下拉方式选择合适的值！";
        String errorBoxStr = "你输入的值不合法，请不要手动输入！";
        setDataValidationByFormula(sheet1, "province", 2, 1, promptBoxStr, errorBoxStr);
        // 这里只是为了设置id值单元格的错误提示语，公式并不会生效
        setDataValidationByFormula(sheet1, "HLOOKUP(A2,provinceValue,2,FALSE)", 2, 2, "", errorBoxStr);
        setDataValidationByFormula(sheet1, "INDIRECT(CONCATENATE(\"_\",sheet1!B2))", 2, 3, promptBoxStr, errorBoxStr);
        setDataValidationByFormula(sheet1, "HLOOKUP(sheet1!C2,INDIRECT(CONCATENATE(\"_\",sheet1!B2,\"Value\")),2,FALSE)", 2, 4, "", errorBoxStr);
        setDataValidationByFormula(sheet1, "INDIRECT(CONCATENATE(\"_\",sheet1!D2))", 2, 5, promptBoxStr, errorBoxStr);
        setDataValidationByFormula(sheet1, "HLOOKUP(sheet1!E2,INDIRECT(CONCATENATE(\"_\",sheet1!D2,\"Value\")),2,FALSE)", 2, 6, "", errorBoxStr);
        setDataValidationByFormula(sheet1, "INDIRECT(CONCATENATE(\"_\",sheet1!F2))", 2, 7, promptBoxStr, errorBoxStr);
        setDataValidationByFormula(sheet1, "HLOOKUP(sheet1!G2,INDIRECT(CONCATENATE(\"_\",sheet1!F2,\"Value\")),2,FALSE)", 2, 8, "", errorBoxStr);
        FileOutputStream os = null;
        try {
            os = new FileOutputStream(OUTPUT_PATH);
            book.write(os);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(os);
        }

    }

    /**
     * 给工作簿添加数据验证
     *
     * @param sheet              工作簿
     * @param formulaString      公式
     * @param naturalRowIndex    验证的行
     * @param naturalColumnIndex 验证的列
     * @param promptBoxStr       单元格提示语
     * @param errorBoxStr        单元格错误提示语
     */
    private static void setDataValidationByFormula(Sheet sheet, String formulaString, int naturalRowIndex,
                                                  int naturalColumnIndex, String promptBoxStr, String errorBoxStr) {
        //加载下拉列表内容
        DVConstraint constraint = DVConstraint.createFormulaListConstraint(formulaString);
        //设置数据有效性加载在哪个单元格上。
        //四个参数分别是：起始行、终止行、起始列、终止列
        int firstRow = naturalRowIndex - 1;
        int lastRow = naturalRowIndex - 1;
        int firstCol = naturalColumnIndex - 1;
        int lastCol = naturalColumnIndex - 1;
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
                .createFormulaListConstraint(formulaString);
        //数据有效性对象
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, regions);
        //设置输入信息提示信息
        validation.createPromptBox("提示", promptBoxStr);
        //设置输入错误提示信息
        validation.createErrorBox("提示", errorBoxStr);
        validation.setSuppressDropDownArrow(true);
        validation.setShowPromptBox(true);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

}
