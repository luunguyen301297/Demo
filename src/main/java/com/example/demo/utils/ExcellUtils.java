package com.viettel.solutions.resource.utils;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.VbaModule;
import com.aspose.cells.Worksheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

public class ExcellUtils {

    public static String code = "Private Sub Worksheet_Change(ByVal Target As Range)\n" +
            "    Dim Oldvalue As String\n" +
            "    Dim Newvalue As String\n" +
            "    Dim items As Variant\n" +
            "    Dim result As String\n" +
            "    Dim isPresent As Boolean\n" +
            "    Dim i As Integer\n" +
            "    Dim ws As Worksheet\n" +
            "    Dim foundID As Variant\n" +
            "    Dim cell As Range\n" +
            "    Dim maxSelections As Integer\n" +
            "    Dim colMapping As Object\n" +
            "    Dim configSheet As Worksheet\n" +
            "    Dim lastRow As Long\n" +
            "    Dim inputRange As Range\n" +
            "    Dim outputColumn As String\n" +
            "    Dim lookupSheetName As String\n" +
            "    Dim lookupRange As String\n" +
            "    \n" +
            "    Application.EnableEvents = False\n" +
            "    On Error GoTo Exitsub\n" +
            "    \n" +
            "    ' Initialize the dictionary for column mappings\n" +
            "    Set colMapping = CreateObject(\"Scripting.Dictionary\")\n" +
            "    \n" +
            "    ' Load configuration from the \"Config\" sheet\n" +
            "    Set configSheet = ThisWorkbook.Sheets(\"Config\")\n" +
            "    lastRow = configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).Row\n" +
            "    \n" +
            "    For i = 1 To lastRow\n" +
            "        colMapping.Add configSheet.Cells(i, 1).Value, Array(configSheet.Cells(i, 2).Value, configSheet.Cells(i, 3).Value, configSheet.Cells(i, 4).Value)\n" +
            "    Next i\n" +
            "    \n" +
            "    ' Iterate over each entry in the dictionary\n" +
            "    For Each Key In colMapping.Keys\n" +
            "        Set inputRange = Range(Key)\n" +
            "        outputColumn = colMapping(Key)(0)\n" +
            "        lookupSheetName = colMapping(Key)(1)\n" +
            "        lookupRange = colMapping(Key)(2)\n" +
            "        \n" +
            "        If Not Intersect(Target, Range(Key)) Is Nothing Then\n" +
            "            For Each cell In Target\n" +
            "                If cell.SpecialCells(xlCellTypeAllValidation) Is Nothing Then GoTo Exitsub\n" +
            "                If cell.Value = \"\" Then GoTo Exitsub\n" +
            "                \n" +
            "                Application.Undo\n" +
            "                Oldvalue = cell.Value\n" +
            "                Application.Undo\n" +
            "                Newvalue = cell.Value\n" +
            "                \n" +
            "                If Oldvalue = \"\" Then\n" +
            "                    cell.Value = Newvalue\n" +
            "                Else\n" +
            "                    items = Split(Oldvalue, \", \")\n" +
            "                    isPresent = False\n" +
            "                    result = \"\"\n" +
            "                    \n" +
            "                    For i = LBound(items) To UBound(items)\n" +
            "                    If Trim(items(i)) = Newvalue Then\n" +
            "                        isPresent = True\n" +
            "                    Else\n" +
            "                        If result = \"\" Then\n" +
            "                            result = Trim(items(i))\n" +
            "                        Else\n" +
            "                            result = result & \", \" & Trim(items(i))\n" +
            "                        End If\n" +
            "                    End If\n" +
            "                Next i\n" +
            "                \n" +
            "                If Not isPresent Then\n" +
            "                    If result = \"\" Then\n" +
            "                        result = Newvalue\n" +
            "                    Else\n" +
            "                        result = result & \", \" & Newvalue\n" +
            "                    End If\n" +
            "                End If\n" +
            "                \n" +
            "                Target.Value = result\n" +
            "            End If\n" +
            "                \n" +
            "                Dim roleItems As Variant\n" +
            "                Dim ids As String\n" +
            "                ids = \"\"\n" +
            "                \n" +
            "                Set ws = ThisWorkbook.Sheets(lookupSheetName)\n" +
            "                roleItems = Split(cell.Value, \", \")\n" +
            "                \n" +
            "                For i = LBound(roleItems) To UBound(roleItems)\n" +
            "                    foundID = Application.VLookup(Trim(roleItems(i)), ws.Range(lookupRange), 2, False)\n" +
            "                    If Not IsError(foundID) Then\n" +
            "                        If ids = \"\" Then\n" +
            "                            ids = foundID\n" +
            "                        Else\n" +
            "                            ids = ids & \", \" & foundID\n" +
            "                        End If\n" +
            "                    End If\n" +
            "                Next i\n" +
            "                \n" +
            "                Cells(cell.Row, outputColumn).Value = ids\n" +
            "            Next cell\n" +
            "        End If\n" +
            "    Next Key\n" +
            "    \n" +
            "Exitsub:\n" +
            "    Application.EnableEvents = True\n" +
            "End Sub\n";

    public static Map<String, CellStyle> genStyleMap(Workbook wb) {
        Map<String, CellStyle> styleMap = new HashMap<>();

        CellStyle titleCss = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontName("Times New Roman");
        f.setFontHeightInPoints((short) 15);
        titleCss.setFont(f);
        titleCss.setAlignment(HorizontalAlignment.CENTER);
        titleCss.setVerticalAlignment(VerticalAlignment.CENTER);
        styleMap.put("titleCss", titleCss);

        CellStyle colHeaderCss = wb.createCellStyle();
        f = wb.createFont();
        f.setBold(true);
        f.setFontName("Times New Roman");
        f.setFontHeightInPoints((short) 13);
        colHeaderCss.setFont(f);
        ((XSSFCellStyle) colHeaderCss).setFillForegroundColor(new XSSFColor(new java.awt.Color(244, 176, 132)));
        colHeaderCss.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        colHeaderCss.setBorderTop(BorderStyle.THIN);
        colHeaderCss.setBorderBottom(BorderStyle.THIN);
        colHeaderCss.setBorderLeft(BorderStyle.THIN);
        colHeaderCss.setBorderRight(BorderStyle.THIN);
        colHeaderCss.setAlignment(HorizontalAlignment.CENTER);
        colHeaderCss.setVerticalAlignment(VerticalAlignment.CENTER);
        styleMap.put("colHeaderCss", colHeaderCss);

        CellStyle colDescriptionCss = wb.createCellStyle();
        f = wb.createFont();
        f.setItalic(true);
        f.setFontName("Times New Roman");
        f.setFontHeightInPoints((short) 13);
        colDescriptionCss.setFont(f);
        ((XSSFCellStyle) colDescriptionCss).setFillForegroundColor(new XSSFColor(new java.awt.Color(252, 228, 214)));
        colDescriptionCss.setWrapText(true);
        colDescriptionCss.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        colDescriptionCss.setBorderTop(BorderStyle.THIN);
        colDescriptionCss.setBorderBottom(BorderStyle.THIN);
        colDescriptionCss.setBorderLeft(BorderStyle.THIN);
        colDescriptionCss.setBorderRight(BorderStyle.THIN);
        colDescriptionCss.setAlignment(HorizontalAlignment.LEFT);
        colDescriptionCss.setVerticalAlignment(VerticalAlignment.CENTER);
        styleMap.put("colDescriptionCss", colDescriptionCss);

        CellStyle normalCss = wb.createCellStyle();
        f = wb.createFont();
        f.setFontName("Times New Roman");
        f.setFontHeightInPoints((short) 13);
        normalCss.setFont(f);
        normalCss.setWrapText(true);
        normalCss.setBorderTop(BorderStyle.THIN);
        normalCss.setBorderBottom(BorderStyle.THIN);
        normalCss.setBorderLeft(BorderStyle.THIN);
        normalCss.setBorderRight(BorderStyle.THIN);
        normalCss.setAlignment(HorizontalAlignment.LEFT);
        normalCss.setVerticalAlignment(VerticalAlignment.CENTER);
        styleMap.put("normalCss", normalCss);

        return styleMap;
    }

    public static void generateCell(Row row, int rowHeight, int rowIdx, XSSFSheet sheet1, int cellIdx, Cell c, CellStyle style, String[] arg) {
        Locale locale = LocaleContextHolder.getLocale();
        if (arg != null && arg.length > 0) {
            String[] params = new String[arg.length];
            row = sheet1.createRow(rowIdx);
            if (rowHeight > 0) {
                row.setHeight((short) rowHeight);
            }
            for (int i = 0; i < arg.length; i++) {
                try {
                    c = row.createCell(cellIdx);
                    if (arg[i] != null && !arg[i].isEmpty()) {
                        c.setCellValue(arg[i]);
                    }
                    c.setCellStyle(style);
                    cellIdx++;
                } catch (Exception e) {
                    params[i] = arg[i];
                }
            }
        }
    }

    public static void generateSheetData(XSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, String constraint) {
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidationConstraint constraintValid = validationHelper.createFormulaListConstraint(constraint);
        DataValidation dataValidation = validationHelper.createValidation(constraintValid, addressList);

        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        }
        sheet.addValidationData(dataValidation);
    }

    public static String formatVLookupCell(String vLookup, String... value) {
        return String.format(vLookup, (Object[]) value);
    }

    public static ByteArrayOutputStream addCodeVBAToExcel(ByteArrayOutputStream byteArrayOutputStream, List<Map<String, Object>> datas) {
        try {
            InputStream byteArrayInputStream = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
            com.aspose.cells.Workbook workbook = new com.aspose.cells.Workbook(byteArrayInputStream);
            byteArrayInputStream.close();
            Worksheet worksheet = workbook.getWorksheets().get(0);

            int idx = workbook.getVbaProject().getModules().add(worksheet);

            VbaModule module = workbook.getVbaProject().getModules().get(idx);
            module.setCodes(code);

            byteArrayOutputStream.reset();
            workbook.save(byteArrayOutputStream, SaveFormat.XLSM);

            ByteArrayInputStream byteArrayInputStreamNew = new ByteArrayInputStream(byteArrayOutputStream.toByteArray());
            Workbook wbNew = new XSSFWorkbook(byteArrayInputStreamNew);
            byteArrayInputStreamNew.close();
            wbNew.removeSheetAt(wbNew.getSheetIndex("Evaluation Warning"));

            List<ColumnMapping> colMappings = new ArrayList<>();
            for (Map<String, Object> data: datas) {
                colMappings.add(new ColumnMapping(data.get("generateRange").toString(), data.get("colMap").toString(), data.get("sheet").toString(), data.get("dataRange").toString()));
            }
            Sheet sheetConfig = wbNew.createSheet("Config");
            int rowNum = 0;
            for (ColumnMapping mapping : colMappings) {
                Row row = sheetConfig.createRow(rowNum++);
                row.createCell(0).setCellValue(mapping.getInputRange());
                row.createCell(1).setCellValue(mapping.getOutputColumn());
                row.createCell(2).setCellValue(mapping.getLookupSheet());
                row.createCell(3).setCellValue(mapping.getLookupRange());
            }
            wbNew.setSheetVisibility(wbNew.getSheetIndex(sheetConfig), SheetVisibility.VERY_HIDDEN);

            byteArrayOutputStream.reset();
            wbNew.write(byteArrayOutputStream);
            wbNew.close();
            return byteArrayOutputStream;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static ByteArrayOutputStream checkExceptionExcel(List<Map<String, Object>> exception, MultipartFile file) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        try {
            InputStream inputStream = file.getInputStream();
            Workbook workbook = WorkbookFactory.create(inputStream);
            inputStream.close();
            Sheet sheet = workbook.getSheetAt(0);
            for (Map<String, Object> e: exception) {
                Row currentRow = sheet.getRow((Integer) e.get("row"));
                Cell currentCell = currentRow.getCell((int) e.get("cell"));
                if (currentCell != null) {
                    currentCell.setCellValue(String.valueOf(e.get("result")));
                } else {
                    currentCell = currentRow.createCell((int) e.get("cell"));
                    currentCell.setCellValue(String.valueOf(e.get("result")));
                }
            }

            workbook.write(byteArrayOutputStream);
            workbook.close();
            return byteArrayOutputStream;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static int findFirstRow(Sheet sheet, String id) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getCell(0).getStringCellValue().equals(id)) {
                return i;
            }
        }
        return -1; // Không tìm thấy
    }

    public static int findLastRow(Sheet sheet, String id) {
        int lastRow = -1;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null && row.getCell(0).getStringCellValue().equals(id)) {
                lastRow = i;
            }
        }
        return lastRow;
    }

    public static void setBorderStyle(Sheet sheet, int rowStart, int rowEnd, int colStart, int colEnd, CellStyle style) {
        for (int rowIndex = rowStart; rowIndex < rowEnd; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            for (int colIndex = colStart; colIndex < colEnd; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }
                cell.setCellStyle(style);
            }
        }
    }

    static class ColumnMapping {
        private String inputRange;
        private String outputColumn;
        private String lookupSheet;
        private String lookupRange;

        public ColumnMapping(String inputRange, String outputColumn, String lookupSheet, String lookupRange) {
            this.inputRange = inputRange;
            this.outputColumn = outputColumn;
            this.lookupSheet = lookupSheet;
            this.lookupRange = lookupRange;
        }

        public String getInputRange() {
            return inputRange;
        }

        public String getOutputColumn() {
            return outputColumn;
        }

        public String getLookupSheet() {
            return lookupSheet;
        }

        public String getLookupRange() {
            return lookupRange;
        }
    }

    public static void addException(List<Map<String, Object>> exception, int rowNum, int cellNum, Object message) {
        Map<String, Object> mapException = new HashMap<>();
        mapException.put("result", message);
        mapException.put("row", rowNum);
        mapException.put("cell", cellNum);
        exception.add(mapException);
    }

    public static HttpHeaders headers(String fileName) {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.ms-excel.sheet.macroEnabled.12"));
        headers.setContentDispositionFormData("attachment", fileName);
        return headers;
    }

}
