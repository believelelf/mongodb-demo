package com.weiquding.mongodb;

import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;

/**
 * mongodb 查询示例
 *
 * @author beliveyourself
 * @version V1.0
 * @date 2020/3/13
 */
public class ImQueryDemo {

    private static final AtomicLong ATOMIC_LONG = new AtomicLong();

    private static final String PATH = "C:\\temp";

    private static final String USER_NAME = "root";
    private static final String PASS = "root";
    private static final String HOST = "localhost";

    private static final int PORT = 27017;
    private static final String SOURCE = "admin";

    private static final String FILE_NAME = "在线客服记录";


    public static void main(String[] args) {
        if(args ==null || args.length == 0){
            System.err.println("请输入每个文件的行数");
            return;
        }

        int fileLineNumber = Integer.parseInt(args[0]);

        MongoCredential testAuth = MongoCredential.createScramSha1Credential(USER_NAME, SOURCE, PASS.toCharArray());
        List<MongoCredential> auths = new ArrayList<>();
        auths.add(testAuth);
        ServerAddress serverAddress = new ServerAddress(HOST, PORT);
        File file = new File(PATH);
        if (!file.exists()) {
            file.mkdirs();
        }

        try (
                MongoClient mongo = new MongoClient(serverAddress, auths);
        ) {
            MongoDatabase db = mongo.getDatabase("im");
            MongoCollection<Document> collections = db.getCollection("im_cc_chat_message");
            long count = collections.count();

            System.out.println("总数：" + count);

            //按5万进行分excel
            long fileNum = count % fileLineNumber == 0 ? count / fileLineNumber : count / fileLineNumber + 1;

            for (int i = 0; i < fileNum; i++) {
                file = new File(PATH + File.separator + FILE_NAME + (i + 1) + ".xlsx");
                if (!file.exists()) {
                    try {
                        boolean newFile = file.createNewFile();
                        System.out.println("创建文件：" + newFile);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }

            MongoCursor<Document> iterator = collections.find().batchSize(100).iterator();

            XSSFWorkbook workbook = null;
            FileOutputStream out = null;
            XSSFSheet sheet = null;
            StyleSet styleSet = null;
            int rowNum = 0;
            int fileIndex = 0;
            while (iterator.hasNext()) {
                Document document = iterator.next();
                long lineNum = ATOMIC_LONG.getAndIncrement();
                int size = document.size();
                //System.out.println("lineNum==>" + lineNum + "; size==>" + size);
                if (lineNum % fileLineNumber == 0) {
                    out = new FileOutputStream(PATH + File.separator + FILE_NAME + (fileIndex + 1) + ".xlsx");
                    workbook = new XSSFWorkbook();
                    sheet = workbook.createSheet(FILE_NAME);
                    Header header1 = new Header(workbook, sheet).invoke();
                    rowNum = header1.getRowNum();
                    styleSet = header1.getStyleSet();
                    fileIndex++;
                    System.out.println("创建第" + fileIndex + "个文件，lineNum=" + lineNum);
                }
                int colNum = 0;
                Row row = sheet.createRow(rowNum++);
                setCellValue(row.createCell(colNum++), document.get("_id"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("sessionId"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("msgId"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("msgType"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("msgBusiType"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("talker"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("sendOutgoing"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("sendStatus"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("sendStatusDesc"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("sendTime"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("msgContent"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("createTime"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("updateTime"), styleSet, false);
                setCellValue(row.createCell(colNum++), document.get("_class"), styleSet, false);
                if (lineNum > 0 && (lineNum + 1) % fileLineNumber == 0 || (lineNum + 1) == count) {
                    System.out.println("写入第" + fileIndex + "个文件，lineNum=" + lineNum + ",写入" + rowNum + "行");
                    workbook.write(out);
                    workbook.close();
                    out.close();
                }
            }
            if (count != ATOMIC_LONG.get()) {
                System.err.println("count error:" + count + "<==>" + ATOMIC_LONG.get());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 设置单元格值<br>
     * 根据传入的styleSet自动匹配样式<br>
     * 当为头部样式时默认赋值头部样式，但是头部中如果有数字、日期等类型，将按照数字、日期样式设置
     *
     * @param cell     单元格
     * @param value    值
     * @param styleSet 单元格样式集，包括日期等样式
     * @param isHeader 是否为标题单元格
     */
    public static void setCellValue(Cell cell, Object value, StyleSet styleSet, boolean isHeader) {
        final CellStyle headCellStyle = styleSet.getHeadCellStyle();
        final CellStyle cellStyle = styleSet.getCellStyle();
        if (isHeader && null != headCellStyle) {
            cell.setCellStyle(headCellStyle);
        } else if (null != cellStyle) {
            cell.setCellStyle(cellStyle);
        }

        if (null == value) {
            cell.setCellValue("");
            cell.setCellType(CellType.BLANK);
        } else if (value instanceof Date) {
            if (null != styleSet.getCellStyleForDate()) {
                cell.setCellStyle(styleSet.getCellStyleForDate());
            }
            cell.setCellType(CellType.STRING);
            cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            cell.setCellType(CellType.STRING);
            cell.setCellValue((Calendar) value);
        } else if (value instanceof Boolean) {
            cell.setCellType(CellType.BOOLEAN);
            cell.setCellValue((Boolean) value);
        } else if (value instanceof RichTextString) {
            cell.setCellType(CellType.STRING);
            cell.setCellValue((RichTextString) value);
        } else if (value instanceof Number) {
            if ((value instanceof Double || value instanceof Float) && null != styleSet.getCellStyleForNumber()) {
                cell.setCellStyle(styleSet.getCellStyleForNumber());
            }
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(((Number) value).doubleValue());
        } else {
            cell.setCellType(CellType.STRING);
            cell.setCellValue(value.toString());
        }
    }

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class StyleSet {

        private CellStyle headCellStyle;
        private CellStyle cellStyle;
        private CellStyle cellStyleForDate;
        private CellStyle cellStyleForNumber;

        public static StyleSet createStyleSet(Workbook workbook) {
            CellStyle headCellStyle = defaultHeaderCellStyle(workbook);

            CellStyle cellStyle = defaultDataCellStyle(workbook);

            CellStyle cellStyleForDate = defaultDataCellStyle(workbook);
            cellStyleForDate.setDataFormat(defaultDateDataFormat(workbook));

            CellStyle cellStyleForNumber = defaultDataCellStyle(workbook);
            cellStyleForNumber.setDataFormat(defaultDoubleDataFormat(workbook));

            return new StyleSet(headCellStyle, cellStyle, cellStyleForDate, cellStyleForNumber);
        }

        /**
         * Returns the default title style. Obtained from:
         * http://svn.apache.org/repos/asf/poi
         * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
         *
         * @param wb the wb
         * @return the cell style
         */
        public static CellStyle defaultTitleCellStyle(final Workbook wb) {
            CellStyle style;
            final Font titleFont = wb.createFont();
            titleFont.setFontHeightInPoints((short) 18);
            titleFont.setBold(true);
            style = wb.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setFont(titleFont);
            return style;
        }

        /**
         * Returns the default header style. Obtained from:
         * http://svn.apache.org/repos/asf/poi
         * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
         *
         * @param wb the wb
         * @return the cell style
         */
        public static CellStyle defaultHeaderCellStyle(final Workbook wb) {
            CellStyle style;
            final Font monthFont = wb.createFont();
            monthFont.setFontHeightInPoints((short) 11);
            monthFont.setColor(IndexedColors.WHITE.getIndex());
            style = wb.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setFont(monthFont);
            style.setWrapText(true);
            return style;
        }

        /**
         * Returns the default data cell style. Obtained from:
         * http://svn.apache.org/repos/asf/poi
         * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
         *
         * @param wb the wb
         * @return the cell style
         */
        public static CellStyle defaultDataCellStyle(final Workbook wb) {
            CellStyle style;
            style = wb.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setWrapText(true);
            style.setBorderRight(BorderStyle.THIN);
            style.setRightBorderColor(IndexedColors.BLACK.getIndex());
            style.setBorderLeft(BorderStyle.THIN);
            style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            style.setBorderTop(BorderStyle.THIN);
            style.setTopBorderColor(IndexedColors.BLACK.getIndex());
            style.setBorderBottom(BorderStyle.THIN);
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            style.setDataFormat(defaultDateDataFormat(wb));
            return style;
        }

        /**
         * Returns the default totals row style for Double data. Obtained from:
         * http://svn.apache.org/repos/asf/poi
         * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
         *
         * @param wb the wb
         * @return the cell style
         */
        public static CellStyle defaultTotalsDoubleCellStyle(final Workbook wb) {
            CellStyle style;
            style = wb.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setDataFormat(defaultDoubleDataFormat(wb));
            return style;
        }

        /**
         * Returns the default totals row style for Integer data. Obtained from:
         * http://svn.apache.org/repos/asf/poi
         * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
         *
         * @param wb the wb
         * @return the cell style
         */
        public static CellStyle defaultTotalsIntegerCellStyle(final Workbook wb) {
            CellStyle style;
            style = wb.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setDataFormat(defaultIntegerDataFormat(wb));
            return style;
        }

        public static short defaultDoubleDataFormat(final Workbook wb) {
            return wb.createDataFormat().getFormat("0.00");
        }

        public static short defaultIntegerDataFormat(final Workbook wb) {
            return wb.createDataFormat().getFormat("0");
        }

        public static short defaultDateDataFormat(final Workbook wb) {
            return wb.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss");
        }


        /**
         * Utility method to determine whether value being put in the Cell is
         * numeric.
         *
         * @param type the type
         * @return true, if is numeric
         */
        public static boolean isNumeric(final Class<?> type) {
            if (isIntegerLongShortOrBigDecimal(type)) {
                return true;
            }
            if (isDoubleOrFloat(type)) {
                return true;
            }
            if (Number.class.equals(type)) {
                return true;
            }
            return false;
        }

        /**
         * Utility method to determine whether value being put in the Cell is
         * integer-like type.
         *
         * @param type the type
         * @return true, if is integer-like
         */
        public static boolean isIntegerLongShortOrBigDecimal(final Class<?> type) {
            if ((Integer.class.equals(type) || (int.class.equals(type)))) {
                return true;
            }
            if ((Long.class.equals(type) || (long.class.equals(type)))) {
                return true;
            }
            if ((Short.class.equals(type)) || (short.class.equals(type))) {
                return true;
            }
            if ((BigDecimal.class.equals(type)) || (BigDecimal.class.equals(type))) {
                return true;
            }
            return false;
        }

        /**
         * Utility method to determine whether value being put in the Cell is
         * double-like type.
         *
         * @param type the type
         * @return true, if is double-like
         */
        public static boolean isDoubleOrFloat(final Class<?> type) {
            if ((Double.class.equals(type)) || (double.class.equals(type))) {
                return true;
            }
            if ((Float.class.equals(type)) || (float.class.equals(type))) {
                return true;
            }
            return false;
        }


    }

    private static class Header {
        private XSSFWorkbook workbook;
        private XSSFSheet sheet;
        private int rowNum;
        private StyleSet styleSet;

        public Header(XSSFWorkbook workbook, XSSFSheet sheet) {
            this.workbook = workbook;
            this.sheet = sheet;
        }

        public int getRowNum() {
            return rowNum;
        }

        public StyleSet getStyleSet() {
            return styleSet;
        }

        public Header invoke() {
            rowNum = 0;
            String[] headers = new String[]{"_id", "sessionId", "msgId", "msgType", "msgBusiType",
                    "talker", "sendOutgoing", "sendStatus", "sendStatusDesc",
                    "sendTime", "msgContent", "createTime", "updateTime", "_class"
            };

            Row row = sheet.createRow(rowNum++);
            styleSet = StyleSet.createStyleSet(workbook);
            int colNum = 0;
            for (String header : headers) {
                Cell cell = row.createCell(colNum++);
                setCellValue(cell, header, styleSet, true);
            }
            return this;
        }
    }
}
