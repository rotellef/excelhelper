package com.rt.excelhelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ExcelHelperImpl implements ExcelHelper {
    private static final int ROW_WINDOWS_TO_FLUSH = 100;
    private static final Logger LOG = Logger.getLogger(ExcelHelperImpl.class);
    private final Workbook wb = new SXSSFWorkbook(ROW_WINDOWS_TO_FLUSH);
    private final DataFormat format = wb.createDataFormat();
    private final Font normalFont = wb.createFont ();
    private final Font fatFont = wb.createFont ();
    private final Font italicFont = wb.createFont ();
    private final Font fatItalicFont = wb.createFont ();
    private final Map<String, CellStyle> styles = new HashMap<>();
    private final Sheet activeSheet;
    private String filePath = "";
    private int rowPointer = 0;
    private Theme theme = Themes.STANDARD;

    public ExcelHelperImpl(final String filePath) {
        LOG.info("Starting");
        this.filePath = WorkbookUtil.createSafeSheetName (filePath);
        activeSheet = wb.createSheet("new sheet");
        // Font setup
        normalFont.setBoldweight (Font.BOLDWEIGHT_NORMAL);
        fatFont.setBoldweight (Font.BOLDWEIGHT_BOLD);
        italicFont.setItalic (true);
        fatItalicFont.setBoldweight (Font.BOLDWEIGHT_BOLD);
        fatItalicFont.setItalic (true);
    }

    @Override
    public void printRow(final RowContent rowContent) {
        try {
            printRow(rowContent.getItems(), rowContent.getFormats ());
        } catch (final Exception e) {
            LOG.error("Failed to write row " + rowPointer, e);
        }
    }

    private void printRow(final List<Object> rowContent, final F...formatOverride) {
        final Row row = activeSheet.createRow(rowPointer);
        int colPointer = 0;
        for (final Object value : rowContent) {
            printCell(row, colPointer, value, formatOverride);
            colPointer++;
        }
        rowPointer++;
    }

    private void printCell(final Row row, final int colPointer, final Object content, final F[] formatOverride) {
        if (content == null) return;
        final List<F> formats = new ArrayList<> (Arrays.asList (formatOverride));
        // Determine formats
        if (content instanceof Date) {
            formats.add (F.DATETIME);
        }
        
        final Cell cell = row.createCell(colPointer);
        cell.setCellStyle(getStyle(formats));
        // Set the cell values
        if (content instanceof Date) {
            cell.setCellValue((Date) content);
        } else if (content instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) content).doubleValue());
        } else if (content instanceof String) {
            cell.setCellValue((String) content);
        } else {
            cell.setCellValue((Integer) content);
        }
    }

    /**
     * Looks for a style with all the formats in the map. If its not found, it creates a new style with the formats as a
     * key and stores it in the map.
     * 
     * @param formats
     * @return
     */
    private CellStyle getStyle (final List<F> formats) {
        final String key = F.listToKey (formats);
        CellStyle style = styles.get (key);
        if (style != null) {
            // Style found stored in map, return it!
            return style;
        } else {
            // Create new style
            style = wb.createCellStyle ();
            styles.put (key, style);
        }
        
        Font font = normalFont;
        for(final F f : formats){
            switch (f) {
                case AMOUNT:
                    style.setDataFormat (format.getFormat ("#,##0.00"));
                    break;
                case DATE:
                    style.setDataFormat(format.getFormat("m/d/yy"));
                    break;
                case DATETIME:
                    style.setDataFormat(format.getFormat("m/d/yy hh:mm"));
                    break;
                case TIME:
                    style.setDataFormat(format.getFormat("hh:mm"));
                    break;
                case INTEGER:
                    style.setDataFormat (format.getFormat ("#,##0"));
                    break;
                case PERCENT:
                    style.setDataFormat (format.getFormat ("#%"));
                    break;
                case ALIGN_CENTER:
                    style.setAlignment (CellStyle.ALIGN_CENTER);
                    break;
                case ALIGN_RIGHT:
                    style.setAlignment (CellStyle.ALIGN_RIGHT);
                    break;
                case ALIGN_LEFT:
                    style.setAlignment (CellStyle.ALIGN_LEFT);
                    break;
                case BOLD:
                    if(font.equals (italicFont))
                        font = fatItalicFont;
                    else
                        font = fatFont;
                    break;
                case BORDER_BOTTOM:
                    style.setBorderBottom (CellStyle.BORDER_MEDIUM);
                    break;
                case ITALIC:
                    if(font.equals (fatFont))
                        font = fatItalicFont;
                    else
                        font = italicFont;
                    break;
                default:
                    break;
            }
        }
        style.setFont (font);
        return style;
    }

    @Override
    public void writeAndClose() {
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(filePath);
            wb.write(fileOut);
            fileOut.close();
        } catch (final IOException e) {
            LOG.error("Failed to write file", e);
        }
        LOG.info("Done");

    }

    @Override
    public void setTheme(final Theme theme) {
        this.theme = theme;
    }

    @Override
    public void printHeader(final RowContent rowContent) {
        printRow(rowContent.getItems(), F.BOLD, F.BORDER_BOTTOM);
    }

    @Override
    public void printBlockWithSum(final Block block) {
        final int blockStartRow = rowPointer;
        for (final RowContent row : block.getBlockContent()) {
            printRow(row);
        }
        printSumRow(blockStartRow, rowPointer - 1, new int[]{2,3,4,5});

    }

    private void printSumRow(final int start, final int end, final int[] colIndexes) {
        final Row row = activeSheet.createRow(rowPointer);
        for (final int colIndex : colIndexes) {
            final Cell cell = row.createCell(colIndex);
            cell.setCellFormula("SUM(" + HelperUtils.getRange(colIndex, start, end) + ")");
        }
    }


}
