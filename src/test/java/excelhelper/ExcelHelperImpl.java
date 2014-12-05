package excelhelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.rt.excelhelper.ExcelHelper;
import com.rt.excelhelper.RowContent;

public class ExcelHelperImpl implements ExcelHelper {
    private static final Logger LOG = Logger.getLogger(ExcelHelperImpl.class);
    private final Workbook wb = new XSSFWorkbook();
    CreationHelper createHelper = wb.getCreationHelper();
    private String filePath = "";
    private Sheet activeSheet;
    private int rowPointer = 0;
    private Theme theme = Themes.STANDARD;
    private final Map<StyleType, CellStyle> styles = new HashMap<>();

    public ExcelHelperImpl(String filePath) {
        LOG.info("Starting");
        this.filePath = filePath;
        activeSheet = wb.createSheet("new sheet");
        creatingCellStyles();
    }

    private void creatingCellStyles() {
        Font fontNormal = wb.createFont();
        fontNormal.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        //
        CellStyle datetime = wb.createCellStyle();
        datetime.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy hh:mm"));
        styles.put(StyleType.DATETIME, datetime);
        CellStyle date = wb.createCellStyle();
        date.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy"));
        styles.put(StyleType.DATE, date);
        CellStyle time = wb.createCellStyle();
        time.setDataFormat(createHelper.createDataFormat().getFormat("hh:mm"));
        styles.put(StyleType.DATETIME, time);
        //
        Font fatFont = wb.createFont();
        fatFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        CellStyle heading1 = wb.createCellStyle();
        heading1.setBorderBottom((short) 1);
        heading1.setFont(fatFont);
        styles.put(StyleType.H1, heading1);
        // TODO Auto-generated method stub

    }

    @Override
    public void printRow(RowContent rowContent) {
        try {
            printRow(rowContent.getItems(), StyleType.NORMAL);
        } catch (Exception e) {
            LOG.error("Failed to write row " + rowPointer, e);
        }
    }

    private void printRow(List<Object> rowContent, StyleType styletype) {
        Row row = activeSheet.createRow(rowPointer);
        int colPointer = 0;
        for (Object value : rowContent) {
            printCell(row, colPointer, value, styletype);
            colPointer++;
        }
        rowPointer++;
    }

    private void printCell(Row row, int colPointer, Object value, StyleType styletype) {
        Cell cell = row.createCell(colPointer);
        if (value instanceof Date) {
            cell.setCellValue((Date) value);
            cell.setCellStyle(styles.get(StyleType.DATETIME));
        } else if (value instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) value).doubleValue());
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else {
            cell.setCellValue((Integer) value);
        }
    }

    @Override
    public void writeAndClose() {
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(filePath);
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            LOG.error("Failed to write file", e);
        }
        LOG.info("Done");

    }

    @Override
    public void setTheme(Theme theme) {
        this.theme = theme;
    }

    @Override
    public void printHeader(RowContent rowContent) {
        printRow(rowContent.getItems(), StyleType.H1);
    }
}