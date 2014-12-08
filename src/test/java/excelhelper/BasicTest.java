package excelhelper;

import java.math.BigDecimal;
import java.util.Date;

import org.apache.log4j.Logger;
import org.junit.Test;

import com.rt.excelhelper.Block;
import com.rt.excelhelper.ExcelHelper;
import com.rt.excelhelper.ExcelHelperImpl;
import com.rt.excelhelper.RowContent;
import com.rt.excelhelper.Themes;

public class BasicTest {
    private static final Logger LOG = Logger.getLogger(BasicTest.class);

    @Test
    public void testSmall() {
        LOG.info("small file test");
        ExcelHelper helper = new ExcelHelperImpl("testSmall.xlsx");
        helper.setTheme(Themes.STANDARD);
        helper.printHeader(new RowContent().item("Navn").item("Nummer").item("Antall"));
        helper.printRow(new RowContent().item("hei").item(1234));
        helper.printRow(new RowContent().item("hei").item(new BigDecimal(400.20)));
        helper.printRow(new RowContent().item(new Date()).item(new BigDecimal(400.20)));
        helper.printRow(new RowContent().item(null).item(new BigDecimal(400.20)));
        helper.printBlockWithSum(new Block().withRow(new RowContent().item("block").item("number").item(888)).withRow(
                        new RowContent().item("more block").item(null).item(999)));
        helper.writeAndClose();
    }

    // @Test
    // public void testBig() {
    // LOG.info("big file test");
    // ExcelHelper helper = new ExcelHelperImpl("testBig.xlsx");
    // helper.setTheme(Themes.STANDARD);
    // helper.printHeader(new
    // RowContent().item("Navn").item("Nummer").item("Antall"));
    // helper.printRow(new RowContent().item("hei").item(1234));
    // helper.printRow(new RowContent().item("hei").item(new
    // BigDecimal(400.20)));
    // helper.printRow(new RowContent().item(new Date()).item(new
    // BigDecimal(400.20)));
    // for (int i = 0; i < 3000; i++) {
    // RowContent rowContent = new RowContent();
    // for (int j = 0; j < 100; j++)
    // rowContent = rowContent.item(new BigDecimal(2452.23));
    // helper.printRow(rowContent);
    // }
    // helper.writeAndClose();
    // }

}
