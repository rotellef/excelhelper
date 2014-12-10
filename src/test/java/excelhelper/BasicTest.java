
package excelhelper;

import java.math.BigDecimal;
import java.util.Date;

import org.apache.log4j.Logger;
import org.junit.Test;

import com.rt.excelhelper.Block;
import com.rt.excelhelper.ExcelHelper;
import com.rt.excelhelper.ExcelHelperImpl;
import com.rt.excelhelper.F;
import com.rt.excelhelper.RowContent;
import com.rt.excelhelper.Themes;

public class BasicTest {
    private static final Logger LOG = Logger.getLogger (BasicTest.class);

    @Test
    public void testSmall () {
        LOG.info ("small file test");
        final ExcelHelper helper = new ExcelHelperImpl ("testSmall.xlsx");
        helper.setTheme (Themes.STANDARD);
        helper.printHeader (new RowContent ().item ("Navn").item ("Nummer").item ("Antall"));
        helper.printRow (new RowContent ().item ("hei").item (1234).withFormats (F.ITALIC));
        helper.printRow (new RowContent ().item ("hei").item (new BigDecimal (400.20)));
        helper.printRow (new RowContent ().item (new Date ()).item (new BigDecimal (400.20)));
        helper.printRow (new RowContent ().item (null).item (new BigDecimal (400.20)));
        helper.printBlockWithSum (new Block ().withRow (new RowContent ().item ("block").item ("number").item (888))
                                              .withRow (new RowContent ().item ("more block").item (null).item (999)));

        final RowContent customRow = new RowContent ();
        for (int i = 0; i < 200; i++) {
            customRow.item ("hei på deg");
        }
        helper.printRow (customRow);
        helper.writeAndClose ();
    }

    @Test
    public void testMultipleSheets () {
        LOG.info ("small file test");
        final ExcelHelper helper = new ExcelHelperImpl ("testMultisheet.xlsx");
        helper.setTheme (Themes.STANDARD);
        helper.printHeader (new RowContent ().item ("Navn").item ("Nummer").item ("Antall"));
        helper.printBlockWithSum (new Block ().withRow (new RowContent ().item ("block").item ("number").item (888))
                                              .withRow (new RowContent ().item ("more block").item (6362).item (999))
                                              .withRow (new RowContent ().item ("another block").item (null).item (999))
                                              .withRow (new RowContent ().item ("Visa").item (1525).item (999))
                                              .withRow (new RowContent ().item ("Mastercard").item (null).item (999))
                                              .withRow (new RowContent ().item ("Onus").item (41515).item (999))
                                              .withRow (new RowContent ().item ("Credit").item (null).item (999))
                                              .withRow (new RowContent ().item ("Debit").item (7234).item (999))
                                              .withRow (new RowContent ().item ("more block").item (null).item (999))
                                              .withRow (new RowContent ().item ("more block").item (62627).item (999)));

        helper.newSheet ("Sheet nr.2");
        helper.printHeader (new RowContent ().item ("2nd sheet").item ("Nummer").item ("Antall"));
        helper.printBlockWithSum (new Block ().withRow (
                                                  new RowContent ().empty (3).item ("block").item ("number").item (888))
                                              .withRow (
                                                  new RowContent ().empty ().item ("more block").item (6362).item (999))
                                              .withRow (new RowContent ().item ("another block").item (null).item (999))
                                              .withRow (new RowContent ().item ("Visa").item (1525).item (999))
                                              .withRow (new RowContent ().item ("Mastercard").item (null).item (999))
                                              .withRow (new RowContent ().item ("Onus").item (41515).item (999))
                                              .withRow (new RowContent ().item ("Credit").item (null).item (999))
                                              .withRow (new RowContent ().item ("Debit").item (7234).item (999))
                                              .withRow (new RowContent ().item ("more block").item (null).item (999))
                                              .withRow (new RowContent ().item ("more block").item (62627).item (999)));

        helper.writeAndClose ();
    }

    @Test
    public void testBig () {
        LOG.info ("big file test");
        final ExcelHelper helper = new ExcelHelperImpl ("testBig.xlsx");
        helper.setTheme (Themes.STANDARD);
        helper.printHeader (new RowContent ().empty (2).item ("Navn").item ("Nummer").item ("Antall"));
        helper.printRow (new RowContent ().item ("hei").item (1234));
        helper.printRow (new RowContent ().item ("hei").item (new BigDecimal (400.20)));
        helper.printRow (new RowContent ().item (new Date ()).item (new BigDecimal (400.20)));
        for (int i = 0; i < 100*1000; i++) {
            RowContent rowContent = new RowContent ().empty (2);
            for (int j = 0; j < 200; j++)
                rowContent = rowContent.item (new BigDecimal (2452.23));
            helper.printRow (rowContent);
        }
        helper.writeAndClose ();
    }

}
