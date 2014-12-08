package excelhelper;

import static org.junit.Assert.*;

import org.junit.Test;

import com.rt.excelhelper.HelperUtils;

public class UtilsTests {

    @Test
    public void test() {
        assertEquals("BA1:BA30",HelperUtils.getRange(52, 0, 29));
        assertEquals("AZ1:AZ30",HelperUtils.getRange(51, 0, 29));
        assertEquals("AA1:AA30",HelperUtils.getRange(26, 0, 29));
        assertEquals("Z1:Z30",HelperUtils.getRange(25, 0, 29));
        assertEquals("A1:A30",HelperUtils.getRange(0, 0, 29));
    }

}
