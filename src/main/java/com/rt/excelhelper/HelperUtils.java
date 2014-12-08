package com.rt.excelhelper;

public class HelperUtils {
    public static String getRange(int colIndex, int start, int end) {
        String colLetter = "";
        colLetter += getCharForNumber(colIndex/26);
        colLetter += getCharForNumber(colIndex%26 + 1);
        return colLetter + (start + 1) + ":" + colLetter + (end + 1);
    }

    /**
     * 1=a 26=z
     * 
     * @param i
     * @return
     */
    private static String getCharForNumber(int i) {
        return i > 0 && i < 27 ? String.valueOf((char) (i + 64)) : "";
    }
}
