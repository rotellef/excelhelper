
package com.rt.excelhelper;

import java.util.List;

public enum F {
    BOLD,
    ITALIC,
    DATE,
    DATETIME,
    TIME,
    PERCENT,
    AMOUNT,
    INTEGER,
    ALIGN_LEFT,
    ALIGN_CENTER,
    ALIGN_RIGHT,
    BORDER_BOTTOM, DEFAULT
    
    ;

    public static String listToKey (final List<F> formats) {
        final StringBuilder sb = new StringBuilder ();
        for (final F f : formats) {
            sb.append (f);
        }
        return sb.toString ();
    }
}
