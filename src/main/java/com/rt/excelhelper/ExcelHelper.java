package com.rt.excelhelper;

import java.util.List;

import excelhelper.Theme;

public interface ExcelHelper {

    void printRow(RowContent row);

    void writeAndClose();

    void setTheme(Theme standard);

    void printHeader(RowContent item);

}
