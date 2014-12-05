package com.rt.excelhelper;

import java.util.ArrayList;
import java.util.List;

public class RowContent {
    private final List<Object> items = new ArrayList<>();
    public RowContent item(Object o){
        getItems().add(o);
        return this;
    }
    public List<Object> getItems() {
        return items;
    }
}
