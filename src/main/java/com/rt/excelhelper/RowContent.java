package com.rt.excelhelper;

import java.util.ArrayList;
import java.util.List;

public class RowContent {
    private final List<Object> items = new ArrayList<>();
    private F[] formats = new F[]{F.DEFAULT};
    public RowContent item(final Object o){
        getItems().add(o);
        return this;
    }
    public List<Object> getItems() {
        return items;
    }
    public RowContent withFormats(final F...formats){
        this.formats = formats;
        return this;
    }
    public F[] getFormats () {
        return formats;
    }
}
