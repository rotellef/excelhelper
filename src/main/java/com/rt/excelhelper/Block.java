package com.rt.excelhelper;

import java.util.ArrayList;
import java.util.List;

public class Block {
    private final List<RowContent> blockContent = new ArrayList<>();

    public Block withRow(RowContent item) {
        getBlockContent().add(item);
        return this;
    }

    public List<RowContent> getBlockContent() {
        return blockContent;
    }

}
