package com.test;

public class CSLCell {

    private int row;
    private int col;
    private String value;
    private String name;
    private String style;

    public CSLCell(int row, int col, String value, String name, String style) {
        this.row = row;
        this.col = col;
        this.value = value;
        this.name = name;
        this.style = style;
    }

    public int getRow() {
        return row;
    }

    public int getCol() {
        return col;
    }

    public String getValue() {
        return value;
    }

    public String getName() {
        return name;
    }

    public String getStyle() {
        return style;
    }
}
