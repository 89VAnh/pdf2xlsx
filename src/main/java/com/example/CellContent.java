package com.example;

public class CellContent {
    public String value;
    public double top;
    public double left;
    public double line_height;
    public String font_family;
    public double font_size;
    public double width;
    public String color;
    public boolean isRotate = false;

    @Override
    public String toString() {
        return this.value + "(" + this.top + "|" + this.left + "|" + this.line_height + "|" + this.font_family + "|"
                + this.font_size + "|" + this.width + "|" + this.color + ")";
    }
}
