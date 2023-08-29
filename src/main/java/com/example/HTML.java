package com.example;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class HTML {
    Document doc;

    public HTML(String filename) {
        try {
            doc = Jsoup.parse(new File(filename), "ISO-8859-1");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getStyleValue(String style, String attName) {

        int attIndex = style.indexOf(attName);
        if (attIndex != -1) {
            int nextSemicolon = style.indexOf(";", attIndex);

            return style.substring(attIndex, nextSemicolon).split(":")[1];

        } else
            return "";

    }

    private static CellContent getElementAtt(Element e) {
        CellContent cellContent = new CellContent();

        cellContent.value = e.text();

        String style = e.attributes().asList().get(2).toString();

        cellContent.top = Double.parseDouble(getStyleValue(style, "top").replace("pt", ""));

        cellContent.left = Double.parseDouble(getStyleValue(style, "left").replace("pt", ""));

        cellContent.line_height = Double.parseDouble(getStyleValue(style, "line-height").replace("pt", ""));

        cellContent.font_family = getStyleValue(style, "font-family");

        cellContent.font_size = Double.parseDouble(getStyleValue(style, "font-size").replace("pt", ""));

        cellContent.width = Double.parseDouble(getStyleValue(style, "width").replace("pt", ""));

        cellContent.color = getStyleValue(style, "color");

        return cellContent;
    }

    public ArrayList<CellContent> getCellContents() {
        Elements content = doc.getElementsByClass("p");
        ArrayList<CellContent> cellContents = new ArrayList<CellContent>();

        for (int i = 0; i < content.size(); i++) {
            CellContent currentCell = getElementAtt(content.get(i));

            if (currentCell.value.equals("DOM") | currentCell.value.equals("CBU")) {
                continue;
            }

            if (cellContents.size() > 1) {
                CellContent lastCell = cellContents.get(cellContents.size() - 1);

                double space = currentCell.left - lastCell.left - lastCell.width;

                if (space == 0) {
                    lastCell.value += currentCell.value;
                    lastCell.isRotate = true;
                    continue;
                }

                if (lastCell.top == currentCell.top && lastCell.line_height == currentCell.line_height
                        && lastCell.font_family.equals(currentCell.font_family)
                        && lastCell.font_size == currentCell.font_size
                        && lastCell.color.equals(currentCell.color)
                        && space > 0 && space < 1.5 * lastCell.font_size) {
                    lastCell.value += " " + currentCell.value;
                    lastCell.width += space + currentCell.width;
                    continue;
                }
            }
            cellContents.add(currentCell);
        }

        return cellContents;
    }
}
