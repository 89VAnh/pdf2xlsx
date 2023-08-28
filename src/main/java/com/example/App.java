package com.example;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.Writer;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.fit.pdfdom.PDFDomTree;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.util.ArrayList;
import java.util.Collections;
import java.util.TreeMap;
import java.util.Map;

class CellComparator implements java.util.Comparator<CellContent> {
    @Override
    public int compare(CellContent a, CellContent b) {
        return (int) (a.left - b.left);
    }
}

public class App {
    private static void generateHTMLFromPDF(String filename) {
        try {

            PDDocument pdf = PDDocument.load(new File(filename));
            Writer output = new PrintWriter("pdf.html", "utf-8");

            new PDFDomTree().writeText(pdf, output);

            output.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    private static String getValueInStyle(String style, String attName) {

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

        cellContent.top = Double.parseDouble(getValueInStyle(style, "top").replace("pt", ""));

        cellContent.left = Double.parseDouble(getValueInStyle(style, "left").replace("pt", ""));

        cellContent.line_height = Double.parseDouble(getValueInStyle(style, "line-height").replace("pt", ""));

        cellContent.font_family = getValueInStyle(style, "font-family");

        cellContent.font_size = Double.parseDouble(getValueInStyle(style, "font-size").replace("pt", ""));

        cellContent.width = Double.parseDouble(getValueInStyle(style, "width").replace("pt", ""));

        cellContent.color = getValueInStyle(style, "color");

        return cellContent;
    }

    private static ArrayList<CellContent> HTMLParse(String filename) {
        Document htmlFile = null;
        try {
            htmlFile = Jsoup.parse(new File(filename), "ISO-8859-1");
        } catch (IOException e) {
            e.printStackTrace();
        }

        Elements content = htmlFile.getElementsByClass("p");
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

    private static void WriteContent(ArrayList<CellContent> cellContents) {
        try {
            FileWriter fWriter = new FileWriter("text.txt");

            for (CellContent cellContent : cellContents) {
                fWriter.write(cellContent.toString() + "\n");
            }

            fWriter.close();
        } catch (IOException e) {

            System.out.print(e.getMessage());
        }
    }

    private static TreeMap<Integer, ArrayList<CellContent>> getTreeMap(ArrayList<CellContent> cellContents) {
        TreeMap<Integer, ArrayList<CellContent>> MapContent = new TreeMap<Integer, ArrayList<CellContent>>();

        for (CellContent cellContent : cellContents) {
            Integer key = (int) Math.floor(cellContent.top / 3);

            if (cellContent.isRotate | cellContent.value.matches("Page[0-9]+of[0-9]+")
                    | cellContent.value.contains("FACTORY")) {
                continue;
            }

            if (MapContent.get(key) == null) {
                ArrayList<CellContent> cellList = new ArrayList<CellContent>();
                cellList.add(cellContent);
                MapContent.put(key, cellList);
            } else {
                MapContent.get(key).add(cellContent);
            }
        }

        for (Map.Entry<Integer, ArrayList<CellContent>> entry : MapContent.entrySet()) {
            Collections.sort(entry.getValue(), new CellComparator());
        }

        // try {
        // FileWriter fWriter = new FileWriter("map.txt");

        // for (Map.Entry<Integer, ArrayList<CellContent>> entry :
        // MapContent.entrySet()) {
        // fWriter.write(entry.getKey() + " : " + entry.getValue() + "\n");
        // System.out.println(entry.getKey() + " : " + entry.getValue().size());
        // }

        // fWriter.close();
        // } catch (IOException e) {

        // System.out.print(e.getMessage());
        // }

        return MapContent;
    }

    public static void main(String[] args) {
        // generateHTMLFromPDF("Customer D.pdf");

        ArrayList<CellContent> cellContents = HTMLParse("pdf.html");

        // WriteContent(cellContents);
        ArrayList<CellContent> rotateList = new ArrayList<CellContent>();

        for (CellContent cellContent : cellContents) {
            if (cellContent.isRotate) {
                rotateList.add(cellContent);
            }
        }

        ArrayList<CellContent> facList = new ArrayList<CellContent>();

        for (CellContent cellContent : cellContents) {
            if (cellContent.value.startsWith("FACTORY")) {
                facList.add(cellContent);
            }
        }

        TreeMap<Integer, ArrayList<CellContent>> treeMap = getTreeMap(cellContents);
        Excel.Write("test.xlsx", treeMap, rotateList, facList);
    }
}
