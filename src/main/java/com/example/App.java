package com.example;

import java.io.FileWriter;
import java.io.IOException;

import java.util.ArrayList;

class CellComparator implements java.util.Comparator<CellContent> {
    @Override
    public int compare(CellContent a, CellContent b) {
        return (int) (a.left - b.left);
    }
}

public class App {

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

    public static void main(String[] args) {

        String pdfFile = "Customer D.pdf";
        String htmlFile = "pdf.html";
        String excelFile = "test.xlsx";

        PDF pdf = new PDF(pdfFile);
        pdf.WriteHTML(htmlFile);

        HTML html = new HTML(htmlFile);
        ArrayList<CellContent> cellContents = html.getCellContents();

        // WriteContent(cellContents);

        Excel excel = new Excel(excelFile);
        excel.LoadCellContents(cellContents);
    }
}
