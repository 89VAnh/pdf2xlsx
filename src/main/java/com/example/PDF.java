package com.example;

import java.io.File;
import java.io.PrintWriter;
import java.io.Writer;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.fit.pdfdom.PDFDomTree;

public class PDF {

    private PDDocument pdf = null;

    public PDF(String filename) {
        try {
            pdf = PDDocument.load(new File(filename));
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    public void WriteHTML(String htmlFile) {
        try {
            Writer output = new PrintWriter(htmlFile, "utf-8");

            new PDFDomTree().writeText(pdf, output);

            output.close();
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
}
