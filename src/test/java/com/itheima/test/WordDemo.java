package com.itheima.test;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.util.List;

public class WordDemo {
    public static void main(String[] args) throws Exception {
        XWPFDocument document = new XWPFDocument(new FileInputStream("d://test.docx"));
        
//        读取正文：
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            System.out.println(paragraph.getText());
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                System.out.println(run.getText(0));

            }
        }
//        读取表格
        XWPFTable xwpfTable = document.getTables().get(0);
        List<XWPFTableRow> rows = xwpfTable.getRows();
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> tableCells = row.getTableCells();
            for (XWPFTableCell tableCell : tableCells) {
                List<XWPFParagraph> paragraphs1 = tableCell.getParagraphs();
                for (XWPFParagraph xwpfParagraph : paragraphs1) {
                    System.out.println(xwpfParagraph.getText());
                }
            }
        }
    }
}
