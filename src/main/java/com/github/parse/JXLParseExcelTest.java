package com.github.parse;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;

// JXL 解析 Excel 文件
public class JXLParseExcelTest {

    public static void main(String[] args) {
        try {
            // 创建工作簿对象
            Workbook workbook = Workbook.getWorkbook(new File("excel/jxl_test.xls"));

            // 通过工作簿对象获取第一张工作表对象
            Sheet sheet = workbook.getSheet(0);
            int rows = sheet.getRows();
            int columns = sheet.getColumns();

            // 获取数据
            for (int i = 0; i < rows; i++) {
                for (int j = 0; j < columns; j++) {
                    Cell cell = sheet.getCell(j, i);
                    System.out.print(cell.getContents() + " ");
                }
                System.out.println();
            }

            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }

    }

}
