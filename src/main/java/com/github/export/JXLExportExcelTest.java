package com.github.export;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import java.io.File;
import java.io.IOException;

// JXL 创建 Excel
public class JXLExportExcelTest {



    public void test() {

    }

    @Test
    public void testExport2File() {

        // 创建目录
        File dir = new File("excel");
        if (!dir.exists()) {
            dir.mkdir();
        }

        // 创建 Excel 文件
        File file = new File("excel/jxl_test.xls");

        // 创建表头
        String[] title = {"id", "name", "gender"};

        try {
            file.createNewFile();
            // 创建工作簿
            WritableWorkbook workbook = Workbook.createWorkbook(file);

            // 生成一个名为 “第一页” 的工作表，“0” 表示第一页
            WritableSheet sheet1 = workbook.createSheet("JXL输出页", 0);

            // 使用 Label 对象创建单元格
            Label label = null;

            // 第一行设置列名
            for (int i = 0; i < title.length; i++) {
                // 第i列，第1行对应坐标为 (0,i)
               label = new Label(i, 0, title[i]);
               sheet1.addCell(label);
            }

            // 追加数据
            for (int i = 1; i < 10; i++) {
                // 第1列，第i行
                label = new Label(0, i, "a" + i);
                sheet1.addCell(label);
                // 第2列，第i行
                label = new Label(1, i, "user" + i);
                sheet1.addCell(label);
                // 第3列，第i行
                label = new Label(2, i, "男");
                sheet1.addCell(label);
            }

            // 写入数据
            workbook.write();
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

}
