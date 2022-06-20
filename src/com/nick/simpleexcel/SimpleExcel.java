package com.nick.simpleexcel;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class SimpleExcel {

    public static void main(String[] args) throws Exception {
        // String directory = System.getProperty("user.dir");

        // this is comment 2 ;
        String file1 = "resource/new.xls";
        String file2 = "resource/test.xls";
        String file3 = "resource/new.xls";
        String file4 = "resource/new.xls";

        String[][] output1 = readFile(file1);
        String[][] output2 = readFile(file2);
        String[][] output3 = readFile(file3);
        String[][] output4 = readFile(file4);

        System.out.println(output1[1][1]);
        System.out.println(output1[2][1]);
        System.out.println(output2[1][1]);
        System.out.println(output3[1][1]);
        System.out.println(output4[1][1]);

    }

    public static String[][] readFile(String fileName) throws Exception {

        File file = new File(fileName);
        FileInputStream fileInputStream = new FileInputStream(file);
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet hssfSheet = hssfWorkbook.getSheet("Country");

        int rowNum = hssfSheet.getLastRowNum() + 1;
        int colNum = hssfSheet.getRow(0).getLastCellNum();
        String[][] data = new String[rowNum][colNum];

        for (int i = 0; i < rowNum; i++) {

            HSSFRow hssfRow = hssfSheet.getRow(i);
            for (int j = 0; j < colNum; j++) {
                HSSFCell hssfCell = hssfRow.getCell(j);
                String value = celltoString(hssfCell);
                data[i][j] = value;
                System.out.println("The value is " + value);
            }

        }

        return data;

    }

    public static String celltoString(HSSFCell hssfCell) {
        int type;
        Object result;
        type = hssfCell.getCellType();

        switch (type) {

        case 0:
            result = hssfCell.getNumericCellValue();
            break;
        case 1:
            result = hssfCell.getStringCellValue();
            break;
        default:
            throw new RuntimeException("No Supporting Type");

        }

        return result.toString();
    }

}
