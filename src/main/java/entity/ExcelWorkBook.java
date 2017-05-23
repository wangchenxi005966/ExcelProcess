package entity;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelWorkBook {

    public Workbook getExcelWorkBook(String fileName) {
        File excelFile = new File(fileName); //创建文件对象
        FileInputStream fileInputStream; //文件流
        Workbook workbook = null;
        try {
            fileInputStream = new FileInputStream(excelFile);
            workbook = WorkbookFactory.create(fileInputStream);
            fileInputStream.close();
        } catch (IOException | org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            e.printStackTrace();
        }
        return workbook;
    }
}
