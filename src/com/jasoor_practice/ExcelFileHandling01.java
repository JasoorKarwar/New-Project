package com.jasoor_practice;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileHandling01 {

    public static void main(String[] args) throws IOException {

        String path = "C:\\Users\\karwa\\eclipse-workspace\\MyFirstProject\\src\\com\\jasoor_practice\\PracticeTest.xlsx";

        System.out.println(path); //path to the file
        FileInputStream fileInputStream=new FileInputStream(path); //creating connection
        XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream); //creating the object of XSSFworkbook to manipulate xlsx files
        Sheet sheet=workbook.getSheet("FirstPage"); //accessing the sheet
        
        for (int row = 0; row<sheet.getPhysicalNumberOfRows();row++);
        int row = 0;
        Row rowData = sheet.getRow(row);

    }

}
