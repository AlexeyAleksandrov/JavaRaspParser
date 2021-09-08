package com.mainclass;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.*;
import java.util.List;
import java.util.Vector;

import static org.apache.poi.ss.usermodel.CellType.*;

public class ExcelEditor
{
    public void createSimpleXlsExcelBook() throws IOException
    {
        Workbook book = new HSSFWorkbook();     // объект работы с Excel 92-2007 (*.xls)
        Sheet sheet = book.createSheet("My_list_28");   // создаём лист



        FileOutputStream ExcelOutStream = new FileOutputStream("testexcel.xls");    // создаём поток для записи файла
        book.write( ExcelOutStream );   // записываем Excel

        ExcelOutStream.close();     // закрываем поток
    }

    public void createSimpleXlsXExcelBook() throws IOException
    {
        Workbook book = new XSSFWorkbook();     // объект работы с *.xlsx форматом

        // создание листов
        Sheet sheet = book.createSheet("My_list_282");   // создаём лист

        // запись ячеек
        Row row = sheet.createRow(2);   // создаём строку в книге
        Cell cell = row.createCell(4);  // создаём ячейку в строке
        cell.setCellValue("Текст в ячейку 2:4");    // записываем текст в ячейку

        // сохранение файла
        book.write(new FileOutputStream("Workbook_X.xlsx"));  // сохряняем файл
    }

    public void readSimpleXlsExcelBook() throws IOException
    {
        Workbook book = new HSSFWorkbook(new FileInputStream("_workbook.xls"));     // объект работы с Excel 92-2007 (*.xls)
        Sheet sheet = book.getSheetAt(0);   // получаем страницу

        for (Row row : sheet)   // прохолим по всем строка
        {
            for (Cell cell : row)   // проходим по всем ячейкам строки
            {
                String celltext = getCellText(cell);    // получаем текст из ячейки
                System.out.println("text: " + celltext);
            }
        }

//        Cell cell = sheet.getRow(2).getCell(4); // получаем ячейку
//        String celltext = getCellText(cell);    // получаем текст из ячейки
//        System.out.println("text: " + celltext);
    }
    public void readSimpleXlsXExcelBook() throws IOException
    {
        Workbook book = new XSSFWorkbook(new FileInputStream("C:\\Users\\ASUS\\IdeaProjects\\JavaExcel\\rasp_data\\ФТИ_4 курс_21-22_осень.xlsx"));     // объект работы с Excel 92-2007 (*.xls)
        Sheet sheet = book.getSheetAt(0);   // получаем страницу

        for (Row row : sheet)   // прохолим по всем строка
        {
            for (Cell cell : row)   // проходим по всем ячейкам строки
            {
                String celltext = getCellText(cell);    // получаем текст из ячейки
                System.out.println("text: " + celltext);
            }
        }
    }


    private String getCellText(Cell cell)
    {
        if(cell == null)
        {
            return "";
        }
        switch (cell.getCellType())
        {
            case STRING:
                return cell.getRichStringCellValue().getString();
            case _NONE:
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell))
                {
                    return cell.getDateCellValue().toString();
                }
                else
                {
                    return Double.toString(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
            case ERROR:
            default:
                return "";
        }
        return "";
    }

    public Vector<String> listFilesForFolder(final @NotNull File folder)
    {
        Vector<String> filesList = new Vector<String>();
        for (final File fileEntry : folder.listFiles())
        {
            if (fileEntry.isDirectory())
            {
                listFilesForFolder(fileEntry);
            }
            else
            {
//                System.out.println(fileEntry.getName());
                filesList.add(fileEntry.getName());
            }
        }
        return filesList;
    }
}
