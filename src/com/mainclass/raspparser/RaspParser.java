package com.mainclass.raspparser;

import com.mainclass.raspparser.subjects.pair_subject;
import com.mainclass.raspparser.subjects.pair_time;
import com.mainclass.raspparser.subjects.univer_group;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;


public class RaspParser {
    public List<univer_group> readRasp(String filename) throws IOException {
        if (!new File(filename).exists()) {
            return new ArrayList<>();
        }
        Workbook book = null;
        if (filename.endsWith(".xlsx")) {
            book = new XSSFWorkbook(new FileInputStream(filename));     // объект работы с Excel 2010 (*.xlsx)
        } else if (filename.endsWith(".xls")) {
            book = new HSSFWorkbook(new FileInputStream(filename));     // объект работы с Excel 92-2007 (*.xls)
        }

        if (book == null)    // если пустой
        {
            return new ArrayList<>();
        }

//        Workbook book = new XSSFWorkbook(new FileInputStream(filename));     // объект работы с Excel 92-2007 (*.xls)
        int sheetsCount = book.getNumberOfSheets(); // получаем количество страниц

        List<univer_group> groups = new ArrayList<>();

        for (int sheetIndex = 0; sheetIndex < sheetsCount; sheetIndex++) {
            Sheet sheet = book.getSheetAt(sheetIndex);   // получаем страницу

            if (sheet == null) {
                continue;
            }

//        int cellnum = -1;
//        int row_index = 0;
            int rows = sheet.getPhysicalNumberOfRows();     // получаем количество строк
//            System.out.println("кол-во строк: " + rows);
            if (rows < 2) {
                continue;
            }
            if (sheet.getRow(0) == null) {
                continue;
            }
            if (sheet.getRow(0).getPhysicalNumberOfCells() < 2) {
                continue;
            }

            // сначала определяем, в какой день сколько пар и какой паре (по номеру) какая строка соответствует
            String day_of_week = "";
            int day_of_week_number = 0;
            String pair_number = "";
            String week_type = "";
            String pair_start = "";
            String pair_end = "";

            List<pair_time> pair_times = new ArrayList<>();

            for (int i = 3; i < rows; i++)  // начинаем с 3, т.к. там первая неделя
            {
                String celltext = getCellText(sheet, i, 4); // идём по 4-й колонке (отвечает за чередование недель)
//            System.out.println("Строка: " + i);
                if (celltext.isEmpty())  // Если дальше идти некуда
                {
//                System.out.println("Неи информации о неделе");
                    break;
                }
                week_type = celltext;
                String p_end = getCellText(sheet, i, 3);    // время окончания пары
                if (!p_end.isEmpty()) {
                    pair_end = p_end;
                }
                String p_start = getCellText(sheet, i, 2);  // время начала пары
                if (!p_start.isEmpty()) {
                    pair_start = p_start;
                }
                String p_num = getCellText(sheet, i, 1);    // номер пары
                if (!p_num.isEmpty()) {
                    pair_number = p_num;
                }
                String d_week = getCellText(sheet, i, 0);   // неделя, к которой привязана пара
                if (!d_week.isEmpty()) {
                    if (!day_of_week.equals(d_week)) // если уже другой день, то меняем индекс
                    {
                        day_of_week_number++;
                    }
                    day_of_week = d_week;
                }

                try {
                    pair_time p_time = new pair_time();
                    p_time.day_of_week = day_of_week;
                    p_time.day_of_week_number = day_of_week_number;
                    p_time.pair_number = (int) Double.parseDouble(pair_number);
                    p_time.pair_start = pair_start;
                    p_time.pair_end = pair_end;
                    p_time.week_type = week_type;

                    pair_times.add(p_time);

//                System.out.println("Данные о паре: " + p_time.toString());
                } catch (Exception error) {
                    System.out.println(error.toString());
                }


//            System.out.println(p_time);
            }

            // получаем теперь колонки, содержащие предметы
            if (sheet.getPhysicalNumberOfRows() < 2)     // если количество строк меньше 2
            {
                continue;
            }
            List<Integer> columnsWithGroups = new ArrayList<>();
            int columns_count = sheet.getRow(2).getPhysicalNumberOfCells(); // получаем количество столбцов во 2й строке (там, где указано слово предмет)
            for (int col = 0; col < columns_count; col++) {
                String celltext = getCellText(sheet, 2, col); // получаем текст в ячейке
                if (celltext.contains("Предмет") || celltext.contains("Дисциплина")) {
                    columnsWithGroups.add(col); // добавляем колонку в список, как колонку с предметами
                }
            }

//            Vector<univer_group> groups = new Vector<univer_group>();

            for (int col : columnsWithGroups) {
                String group_name = getCellText(sheet, 1, col); // получаем название группы (всегда записано в 1 строке

                univer_group group = new univer_group();
                group.setGroup_name(group_name);

                for (int row = 3; row < pair_times.size() + 3; row++) // проходим по всем строкам, содержащим предметы
                {
//                String subject_name = getCellText(sheet, row, col).replace("\n", "\t");     // название предмета
                    String[] subject_names = getCellText(sheet, row, col).split("\n");
                    for (int i = 0; i < subject_names.length; i++) {
                        String subject_name = subject_names[i];
                        if (!subject_name.isEmpty()) {
                            try {
                                pair_subject pairSubject = new pair_subject();
                                pairSubject.time = pair_times.get(row - 3);   // получаем время пары
                                pairSubject.subject_name = subject_name;
                                pairSubject.subject_type = getCellText(sheet, row, col + 1).split("\n")[i];   // тип занятий
                                pairSubject.subject_lecturer = getCellText(sheet, row, col + 2).split("\n")[i];   // ФИО преподавателя
                                pairSubject.subject_classroom = getCellText(sheet, row, col + 3).split("\n")[i];  // номер аудитории
                                pairSubject.subject_groupName = group_name;

                                group.addPair(pairSubject); // добавляем пару к группе
                            } catch (Exception error) {
//                            System.out.println("Exception: " + error.toString() + " при обработке " + subject_names[i]);
                            }
//                    System.out.println(group_name + " " + subject_name + " " + subject_type + " " + subject_lecturer + " " + subject_classroom);
                        }
                    }
                }

                group.sortPairs(group.getPairs());

                groups.add(group);  // добавляем группу в список
            }
        }


//        // в конце выводим то, что получилось:
//        for (univer_group group : groups)
//        {
//            for (pair_subject pair : group.getPairs())
//            {
//                System.out.println("Группа: " + group.getGroup_name() + " Предмет: " + pair);
//            }
//        }

        return groups;
    }

    private String getCellText(Sheet sheet, int row, int cell) {
        Row sheet_row = sheet.getRow(row);
        if (sheet_row == null) {
//            System.out.println("row is null - " + row + " col - " + cell + " sheet " + sheet.getSheetName());
            return "";
        }
        return getCellText(sheet_row.getCell(cell));
    }

    private String getCellText(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getRichStringCellValue().getString();
            case _NONE:
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
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

    /**
     * Генерирует Excel файл на основе списка пар
     *
     * @param pairs список пар, которые будут записаны в *.xlsx файл
     */
    public void writePairsToFile(List<pair_subject> pairs) throws IOException
    {
        Workbook book = new XSSFWorkbook();     // объект работы с *.xlsx форматом

        // создание листов
        Sheet sheet = book.createSheet("Расписание");   // создаём лист

        // генерируем внешний вид
        // генерируем ячейки

        String[] header = {"День", "№", "Н", "Время", "Предмет"};
        String[] daysNames = {"Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"};
        int pairsInDayCount = 7;

        int rows = 1 + pairsInDayCount * 2 * daysNames.length;   // заголовок + 7 пар * 2 недели * 6 дней
        int cells = header.length;  // день недели, номер пары, номер недели, время, предмет

        for (int i = 0; i < rows; i++)
        {
            Row row = sheet.createRow(i);   // создаём строку в книге
            for (int j = 0; j < cells; j++)
            {
                row.createCell(j);
            }
        }

        // заполняем статичные данные
        // заполняем заголовок
        for (int j = 0; j < header.length; j++)
        {
            sheet.getRow(0).getCell(j).setCellValue(header[j]);
        }

        // заполняем дни недели
        for (int i = 0; i < daysNames.length; i++)
        {
            int rowIndex = i * pairsInDayCount*2 + 1;  // строка, в которую будет записан день недели
            sheet.getRow(rowIndex).getCell(0).setCellValue(daysNames[i]);    // записываем название дня недели
            for (int j = 0; j < pairsInDayCount; j++)
            {
                int pairRowIndex =  rowIndex + j*2;
                sheet.getRow(pairRowIndex).getCell(1).setCellValue(j+1);  // записываем номер пары для нечётной недели
                sheet.getRow(pairRowIndex+1).getCell(1).setCellValue(j+1);  // записываем номер пары для чётной недели

                sheet.getRow(pairRowIndex).getCell(2).setCellValue("I");  // записываем тип недели для нечётной недели
                sheet.getRow(pairRowIndex+1).getCell(2).setCellValue("II");  // записываем тип недели для чётной недели
            }
        }

        // заполняем данными из массива данных
        if(pairs != null)
        {

        }

//        // запись ячеек
//        Row row = sheet.createRow(2);   // создаём строку в книге
//        Cell cell = row.createCell(4);  // создаём ячейку в строке
//        cell.setCellValue("Текст в ячейку 2:4");    // записываем текст в ячейку

        // сохранение файла
        book.write(new FileOutputStream("Workbook_X.xlsx"));  // сохряняем файл
    }
}
