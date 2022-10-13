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
    public List<univer_group> parceFiles(String path) throws IOException
    {
        String files;
        File folder = new File(path);
        File[] listOfFiles = folder.listFiles();

        // создаём парсер
//        RaspParser parser = new RaspParser();
        List<univer_group> groups = new ArrayList<>();

        // проходим по всем файлам
        for (int i = 0; i < listOfFiles.length; i++)
        {
            if (listOfFiles[i].isFile())
            {
                files = listOfFiles[i].getName();
                System.out.println("Файл " + Integer.toString(i+1) + "/" + Integer.toString(listOfFiles.length) + ": " + files);
                if((!files.startsWith("~$")) && files.endsWith(".xlsx"))  // если это файл, который нам нужен
                {
                    groups.addAll(this.readRasp(path + files));  // читаем и парсим его
                }
            }
        }

        groups.sort(new Comparator<univer_group>()
        {
            @Override
            public int compare(final univer_group group1, final univer_group group2)
            {
//                return group1.getGroup_name().compareTo(group2.getGroup_name());
                return group1.compare(group1, group2);
            }
        });

        return groups;
    }

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
                    if (!day_of_week.equals(d_week)) // если уже другой день, то меняем индекс. Если ячейка пустая, значит день недели остаётся тот же самый
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
     * @param headerText заголовок, который будет выведен над таблицей
     * @param startCellIndex номер столбца, с которого начинается заполнение. Нумерация начинается с 0
     */
    public void writePairsToFile(Sheet sheet, List<pair_subject> pairs, String headerText, int startCellIndex) throws IOException
    {
        // генерируем внешний вид
        // генерируем ячейки

        String[] header = {"День", "№", "Н", "Время", "Предмет"};
        String[] daysNames = {"Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"};
        int pairsInDayCount = 7;    // кол-во пар в день
        int headerRows = 2; // кол-во строк, выделенное под заголовок

        int rows = headerRows + pairsInDayCount * 2 * daysNames.length;   // заголовок + 7 пар * 2 недели * 6 дней
        int cells = header.length;  // день недели, номер пары, номер недели, время, предмет

        for (int i = 0; i < rows; i++)
        {
            Row row = sheet.getRow(i);
            if(row == null)
            {
                row = sheet.createRow(i);   // создаём строку в книге
            }
            for (int j = 0; j < cells; j++)
            {
                row.createCell(startCellIndex + j);
            }
        }

        // заполняем статичные данные
        // заполняем заголовок
        sheet.getRow(0).getCell(startCellIndex + 3).setCellValue(headerText);    // записываем название заголовка
        for (int j = 0; j < header.length; j++)
        {
            sheet.getRow(headerRows-1).getCell(startCellIndex + j).setCellValue(header[j]);
        }

        // заполняем дни недели
        for (int i = 0; i < daysNames.length; i++)
        {
            int rowIndex = i * pairsInDayCount*2 + headerRows;  // строка, в которую будет записан день недели
            sheet.getRow(rowIndex).getCell(startCellIndex).setCellValue(daysNames[i]);    // записываем название дня недели
            for (int j = 0; j < pairsInDayCount; j++)
            {
                int pairRowIndex =  rowIndex + j*2;
                sheet.getRow(pairRowIndex).getCell(startCellIndex + 1).setCellValue(j+1);  // записываем номер пары для нечётной недели
                sheet.getRow(pairRowIndex+1).getCell(startCellIndex + 1).setCellValue(j+1);  // записываем номер пары для чётной недели

                sheet.getRow(pairRowIndex).getCell(startCellIndex + 2).setCellValue("I");  // записываем тип недели для нечётной недели
                sheet.getRow(pairRowIndex+1).getCell(startCellIndex + 2).setCellValue("II");  // записываем тип недели для чётной недели
            }
        }

        // заполняем данными из массива данных
        if(pairs != null)
        {
            for (pair_subject pair : pairs)     // перебор всех пар
            {
                // (pair.time.week_type.equals("II") ? 0 : 1)
                int pair_row = (pair.time.day_of_week_number-1) * pairsInDayCount*2 + (pair.time.pair_number - 1) * 2 + (pair.time.week_type.equals("II") ? 0 : 1) + headerRows;     // номер строки, в которую будет записана пара

                String pair_time = pair.time.pair_start + " - " + pair.time.pair_end;   // формат записи времени пары
                String subject = pair.subject_name + " " + pair.subject_type + " " + pair.subject_lecturer;     // формат записи данных о занятии

                sheet.getRow(pair_row).getCell(startCellIndex + 3).setCellValue(pair_time);  // записываем время
                sheet.getRow(pair_row).getCell(startCellIndex + 4).setCellValue(subject);    // записываем предмет
            }
        }
    }

    /**
     * Генерирует Excel файл на основе списка пар
     *
     * @param pairs список пар, которые будут записаны в *.xlsx файл
     * @param headerText заголовок, который будет выведен над таблицей
     * @param startCellIndex номер столбца, с которого начинается заполнение. Нумерация начинается с 0
     * @param outputFileName название файла на выходе
     */
    public void writePairsToFile(List<pair_subject> pairs, String headerText, int startCellIndex, String outputFileName) throws IOException
    {
        Workbook book = new XSSFWorkbook();     // объект работы с *.xlsx форматом
        Sheet sheet = book.createSheet("Расписание");   // создаём лист
        writePairsToFile(sheet, pairs, headerText, startCellIndex);     // выполняем запись
        book.write(new FileOutputStream(outputFileName));  // сохраняем файл
    }
}
