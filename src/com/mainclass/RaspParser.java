package com.mainclass;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Comparator;
import java.util.Vector;
import java.util.Comparator;

class pair_time
{
    public String day_of_week = "";
    public int day_of_week_number = 1;
    public int pair_number = 0;
    public String week_type = "";
    public String pair_start = "";
    public String pair_end = "";


//    @Override
//    public String toString()
//    {
//        return day_of_week + " " + Integer.toString(pair_number) + " " + pair_start + " " + pair_end + " " + week_type;
//    }


    @Override
    public String toString()
    {
        return "pair_time{" +
                "day_of_week='" + day_of_week + '\'' +
                ", pair_number=" + pair_number +
                ", week_type='" + week_type + '\'' +
                ", pair_start='" + pair_start + '\'' +
                ", pair_end='" + pair_end + '\'' +
                '}';
    }
}

class pair_subject
{
    public pair_time time = null;
    public String subject_name;    // название предмета
    public String subject_type = "";   // тип занятий
    public String subject_lecturer = "";   // ФИО преподавателя
    public String subject_classroom = "";  // аудитория, в которой проводятся занятия
    public String subject_groupName = "";

    @Override
    public String toString()
    {
        return "pair_subject{" +
                "time=" + time +
                ", subject_name='" + subject_name + '\'' +
                ", subject_type='" + subject_type + '\'' +
                ", subject_lecturer='" + subject_lecturer + '\'' +
                ", subject_classroom='" + subject_classroom + '\'' +
                '}';
    }
}

class univer_group
{
    private String group_name = null;
    private final Vector<pair_subject> pairs = new Vector<pair_subject>();

    public void setGroup_name(String group_name)
    {
        this.group_name = group_name;
    }

    public String getGroup_name()
    {
        return group_name;
    }

    public Vector<pair_subject> getPairs()
    {
        return pairs;
    }

    public void addPair(pair_subject pair)
    {
        pairs.add(pair);
    }

    public static void sortPairs(Vector<pair_subject> pairs)
    {
        pairs.sort(new Comparator<pair_subject>()
        {
            @Override
            public int compare(pair_subject pair1, pair_subject pair2)
            {
                int compare_day_of_week = pair1.time.day_of_week_number - pair2.time.day_of_week_number;
                if (compare_day_of_week == 0)
                {
                    int compare_pair_number = pair1.time.pair_number - pair2.time.pair_number;
                    if(compare_pair_number == 0)
                    {
                        int compare_week_type = pair1.time.week_type.compareTo(pair2.time.week_type);
                        if(compare_week_type == 0)
                        {
                            int compare_subject_name = pair1.subject_name.compareTo(pair2.subject_name);
                            return compare_subject_name;
                        }
                        else
                        {
                            return compare_week_type;
                        }
                    }
                    else
                    {
                        return compare_pair_number;
                    }
                }
                else
                {
                    return compare_day_of_week;
                }
            }
        });
    };

    @Override
    public String toString()
    {
        return "univer_group{" +
                "group_name='" + group_name + '\'' +
                ", pairs=" + pairs +
                '}';
    }
}


public class RaspParser
{
    public Vector<univer_group> readRasp(String filename) throws IOException
    {
        if(!new File(filename).exists())
        {
            return new Vector<univer_group>();
        }
        Workbook book = new XSSFWorkbook(new FileInputStream(filename));     // объект работы с Excel 92-2007 (*.xls)
        Sheet sheet = book.getSheetAt(0);   // получаем страницу

//        int cellnum = -1;
//        int row_index = 0;
        int rows = sheet.getPhysicalNumberOfRows();     // получаем количество строк

        // сначала определяем, в какой день сколько пар и какой паре (по номеру) какая строка соответствует
        String day_of_week = "";
        int day_of_week_number = 0;
        String pair_number = "";
        String week_type = "";
        String pair_start = "";
        String pair_end = "";

        Vector<pair_time> pair_times = new Vector<pair_time>();

        for (int i=3; i<rows; i++)  // начинаем с 3, т.к. там первая неделя
        {
            String celltext = getCellText(sheet, i, 4); // идём по 4-й колонке (отвечает за чередование недель)
            if(celltext.isEmpty())  // Если дальше идти некуда
            {
                break;
            }
            week_type = celltext;
            String p_end = getCellText(sheet, i, 3);
            if(!p_end.isEmpty())
            {
                pair_end = p_end;
            }
            String p_start = getCellText(sheet, i, 2);
            if(!p_start.isEmpty())
            {
                pair_start = p_start;
            }
            String p_num = getCellText(sheet, i, 1);
            if(!p_num.isEmpty())
            {
                pair_number = p_num;
            }
            String d_week = getCellText(sheet, i, 0);
            if(!d_week.isEmpty())
            {
                if(!day_of_week.equals(d_week)) // если уже другой день, то меняем индекс
                {
                    day_of_week_number++;
                }
                day_of_week = d_week;
            }

            try
            {
                pair_time p_time = new pair_time();
                p_time.day_of_week = day_of_week;
                p_time.day_of_week_number = day_of_week_number;
                p_time.pair_number = (int)Double.parseDouble(pair_number);
                p_time.pair_start = pair_start;
                p_time.pair_end = pair_end;
                p_time.week_type = week_type;

                pair_times.add(p_time);
            }
            catch (Exception error)
            {
                System.out.println(error.toString());
            }


//            System.out.println(p_time);
        }

        // получаем теперь колонки, содержащие предметы
        Vector<Integer> columnsWithGroups = new Vector<Integer>();
        int columns_count = sheet.getRow(2).getPhysicalNumberOfCells(); // получаем количество столбцов во 2й строке (там, где указано слово предмет)
        for (int col=0; col<columns_count; col++)
        {
            String celltext = getCellText(sheet, 2, col); // получаем текст в ячейке
            if (celltext.contains("Предмет"))
            {
                columnsWithGroups.add(col); // добавляем колонку в список, как колонку с предметами
            }
        }

        Vector<univer_group> groups = new Vector<univer_group>();

        for (int col : columnsWithGroups)
        {
            String group_name = getCellText(sheet, 1, col); // получаем название группы (всегда записано в 1 строке

            univer_group group = new univer_group();
            group.setGroup_name(group_name);

            for (int row=3; row<pair_times.size()+3; row++) // проходим по всем строкам, содержащим предметы
            {
//                String subject_name = getCellText(sheet, row, col).replace("\n", "\t");     // название предмета
                String[] subject_names = getCellText(sheet, row, col).split("\n");
                for (int i=0; i<subject_names.length; i++)
                {
                    String subject_name = subject_names[i];
                    if(!subject_name.isEmpty())
                    {
                        try
                        {
                            pair_subject pairSubject = new pair_subject();
                            pairSubject.time = pair_times.get(row - 3);   // получаем время пары
                            pairSubject.subject_name = subject_name;
                            pairSubject.subject_type = getCellText(sheet, row, col + 1).split("\n")[i];   // тип занятий
                            pairSubject.subject_lecturer = getCellText(sheet, row, col + 2).split("\n")[i];   // ФИО преподавателя
                            pairSubject.subject_classroom = getCellText(sheet, row, col + 3).split("\n")[i];  // номер аудитории
                            pairSubject.subject_groupName = group_name;

                            group.addPair(pairSubject); // добавляем пару к группе
                        }
                        catch (Exception error)
                        {
//                            System.out.println("Exception: " + error.toString() + " при обработке " + subject_names[i]);
                        }
//                    System.out.println(group_name + " " + subject_name + " " + subject_type + " " + subject_lecturer + " " + subject_classroom);
                    }
                }
            }

            group.sortPairs(group.getPairs());

            groups.add(group);  // добавляем группу в список
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

    private String getCellText(Sheet sheet, int row, int cell)
    {
        return getCellText(sheet.getRow(row).getCell(cell));
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
}
