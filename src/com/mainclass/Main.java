package com.mainclass;


import com.mainclass.raspparser.RaspParser;
import com.mainclass.raspparser.subjects.pair_subject;
import com.mainclass.raspparser.subjects.univer_group;

import java.io.File;
import java.util.*;
import java.util.stream.Collectors;

public class Main {

    public static void main(String[] args) throws Exception
    {
        // ищем все файлы в папке
        String path = "C:\\Users\\ASUS\\Downloads\\rasp_20sep\\";
        String outputFile = "";

//        ExcelEditor excelEditor = new ExcelEditor();
//        excelEditor.createSimpleXlsXExcelBook();

        String files;
        File folder = new File(path);
        File[] listOfFiles = folder.listFiles();

        // создаём парсер
        RaspParser parser = new RaspParser();
        List<univer_group> groups = new ArrayList<>();

        parser.writePairsToFile(null);

        // проходим по всем файлам
        for (int i = 0; i < listOfFiles.length; i++)
        {
            if (listOfFiles[i].isFile())
            {
                files = listOfFiles[i].getName();
                System.out.println("Файл " + Integer.toString(i+1) + "/" + Integer.toString(listOfFiles.length) + ": " + files);
                if((!files.startsWith("~$")) && files.endsWith(".xlsx"))  // если это файл, который нам нужен
                {
                    groups.addAll(parser.readRasp(path + files));  // читаем и парсим его
                }
            }
        }

        groups.sort(new Comparator<univer_group>()
        {
            @Override
            public int compare(final univer_group group1, final univer_group group2)
            {
                return group1.getGroup_name().compareTo(group2.getGroup_name());
            }
        });

//        getGroupRasp(groups, "ЭЛБО-01-18");
//        getGroupRasp(groups, "ЭОСО-01-18");
        List<pair_subject> pairs_by_group = getGroupRasp(groups, "ЭФМО-01-22");
        printPairs(pairs_by_group);

//        List<pair_subject> pairs_in_classroom = getPairsInClassroom(groups, "В-404");
//        printPairs(pairs_in_classroom);

//        List<pair_subject> pairs_by_lecturer = getPairsByLecturer(groups, "Мильчакова Н.Е.");
//        printPairs(pairs_by_lecturer);

//        List<pair_subject> matchingPairs = getMatchersPairs(groups);
//        printPairs(matchingPairs);

    }

    public static void printPairs(List<pair_subject> pairs)
    {
        String last_day = "";
        for (pair_subject pair : pairs)
        {
            if(!pair.time.day_of_week.equals(last_day))
            {
                last_day = pair.time.day_of_week;
                System.out.println("================================================================================================");
            }
//            System.out.print("Аудитория: " + pair.subject_classroom + "\t");
//            System.out.print("День недели: " + pair.time.day_of_week + "\t");
//            System.out.print("Номер пары: " + pair.time.pair_number + "\t");
//            System.out.print("Тип недели: " + pair.time.week_type + "\t");
//            System.out.print("Группа: " + pair.subject_groupName + "\t");
//            System.out.print("Предмет: " + pair.subject_name + "\t");
//            System.out.print("Тип пары: " + pair.subject_type + "\t");
//            System.out.print("Преподаватель: " + pair.subject_lecturer + "\t");
//            System.out.println("");
            System.out.println(pair.subject_classroom +
                    "\t" + pair.time.day_of_week +
                    "\t" + pair.time.pair_number  +
                    "\t" + pair.time.pair_start  +
                    " - " + pair.time.pair_end  +
                    "\t" + pair.time.week_type +
                    "\t" + pair.subject_groupName +
                    "\t" + pair.subject_name +
                    "\t" + pair.subject_type +
                    "\t" + pair.subject_lecturer +
                    "\t");
        }
    }

    public static List<pair_subject> getGroupRasp(List<univer_group> groups, String group_name)
    {
        for (univer_group group : groups)
        {
            if(group.getGroup_name().contains(group_name))
            {
                return group.getPairs();
//                for (pair_subject pair : group.getPairs())
//                {
//                    System.out.print("Аудитория: " + pair.subject_classroom + "\t");
//                    System.out.print("День недели: " + pair.time.day_of_week + "\t");
//                    System.out.print("Номер пары: " + pair.time.pair_number + "\t");
//                    System.out.print("Тип недели: " + pair.time.week_type + "\t");
//                    System.out.print("Группа: " + group.getGroup_name() + "\t");
//                    System.out.print("Предмет: " + pair.subject_name + "\t");
//                    System.out.print("Тип пары: " + pair.subject_type + "\t");
//                    System.out.print("Преподаватель: " + pair.subject_lecturer + "\t");
//                    System.out.println("");
//                }
//                return group;
            }
        }
        return new ArrayList<>();
    }

    public static List<pair_subject> getPairsInClassroom(List<univer_group> groups, String classroom_name)
    {
        List<pair_subject> pairs = new ArrayList<>();
        for (univer_group group : groups)
        {
            for (pair_subject pair : group.getPairs())
            {
                if(pair.subject_classroom.toLowerCase().contains(classroom_name.toLowerCase()))
                {
                    pairs.add(pair);
                }
            }
        }
        univer_group.sortPairs(pairs);  // сортируем пары
        return pairs;
    }

    public static List<pair_subject> getPairsByLecturer(List<univer_group> groups, String lecturer_name)
    {
        List<pair_subject> pairs = new ArrayList<>();
        for (univer_group group : groups)
        {
            for (pair_subject pair : group.getPairs())
            {
                if(pair.subject_lecturer.toLowerCase().contains(lecturer_name.toLowerCase()))
                {
                    pairs.add(pair);
                }
            }
        }
        univer_group.sortPairs(pairs);  // сортируем пары
        return pairs;
    }

    public static List<pair_subject> getMatchersPairs(List<univer_group> groups)
    {
        // совпадающие
        List<pair_subject> matchingPairs = new ArrayList<>();

        for (int i = 0; i < groups.size(); i++)
        {
            List<pair_subject> pairs_1 = groups.get(i).getPairs();
            for (int j = 0; j < i; j++)
            {
                List<pair_subject> pairs_2 = groups.get(j).getPairs();

                for (int k = 0; k < pairs_1.size(); k++)
                {
                    pair_subject pair_1 = pairs_1.get(k);
                    for (int l = 0; l < pairs_2.size(); l++)
                    {
                        pair_subject pair_2 = pairs_2.get(l);

                        if(
                                (pair_1.subject_classroom.equals(pair_2.subject_classroom)
                                        && (!pair_1.subject_classroom.isEmpty())
                                        && (!pair_1.subject_classroom.toLowerCase().contains("фок"))
                                        && (!pair_1.subject_classroom.toLowerCase().contains("каф"))
                                        && (!pair_1.subject_classroom.toLowerCase().contains("база"))
                                        && (pair_1.time.week_type.equals(pair_2.time.week_type))
                                        && (pair_1.time.day_of_week_number == pair_2.time.day_of_week_number)
                                        && (pair_1.time.pair_number == pair_2.time.pair_number)
                                        && (!pair_1.subject_name.equals(pair_2.subject_name))
                                )
                                        && (!pair_1.subject_type.toLowerCase().equals("л"))
                                        && (!pair_1.subject_type.toLowerCase().equals("лк")))
                        {
                            matchingPairs.add(pair_1);
                            matchingPairs.add(pair_2);
                        }
                    }
                }
            }
        }

        matchingPairs = matchingPairs.stream().distinct().collect(Collectors.toList());

        matchingPairs.sort(new Comparator<pair_subject>() {
            @Override
            public int compare(pair_subject o1, pair_subject o2)
            {
//                return o1.subject_classroom.compareTo(o2.subject_classroom);
                if(o1.time.day_of_week_number != o2.time.day_of_week_number)
                {
                    return Integer.compare(o1.time.day_of_week_number, o2.time.day_of_week_number);
                }
                else if(o1.time.pair_number != o2.time.pair_number)
                {
                    return Integer.compare(o1.time.pair_number, o2.time.pair_number);
                }
                else if(!o1.time.week_type.equals(o2.time.week_type))
                {
                    return o1.time.week_type.compareTo(o2.time.week_type);
                }
                else
                {
                    return o1.subject_classroom.compareTo(o2.subject_classroom);
                }
            }
        });

        return matchingPairs;
    }
}
