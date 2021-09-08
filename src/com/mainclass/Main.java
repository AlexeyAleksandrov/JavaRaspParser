package com.mainclass;


import java.io.File;
import java.util.Collections;
import java.util.Comparator;
import java.util.Vector;

public class Main {

    public static void main(String[] args) throws Exception
    {
        // ищем все файлы в папке
        String path = "C:\\Users\\ASUS\\IdeaProjects\\JavaExcel\\rasp_data\\";

        String files;
        File folder = new File(path);
        File[] listOfFiles = folder.listFiles();

        // создаём парсер
        RaspParser parser = new RaspParser();
        Vector<univer_group> groups = new Vector<univer_group>();

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
//        Vector<pair_subject> pairs_by_group = getGroupRasp(groups, "КСБО-01-19");
//        printPairs(pairs_by_group);

//        Vector<pair_subject> pairs_in_classroom = getPairsInClassroom(groups, "Б-217");
//        printPairs(pairs_in_classroom);

        Vector<pair_subject> pairs_by_lecturer = getPairsByLecturer(groups, "Карпов");
        printPairs(pairs_by_lecturer);

    }

    public static void printPairs(Vector<pair_subject> pairs)
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
            System.out.println(pair.subject_classroom + "\t" + pair.time.day_of_week + "\t" + pair.time.pair_number + "\t" + pair.time.week_type + "\t" + pair.subject_groupName + "\t" + pair.subject_name + "\t");
        }
    }

    public static Vector<pair_subject> getGroupRasp(Vector<univer_group> groups, String group_name)
    {
        for (univer_group group : groups)
        {
            if(group.getGroup_name().equals(group_name))
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
        return null;
    }

    public static Vector<pair_subject> getPairsInClassroom(Vector<univer_group> groups, String classroom_name)
    {
        Vector<pair_subject> pairs = new Vector<pair_subject>();
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

    public static Vector<pair_subject> getPairsByLecturer(Vector<univer_group> groups, String lecturer_name)
    {
        Vector<pair_subject> pairs = new Vector<pair_subject>();
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
}