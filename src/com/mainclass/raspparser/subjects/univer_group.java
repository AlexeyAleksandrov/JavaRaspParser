package com.mainclass.raspparser.subjects;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class univer_group {
    private String group_name = null;
    private final List<pair_subject> pairs = new ArrayList<>();

    public void setGroup_name(String group_name) {
        this.group_name = group_name;
    }

    public String getGroup_name() {
        return group_name;
    }

    public List<pair_subject> getPairs() {
        return pairs;
    }

    public void addPair(pair_subject pair) {
        pairs.add(pair);
    }

    public static void sortPairs(List<pair_subject> pairs) {
        pairs.sort(new Comparator<pair_subject>() {
            @Override
            public int compare(pair_subject pair1, pair_subject pair2) {
                int compare_day_of_week = pair1.time.day_of_week_number - pair2.time.day_of_week_number;
                if (compare_day_of_week == 0) {
                    int compare_pair_number = pair1.time.pair_number - pair2.time.pair_number;
                    if (compare_pair_number == 0) {
                        int compare_week_type = pair1.time.week_type.compareTo(pair2.time.week_type);
                        if (compare_week_type == 0) {
                            int compare_subject_name = pair1.subject_name.compareTo(pair2.subject_name);
                            return compare_subject_name;
                        } else {
                            return compare_week_type;
                        }
                    } else {
                        return compare_pair_number;
                    }
                } else {
                    return compare_day_of_week;
                }
            }
        });
    }

    ;

    @Override
    public String toString() {
        return "univer_group{" +
                "group_name='" + group_name + '\'' +
                ", pairs=" + pairs +
                '}';
    }
}
