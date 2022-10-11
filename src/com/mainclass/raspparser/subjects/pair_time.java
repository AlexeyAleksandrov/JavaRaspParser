package com.mainclass.raspparser.subjects;

public class pair_time {
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
    public String toString() {
        return "pair_time{" +
                "day_of_week='" + day_of_week + '\'' +
                ", pair_number=" + pair_number +
                ", week_type='" + week_type + '\'' +
                ", pair_start='" + pair_start + '\'' +
                ", pair_end='" + pair_end + '\'' +
                '}';
    }
}
