package com.mainclass.raspparser.subjects;

public class pair_subject
{
    public pair_time time = null;
    public String subject_name = "";    // название предмета
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
